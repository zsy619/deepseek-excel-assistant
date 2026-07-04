/**
 * ============================================================================
 * DeepSeek API Service
 * ----------------------------------------------------------------------------
 * Thin wrapper around the OpenAI-compatible /chat/completions endpoint
 * exposed by DeepSeek (and any compatible proxy). Provides:
 *
 *   - chatCompletion()       non-streaming call returning the full reply
 *   - chatCompletionStream() streaming call with onChunk / onDone / onError
 *
 * Streaming uses fetch + ReadableStream so it works inside the Office
 * webview where EventSource is unreliable. An AbortController is returned
 * to support user-initiated cancellation mid-stream.
 *
 * Error handling maps upstream HTTP codes + network conditions into a
 * normalized ApiError so the UI layer can render consistent messages.
 * ============================================================================
 */

import type {
  ApiError,
  ChatCompletionChunk,
  ChatCompletionRequest,
  ChatCompletionResponse,
  ChatMessage,
  ChatRole,
  DeepSeekConfig,
  ToolDefinition,
  ToolCall,
} from "../types";
import { REQUEST_TIMEOUT_MS } from "../utils/constants";
import { toApiError } from "../utils/helpers";

/* ----------------------------------------------------------------- *
 * Tool registry (PRD-09)                                             *
 * ----------------------------------------------------------------- */

/** Tools exposed to the model. The list is intentionally short and
 *  scoped to Excel-side actions the assistant can perform safely. */
export const TOOL_REGISTRY: ToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "writeFormula",
      description:
        "写入一个 Excel 公式到活动单元格（或 address 参数指定的单元格）。参数：address (可选, 默认活动单元格), formula (必填, 以 = 开头的公式)。",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string", description: "目标单元格地址, 如 'B3', 留空则使用活动单元格" },
          formula: { type: "string", description: "Excel 公式, 必须以 = 开头" },
        },
        required: ["formula"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insertChart",
      description:
        "在当前选区旁边插入一个 Excel 图表。参数：chartType (ColumnClustered|Line|Pie|XYScatter|Area|Radar), title (可选)。",
      parameters: {
        type: "object",
        properties: {
          chartType: {
            type: "string",
            description: "图表类型",
            enum: ["ColumnClustered", "Line", "Pie", "XYScatter", "Area", "Radar"],
          },
          title: { type: "string", description: "图表标题 (可选)" },
        },
        required: ["chartType"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "scanPII",
      description:
        "扫描当前选区中的敏感信息 (手机号、邮箱、身份证、银行卡等), 返回命中数量与位置。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
  {
    type: "function",
    function: {
      name: "getSelectionInfo",
      description:
        "读取当前选区的基本信息: 地址、表名、数据预览 (前 3 行 × 5 列), 用于让 AI 理解用户当前操作上下文。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
];

/** Quick lookup so the dispatcher can validate the model didn't hallucinate
 *  a function name. */
const TOOL_NAMES = new Set(TOOL_REGISTRY.map((t) => t.function.name));

/** Strip the persisted message envelope down to what the API expects:
 *  role + content (+ optional tool plumbing). We never send internal
 *  fields like `tokens` / `timestamp`. */
export function toApiMessages(
  messages: ChatMessage[]
): ChatCompletionRequest["messages"] {
  return messages
    .filter((m) => m.role !== "system")
    .map((m) => {
      // Tool-result messages need tool_call_id + name on the wire.
      if (m.role === "tool") {
        return {
          role: "tool" as ChatRole,
          content: m.content || "",
          tool_call_id: m.toolCallId,
          name: m.toolName,
        };
      }
      // Assistant messages that triggered tool calls need to replay
      // the tool_calls block so the model knows the next tool message
      // is answering them.
      if (m.role === "assistant" && m.toolCalls && m.toolCalls.length) {
        return {
          role: "assistant" as ChatRole,
          content: m.content || "",
          tool_calls: m.toolCalls.map((tc) => ({
            id: tc.id,
            type: "function",
            function: { name: tc.name, arguments: tc.arguments },
          })),
        };
      }
      return { role: m.role, content: m.content };
    });
}

/** Normalize an HTTP status to a stable error code. */
export function classifyHttpStatus(status: number): ApiError["code"] {
  if (status === 401 || status === 403) return "AUTH";
  if (status === 429) return "RATE_LIMIT";
  if (status >= 500) return "SERVER";
  if (status >= 400) return "VALIDATION";
  return "UNKNOWN";
}

/** Build the friendly user-facing message for a given status. */
export function friendlyMessage(status: number, fallback: string): string {
  switch (classifyHttpStatus(status)) {
    case "AUTH":
      return "API Key 无效或未通过鉴权，请到设置中检查。";
    case "RATE_LIMIT":
      return "请求过于频繁，请稍后再试。";
    case "SERVER":
      return `DeepSeek 服务异常 (HTTP ${status})，请稍后再试。`;
    case "VALIDATION":
      return `请求参数有误 (HTTP ${status})：${fallback}`;
    default:
      return fallback || `请求失败 (HTTP ${status})`;
  }
}

/** Wrap fetch with an AbortController-driven timeout. */
function fetchWithTimeout(url: string, init: RequestInit, timeoutMs: number): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  // Forward caller-supplied abort if present
  const upstream = init.signal;
  if (upstream) {
    if (upstream.aborted) controller.abort();
    else upstream.addEventListener("abort", () => controller.abort(), { once: true });
  }
  return fetch(url, { ...init, signal: controller.signal }).finally(() => clearTimeout(timer));
}

/** Common body builder - guarantees the same shape is sent for streaming
 *  and non-streaming requests. Pass `withTools=true` to expose the
 *  function-calling tool list (PRD-09). */
export function buildRequestBody(
  messages: ChatMessage[],
  config: DeepSeekConfig,
  stream: boolean,
  withTools: boolean = false
): ChatCompletionRequest {
  const body: ChatCompletionRequest = {
    model: config.model,
    messages: toApiMessages(messages),
    temperature: config.temperature,
    max_tokens: config.maxTokens,
    top_p: config.topP,
    stream,
  };
  if (withTools) {
    body.tools = TOOL_REGISTRY;
    body.tool_choice = "auto";
  }
  return body;
}

/**
 * Non-streaming chat completion. Returns the full assistant message.
 *
 * @throws ApiError when the request fails for any reason.
 */
export async function chatCompletion(
  messages: ChatMessage[],
  config: DeepSeekConfig
): Promise<ChatCompletionResponse> {
  const url = joinUrl(config.baseUrl, "/chat/completions");

  try {
    const res = await fetchWithTimeout(
      url,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${config.apiKey}`,
          Accept: "application/json",
        },
        body: JSON.stringify(buildRequestBody(messages, config, false)),
      },
      REQUEST_TIMEOUT_MS
    );

    if (!res.ok) {
      const errBody = await safeReadError(res);
      throw {
        code: classifyHttpStatus(res.status),
        message: friendlyMessage(res.status, errBody),
        status: res.status,
        retryable: res.status >= 500 || res.status === 429,
      } as ApiError;
    }

    return (await res.json()) as ChatCompletionResponse;
  } catch (err) {
    throw toApiError(err);
  }
}

/**
 * Streaming chat completion. Calls `onChunk` for every token the model emits,
 * `onDone` once the stream ends cleanly, `onToolCalls` if the model emitted
 * any function calls (PRD-09), and `onError` on any failure.
 *
 * Returns the AbortController so the caller can cancel mid-flight.
 */
export async function chatCompletionStream(
  messages: ChatMessage[],
  config: DeepSeekConfig,
  onChunk: (text: string, accumulated: string) => void,
  onDone: (fullText: string) => void,
  onError: (err: ApiError) => void,
  opts: { withTools?: boolean; onToolCalls?: (calls: ToolCall[]) => void } = {}
): Promise<AbortController> {
  const controller = new AbortController();
  const url = joinUrl(config.baseUrl, "/chat/completions");
  const withTools = !!opts.withTools;
  const onToolCalls = opts.onToolCalls;

  let accumulated = "";
  // Tool calls stream in piece-by-piece: name first, then arguments
  // appended in chunks. We accumulate per-index and emit only when the
  // stream ends.
  const toolCallMap: Map<number, { id: string; name: string; args: string }> = new Map();

  // Kick off the async processing - callers don't await this; they
  // observe progress through the callbacks.
  (async () => {
    try {
      const res = await fetchWithTimeout(
        url,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${config.apiKey}`,
            Accept: "text/event-stream",
          },
          body: JSON.stringify(buildRequestBody(messages, config, true, withTools)),
        },
        REQUEST_TIMEOUT_MS,
        // attach the caller-visible controller so chat UI can cancel
      );

      if (!res.ok || !res.body) {
        const errBody = await safeReadError(res);
        const err: ApiError = {
          code: classifyHttpStatus(res.status),
          message: friendlyMessage(res.status, errBody),
          status: res.status,
          retryable: res.status >= 500 || res.status === 429,
        };
        onError(err);
        return;
      }

      const reader = res.body.getReader();
      const decoder = new TextDecoder("utf-8");
      let buffer = "";

      // Outer loop: read until done.
      // Inner loop: split buffer by SSE delimiter ("\n\n").
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        let sepIdx: number;
        // eslint-disable-next-line no-cond-assign
        while ((sepIdx = buffer.indexOf("\n\n")) >= 0) {
          const rawEvent = buffer.slice(0, sepIdx);
          buffer = buffer.slice(sepIdx + 2);
          const piece = parseSseEvent(rawEvent);
          if (!piece) continue;
          if (piece === "[DONE]") {
            finishToolCalls();
            onDone(accumulated);
            return;
          }
          try {
            const parsed = JSON.parse(piece) as ChatCompletionChunk;
            const choice = parsed.choices?.[0];
            const delta = choice?.delta;
            const text = delta?.content;
            if (text) {
              accumulated += text;
              onChunk(text, accumulated);
            }
            // Tool calls: stream may emit a single chunk with id+name,
            // then many chunks with arguments delta. Aggregate them.
            if (withTools && delta?.tool_calls) {
              for (const tc of delta.tool_calls) {
                const slot =
                  toolCallMap.get(tc.index) ??
                  { id: tc.id || "", name: "", args: "" };
                if (tc.id) slot.id = tc.id;
                if (tc.function?.name) slot.name += tc.function.name;
                if (tc.function?.arguments) slot.args += tc.function.arguments;
                toolCallMap.set(tc.index, slot);
              }
            }
            // Some providers emit finish_reason="tool_calls" without a
            // tool_calls payload in the same chunk. Treat it as a hint.
            if (withTools && choice?.finish_reason === "tool_calls") {
              finishToolCalls();
            }
          } catch {
            // Ignore malformed chunks - many providers emit keepalive
            // comments or partial JSON during disconnects.
          }
        }
      }

      finishToolCalls();
      onDone(accumulated);
    } catch (err) {
      const apiErr = toApiError(err);
      // ABORTED is a normal user action, not a real failure.
      if (apiErr.code !== "ABORTED") {
        onError(apiErr);
      } else {
        finishToolCalls();
        onDone(accumulated);
      }
    }
  })();

  function finishToolCalls() {
    if (!withTools || !onToolCalls || toolCallMap.size === 0) return;
    const calls: ToolCall[] = [];
    const indices = [...toolCallMap.keys()].sort((a, b) => a - b);
    for (const idx of indices) {
      const slot = toolCallMap.get(idx)!;
      if (!slot.id || !slot.name) continue;
      // Validate: drop calls to functions we never advertised.
      if (!TOOL_NAMES.has(slot.name)) continue;
      calls.push({
        id: slot.id,
        name: slot.name,
        arguments: slot.args || "{}",
      });
    }
    if (calls.length) onToolCalls(calls);
  }

  return controller;
}

/** Normalize URL joining so users can paste "https://x.com" or "https://x.com/". */
function joinUrl(base: string, path: string): string {
  if (!base) throw { code: "VALIDATION", message: "API 端点未配置", retryable: false } as ApiError;
  const trimmed = base.replace(/\/+$/, "");
  return trimmed + path;
}

/** Best-effort error body reader. Returns empty string on failure. */
async function safeReadError(res: Response): Promise<string> {
  try {
    const txt = await res.text();
    try {
      const json = JSON.parse(txt);
      return json?.error?.message || json?.message || txt;
    } catch {
      return txt;
    }
  } catch {
    return "";
  }
}

/** Parse a single SSE event block into the raw data payload (or [DONE]).
 *  Skips comments and event-name lines. */
export function parseSseEvent(rawEvent: string): string | null {
  const lines = rawEvent.split(/\r?\n/);
  const dataLines: string[] = [];
  for (const line of lines) {
    if (!line || line.startsWith(":")) continue;
    if (line.startsWith("data:")) {
      dataLines.push(line.slice(5).trimStart());
    }
  }
  if (dataLines.length === 0) return null;
  return dataLines.join("\n");
}

/** Public helper exposed so the UI can show the model/user friendly
 *  error text without duplicating the classifier. */
export function describeApiError(err: ApiError): string {
  if (!err) return "";
  return err.message || "请求失败";
}

/* ----------------------------------------------------------------- */
/* Formula diagnostics (PRD-01)                                       */
/* ----------------------------------------------------------------- */

export interface FormulaDiagnosis {
  /** Cell address within the scanned range (e.g. "C3"). */
  address: string;
  /** Excel error sentinel (#REF!, #DIV/0!, etc.). */
  error: string;
  /** One-sentence cause of the error in Chinese. */
  cause: string;
  /** Concrete fix suggestion in Chinese. */
  suggestion: string;
  /** Replacement formula starting with "=" (or empty if unfixable). */
  fixedFormula: string;
  /** Confidence score 0-1. */
  confidence: number;
}

export interface FormulaDiagnosisRequest {
  sheetName: string;
  rangeAddress: string;
  errors: Array<{
    address: string;
    fullAddress: string;
    formula: string;
    value: string;
    error: string;
  }>;
}

export function buildDiagnosisMessages(req: FormulaDiagnosisRequest): ChatMessage[] {
  const errorList = req.errors
    .map(
      (e, i) =>
        `${i + 1}. 单元格 ${e.address}（${e.fullAddress}）\n` +
        `   当前公式：${e.formula}\n` +
        `   错误：${e.value}`
    )
    .join("\n\n");

  const systemPrompt =
    "你是一个 Excel 公式调试专家。请针对每个错误公式给出：\n" +
    "1. 错误原因（一句话中文）\n" +
    "2. 修复建议（一句话中文）\n" +
    "3. 修复后的公式（以 = 开头；如果无法修复则留空字符串）\n" +
    "4. 置信度（0 到 1 之间的小数）\n\n" +
    "请严格按 JSON 数组格式输出，不要包含 Markdown 代码块标记，每个对象包含 address / error / cause / suggestion / fixedFormula / confidence 六个字段。";

  const userPrompt = `工作表：${req.sheetName}\n选区地址：${req.rangeAddress}\n\n发现 ${req.errors.length} 个错误公式：\n\n${errorList}`;

  return [
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ];
}

/**
 * Streamed diagnostics for a batch of formula errors. Calls `onPartial`
 * with the accumulated JSON text as it streams in, then `onDone` with the
 * parsed array. On parse failure we fall back to wrapping the raw text in
 * a single diagnosis entry so the UI can still surface something.
 */
export async function diagnoseFormulasStream(
  req: FormulaDiagnosisRequest,
  config: DeepSeekConfig,
  callbacks: {
    onPartial?: (raw: string) => void;
    onDone: (results: FormulaDiagnosis[]) => void;
    onError: (err: ApiError) => void;
  }
): Promise<AbortController> {
  const messages = buildDiagnosisMessages(req);
  let accumulated = "";

  return chatCompletionStream(
    messages,
    config,
    (_chunk, total) => {
      accumulated = total;
      callbacks.onPartial?.(accumulated);
    },
    (fullText) => {
      const parsed = tryParseDiagnosisJson(fullText);
      callbacks.onDone(parsed);
    },
    callbacks.onError
  );
}

/** Extract the first JSON array from arbitrary LLM output. Tolerant of
 *  stray prose and markdown fences. */
export function tryParseDiagnosisJson(text: string): FormulaDiagnosis[] {
  if (!text) return [];
  const cleaned = text
    .replace(/^```(?:json)?/im, "")
    .replace(/```$/m, "")
    .trim();
  try {
    const v = JSON.parse(cleaned);
    if (Array.isArray(v)) return v.filter(isValidDiagnosis);
  } catch {
    /* fall through to bracket extraction */
  }
  const start = cleaned.indexOf("[");
  const end = cleaned.lastIndexOf("]");
  if (start >= 0 && end > start) {
    const slice = cleaned.slice(start, end + 1);
    try {
      const v = JSON.parse(slice);
      if (Array.isArray(v)) return v.filter(isValidDiagnosis);
    } catch {
      /* noop */
    }
  }
  return [
    {
      address: "",
      error: "",
      cause: "AI 响应无法解析为 JSON",
      suggestion: "请重试或简化选区",
      fixedFormula: "",
      confidence: 0,
    },
  ];
}

export function isValidDiagnosis(x: any): x is FormulaDiagnosis {
  return !!x && typeof x === "object" && typeof x.address === "string";
}
/* ----------------------------------------------------------------- *
 * Formula -> VBA / Office Scripts translation (PRD-04)             *
 * ----------------------------------------------------------------- */

export type ScriptFlavor = "vba" | "office-scripts";

/** Strip markdown code fences around a ``` block so we can hand the
 *  assistant output back to the user verbatim while still letting the
 *  LLM wrap its answer for readability. */
export function stripCodeFences(s: string): string {
  if (!s) return s;
  return s
    .replace(/^\s*```(?:vba|typescript|vb|ts)?\s*\n/i, "")
    .replace(/\n```\s*$/i, "")
    .trim();
}

/** Build the system + user prompt for translating N formulas into
 *  either VBA or Office Scripts. */
export function buildCodeGenMessages(
  sheet: string,
  block: string,
  flavor: ScriptFlavor
): { system: string; user: string } {
  const target = flavor === "vba" ? "VBA Sub 函数" : "Office Script (TypeScript, ExcelScript API)";
  const cellsPerLine = block.split("\n").filter(Boolean);
  const system =
    `你是一位资深的 Excel ${flavor === "vba" ? "VBA" : "Office Scripts"} 专家。` +
    `请把用户提供的公式转换为等价的 ${target}。` +
    `要求：\n` +
    `1. 保持公式之间的依赖顺序\n` +
    `2. 使用 Worksheets("...").Range("...") / ExcelScript API 引用单元格\n` +
    (flavor === "vba"
      ? `3. Excel 内置函数用 Application.WorksheetFunction 或简单表达式\n`
      : `3. 数组结果写到 Range.setValues()\n`) +
    `4. 输出完整可运行的脚本，使用 Markdown 的 \`\`\` 代码块包裹\n` +
    `5. 不要做任何解释、不要客套话，只输出代码块`;

  const user =
    `工作表：${sheet}\n` +
    `单元格地址 + 公式（共 ${cellsPerLine.length} 个）：\n` +
    `${block}\n\n` +
    `请生成${target}。`;

  return { system, user };
}

/** Build the streaming controller for code-gen. Returns an AbortController
 *  the caller can use to cancel mid-stream. Tokens are delivered through
 *  `onPartial`; the final accumulated text is in `onDone`. Errors hit
 *  `onError`. Mirrors the API shape of `chatCompletionStream` so the
 *  caller code stays uniform. */
export async function translateToScript(
  formulas: { sheet: string; block: string },
  flavor: ScriptFlavor,
  config: DeepSeekConfig,
  callbacks: {
    onPartial: (text: string, accumulated: string) => void;
    onDone: (fullText: string) => void;
    onError: (err: ApiError) => void;
  }
): Promise<AbortController> {
  const { system, user } = buildCodeGenMessages(formulas.sheet, formulas.block, flavor);
  const messages: ChatMessage[] = [
    { id: "cg-sys", role: "system", content: system, timestamp: Date.now() },
    { id: "cg-usr", role: "user", content: user, timestamp: Date.now() },
  ];
  void stripCodeFences;
  return await chatCompletionStream(messages, config, callbacks.onPartial, callbacks.onDone, callbacks.onError);
}

/* ----------------------------------------------------------------- *
 * Chart recommender (PRD-05)                                        *
 * ----------------------------------------------------------------- */

import type { ExcelChartType } from "../types";
import { SUPPORTED_CHART_TYPES, DEFAULT_CHART_TYPE } from "../types";

export interface ChartRecommendation {
  type: ExcelChartType;
  title: string;
  reason: string;
}

export interface ChartRecommendResult {
  recommendations: ChartRecommendation[];
  /** Heuristic guess of the X-axis label (first column) for fallback use. */
  inferredCategory: string;
  /** True if we could not parse a structured response and used local
   *  heuristics to provide a default. */
  usedFallback: boolean;
}

export interface ChartRecommendRequest {
  sheet: string;
  headers: string[];
  preview: string[][];
  rowCount: number;
  columnCount: number;
}

/** Build a JSON-only prompt that asks the model to pick from the
 *  whitelisted chart types. */
export function buildChartRecommendMessages(req: ChartRecommendRequest): {
  system: string;
  user: string;
} {
  const listStr = SUPPORTED_CHART_TYPES.join(", ");
  const previewStr = req.preview
    .map((row) => row.join(" | "))
    .slice(0, 6)
    .join("\n");
  const system =
    `你是一位 Excel 图表专家。基于用户提供的数据，挑选出 3 个最合适的图表类型。` +
    `可用类型：${listStr}。\n` +
    `要求：\n` +
    `1. 输出严格 JSON 数组，不要任何解释、不要代码块包裹\n` +
    `2. 每项结构：{ "type": "<类型>", "title": "<中文标题>", "reason": "<选择理由，1 句话>" }\n` +
    `3. type 必须是可用类型列表里的字符串\n` +
    `4. 3 个推荐中第一个为最推荐`;
  const user =
    `工作表：${req.sheet}\n` +
    `列数：${req.columnCount}，行数：${req.rowCount}\n` +
    `表头：${req.headers.join(" | ")}\n` +
    `数据预览：\n${previewStr}\n\n` +
    `请给出 3 个推荐。`;
  return { system, user };
}

/** Best-effort JSON object array parser. Tolerates ```json fences and
 *  partial trailing commas. Returns the first N valid recommendations. */
export function tryParseChartList(raw: string, max = 3): ChartRecommendation[] {
  if (!raw) return [];
  let txt = raw.trim();
  // Strip ``` fences
  const fence = txt.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) txt = fence[1].trim();
  // Find first '[' and matching ']'
  const a = txt.indexOf("[");
  const b = txt.lastIndexOf("]");
  if (a < 0 || b <= a) return [];
  const inner = txt.slice(a, b + 1);
  try {
    const arr = JSON.parse(inner);
    if (!Array.isArray(arr)) return [];
    const out: ChartRecommendation[] = [];
    for (const item of arr) {
      if (!item || typeof item !== "object") continue;
      const t = (item as any).type;
      const title = String((item as any).title ?? "").trim();
      const reason = String((item as any).reason ?? "").trim();
      if (!SUPPORTED_CHART_TYPES.includes(t)) continue;
      out.push({ type: t as ExcelChartType, title, reason });
      if (out.length >= max) break;
    }
    return out;
  } catch {
    return [];
  }
}

/** Local fallback recommender. Picks based on simple heuristics. */
export function fallbackChartRecommend(req: ChartRecommendRequest): ChartRecommendation[] {
  const out: ChartRecommendation[] = [];
  const numericCols = req.headers.filter((h) => true).length;
  const hasManyRows = req.rowCount >= 8;
  // Default: bar chart for category + value
  out.push({
    type: "ColumnClustered",
    title: "柱状图（默认）",
    reason: `检测到 ${req.columnCount} 列数据，柱状图是最通用的对比视图。`,
  });
  // If more than 1 numeric column or many rows, suggest line
  if (numericCols >= 2 || hasManyRows) {
    out.push({
      type: "Line",
      title: "折线图",
      reason: "数据点较多或多系列数值,折线图能更好地呈现趋势。",
    });
  } else {
    out.push({
      type: "Line",
      title: "折线图",
      reason: "折线图可以显示数据随行的变化趋势。",
    });
  }
  // Pie for small sets
  if (req.rowCount <= 8 && req.columnCount === 2) {
    out.push({
      type: "Pie",
      title: "饼图",
      reason: "数据点较少且只有一组数值,饼图直观展示占比。",
    });
  } else {
    out.push({
      type: "Area",
      title: "面积图",
      reason: "面积图强调数值随类别的累计变化。",
    });
  }
  return out.slice(0, 3);
}

/** Stream chart recommendations. Returns an AbortController so the
 *  caller can cancel. Calls `onPartial` as JSON arrives; the FINAL
 *  parsed list lands in `onDone`. */
export async function recommendChartStream(
  config: DeepSeekConfig,
  req: ChartRecommendRequest,
  callbacks: {
    onPartial?: (raw: string) => void;
    onDone: (result: ChartRecommendResult) => void;
    onError: (err: ApiError) => void;
  }
): Promise<AbortController> {
  const { system, user } = buildChartRecommendMessages(req);
  const messages: ChatMessage[] = [
    { id: "ch-sys", role: "system", content: system, timestamp: Date.now() },
    { id: "ch-usr", role: "user", content: user, timestamp: Date.now() },
  ];
  let latestRaw = "";
  let parsed: ChartRecommendation[] = [];
  return chatCompletionStream(
    messages,
    config,
    (text, accumulated) => {
      latestRaw = accumulated;
      callbacks.onPartial?.(accumulated);
      const candidate = tryParseChartList(accumulated, 3);
      if (candidate.length > 0 && candidate.length >= parsed.length) {
        parsed = candidate;
      }
    },
    () => {
      parsed = tryParseChartList(latestRaw, 3);
      const usedFallback = parsed.length === 0;
      const list = usedFallback ? fallbackChartRecommend(req) : parsed;
      const result: ChartRecommendResult = {
        recommendations: list,
        inferredCategory: req.headers[0] ?? "",
        usedFallback,
      };
      callbacks.onDone(result);
    },
    callbacks.onError
  );
}
