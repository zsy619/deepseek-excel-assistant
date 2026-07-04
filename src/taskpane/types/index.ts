/**
 * ============================================================================
 * DeepSeek Excel Assistant - Type Definitions
 * ----------------------------------------------------------------------------
 * All shared TypeScript interfaces used across services, components, and UI.
 * Keep this file free of runtime code so it can be safely imported anywhere.
 * ============================================================================
 */

/** Roles supported by the OpenAI-compatible chat completion format. */
export type ChatRole = "system" | "user" | "assistant" | "tool";

/**
 * A single message exchanged between the user and the model.
 * `tokens` is populated when the API reports usage (best-effort, may be
 * undefined for streaming responses).
 */
export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  timestamp: number;
  tokens?: number;
  /** Tool calls the model requested for this assistant turn (PRD-09). */
  toolCalls?: ToolCall[];
  /** For `role: "tool"` messages: which tool_call.id this is the result of. */
  toolCallId?: string;
  /** Friendly label like "插入图表" stored alongside a tool message. */
  toolName?: string;
  /** Optional branch metadata (PRD-08). When set, this message is a
   *  sibling of the message with id `parentId` - i.e. user wanted a
   *  different AI answer and branched from that earlier turn. */
  branch?: {
    parentId: string;
    branchId: string;
  };
}

/** Persisted conversation. */
export interface ChatSession {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  messages: ChatMessage[];
}

/** Runtime configuration persisted in localStorage. */
export interface DeepSeekConfig {
  apiKey: string;
  baseUrl: string;
  model: DeepSeekModel;
  temperature: number;
  maxTokens: number;
  topP: number;
  systemPrompt: string;
}

/** Models supported out of the box. The baseUrl override lets users
 *  plug in any OpenAI-compatible proxy that exposes these names. */
export type DeepSeekModel = "deepseek-chat" | "deepseek-reasoner";

/** Payload sent to /chat/completions. Mirrors the OpenAI schema so any
 *  compatible server (DeepSeek, Azure OpenAI, OpenRouter, local proxy…)
 *  can be used without code changes. */
export interface ChatCompletionRequest {
  model: DeepSeekModel;
  messages: Array<{ role: ChatRole; content: string; tool_call_id?: string; name?: string; tool_calls?: any[] }>;
  temperature?: number;
  max_tokens?: number;
  top_p?: number;
  stream?: boolean;
  tools?: ToolDefinition[];
  tool_choice?: "auto" | "none" | { type: "function"; function: { name: string } };
}

/** Non-streaming response payload (subset we actually consume). */
export interface ChatCompletionResponse {
  id: string;
  object: string;
  created: number;
  model: string;
  choices: Array<{
    index: number;
    message: { role: "assistant"; content: string };
    finish_reason: string | null;
  }>;
  usage?: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
}

/** A single SSE chunk from a streaming response. */
export interface ChatCompletionChunk {
  id: string;
  object: string;
  created: number;
  model: string;
  choices: Array<{
    index: number;
    delta: {
      role?: "assistant";
      content?: string;
      tool_calls?: Array<{
        index: number;
        id?: string;
        type?: "function";
        function?: { name?: string; arguments?: string };
      }>;
    };
    finish_reason: string | null;
  }>;
}

/* ----------------------------------------------------------------- *
 * Function calling (PRD-09)                                         *
 * ----------------------------------------------------------------- */

/** JSON-schema-style tool definition the model can call. Mirrors the
 *  OpenAI function-calling payload so any compatible server works. */
export interface ToolDefinition {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: {
      type: "object";
      properties: Record<string, { type: string; description?: string; enum?: string[] }>;
      required?: string[];
    };
  };
}

/** A single tool invocation emitted by the model. `arguments` is the raw
 *  JSON string the model streamed; parse on use. */
export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
  /** Parsed args, populated by applyToolCall(). */
  parsed?: Record<string, any>;
  /** Execution outcome (success / error message). */
  result?: string;
  /** True once we've actually invoked the tool locally. */
  executed?: boolean;
}

/** Shape of data returned from Excel via the Office.js API. */
export interface ExcelSelection {
  address: string;
  sheetName: string;
  values: any[][];
  formulas?: string[][];
  numberFormats?: string[][];
  rowCount: number;
  columnCount: number;
}

/** Result of attempting to talk to the DeepSeek API. */
export type ApiResult<T> =
  | { ok: true; data: T }
  | { ok: false; error: ApiError };

/** Normalized error envelope used by the UI layer. */
export interface ApiError {
  code:
    | "NETWORK"
    | "AUTH"
    | "RATE_LIMIT"
    | "SERVER"
    | "TIMEOUT"
    | "ABORTED"
    | "VALIDATION"
    | "UNKNOWN";
  message: string;
  status?: number;
  retryable: boolean;
}

/** User-facing view-model for a chat bubble. */
export interface MessageView {
  message: ChatMessage;
  isUser: boolean;
  isStreaming?: boolean;
}
/* ----------------------------------------------------------------- *
 * Chart types (PRD-05)                                               *
 * ----------------------------------------------------------------- */

/** Excel chart types supported by the AI recommender. Names map 1:1 to
 *  ExcelScript.ChartType values. */
export type ExcelChartType =
  | "ColumnClustered"
  | "Line"
  | "Pie"
  | "XYScatter"
  | "Area"
  | "Radar";

/** Per-locale display name + emoji for the picker UI. */
export const CHART_TYPE_INFO: Record<ExcelChartType, { label: string; icon: string }> = {
  ColumnClustered: { label: "柱状图",   icon: "📊" },
  Line:            { label: "折线图",   icon: "📈" },
  Pie:             { label: "饼图",     icon: "🥧" },
  XYScatter:       { label: "散点图",   icon: "✨" },
  Area:            { label: "面积图",   icon: "🟦" },
  Radar:           { label: "雷达图",   icon: "🕸️" },
};

/** Fallback mapping if the AI returns an unknown / null chart type. */
export const DEFAULT_CHART_TYPE: ExcelChartType = "ColumnClustered";

export const SUPPORTED_CHART_TYPES: ExcelChartType[] = Object.keys(CHART_TYPE_INFO) as ExcelChartType[];
