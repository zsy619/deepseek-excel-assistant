import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  TOOL_REGISTRY,
  describeApiError,
  chatCompletion,
  chatCompletionStream,
  buildRequestBody,
  toApiMessages,
  classifyHttpStatus,
  friendlyMessage,
  parseSseEvent,
  buildDiagnosisMessages,
  tryParseDiagnosisJson,
  isValidDiagnosis,
  stripCodeFences,
  buildCodeGenMessages,
  buildChartRecommendMessages,
  tryParseChartList,
  fallbackChartRecommend,
  diagnoseFormulasStream,
  translateToScript,
  recommendChartStream,
  type FormulaDiagnosisRequest,
  type ChartRecommendRequest,
} from "./deepseek";
import type { ApiError, ChatMessage, DeepSeekConfig } from "../types";

/* ----------------------------------------------------------------- */
/*  Sample data used across tests                                     */
/* ----------------------------------------------------------------- */

const sampleConfig: DeepSeekConfig = {
  apiKey: "sk-test-key",
  baseUrl: "https://api.deepseek.com",
  model: "deepseek-chat",
  temperature: 0.7,
  maxTokens: 2048,
  topP: 0.9,
  systemPrompt: "You are a helpful assistant.",
};

const sampleMessages: ChatMessage[] = [
  { id: "m1", role: "user", content: "Hello", timestamp: Date.now() },
  { id: "m2", role: "assistant", content: "Hi there!", timestamp: Date.now() },
];

/* ----------------------------------------------------------------- */
/*  TOOL_REGISTRY                                                     */
/* ----------------------------------------------------------------- */

describe("TOOL_REGISTRY", () => {
  it("has at least 3 tools", () => {
    expect(TOOL_REGISTRY.length).toBeGreaterThanOrEqual(3);
  });

  it("every tool has name, description, and parameters", () => {
    for (const t of TOOL_REGISTRY) {
      expect(t.type).toBe("function");
      expect(t.function.name).toBeTruthy();
      expect(t.function.description).toBeTruthy();
      expect(t.function.parameters).toBeDefined();
    }
  });

  it("all tool names are unique", () => {
    const names = TOOL_REGISTRY.map((t) => t.function.name);
    expect(new Set(names).size).toBe(names.length);
  });
});

/* ----------------------------------------------------------------- */
/*  describeApiError                                                  */
/* ----------------------------------------------------------------- */

describe("describeApiError", () => {
  it("returns message when present", () => {
    const err: ApiError = { code: "AUTH", message: "bad key", status: 401, retryable: false };
    expect(describeApiError(err)).toBe("bad key");
  });

  it("returns fallback when message is empty", () => {
    const err: ApiError = { code: "UNKNOWN", message: "", status: 0, retryable: false };
    expect(describeApiError(err)).toBe("请求失败");
  });

  it("returns empty string for null/undefined input", () => {
    expect(describeApiError(null as any)).toBe("");
    expect(describeApiError(undefined as any)).toBe("");
  });
});

/* ----------------------------------------------------------------- */
/*  classifyHttpStatus                                                */
/* ----------------------------------------------------------------- */

describe("classifyHttpStatus", () => {
  it("classifies 401 and 403 as AUTH", () => {
    expect(classifyHttpStatus(401)).toBe("AUTH");
    expect(classifyHttpStatus(403)).toBe("AUTH");
  });

  it("classifies 429 as RATE_LIMIT", () => {
    expect(classifyHttpStatus(429)).toBe("RATE_LIMIT");
  });

  it("classifies 500+ as SERVER", () => {
    expect(classifyHttpStatus(500)).toBe("SERVER");
    expect(classifyHttpStatus(502)).toBe("SERVER");
    expect(classifyHttpStatus(503)).toBe("SERVER");
  });

  it("classifies other 4xx as VALIDATION", () => {
    expect(classifyHttpStatus(400)).toBe("VALIDATION");
    expect(classifyHttpStatus(404)).toBe("VALIDATION");
    expect(classifyHttpStatus(422)).toBe("VALIDATION");
  });

  it("classifies everything else as UNKNOWN", () => {
    expect(classifyHttpStatus(200)).toBe("UNKNOWN");
    expect(classifyHttpStatus(301)).toBe("UNKNOWN");
  });
});

/* ----------------------------------------------------------------- */
/*  friendlyMessage                                                   */
/* ----------------------------------------------------------------- */

describe("friendlyMessage", () => {
  it("returns Chinese AUTH message for 401", () => {
    const msg = friendlyMessage(401, "nope");
    expect(msg).toContain("API Key");
  });

  it("returns Chinese RATE_LIMIT message for 429", () => {
    const msg = friendlyMessage(429, "slow down");
    expect(msg).toContain("频繁");
  });

  it("returns Chinese SERVER message for 5xx", () => {
    const msg = friendlyMessage(502, "bad gateway");
    expect(msg).toContain("服务异常");
  });

  it("includes the fallback for 4xx VALIDATION", () => {
    const msg = friendlyMessage(400, "invalid params");
    expect(msg).toContain("invalid params");
  });

  it("returns fallback for other status codes", () => {
    const msg = friendlyMessage(200, "unexpected");
    expect(msg).toContain("unexpected");
  });
});

/* ----------------------------------------------------------------- */
/*  parseSseEvent                                                     */
/* ----------------------------------------------------------------- */

describe("parseSseEvent", () => {
  it("parses a single data line", () => {
    const result = parseSseEvent("data: hello world");
    expect(result).toBe("hello world");
  });

  it("parses data after trimming leading whitespace", () => {
    const result = parseSseEvent("data:  {\"key\": 1}");
    expect(result).toBe('{"key": 1}');
  });

  it("joins multiple data lines", () => {
    const result = parseSseEvent("data: line1\ndata: line2");
    expect(result).toBe("line1\nline2");
  });

  it("returns null for empty input", () => {
    expect(parseSseEvent("")).toBeNull();
  });

  it("skips comment lines (starting with :)", () => {
    const result = parseSseEvent(": comment\ndata: payload");
    expect(result).toBe("payload");
  });

  it("skips lines without data: prefix", () => {
    const result = parseSseEvent("event: complete\ndata: hello");
    expect(result).toBe("hello");
  });

  it("returns null when no data lines exist", () => {
    const result = parseSseEvent("event: complete\n:just a comment");
    expect(result).toBeNull();
  });
});

/* ----------------------------------------------------------------- */
/*  toApiMessages                                                     */
/* ----------------------------------------------------------------- */

describe("toApiMessages", () => {
  it("filters out system messages", () => {
    const msgs: ChatMessage[] = [
      { id: "1", role: "system", content: "be nice", timestamp: 0 },
      { id: "2", role: "user", content: "ok", timestamp: 0 },
    ];
    const result = toApiMessages(msgs);
    expect(result).toHaveLength(1);
    expect(result[0].role).toBe("user");
  });

  it("includes tool_call_id and name for tool messages", () => {
    const msgs: ChatMessage[] = [
      { id: "1", role: "tool", content: "42", toolCallId: "call_1", toolName: "getAnswer", timestamp: 0 },
    ];
    const result = toApiMessages(msgs);
    expect(result[0]).toHaveProperty("tool_call_id", "call_1");
    expect(result[0]).toHaveProperty("name", "getAnswer");
  });

  it("includes tool_calls for assistant messages with toolCalls", () => {
    const msgs: ChatMessage[] = [
      {
        id: "1",
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call_1", name: "getAnswer", arguments: "{}" }],
        timestamp: 0,
      },
    ];
    const result = toApiMessages(msgs);
    expect(result[0]).toHaveProperty("tool_calls");
    expect((result[0] as any).tool_calls[0].function.name).toBe("getAnswer");
  });
});

/* ----------------------------------------------------------------- */
/*  buildRequestBody                                                  */
/* ----------------------------------------------------------------- */

describe("buildRequestBody", () => {
  it("includes stream: false by default", () => {
    const body = buildRequestBody(sampleMessages, sampleConfig, false);
    expect(body.stream).toBe(false);
    expect(body.model).toBe("deepseek-chat");
    expect(body.temperature).toBe(0.7);
  });

  it("includes tools when withTools is true", () => {
    const body = buildRequestBody(sampleMessages, sampleConfig, true, true);
    expect(body.tools).toBeDefined();
    expect(body.tool_choice).toBe("auto");
  });

  it("strips system messages from messages array", () => {
    const msgs: ChatMessage[] = [
      { id: "s1", role: "system", content: "prompt", timestamp: 0 },
      { id: "u1", role: "user", content: "hello", timestamp: 0 },
    ];
    const body = buildRequestBody(msgs, sampleConfig, false);
    expect(body.messages).toHaveLength(1);
    expect(body.messages[0].role).toBe("user");
  });
});

/* ----------------------------------------------------------------- */
/*  API calls (mocked fetch)                                          */
/* ----------------------------------------------------------------- */

describe("chatCompletion (mocked)", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("resolves with the response body on 200", async () => {
    const fakeResponse = {
      id: "cmpl-xxx",
      choices: [{ message: { role: "assistant", content: "Hello!" } }],
      usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
    };
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(JSON.stringify(fakeResponse), { status: 200, headers: { "Content-Type": "application/json" } })
    );
    const result = await chatCompletion(sampleMessages, sampleConfig);
    expect(result.choices[0].message.content).toBe("Hello!");
  });

  it("throws ApiError with AUTH code on 401", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(JSON.stringify({ error: { message: "unauthorized" } }), { status: 401 })
    );
    await expect(chatCompletion(sampleMessages, sampleConfig)).rejects.toMatchObject({
      code: "AUTH",
      status: 401,
      retryable: false,
    });
  });

  it("throws ApiError with RATE_LIMIT on 429", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response("Too Many Requests", { status: 429 })
    );
    await expect(chatCompletion(sampleMessages, sampleConfig)).rejects.toMatchObject({
      code: "RATE_LIMIT",
      status: 429,
      retryable: true,
    });
  });

  it("throws ApiError with SERVER on 503", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response("Service Unavailable", { status: 503 })
    );
    await expect(chatCompletion(sampleMessages, sampleConfig)).rejects.toMatchObject({
      code: "SERVER",
      status: 503,
      retryable: true,
    });
  });

  it("wraps network error as ApiError", async () => {
    vi.spyOn(globalThis, "fetch").mockRejectedValueOnce(new TypeError("network failure"));
    await expect(chatCompletion(sampleMessages, sampleConfig)).rejects.toMatchObject({
      code: "NETWORK",
    });
  });
});

describe("chatCompletionStream (mocked)", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  function mockSseStream(chunks: string[]): void {
    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        for (const c of chunks) {
          controller.enqueue(encoder.encode(c));
        }
        controller.close();
      },
    });
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(stream, { status: 200, headers: { "Content-Type": "text/event-stream" } })
    );
  }

  it("streams chunks and calls onDone", async () => {
    const chunks = [
      "data: {\"choices\":[{\"delta\":{\"content\":\"Hello\"},\"index\":0}]}\n\n",
      "data: {\"choices\":[{\"delta\":{\"content\":\" World\"},\"index\":0}]}\n\n",
      "data: [DONE]\n\n",
    ];
    mockSseStream(chunks);

    const onChunk = vi.fn();
    const onDone = vi.fn();
    const onError = vi.fn();

    await chatCompletionStream(sampleMessages, sampleConfig, onChunk, onDone, onError);

    // Wait for the async processing to complete
    await vi.waitFor(() => {
      expect(onChunk).toHaveBeenCalledTimes(2);
    });
    expect(onChunk).toHaveBeenNthCalledWith(1, "Hello", "Hello");
    expect(onChunk).toHaveBeenNthCalledWith(2, " World", "Hello World");
    expect(onDone).toHaveBeenCalledWith("Hello World");
    expect(onError).not.toHaveBeenCalled();
  });

  it("calls onError on server error", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response("Internal Error", { status: 500 })
    );

    const onChunk = vi.fn();
    const onDone = vi.fn();
    const onError = vi.fn();

    await chatCompletionStream(sampleMessages, sampleConfig, onChunk, onDone, onError);

    await vi.waitFor(() => {
      expect(onError).toHaveBeenCalled();
    });
    expect(onError).toHaveBeenCalledWith(
      expect.objectContaining({ code: "SERVER", status: 500 })
    );
    expect(onDone).not.toHaveBeenCalled();
  });

  it("streams tool calls and calls onToolCalls", async () => {
    const chunks = [
      "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"call_1\",\"function\":{\"name\":\"writeFormula\",\"arguments\":\"\"}}]},\"index\":0}]}\n\n",
      "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"arguments\":\"{\\\"f\"}}]},\"index\":0}]}\n\n",
      "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"arguments\":\"ormula\\\":\\\"=SUM(A1:A10)\\\"}\"}}]},\"index\":0}]}\n\n",
      "data: {\"choices\":[{\"delta\":{},\"finish_reason\":\"tool_calls\"}],\"index\":0}\n\n",
      "data: [DONE]\n\n",
    ];
    mockSseStream(chunks);

    const onChunk = vi.fn();
    const onDone = vi.fn();
    const onError = vi.fn();
    const onToolCalls = vi.fn();

    await chatCompletionStream(sampleMessages, sampleConfig, onChunk, onDone, onError, {
      withTools: true,
      onToolCalls,
    });

    await vi.waitFor(() => {
      expect(onToolCalls).toHaveBeenCalled();
    });
    expect(onToolCalls).toHaveBeenCalledWith(
      expect.arrayContaining([
        expect.objectContaining({ id: "call_1", name: "writeFormula" }),
      ])
    );
    expect(onDone).toHaveBeenCalled();
    expect(onError).not.toHaveBeenCalled();
  });
});

/* ----------------------------------------------------------------- */
/*  Formula diagnosis                                                 */
/* ----------------------------------------------------------------- */

describe("buildDiagnosisMessages", () => {
  it("builds system + user messages with error details", () => {
    const req: FormulaDiagnosisRequest = {
      sheetName: "Sheet1",
      rangeAddress: "A1:C10",
      errors: [
        { address: "B3", fullAddress: "Sheet1!B3", formula: "=A1/0", value: "#DIV/0!", error: "#DIV/0!" },
      ],
    };
    const messages = buildDiagnosisMessages(req);
    expect(messages).toHaveLength(2);
    expect(messages[0].role).toBe("system");
    expect(messages[1].role).toBe("user");
    expect(messages[1].content).toContain("Sheet1");
    expect(messages[1].content).toContain("#DIV/0!");
  });
});

describe("tryParseDiagnosisJson", () => {
  it("parses a clean JSON array", () => {
    const input = '[{"address":"B3","error":"#DIV/0!","cause":"除以零","suggestion":"检查分母","fixedFormula":"=IF(B1=0,\\"\\",A1/B1)","confidence":0.9}]';
    const result = tryParseDiagnosisJson(input);
    expect(result).toHaveLength(1);
    expect(result[0].address).toBe("B3");
    expect(result[0].error).toBe("#DIV/0!");
  });

  it("parses JSON inside markdown fences", () => {
    const input = "```json\n[{\"address\":\"C5\",\"error\":\"#REF!\",\"cause\":\"引用已删除\",\"suggestion\":\"更新引用\",\"fixedFormula\":\"\",\"confidence\":0.7}]\n```";
    const result = tryParseDiagnosisJson(input);
    expect(result).toHaveLength(1);
    expect(result[0].error).toBe("#REF!");
  });

  it("extracts JSON array from surrounding prose", () => {
    const input = "Here are the errors:\n[{\"address\":\"D2\",\"error\":\"#N/A\",\"cause\":\"VLOOKUP not found\",\"suggestion\":\"Check lookup value\",\"fixedFormula\":\"\",\"confidence\":0.6}]\nEnd.";
    const result = tryParseDiagnosisJson(input);
    expect(result).toHaveLength(1);
  });

  it("returns empty fallback entry for unparseable input", () => {
    const result = tryParseDiagnosisJson("totally invalid garbage");
    expect(result).toHaveLength(1);
    expect(result[0].cause).toContain("无法解析");
  });

  it("returns empty array for empty input", () => {
    expect(tryParseDiagnosisJson("")).toEqual([]);
  });
});

describe("isValidDiagnosis", () => {
  it("returns true for valid object", () => {
    expect(isValidDiagnosis({ address: "A1" })).toBe(true);
  });

  it("returns false for null", () => {
    expect(isValidDiagnosis(null)).toBe(false);
  });

  it("returns false for object without address", () => {
    expect(isValidDiagnosis({ error: "#REF!" })).toBe(false);
  });
});

/* ----------------------------------------------------------------- */
/*  stripCodeFences                                                   */
/* ----------------------------------------------------------------- */

describe("stripCodeFences", () => {
  it("removes code fences around VBA", () => {
    const input = "```vba\nSub Test()\nEnd Sub\n```";
    expect(stripCodeFences(input)).toBe("Sub Test()\nEnd Sub");
  });

  it("removes code fences around TypeScript", () => {
    const input = "```typescript\nlet x = 1;\n```";
    expect(stripCodeFences(input)).toBe("let x = 1;");
  });

  it("returns the string as-is when no fences", () => {
    expect(stripCodeFences("hello")).toBe("hello");
  });

  it("handles null / undefined", () => {
    expect(stripCodeFences(null as any)).toBeNull();
    expect(stripCodeFences(undefined as any)).toBeUndefined();
  });
});

/* ----------------------------------------------------------------- */
/*  Code-gen message building                                         */
/* ----------------------------------------------------------------- */

describe("buildCodeGenMessages", () => {
  it("builds system + user for VBA", () => {
    const { system, user } = buildCodeGenMessages("Sheet1", "B3 =SUM(A1:A10)", "vba");
    expect(system).toContain("VBA");
    expect(user).toContain("Sheet1");
    expect(user).toContain("B3");
  });

  it("builds system + user for Office Scripts", () => {
    const { system, user } = buildCodeGenMessages("Sheet1", "C5 =AVERAGE(D1:D10)", "office-scripts");
    expect(system).toContain("Office Script");
    expect(user).toContain("Sheet1");
    expect(user).toContain("C5");
  });
});

/* ----------------------------------------------------------------- */
/*  Chart recommendation                                              */
/* ----------------------------------------------------------------- */

describe("buildChartRecommendMessages", () => {
  it("builds system + user with data preview", () => {
    const req: ChartRecommendRequest = {
      sheet: "Sales",
      headers: ["Month", "Revenue"],
      preview: [["Jan", "100"], ["Feb", "200"]],
      rowCount: 12,
      columnCount: 2,
    };
    const { system, user } = buildChartRecommendMessages(req);
    expect(system).toContain("ColumnClustered");
    expect(user).toContain("Sales");
    expect(user).toContain("Month");
    expect(user).toContain("Revenue");
  });
});

describe("tryParseChartList", () => {
  it("parses valid chart recommendations", () => {
    const input = '[{"type":"ColumnClustered","title":"柱状图","reason":"最佳对比视图"}]';
    const result = tryParseChartList(input);
    expect(result).toHaveLength(1);
    expect(result[0].type).toBe("ColumnClustered");
  });

  it("strips markdown fences before parsing", () => {
    const input = '```json\n[{"type":"Line","title":"折线图","reason":"展示趋势"}]\n```';
    const result = tryParseChartList(input);
    expect(result).toHaveLength(1);
    expect(result[0].type).toBe("Line");
  });

  it("returns empty array for unparseable input", () => {
    expect(tryParseChartList("nope")).toEqual([]);
  });

  it("filters out unsupported chart types", () => {
    const input = '[{"type":"Bubble","title":"Bubble","reason":"test"},{"type":"Pie","title":"饼图","reason":"占比"}]';
    const result = tryParseChartList(input);
    expect(result).toHaveLength(1);
    expect(result[0].type).toBe("Pie");
  });
});

describe("fallbackChartRecommend", () => {
  function makeReq(overrides: Partial<ChartRecommendRequest> = {}): ChartRecommendRequest {
    return {
      sheet: "Sheet1",
      headers: ["Category", "Value"],
      preview: [["A", "1"], ["B", "2"]],
      rowCount: 5,
      columnCount: 2,
      ...overrides,
    };
  }

  it("always returns 3 recommendations", () => {
    const result = fallbackChartRecommend(makeReq());
    expect(result).toHaveLength(3);
  });

  it("returns Pie for small datasets with 2 columns", () => {
    const result = fallbackChartRecommend(makeReq({ rowCount: 5, columnCount: 2 }));
    const types = result.map((r) => r.type);
    expect(types).toContain("Pie");
  });

  it("returns Area for large datasets", () => {
    const result = fallbackChartRecommend(makeReq({ rowCount: 20, columnCount: 4 }));
    const types = result.map((r) => r.type);
    expect(types).toContain("Area");
    expect(types).not.toContain("Pie");
  });

  it("every recommendation has type, title, reason", () => {
    const result = fallbackChartRecommend(makeReq());
    for (const r of result) {
      expect(r.type).toBeTruthy();
      expect(r.title).toBeTruthy();
      expect(r.reason).toBeTruthy();
    }
  });
});

/* ----------------------------------------------------------------- */
/*  diagnoseFormulasStream (mocked)                                   */
/* ----------------------------------------------------------------- */

describe("diagnoseFormulasStream (mocked)", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("calls onDone with parsed results", async () => {
    const req: FormulaDiagnosisRequest = {
      sheetName: "Sheet1",
      rangeAddress: "A1:C10",
      errors: [{ address: "B3", fullAddress: "Sheet1!B3", formula: "=A1/0", value: "#DIV/0!", error: "#DIV/0!" }],
    };

    // Mock the underlying fetch that chatCompletionStream uses
    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        const chunk = 'data: {"choices":[{"delta":{"content":"[{\\\"address\\\":\\\"B3\\\",\\\"error\\\":\\\"#DIV/0!\\\",\\\"cause\\\":\\\"除以零\\\",\\\"suggestion\\\":\\\"检查分母\\\",\\\"fixedFormula\\\":\\\"=IF(B1=0,\\\"\\\",A1/B1)\\\",\\\"confidence\\\":0.9}]"},"index":0}]}\n\n';
        controller.enqueue(encoder.encode(chunk));
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        controller.close();
      },
    });
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(stream, { status: 200, headers: { "Content-Type": "text/event-stream" } })
    );

    const onPartial = vi.fn();
    const onDone = vi.fn();
    const onError = vi.fn();

    await diagnoseFormulasStream(req, sampleConfig, { onPartial, onDone, onError });

    await vi.waitFor(() => {
      expect(onDone).toHaveBeenCalled();
    });
    const results = onDone.mock.calls[0][0];
    expect(Array.isArray(results)).toBe(true);
    expect(onError).not.toHaveBeenCalled();
  });
});

/* ----------------------------------------------------------------- */
/*  translateToScript (mocked)                                        */
/* ----------------------------------------------------------------- */

describe("translateToScript (mocked)", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("calls onDone with generated code", async () => {
    const formulas = { sheet: "Sheet1", block: "A1 =SUM(B1:B10)" };
    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        controller.enqueue(encoder.encode("data: {\"choices\":[{\"delta\":{\"content\":\"Sub Test()\\nEnd Sub\"},\"index\":0}]}\n\n"));
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        controller.close();
      },
    });
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(stream, { status: 200, headers: { "Content-Type": "text/event-stream" } })
    );

    const onPartial = vi.fn();
    const onDone = vi.fn();
    const onError = vi.fn();

    await translateToScript(formulas, "vba", sampleConfig, { onPartial, onDone, onError });

    await vi.waitFor(() => {
      expect(onDone).toHaveBeenCalled();
    });
    expect(onDone.mock.calls[0][0]).toContain("Sub Test");
  });
});

/* ----------------------------------------------------------------- */
/*  recommendChartStream (mocked)                                     */
/* ----------------------------------------------------------------- */

describe("recommendChartStream (mocked)", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("returns recommendations on successful stream", async () => {
    const req: ChartRecommendRequest = {
      sheet: "Sales",
      headers: ["Month", "Revenue"],
      preview: [["Jan", "100"]],
      rowCount: 3,
      columnCount: 2,
    };

    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        controller.enqueue(encoder.encode(
          "data: {\"choices\":[{\"delta\":{\"content\":\"[{\\\"type\\\":\\\"ColumnClustered\\\",\\\"title\\\":\\\"柱状图\\\",\\\"reason\\\":\\\"对比各月\\\"},{\\\"type\\\":\\\"Pie\\\",\\\"title\\\":\\\"饼图\\\",\\\"reason\\\":\\\"占比\\\"}]\"},\"index\":0}]}\n\n"
        ));
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        controller.close();
      },
    });
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(stream, { status: 200, headers: { "Content-Type": "text/event-stream" } })
    );

    const onPartial = vi.fn();
    const onDone = vi.fn();
    const onError = vi.fn();

    await recommendChartStream(sampleConfig, req, { onPartial, onDone, onError });

    await vi.waitFor(() => {
      expect(onDone).toHaveBeenCalled();
    });
    const result = onDone.mock.calls[0][0];
    expect(result.recommendations.length).toBeGreaterThanOrEqual(1);
    expect(result).toHaveProperty("inferredCategory");
    expect(result).toHaveProperty("usedFallback");
  });

  it("uses fallback when no JSON is parsed", async () => {
    const req: ChartRecommendRequest = {
      sheet: "Sheet1",
      headers: ["X", "Y"],
      preview: [["a", "1"]],
      rowCount: 5,
      columnCount: 2,
    };

    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        controller.enqueue(encoder.encode("data: {\"choices\":[{\"delta\":{\"content\":\"nope\"},\"index\":0}]}\n\n"));
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        controller.close();
      },
    });
    vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(
      new Response(stream, { status: 200, headers: { "Content-Type": "text/event-stream" } })
    );

    const onDone = vi.fn();
    const onError = vi.fn();

    await recommendChartStream(sampleConfig, req, { onDone, onError });

    await vi.waitFor(() => {
      expect(onDone).toHaveBeenCalled();
    });
    const result = onDone.mock.calls[0][0];
    expect(result.usedFallback).toBe(true);
    expect(result.recommendations.length).toBeGreaterThanOrEqual(1);
  });
});
