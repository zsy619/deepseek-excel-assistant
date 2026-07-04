/**
 * ============================================================================
 * helpers.test.ts
 * ----------------------------------------------------------------------------
 * Tests for pure utility functions. No mocks needed — these functions have
 * zero DOM / Office.js dependencies.
 * ============================================================================
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  generateId,
  estimateTokens,
  formatTokens,
  formatRelativeTime,
  deriveSessionTitle,
  escapeHtml,
  toApiError,
  excelValuesToMarkdown,
  sleep,
  clampInt,
  clampFloat,
  copyToClipboard,
  isValidApiKeyFormat,
  safeJsonParse,
} from "./helpers";

/* ------------------------------------------------------------------ *
 *  generateId                                                        *
 * ------------------------------------------------------------------ */

describe("generateId", () => {
  it("returns a non-empty string", () => {
    const id = generateId();
    expect(id).toBeTruthy();
    expect(typeof id).toBe("string");
  });

  it("produces unique values on successive calls", () => {
    const ids = new Set(Array.from({ length: 100 }, () => generateId()));
    expect(ids.size).toBe(100);
  });
});

/* ------------------------------------------------------------------ *
 *  estimateTokens                                                    *
 * ------------------------------------------------------------------ */

describe("estimateTokens", () => {
  it("returns 0 for empty / falsy input", () => {
    expect(estimateTokens("")).toBe(0);
    expect(estimateTokens(undefined as any)).toBe(0);
    expect(estimateTokens(null as any)).toBe(0);
  });

  it("counts CJK characters as ~1 token each", () => {
    // 4 CJK characters
    expect(estimateTokens("你好世界")).toBe(4);
  });

  it("counts ASCII letters at ~0.25 token each", () => {
    // 4 ASCII chars → ~1 token
    const t = estimateTokens("abcd");
    expect(t).toBe(1);
  });

  it("adds 5% overhead for code fences", () => {
    const without = estimateTokens("hello world");
    const withCode = estimateTokens("```\nhello world\n```");
    // withCode should be ~5% larger
    expect(withCode).toBeGreaterThan(without);
  });

  it("returns at least 1", () => {
    // A single ASCII char should round up to 1 token
    expect(estimateTokens("a")).toBe(1);
    expect(estimateTokens(" ")).toBe(1);
  });
});

/* ------------------------------------------------------------------ *
 *  formatTokens                                                      *
 * ------------------------------------------------------------------ */

describe("formatTokens", () => {
  it("returns raw number for < 1000", () => {
    expect(formatTokens(0)).toBe("0");
    expect(formatTokens(42)).toBe("42");
    expect(formatTokens(999)).toBe("999");
  });

  it("formats with one decimal for 1k-9.9k", () => {
    expect(formatTokens(1000)).toBe("1.0k");
    expect(formatTokens(1500)).toBe("1.5k");
    expect(formatTokens(9999)).toBe("10.0k");
  });

  it("rounds to integer k for >= 10000", () => {
    expect(formatTokens(10000)).toBe("10k");
    expect(formatTokens(12345)).toBe("12k");
    expect(formatTokens(99999)).toBe("100k");
  });
});

/* ------------------------------------------------------------------ *
 *  formatRelativeTime                                                *
 * ------------------------------------------------------------------ */

describe("formatRelativeTime", () => {
  const now = 1_000_000_000_000;

  it('returns "刚刚" for < 1 minute', () => {
    expect(formatRelativeTime(now - 30_000, now)).toBe("刚刚");
  });

  it("returns N 分钟前 for < 1 hour", () => {
    expect(formatRelativeTime(now - 5 * 60_000, now)).toBe("5 分钟前");
  });

  it("returns N 小时前 for < 24 hours", () => {
    expect(formatRelativeTime(now - 3 * 3600_000, now)).toBe("3 小时前");
  });

  it("returns N 天前 for < 7 days", () => {
    expect(formatRelativeTime(now - 2 * 86400_000, now)).toBe("2 天前");
  });

  it("returns absolute date for >= 7 days", () => {
    const ts = new Date("2024-01-05T10:30:00").getTime();
    const later = new Date("2024-01-15T10:30:00").getTime();
    const result = formatRelativeTime(ts, later);
    expect(result).toContain("2024-01-05");
    expect(result).toContain("10:30");
  });
});

/* ------------------------------------------------------------------ *
 *  deriveSessionTitle                                                *
 * ------------------------------------------------------------------ */

describe("deriveSessionTitle", () => {
  it('returns "新对话" for empty input', () => {
    expect(deriveSessionTitle("")).toBe("新对话");
    expect(deriveSessionTitle("   ")).toBe("新对话");
  });

  it("truncates long messages", () => {
    const long = "这是一个非常长的用户消息，用来测试标题截取功能是否正常工作";
    const title = deriveSessionTitle(long);
    expect(title.length).toBeLessThanOrEqual(21); // 20 + ellipsis
    expect(title).toMatch(/…$/);
  });

  it("returns short messages as-is", () => {
    expect(deriveSessionTitle("你好")).toBe("你好");
  });
});

/* ------------------------------------------------------------------ *
 *  escapeHtml                                                        *
 * ------------------------------------------------------------------ */

describe("escapeHtml", () => {
  it("escapes & < > \" '", () => {
    const result = escapeHtml(`<script>alert("xss")</script>`);
    expect(result).not.toContain("<");
    expect(result).not.toContain(">");
    expect(result).not.toContain('"');
    expect(result).toContain("&lt;");
    expect(result).toContain("&gt;");
    expect(result).toContain("&quot;");
  });

  it("passes safe strings through", () => {
    expect(escapeHtml("hello world")).toBe("hello world");
  });
});

/* ------------------------------------------------------------------ *
 *  toApiError                                                        *
 * ------------------------------------------------------------------ */

describe("toApiError", () => {
  it("passes through valid ApiError objects", () => {
    const err = { code: "RATE_LIMIT", message: "too fast", retryable: true };
    expect(toApiError(err)).toBe(err);
  });

  it("converts AbortError", () => {
    const err = new Error("aborted");
    err.name = "AbortError";
    const result = toApiError(err);
    expect(result.code).toBe("ABORTED");
    expect(result.retryable).toBe(false);
  });

  it("converts timeout errors", () => {
    const err = new Error("request timeout");
    const result = toApiError(err);
    expect(result.code).toBe("TIMEOUT");
    expect(result.retryable).toBe(true);
  });

  it("wraps generic errors", () => {
    const result = toApiError(new Error("network failure"));
    expect(result.code).toBe("NETWORK");
    expect(result.retryable).toBe(true);
  });

  it("wraps unknown throwables", () => {
    const result = toApiError("just a string");
    expect(result.code).toBe("UNKNOWN");
    expect(result.retryable).toBe(true);
  });

  it("wraps null / undefined", () => {
    const result = toApiError(null);
    expect(result.code).toBe("UNKNOWN");
  });
});

/* ------------------------------------------------------------------ *
 *  excelValuesToMarkdown                                             *
 * ------------------------------------------------------------------ */

describe("excelValuesToMarkdown", () => {
  it("returns '(空)' for empty input", () => {
    expect(excelValuesToMarkdown([])).toBe("(空)");
    expect(excelValuesToMarkdown(null as any)).toBe("(空)");
  });

  it("produces a markdown table", () => {
    const data = [
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 25],
    ];
    const md = excelValuesToMarkdown(data);
    expect(md).toContain("| Name | Age |");
    expect(md).toContain("| Alice | 30 |");
    expect(md).toContain("| Bob | 25 |");
    expect(md).toContain("| --- | --- |");
  });

  it("handles null/undefined cells", () => {
    const data = [["A", "B"], [null, undefined]];
    const md = excelValuesToMarkdown(data);
    expect(md).toContain("|  |  |");
  });

  it("escapes pipe characters in cell values", () => {
    const data = [["A | B"]];
    const md = excelValuesToMarkdown(data);
    expect(md).toContain("A \\| B");
  });
});

/* ------------------------------------------------------------------ *
 *  sleep                                                             *
 * ------------------------------------------------------------------ */

describe("sleep", () => {
  it("resolves after approximately the given time", async () => {
    vi.useFakeTimers();
    const p = sleep(100);
    await vi.advanceTimersByTimeAsync(100);
    await expect(p).resolves.toBeUndefined();
    vi.useRealTimers();
  });
});

/* ------------------------------------------------------------------ *
 *  clampInt / clampFloat                                             *
 * ------------------------------------------------------------------ */

describe("clampInt", () => {
  it("clamps within bounds", () => {
    expect(clampInt(5, 0, 10, 0)).toBe(5);
    expect(clampInt(-1, 0, 10, 0)).toBe(0);
    expect(clampInt(15, 0, 10, 0)).toBe(10);
  });

  it("floors floating-point input", () => {
    expect(clampInt(3.9, 0, 10, 0)).toBe(3);
  });

  it("returns fallback for NaN", () => {
    expect(clampInt("abc", 0, 10, 5)).toBe(5);
  });

  it("returns 0 for null (Number(null) = 0), fallback for undefined", () => {
    // null → Number(null) = 0, which is valid within [0,10]
    expect(clampInt(null, 0, 10, 5)).toBe(0);
    // undefined → Number(undefined) = NaN → returns fallback
    expect(clampInt(undefined, 0, 10, 5)).toBe(5);
  });
});

describe("clampFloat", () => {
  it("clamps within bounds", () => {
    expect(clampFloat(0.5, 0, 1, 0)).toBe(0.5);
    expect(clampFloat(-0.1, 0, 1, 0)).toBe(0);
    expect(clampFloat(1.5, 0, 1, 0)).toBe(1);
  });

  it("returns fallback for NaN", () => {
    expect(clampFloat("abc", 0, 1, 0.5)).toBe(0.5);
  });
});

/* ------------------------------------------------------------------ *
 *  copyToClipboard                                                   *
 * ------------------------------------------------------------------ */

describe("copyToClipboard", () => {
  beforeEach(() => {
    Object.assign(navigator, {
      clipboard: { writeText: vi.fn().mockResolvedValue(undefined) },
    });
    window.isSecureContext = true;
  });

  it("uses clipboard API when available", async () => {
    await expect(copyToClipboard("test")).resolves.toBe(true);
    expect(navigator.clipboard.writeText).toHaveBeenCalledWith("test");
  });

  it("falls back gracefully when clipboard API fails", async () => {
    (navigator.clipboard.writeText as any).mockRejectedValue(new Error("fail"));
    // In jsdom, document.execCommand('copy') returns false, so expect fallback
    // to also fail gracefully
    const result = await copyToClipboard("fallback");
    expect(result).toBe(false);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });
});

/* ------------------------------------------------------------------ *
 *  isValidApiKeyFormat                                               *
 * ------------------------------------------------------------------ */

describe("isValidApiKeyFormat", () => {
  it("rejects empty / short keys", () => {
    expect(isValidApiKeyFormat("")).toBe(false);
    expect(isValidApiKeyFormat("short")).toBe(false);
  });

  it("accepts keys >= 20 chars", () => {
    expect(isValidApiKeyFormat("sk-" + "a".repeat(17))).toBe(true);
  });

  it("rejects null / undefined", () => {
    expect(isValidApiKeyFormat(null as any)).toBe(false);
    expect(isValidApiKeyFormat(undefined as any)).toBe(false);
  });
});

/* ------------------------------------------------------------------ *
 *  safeJsonParse                                                     *
 * ------------------------------------------------------------------ */

describe("safeJsonParse", () => {
  it("parses valid JSON", () => {
    expect(safeJsonParse('{"a":1}', null)).toEqual({ a: 1 });
  });

  it("returns fallback for invalid JSON", () => {
    expect(safeJsonParse("{bad}", { fallback: true })).toEqual({ fallback: true });
  });

  it("returns fallback for empty string", () => {
    expect(safeJsonParse("", "default")).toBe("default");
  });

  it("returns fallback for null parsed", () => {
    expect(safeJsonParse("null", "fallback")).toBe("fallback");
  });
});
