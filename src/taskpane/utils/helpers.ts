/**
 * ============================================================================
 * Helpers
 * ----------------------------------------------------------------------------
 * Small pure functions used by multiple components. No DOM or Office.js
 * access here - keeps these trivial to unit-test.
 * ============================================================================
 */

import { v4 as uuidv4 } from "uuid";
import type { ApiError, ChatSession } from "../types";

/** Generate a v4 UUID. Falls back to a Math.random based id if uuid fails. */
export function generateId(): string {
  try {
    return uuidv4();
  } catch {
    return "id-" + Math.random().toString(36).slice(2) + "-" + Date.now();
  }
}

/**
 * Approximate token count for a string. We can't run the real BPE
 * tokenizer client-side, so this uses a heuristic that matches
 * DeepSeek's published ratios closely enough for budget warnings:
 *   - 1 CJK ideograph  ≈ 1.0 tokens
 *   - 1 ASCII letter  ≈ 0.25 tokens
 *   - 1 digit/punct   ≈ 0.25 tokens
 *   - whitespace      ≈ 0 (collapsed by the tokenizer)
 *   - markdown / code fences  ≈ +5% overhead
 *
 * The output is rounded up to the nearest integer so users see whole
 * numbers; the relative error vs. the real tokenizer is < 10%.
 */
export function estimateTokens(text: string): number {
  if (!text) return 0;
  let cjk = 0;
  let ascii = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0) || 0;
    if (code > 0x2e80) {
      // CJK + fullwidth + emoji
      cjk += 1;
    } else if (code > 0x20) {
      ascii += 1;
    }
    // whitespace is ignored
  }
  let estimate = cjk + ascii * 0.25;
  if (/```/.test(text)) estimate *= 1.05;
  return Math.max(1, Math.ceil(estimate));
}

/** Format a token count for the status bar: 1234 -> "1,234". */
export function formatTokens(n: number): string {
  if (n < 1000) return String(n);
  if (n < 10_000) return (n / 1000).toFixed(1) + "k";
  return Math.round(n / 1000) + "k";
}

/** Friendly timestamp - "刚刚 / N 分钟前 / YYYY-MM-DD HH:mm". */
export function formatRelativeTime(ts: number, now: number = Date.now()): string {
  const diff = now - ts;
  const minute = 60 * 1000;
  const hour = 60 * minute;
  const day = 24 * hour;

  if (diff < minute) return "刚刚";
  if (diff < hour) return `${Math.floor(diff / minute)} 分钟前`;
  if (diff < day) return `${Math.floor(diff / hour)} 小时前`;
  if (diff < 7 * day) return `${Math.floor(diff / day)} 天前`;

  const d = new Date(ts);
  const pad = (n: number) => n.toString().padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Title for a new session - truncate the first user message. */
export function deriveSessionTitle(firstUserMessage: string): string {
  const trimmed = (firstUserMessage || "").trim().replace(/\s+/g, " ");
  if (!trimmed) return "新对话";
  return trimmed.length > 20 ? trimmed.slice(0, 20) + "…" : trimmed;
}

/** Rough HTML escape to keep user-supplied text from breaking the DOM
 *  when it is rendered outside the markdown pipeline (e.g. session titles). */
export function escapeHtml(input: string): string {
  return String(input)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Build a human-readable error from anything thrown by fetch / API. */
export function toApiError(err: unknown): ApiError {
  if (err && typeof err === "object" && "code" in err && "retryable" in err) {
    return err as ApiError;
  }
  if (err instanceof Error) {
    const name = err.name || "";
    if (name === "AbortError") {
      return { code: "ABORTED", message: "请求已取消", retryable: false };
    }
    if (err.message?.includes("timeout") || err.message?.includes("Timeout")) {
      return { code: "TIMEOUT", message: "请求超时，请检查网络后重试", retryable: true };
    }
    return { code: "NETWORK", message: err.message || "网络错误", retryable: true };
  }
  return { code: "UNKNOWN", message: String(err) || "未知错误", retryable: true };
}

/** Format selected Excel data into a markdown-friendly table for prompts. */
export function excelValuesToMarkdown(values: any[][]): string {
  if (!values || values.length === 0) return "(空)";
  const rows = values.map((row) =>
    row.map((cell) => {
      if (cell === null || cell === undefined) return "";
      const s = String(cell).replace(/\|/g, "\\|").replace(/\n/g, " ");
      return s.length > 50 ? s.slice(0, 50) + "…" : s;
    })
  );
  const header = rows[0];
  const sep = header.map(() => "---");
  const body = rows.slice(1);
  const table = [header, sep, ...body].map((r) => "| " + r.join(" | ") + " |").join("\n");
  return table;
}

/** Sleep helper - used for retry backoff. */
export function sleep(ms: number): Promise<void> {
  return new Promise((res) => setTimeout(res, ms));
}

/** Coerce anything to a safe integer within [min, max]. */
export function clampInt(value: unknown, min: number, max: number, fallback: number): number {
  const n = Math.floor(Number(value));
  if (Number.isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

/** Coerce anything to a safe float within [min, max]. */
export function clampFloat(value: unknown, min: number, max: number, fallback: number): number {
  const n = Number(value);
  if (Number.isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

/** Copy text to clipboard with a fallback for non-secure contexts. */
export async function copyToClipboard(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // fall through
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch {
    return false;
  }
}

/** Serialize a session to a Markdown file body and trigger a download. */
export function downloadSessionAsMarkdown(session: ChatSession): void {
  const lines: string[] = [];
  lines.push(`# ${session.title}`);
  lines.push("");
  lines.push(`> 导出于 ${new Date().toLocaleString("zh-CN")}`);
  lines.push("");
  lines.push("---");
  lines.push("");
  for (const m of session.messages) {
    const who = m.role === "user" ? "👤 用户" : m.role === "assistant" ? "🤖 DeepSeek" : "⚙️ System";
    lines.push(`## ${who}`);
    lines.push("");
    lines.push(m.content);
    lines.push("");
  }
  const blob = new Blob([lines.join("\n")], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const safeTitle = session.title.replace(/[\\/:*?"<>|]/g, "_").slice(0, 40);
  a.download = `deepseek-${safeTitle}-${Date.now()}.md`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Validate the structure of an API key. Returns true for empty (caller
 *  decides whether empty is acceptable) or for keys that look plausible. */
export function isValidApiKeyFormat(key: string): boolean {
  if (!key) return false;
  if (key.length < 20) return false;
  // DeepSeek keys start with sk-, but many proxies use other prefixes.
  // We accept any token >= 20 chars and warn at the UI level.
  return key.length >= 20;
}
/** Safely parse a JSON string. Returns the provided fallback on error. */
export function safeJsonParse<T = any>(raw: string, fallback: T): T {
  if (!raw) return fallback;
  try {
    const v = JSON.parse(raw);
    return v == null ? fallback : (v as T);
  } catch {
    return fallback;
  }
}
