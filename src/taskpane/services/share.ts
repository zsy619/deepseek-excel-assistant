/**
 * ============================================================================
 * Share / collaborate service (PRD-11)
 * ----------------------------------------------------------------------------
 * Two delivery channels for sharing a chat session with someone else:
 *
 *   1) URL fragment: encode → base64 → "#share=..." appended to the current
 *      URL. Recipient opens the same add-in and we auto-detect + restore.
 *      Best for "send me the link" workflows.
 *
 *   2) File export / import: download as .json or load .json via a file
 *      picker. Best for "archive this thread" / attach-in-email workflows.
 *
 * Compression: base64-encoded JSON. We don't need gzip because the Office
 * taskpane URL length budget is generous and the JSON is short (typically
 * <30KB for a few hundred messages).
 * ============================================================================
 */

import type { ChatSession } from "../types";

const URL_KEY = "share";
const FILE_VERSION = 1;

export interface ShareableSession {
  v: number;
  title: string;
  createdAt: number;
  messages: ChatSession["messages"];
}

/* ---------------- Encode / decode ---------------- */

export function encodeSession(session: ChatSession): string {
  const payload: ShareableSession = {
    v: FILE_VERSION,
    title: session.title,
    createdAt: session.createdAt,
    messages: session.messages,
  };
  const json = JSON.stringify(payload);
  return base64UrlEncode(json);
}

export function decodeSession(encoded: string): ShareableSession | null {
  try {
    const json = base64UrlDecode(encoded);
    if (!json) return null;
    const parsed = JSON.parse(json);
    if (!parsed || typeof parsed !== "object") return null;
    if (parsed.v !== FILE_VERSION) return null;
    if (!Array.isArray(parsed.messages)) return null;
    return parsed as ShareableSession;
  } catch {
    return null;
  }
}

/* ---------------- URL helpers ---------------- */

/** Build a shareable URL pointing to the current add-in. The recipient
 *  must be on the same host (Office add-ins run from a fixed URL). */
export function buildShareUrl(encoded: string): string {
  const base = window.location.origin + window.location.pathname;
  return `${base}#${URL_KEY}=${encoded}`;
}

/** Look at the current URL fragment and return any decoded share, or null. */
export function readShareFromUrl(): ShareableSession | null {
  const hash = window.location.hash || "";
  const m = hash.match(new RegExp(`#${URL_KEY}=([^&]+)`));
  if (!m) return null;
  return decodeSession(decodeURIComponent(m[1]));
}

/** Strip the share fragment from the URL so refreshing doesn't re-trigger. */
export function clearShareFromUrl(): void {
  if (window.history && window.history.replaceState) {
    const url = window.location.origin + window.location.pathname + window.location.search;
    window.history.replaceState(null, document.title, url);
  } else {
    window.location.hash = "";
  }
}

/* ---------------- File helpers ---------------- */

export function sessionToShareableFile(session: ChatSession): Blob {
  const payload: ShareableSession = {
    v: FILE_VERSION,
    title: session.title,
    createdAt: session.createdAt,
    messages: session.messages,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json",
  });
  return blob;
}

export async function fileToShareable(file: File): Promise<ShareableSession> {
  const text = await file.text();
  const parsed = JSON.parse(text);
  if (!parsed || parsed.v !== FILE_VERSION || !Array.isArray(parsed.messages)) {
    throw new Error("文件格式不匹配：请确认是 DeepSeek Excel Assistant 导出的会话");
  }
  return parsed as ShareableSession;
}

/** Trigger a browser download of the session JSON. */
export function downloadSessionFile(session: ChatSession, fileName?: string): void {
  const blob = sessionToShareableFile(session);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName || `chat-${slugify(session.title)}-${Date.now()}.json`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 1000);
}

/* ---------------- Base64URL helpers ---------------- */

function base64UrlEncode(s: string): string {
  // First encode as UTF-8 bytes, then base64, then URL-safe.
  const bytes = new TextEncoder().encode(s);
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  const b64 = typeof btoa !== "undefined" ? btoa(bin) : Buffer.from(bin, "binary").toString("base64");
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function base64UrlDecode(s: string): string | null {
  try {
    const b64 = s.replace(/-/g, "+").replace(/_/g, "/");
    const padded = b64 + "===".slice((b64.length + 3) % 4);
    const bin = typeof atob !== "undefined" ? atob(padded) : Buffer.from(padded, "base64").toString("binary");
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

function slugify(s: string): string {
  return (s || "untitled")
    .toLowerCase()
    .replace(/[^a-z0-9一-龥]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40) || "chat";
}