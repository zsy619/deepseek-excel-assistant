/**
 * ============================================================================
 * localStorage wrapper
 * ----------------------------------------------------------------------------
 * Single chokepoint for all read/write access to localStorage. Two reasons:
 *
 *   1. JSON parse/stringify with a typed envelope - never sprinkle
 *      try/catch around the codebase.
 *   2. Defensive trimming of stale sessions so we never exceed MAX_SESSIONS.
 *
 * If the browser refuses to write (quota / private mode) we swallow the
 * error and return false so the UI can fall back to in-memory state.
 * ============================================================================
 */

import { createDefaultConfig, MAX_SESSIONS, STORAGE_KEY_ACTIVE_SESSION, STORAGE_KEY_CONFIG, STORAGE_KEY_SESSIONS, STORAGE_KEY_THEME } from "../utils/constants";
import type { ChatSession, DeepSeekConfig } from "../types";

/** Try to read a JSON value from localStorage. Returns undefined when the
 *  key is missing, the value is malformed, or localStorage is unavailable. */
function readJson<T>(key: string): T | undefined {
  try {
    if (typeof localStorage === "undefined") return undefined;
    const raw = localStorage.getItem(key);
    if (!raw) return undefined;
    return JSON.parse(raw) as T;
  } catch {
    return undefined;
  }
}

/** Persist a JSON-serializable value. Returns true on success. */
function writeJson(key: string, value: unknown): boolean {
  try {
    if (typeof localStorage === "undefined") return false;
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

/** Remove a key. */
function remove(key: string): void {
  try {
    if (typeof localStorage === "undefined") return;
    localStorage.removeItem(key);
  } catch {
    /* noop */
  }
}

/* ---------------------------------------------------------------- *
 * Config                                                          *
 * ---------------------------------------------------------------- */

export function saveConfig(config: DeepSeekConfig): boolean {
  return writeJson(STORAGE_KEY_CONFIG, config);
}

export function loadConfig(): DeepSeekConfig {
  const stored = readJson<Partial<DeepSeekConfig>>(STORAGE_KEY_CONFIG);
  // Merge over defaults so newly added fields always have a value.
  return { ...createDefaultConfig(), ...(stored || {}) };
}

export function clearConfig(): void {
  remove(STORAGE_KEY_CONFIG);
}

/* ---------------------------------------------------------------- *
 * Sessions                                                        *
 * ---------------------------------------------------------------- */

export function saveSessions(sessions: ChatSession[]): boolean {
  // Trim to MAX_SESSIONS before persisting to enforce the cap.
  const trimmed = sessions.slice(0, MAX_SESSIONS);
  return writeJson(STORAGE_KEY_SESSIONS, trimmed);
}

export function loadSessions(): ChatSession[] {
  const stored = readJson<ChatSession[]>(STORAGE_KEY_SESSIONS);
  if (!Array.isArray(stored)) return [];
  // Defensive shape check - corrupt entries are dropped.
  return stored
    .filter((s) => s && typeof s.id === "string" && Array.isArray(s.messages))
    .slice(0, MAX_SESSIONS);
}

export function clearSessions(): void {
  remove(STORAGE_KEY_SESSIONS);
  remove(STORAGE_KEY_ACTIVE_SESSION);
}

/* ---------------------------------------------------------------- *
 * Active session pointer                                          *
 * ---------------------------------------------------------------- */

export function saveActiveSessionId(id: string | null): boolean {
  if (id === null) {
    remove(STORAGE_KEY_ACTIVE_SESSION);
    return true;
  }
  return writeJson(STORAGE_KEY_ACTIVE_SESSION, id);
}

export function loadActiveSessionId(): string | null {
  const v = readJson<string>(STORAGE_KEY_ACTIVE_SESSION);
  return typeof v === "string" && v ? v : null;
}

/* ---------------------------------------------------------------- *
 * Theme                                                            *
 * ---------------------------------------------------------------- */

export type Theme = "light" | "dark" | "auto";

export function saveTheme(theme: Theme): boolean {
  return writeJson(STORAGE_KEY_THEME, theme);
}

export function loadTheme(): Theme {
  const v = readJson<Theme>(STORAGE_KEY_THEME);
  if (v === "light" || v === "dark" || v === "auto") return v;
  return "auto";
}

/* ---------------------------------------------------------------- *
 * Word counter                                                     *
 * ---------------------------------------------------------------- */

/** Lightweight export-to-JSON utility - stores a complete backup
 *  of all sessions in a single file so users can move data between
 *  machines. */
export function exportSessionsAsJson(sessions: ChatSession[]): string {
  const payload = {
    app: "deepseek-excel-assistant",
    version: 1,
    exportedAt: new Date().toISOString(),
    sessions,
  };
  return JSON.stringify(payload, null, 2);
}

/** Restore sessions from a JSON export. Returns null on parse failure. */
export function importSessionsFromJson(text: string): ChatSession[] | null {
  try {
    const parsed = JSON.parse(text);
    if (!parsed || !Array.isArray(parsed.sessions)) return null;
    return parsed.sessions.filter(
      (s: any) => s && typeof s.id === "string" && Array.isArray(s.messages)
    );
  } catch {
    return null;
  }
}