/**
 * ============================================================================
 * storage.test.ts
 * ----------------------------------------------------------------------------
 * Unit tests for the localStorage wrapper. Since vitest uses jsdom, we get
 * a real localStorage implementation that we can prime and inspect directly.
 * ============================================================================
 */

import { describe, it, expect, beforeEach } from "vitest";
import type { ChatSession, DeepSeekConfig } from "../types";
import {
  saveConfig,
  loadConfig,
  clearConfig,
  saveSessions,
  loadSessions,
  clearSessions,
  saveActiveSessionId,
  loadActiveSessionId,
  saveTheme,
  loadTheme,
  exportSessionsAsJson,
  importSessionsFromJson,
  type Theme,
} from "./storage";

/* ------------------------------------------------------------------ *
 *  Fixtures                                                          *
 * ------------------------------------------------------------------ */

const defaultConfig: DeepSeekConfig = {
  apiKey: "",
  baseUrl: "https://api.deepseek.com",
  model: "deepseek-chat",
  temperature: 0.7,
  maxTokens: 2048,
  topP: 0.9,
  systemPrompt: expect.any(String),
};

const sampleSession: ChatSession = {
  id: "sess-001",
  title: "Test Session",
  createdAt: 1000,
  updatedAt: 2000,
  messages: [
    { id: "msg-1", role: "user", content: "Hello", timestamp: 1500 },
    { id: "msg-2", role: "assistant", content: "Hi!", timestamp: 1600 },
  ],
};

const sampleSessions: ChatSession[] = [
  sampleSession,
  {
    id: "sess-002",
    title: "Second Session",
    createdAt: 3000,
    updatedAt: 4000,
    messages: [{ id: "msg-3", role: "user", content: "Analyze this", timestamp: 3500 }],
  },
];

/* ------------------------------------------------------------------ *
 *  Config tests                                                      *
 * ------------------------------------------------------------------ */

describe("storage - config", () => {
  it("returns default config when nothing is stored", () => {
    const config = loadConfig();
    expect(config).toMatchObject(defaultConfig);
    expect(config.systemPrompt).toBeTruthy();
  });

  it("persists and retrieves a custom config", () => {
    const custom: DeepSeekConfig = {
      apiKey: "sk-test-key-12345",
      baseUrl: "https://proxy.example.com",
      model: "deepseek-reasoner",
      temperature: 0.3,
      maxTokens: 4096,
      topP: 0.5,
      systemPrompt: "Be concise",
    };
    expect(saveConfig(custom)).toBe(true);
    const loaded = loadConfig();
    expect(loaded).toEqual(custom);
  });

  it("merges partial config with defaults", () => {
    // Simulate an older version of the code that stored only a subset
    localStorage.setItem("deepseek_excel_config_v1", JSON.stringify({ apiKey: "sk-old" }));
    const loaded = loadConfig();
    expect(loaded.apiKey).toBe("sk-old");
    expect(loaded.baseUrl).toBe("https://api.deepseek.com"); // from default
    expect(loaded.temperature).toBe(0.7);
  });

  it("returns default config after clearing", () => {
    saveConfig({ ...defaultConfig, apiKey: "sk-tmp" });
    clearConfig();
    const loaded = loadConfig();
    expect(loaded.apiKey).toBe("");
  });

  it("handles corrupted JSON gracefully", () => {
    localStorage.setItem("deepseek_excel_config_v1", "not-json");
    const loaded = loadConfig();
    expect(loaded).toMatchObject(defaultConfig);
  });
});

/* ------------------------------------------------------------------ *
 *  Sessions tests                                                    *
 * ------------------------------------------------------------------ */

describe("storage - sessions", () => {
  it("returns empty array when no sessions are stored", () => {
    expect(loadSessions()).toEqual([]);
  });

  it("persists and retrieves sessions", () => {
    expect(saveSessions(sampleSessions)).toBe(true);
    const loaded = loadSessions();
    expect(loaded).toHaveLength(2);
    expect(loaded[0].id).toBe("sess-001");
    expect(loaded[1].id).toBe("sess-002");
  });

  it("trims sessions beyond MAX_SESSIONS", () => {
    // MAX_SESSIONS is 50; save 55
    const many = Array.from({ length: 55 }, (_, i) => ({
      ...sampleSession,
      id: `sess-${i}`,
    }));
    saveSessions(many);
    const loaded = loadSessions();
    expect(loaded).toHaveLength(50);
  });

  it("filters out corrupt session entries", () => {
    localStorage.setItem(
      "deepseek_excel_sessions_v1",
      JSON.stringify([
        { id: "valid", title: "OK", createdAt: 0, updatedAt: 0, messages: [] },
        { id: null, messages: [] },           // bad id
        { id: "no-messages", title: "Nope" }, // missing messages[]
        "just-a-string",                       // not even an object
      ])
    );
    const loaded = loadSessions();
    expect(loaded).toHaveLength(1);
    expect(loaded[0].id).toBe("valid");
  });

  it("clears sessions and active pointer", () => {
    saveSessions(sampleSessions);
    saveActiveSessionId("sess-001");
    clearSessions();
    expect(loadSessions()).toEqual([]);
    expect(loadActiveSessionId()).toBeNull();
  });
});

/* ------------------------------------------------------------------ *
 *  Active session pointer tests                                      *
 * ------------------------------------------------------------------ */

describe("storage - active session id", () => {
  it("returns null when not set", () => {
    expect(loadActiveSessionId()).toBeNull();
  });

  it("persists and retrieves a session id", () => {
    saveActiveSessionId("sess-001");
    expect(loadActiveSessionId()).toBe("sess-001");
  });

  it("removes the key when saving null", () => {
    saveActiveSessionId("sess-001");
    saveActiveSessionId(null);
    expect(loadActiveSessionId()).toBeNull();
  });
});

/* ------------------------------------------------------------------ *
 *  Theme tests                                                        *
 * ------------------------------------------------------------------ */

describe("storage - theme", () => {
  it("returns 'auto' by default", () => {
    expect(loadTheme()).toBe("auto");
  });

  it("persists and retrieves light theme", () => {
    saveTheme("light");
    expect(loadTheme()).toBe("light");
  });

  it("persists and retrieves dark theme", () => {
    saveTheme("dark");
    expect(loadTheme()).toBe("dark");
  });

  it("falls back to auto for invalid values", () => {
    localStorage.setItem("deepseek_excel_theme_v1", JSON.stringify("neon"));
    expect(loadTheme()).toBe("auto");
  });
});

/* ------------------------------------------------------------------ *
 *  Export / import tests                                              *
 * ------------------------------------------------------------------ */

describe("storage - export / import", () => {
  it("exports sessions as formatted JSON", () => {
    const json = exportSessionsAsJson(sampleSessions);
    const parsed = JSON.parse(json);
    expect(parsed.app).toBe("deepseek-excel-assistant");
    expect(parsed.version).toBe(1);
    expect(parsed.sessions).toHaveLength(2);
  });

  it("imports valid JSON export", () => {
    const json = exportSessionsAsJson(sampleSessions);
    const imported = importSessionsFromJson(json);
    expect(imported).toHaveLength(2);
    expect(imported![0].id).toBe("sess-001");
  });

  it("returns null for invalid JSON", () => {
    expect(importSessionsFromJson("not-json")).toBeNull();
  });

  it("returns null for non-object payload", () => {
    expect(importSessionsFromJson('"hello"')).toBeNull();
  });

  it("returns null for missing sessions array", () => {
    expect(importSessionsFromJson(JSON.stringify({ app: "test" }))).toBeNull();
  });

  it("filters corrupt entries during import", () => {
    const json = JSON.stringify({
      sessions: [
        { id: "good", messages: [] },
        { messages: [] },       // missing id → filtered
        "bad",                   // not an object → filtered
      ],
    });
    const imported = importSessionsFromJson(json);
    expect(imported).toHaveLength(1);
  });
});
