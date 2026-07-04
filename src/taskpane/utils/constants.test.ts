import { describe, it, expect } from "vitest";
import {
  STORAGE_KEY_CONFIG,
  STORAGE_KEY_SESSIONS,
  STORAGE_KEY_ACTIVE_SESSION,
  STORAGE_KEY_THEME,
  MAX_SESSIONS,
  FORMULA_DIAGNOSIS_LIMIT,
  FORMULA_DIAGNOSIS_WARN,
  FORMULA_ERROR_INFO,
  REQUEST_TIMEOUT_MS,
  DEFAULT_SYSTEM_PROMPT,
  AVAILABLE_MODELS,
  createDefaultConfig,
  QUICK_ACTION_TEMPLATES,
  QUICK_TEMPLATES,
  APP_TITLE,
  APP_VERSION,
  SLASH_COMMANDS,
  COMMAND_PALETTE_ITEMS,
  SHORTCUTS,
  FORMULA_CARDS,
} from "./constants";

describe("storage keys", () => {
  it("STORAGE_KEY_CONFIG has correct value", () => {
    expect(STORAGE_KEY_CONFIG).toBe("deepseek_excel_config_v1");
  });

  it("STORAGE_KEY_SESSIONS has correct value", () => {
    expect(STORAGE_KEY_SESSIONS).toBe("deepseek_excel_sessions_v1");
  });

  it("STORAGE_KEY_ACTIVE_SESSION has correct value", () => {
    expect(STORAGE_KEY_ACTIVE_SESSION).toBe("deepseek_excel_active_session_v1");
  });

  it("STORAGE_KEY_THEME has correct value", () => {
    expect(STORAGE_KEY_THEME).toBe("deepseek_excel_theme_v1");
  });
});

describe("limits", () => {
  it("MAX_SESSIONS is 50", () => {
    expect(MAX_SESSIONS).toBe(50);
  });

  it("FORMULA_DIAGNOSIS_LIMIT is 500", () => {
    expect(FORMULA_DIAGNOSIS_LIMIT).toBe(500);
  });

  it("FORMULA_DIAGNOSIS_WARN is 200", () => {
    expect(FORMULA_DIAGNOSIS_WARN).toBe(200);
  });
});

describe("FORMULA_ERROR_INFO", () => {
  it("contains all expected error keys", () => {
    const keys = Object.keys(FORMULA_ERROR_INFO);
    expect(keys).toEqual([
      "#REF!",
      "#DIV/0!",
      "#N/A",
      "#VALUE!",
      "#NAME?",
      "#NUM!",
      "#NULL!",
      "#SPILL!",
      "#CALC!",
    ]);
  });

  it("every entry has code, label, reason", () => {
    for (const [key, entry] of Object.entries(FORMULA_ERROR_INFO)) {
      expect(entry.code).toBeTruthy();
      expect(entry.label).toBeTruthy();
      expect(entry.reason).toBeTruthy();
      // code matches the expected pattern
      expect(entry).toHaveProperty("code");
      expect(entry).toHaveProperty("label");
      expect(entry).toHaveProperty("reason");
    }
  });

  it('"#REF!" has the correct metadata', () => {
    const ref = FORMULA_ERROR_INFO["#REF!"];
    expect(ref.code).toBe("REF");
    expect(ref.label).toBe("引用无效");
    expect(ref.reason).toContain("引用断裂");
  });
});

describe("REQUEST_TIMEOUT_MS", () => {
  it("is 60 seconds", () => {
    expect(REQUEST_TIMEOUT_MS).toBe(60_000);
  });
});

describe("DEFAULT_SYSTEM_PROMPT", () => {
  it("is a non-empty string", () => {
    expect(DEFAULT_SYSTEM_PROMPT).toBeTruthy();
    expect(DEFAULT_SYSTEM_PROMPT.length).toBeGreaterThan(50);
  });

  it("mentions Excel assistance", () => {
    expect(DEFAULT_SYSTEM_PROMPT).toContain("Excel");
  });
});

describe("AVAILABLE_MODELS", () => {
  it("has chat and reasoner models", () => {
    const values = AVAILABLE_MODELS.map((m) => m.value);
    expect(values).toContain("deepseek-chat");
    expect(values).toContain("deepseek-reasoner");
  });

  it("every model has label and description", () => {
    for (const m of AVAILABLE_MODELS) {
      expect(m.label).toBeTruthy();
      expect(m.description).toBeTruthy();
    }
  });
});

describe("createDefaultConfig", () => {
  it("returns a fresh object each call", () => {
    const a = createDefaultConfig();
    const b = createDefaultConfig();
    expect(a).not.toBe(b);
  });

  it("has expected defaults", () => {
    const cfg = createDefaultConfig();
    expect(cfg.apiKey).toBe("");
    expect(cfg.baseUrl).toBe("https://api.deepseek.com");
    expect(cfg.model).toBe("deepseek-chat");
    expect(cfg.temperature).toBe(0.7);
    expect(cfg.maxTokens).toBe(2048);
    expect(cfg.topP).toBe(0.9);
    expect(cfg.systemPrompt).toBe(DEFAULT_SYSTEM_PROMPT);
  });

  it("values are within valid ranges", () => {
    const cfg = createDefaultConfig();
    expect(cfg.temperature).toBeGreaterThanOrEqual(0);
    expect(cfg.temperature).toBeLessThanOrEqual(2);
    expect(cfg.maxTokens).toBeGreaterThanOrEqual(256);
    expect(cfg.maxTokens).toBeLessThanOrEqual(8192);
    expect(cfg.topP).toBeGreaterThanOrEqual(0);
    expect(cfg.topP).toBeLessThanOrEqual(1);
  });
});

describe("QUICK_ACTION_TEMPLATES", () => {
  it("has analyze, formula, clean keys", () => {
    expect(Object.keys(QUICK_ACTION_TEMPLATES)).toEqual(["analyze", "formula", "clean"]);
  });

  it("each template contains {CONTEXT} placeholder", () => {
    for (const [key, tmpl] of Object.entries(QUICK_ACTION_TEMPLATES)) {
      expect(tmpl).toContain("{CONTEXT}");
    }
  });

  it("formula template also contains {USER_INPUT}", () => {
    expect(QUICK_ACTION_TEMPLATES.formula).toContain("{USER_INPUT}");
  });
});

describe("QUICK_TEMPLATES", () => {
  it("every template has all required fields", () => {
    for (const t of QUICK_TEMPLATES) {
      expect(t.id).toBeTruthy();
      expect(t.icon).toBeTruthy();
      expect(t.label).toBeTruthy();
      expect(t.prompt).toBeTruthy();
      expect(t.prompt.length).toBeGreaterThan(10);
      // requiresSelection is optional but should be boolean when present
      if (t.requiresSelection !== undefined) {
        expect(typeof t.requiresSelection).toBe("boolean");
      }
    }
  });

  it("icons do not contain emoji zwj sequences (unintended multi-emojis)", () => {
    for (const t of QUICK_TEMPLATES) {
      // Each icon should be a single emoji codepoint
      expect([...t.icon].length).toBeLessThanOrEqual(2);
    }
  });
});

describe("app metadata", () => {
  it("APP_TITLE is correct", () => {
    expect(APP_TITLE).toBe("DeepSeek Excel Assistant");
  });

  it("APP_VERSION is 1.0.0", () => {
    expect(APP_VERSION).toBe("1.0.0");
  });
});

describe("SLASH_COMMANDS", () => {
  it("has at least 8 commands", () => {
    expect(SLASH_COMMANDS.length).toBeGreaterThanOrEqual(8);
  });

  it("every command has all required fields", () => {
    const categories = new Set<string>();
    for (const cmd of SLASH_COMMANDS) {
      expect(cmd.id).toBeTruthy();
      expect(cmd.trigger).toMatch(/^\//);
      expect(cmd.icon).toBeTruthy();
      expect(cmd.label).toBeTruthy();
      expect(cmd.hint).toBeTruthy();
      expect(cmd.template).toBeTruthy();
      expect(cmd.template.length).toBeGreaterThan(10);
      expect(["analyze", "formula", "clean", "explain", "visualize", "summary"]).toContain(cmd.category);
      categories.add(cmd.category);
    }
    // at least 4 categories represented
    expect(categories.size).toBeGreaterThanOrEqual(4);
  });

  it("every command trigger starts with / and is lower-case", () => {
    for (const cmd of SLASH_COMMANDS) {
      expect(cmd.trigger).toMatch(/^\/[a-z]/);
    }
  });

  it("commands that requireSelection have {CONTEXT} or {FORMULA} in template", () => {
    for (const cmd of SLASH_COMMANDS) {
      if (cmd.requiresSelection) {
        const hasContext = cmd.template.includes("{CONTEXT}");
        const hasFormula = cmd.template.includes("{FORMULA}");
        expect(hasContext || hasFormula).toBe(true);
      }
    }
  });
});

describe("COMMAND_PALETTE_ITEMS", () => {
  it("has at least 12 items", () => {
    expect(COMMAND_PALETTE_ITEMS.length).toBeGreaterThanOrEqual(12);
  });

  it("every item has all required fields", () => {
    const categories = new Set<string>();
    for (const item of COMMAND_PALETTE_ITEMS) {
      expect(item.id).toBeTruthy();
      expect(item.icon).toBeTruthy();
      expect(item.label).toBeTruthy();
      expect(item.hint).toBeTruthy();
      expect(item.action).toBeTruthy();
      expect(["ribbon", "session", "view", "help"]).toContain(item.category);
      expect(item.shortcut).toBeTruthy();
      categories.add(item.category);
    }
    expect(categories.size).toBeGreaterThanOrEqual(3);
  });

  it("shortcuts follow the ⌘+<key> pattern", () => {
    for (const item of COMMAND_PALETTE_ITEMS) {
      expect(item.shortcut).toMatch(/^[⌘⌃⇧]\+/);
    }
  });
});

describe("SHORTCUTS", () => {
  it("has at least 10 shortcuts", () => {
    expect(SHORTCUTS.length).toBeGreaterThanOrEqual(10);
  });

  it("every shortcut has combo, label, and valid scope", () => {
    for (const s of SHORTCUTS) {
      expect(s.combo).toBeTruthy();
      expect(s.label).toBeTruthy();
      expect(["global", "input"]).toContain(s.scope);
    }
  });

  it("no duplicate combos across global scope", () => {
    const combos = SHORTCUTS.filter((s) => s.scope === "global").map((s) => s.combo);
    expect(new Set(combos).size).toBe(combos.length);
  });
});

describe("FORMULA_CARDS", () => {
  it("has at least 8 formula cards", () => {
    expect(FORMULA_CARDS.length).toBeGreaterThanOrEqual(8);
  });

  it("every card has id, name, desc, prompt", () => {
    for (const c of FORMULA_CARDS) {
      expect(c.id).toBeTruthy();
      expect(c.name).toBeTruthy();
      expect(c.desc).toBeTruthy();
      expect(c.prompt).toBeTruthy();
      expect(c.prompt.length).toBeGreaterThan(10);
    }
  });

  it("all ids are unique", () => {
    const ids = FORMULA_CARDS.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
