/**
 * ============================================================================
 * DeepSeek Excel Assistant - Taskpane Entry Point
 * ----------------------------------------------------------------------------
 * Responsibilities:
 *   1. Wait for Office.js to finish initializing
 *   2. Hide the sideload splash, reveal the app shell
 *   3. Wire up persisted state (config + sessions) with the UI components
 *   4. Bridge the ribbon command (analyzeSelection) into the chat window
 *
 * The file is intentionally small. The heavy lifting lives in:
 *   - components/ChatWindow.ts   UI + streaming orchestration
 *   - components/SettingsPanel  Settings overlay
 *   - components/HistoryPanel   Sidebar of past sessions
 *   - services/deepseek.ts      Network
 *   - services/storage.ts       localStorage
 *   - services/excel.ts         Office.js wrappers
 * ============================================================================
 */

import "./taskpane.css";

// highlight.js github-dark theme is loaded via a runtime <link> in
// start() because the Office webview can drop webpack-injected styles
// for sibling packages. Keeping it here as a no-op import for grep.

import { ChatWindow } from "./components/ChatWindow";
import { SettingsPanel } from "./components/SettingsPanel";
import { HistoryPanel, downloadSessionAsMarkdown } from "./components/HistoryPanel";
import { CommandPaletteView, type CommandPaletteSelectDetail } from "./components/CommandPalette";
import { ShortcutHelpView } from "./components/ShortcutHelp";
import { KnowledgeBaseView } from "./components/KnowledgeBase";
import { ShareDialogView, type ShareDialogDetail } from "./components/ShareDialog";
import { UsageDashboardView } from "./components/UsageDashboard";
import {
  readShareFromUrl,
  clearShareFromUrl,
  type ShareableSession,
} from "./services/share";

import {
  loadActiveSessionId,
  loadConfig,
  loadSessions,
  saveActiveSessionId,
  saveConfig,
  saveSessions,
} from "./services/storage";
import { excelValuesToMarkdown } from "./utils/helpers";

import {
  APP_TITLE,
  APP_VERSION,
  DEFAULT_SYSTEM_PROMPT,
  MAX_SESSIONS,
} from "./utils/constants";

import { describeApiError } from "./services/deepseek";
import { getSelectedData } from "./services/excel";
import type { ChatSession, DeepSeekConfig } from "./types";

/** Top-level controller object - holds shared state and brokers between
 *  the various UI components. */
class AppController {
  /** Persisted user configuration. */
  public config: DeepSeekConfig;

  /** Persisted list of chat sessions (sorted newest-first). */
  public sessions: ChatSession[];

  /** Active session id, or null when no session is active. */
  public activeSessionId: string | null;

  /** UI components. */
  public chat: ChatWindow | null = null;
  public settings: SettingsPanel | null = null;
  public history: HistoryPanel | null = null;
  /** Modal ⌘K command palette (PRD-03). */
  public palette: CommandPaletteView | null = null;
  /** Modal knowledge-base manager (PRD-10). */
  public knowledgeBase: KnowledgeBaseView | null = null;
  /** Modal share dialog (PRD-11). */
  public shareDialog: ShareDialogView | null = null;
  /** Modal usage dashboard (PRD-12). */
  public usageDashboard: UsageDashboardView | null = null;
  /** Modal shortcut-key help dialog (PRD-03). */
  public shortcutHelp: ShortcutHelpView | null = null;
  /** Internal flag so we only bind window.message once. */
  private _ribbonBound: boolean = false;

  constructor() {
    this.config = loadConfig();
    this.sessions = loadSessions();
    this.activeSessionId = loadActiveSessionId();

    // If the active session id no longer exists (storage was cleared),
    // drop it so the chat window starts on a clean slate.
    if (this.activeSessionId && !this.sessions.find((s) => s.id === this.activeSessionId)) {
      this.activeSessionId = null;
    }
  }

  /** Boot the UI components and connect them. */
  public start(): void {
    this.injectHljsTheme();
    this.applyInitialTheme();

    const root = document.getElementById("app-root");
    if (!root) {
      console.error("[DeepSeek] Missing #app-root");
      return;
    }

    // Chat window first - it owns the primary surface.
    this.chat = new ChatWindow(root, {
      getConfig: () => this.config,
      getSessions: () => this.sessions,
      setSessions: (sessions) => this.setSessions(sessions),
      upsertSession: (session) => this.upsertSession(session),
      getActiveSessionId: () => this.activeSessionId,
    });
    this.chat.onSettingsToggle(() => this.toggleSettings());
    this.chat.onHistoryToggle(() => this.toggleHistory());
    this.chat.onSessionCreated(() => {
      // After creating a new session make sure the history sidebar refreshes.
      this.refreshHistory();
    });

    // Settings overlay - hidden by default.
    this.settings = new SettingsPanel(root, this.config);
    this.settings.element.addEventListener("settings-change", (ev) => {
      const e = ev as CustomEvent;
      this.applyConfig(e.detail.config, e.detail.reset);
    });

    // History sidebar - hidden by default.
    this.history = new HistoryPanel(root, {
      select: (id) => this.selectSession(id),
      create: () => this.createNewSession(),
      rename: (id, title) => this.renameSession(id, title),
      delete: (ids) => this.deleteSessions(ids),
      export: (id) => this.exportSession(id),
    });
    this.refreshHistory();

    // Render the active session into the chat.
    const active = this.activeSessionId
      ? this.sessions.find((s) => s.id === this.activeSessionId) || null
      : null;
    this.chat.renderSession(active);
    this.shareDialog?.setSession(active);

    // Mount the ⌘K command palette and shortcut-key help dialog (PRD-03).
    // They are mounted at root level so they overlay the whole app.
    this.palette = new CommandPaletteView();
    this.palette.element.addEventListener("command-palette-select", (ev) => {
      const e = ev as CustomEvent<CommandPaletteSelectDetail>;
      this.runPaletteAction(e.detail.item);
    });
    document.body.appendChild(this.palette.element);

    this.shortcutHelp = new ShortcutHelpView();
    document.body.appendChild(this.shortcutHelp.element);

    // Knowledge-base manager (PRD-10). ⌘B toggles it.
    this.knowledgeBase = new KnowledgeBaseView();
    this.knowledgeBase.element.addEventListener("kb-toast", (ev) => {
      const e = ev as CustomEvent;
      const { msg, kind } = e.detail || {};
      this.chat.toast(msg, kind);
    });
    document.body.appendChild(this.knowledgeBase.element);

    // Share dialog (PRD-11). ⌘⇧S toggles it.
    this.shareDialog = new ShareDialogView();
    this.shareDialog.element.addEventListener("share-import", (ev) => {
      const e = ev as CustomEvent<ShareDialogDetail>;
      const imported = e.detail?.imported;
      if (!imported) return;
      this.importSharedSession(imported);
    });
    this.shareDialog.element.addEventListener("share-toast", (ev) => {
      const e = ev as CustomEvent;
      const { msg, kind } = e.detail || {};
      this.chat.toast(msg, kind);
    });
    document.body.appendChild(this.shareDialog.element);

    // Usage dashboard (PRD-12). ⌘+D toggles it.
    this.usageDashboard = new UsageDashboardView();
    this.usageDashboard.element.addEventListener("usage-toast", (ev) => {
      const e = ev as CustomEvent;
      const { msg, kind } = e.detail || {};
      this.chat.toast(msg, kind);
    });
    document.body.appendChild(this.usageDashboard.element);

    // Auto-restore if the URL has a share fragment.
    this.maybeRestoreFromShareUrl();

    // Listen for ribbon commands.
    this.setupRibbonBridge();

    // Global keyboard shortcuts (PRD-03). Use capture phase so we beat
    // any element-local handlers in the chat input / settings panel.
    this.installGlobalShortcuts();
  }

  /** Run a ⌘K-selected action by routing it through the same dispatcher
   *  the ribbon uses. Keeps everything single-source-of-truth. */
  private runPaletteAction(item: { action: string; requiresSelection?: boolean }): void {
    const action = item.action;
    // Map palette-only actions that don't have a ribbon command.
    if (action === "toggleHistory") {
      this.toggleHistory();
      return;
    }
    if (action === "newSession") {
      this.createNewSession();
      return;
    }
    if (action === "showShortcuts") {
      this.shortcutHelp?.show();
      return;
    }
    if (action === "toggleKnowledgeBase") {
      this.knowledgeBase?.toggle();
      return;
    }
    // Otherwise synthesize a command envelope and run it through
    // the same `handle(...)` path so banner / ack / ack-all still fire.
    const env = {
      id: "palette-" + Date.now().toString(36),
      type: "deepseek:command",
      command: action,
      payload: null,
      t: Date.now(),
      via: "palette",
    };
    // Reuse the dispatcher we already set up in setupRibbonBridge.
    (window as any).__deepseekApp?.runCommand?.(env);
  }

  /** Wire ⌘K, ⌘+/, ⌘+1..9, ⌘+N, ⌘+H, ⌘+Enter and Esc. Done at the
   *  document level in capture phase so the chat textarea does not
   *  swallow them. */
  private installGlobalShortcuts(): void {
    const isMod = (ev: KeyboardEvent): boolean => ev.metaKey || ev.ctrlKey;

    document.addEventListener(
      "keydown",
      (ev) => {
        const target = ev.target as HTMLElement | null;
        const inInput = !!target && (
          target.tagName === "TEXTAREA" ||
          target.tagName === "INPUT" ||
          (target as any).isContentEditable
        );

        // Esc always works - closes palette / help / dialogs.
        if (ev.key === "Escape") {
          if (this.palette?.visible) { this.palette.hide(); ev.preventDefault(); return; }
          if (this.shortcutHelp && (this.shortcutHelp as any).root?.style?.display !== "none") {
            this.shortcutHelp.hide(); ev.preventDefault(); return;
          }
        }

        // While the palette is open, don't double-fire.
        if (this.palette?.visible) return;

        // ⌘+K - toggle palette.
        if (isMod(ev) && (ev.key === "k" || ev.key === "K")) {
          ev.preventDefault();
          this.palette?.show();
          return;
        }

        // ⌘+/ - shortcuts help.
        if (isMod(ev) && ev.key === "/") {
          ev.preventDefault();
          this.shortcutHelp?.show();
          return;
        }

        // ⌘+Enter inside input -> send.
        if (isMod(ev) && ev.key === "Enter" && inInput) {
          // ChatWindow already handles Enter. We only intervene on the
          // modded variant. Let ChatWindow send if it wants; we just
          // prevent default shift-handling.
          // (the textarea's own listener sends on Enter; ⌘+Enter is
          // treated the same as Enter by our existing handler.)
          return;
        }

        // ⌘+N - new session, ⌘+H - history, ⌘+B - knowledge base, ⌘+1..9 - ribbon commands.
        if (isMod(ev) && !ev.shiftKey && !ev.altKey) {
          if (ev.key === "n" || ev.key === "N") {
            ev.preventDefault();
            this.createNewSession();
            return;
          }
          if (ev.key === "h" || ev.key === "H") {
            ev.preventDefault();
            this.toggleHistory();
            return;
          }
          if (ev.key === "b" || ev.key === "B") {
            ev.preventDefault();
            this.knowledgeBase?.toggle();
            return;
          }
          if ((ev.key === "s" || ev.key === "S") && ev.shiftKey) {
            // ⌘+⇧+S - share dialog (without shift, ⌘+S is reserved for
            // browser save / Excel-internal shortcuts).
            ev.preventDefault();
            this.toggleShareDialog();
            return;
          }
          if (ev.key === "d" || ev.key === "D") {
            ev.preventDefault();
            this.usageDashboard?.toggle();
            return;
          }
          // ⌘+1..9 maps to ribbon buttons 2..9 in the spec, plus diagnose=0.
          // Spec table: ⌘+1 chat focus, ⌘+2 analyze, ..., ⌘+9 settings.
          const map: Record<string, string> = {
            "1": "showTaskpane",
            "2": "analyzeSelection",
            "3": "generateFormula",
            "4": "cleanData",
            "5": "insertLastReply",
            "6": "exportCurrentSession",
            "7": "clearCurrentChat",
            "8": "toggleTheme",
            "9": "openSettings",
            "0": "diagnoseFormulas",
          };
          const cmd = map[ev.key];
          if (cmd) {
            ev.preventDefault();
            if (cmd === "showTaskpane") {
              // ⌘+1 - focus the chat input for typing.
              try { (this.chat as any)?.inputEl?.focus?.(); } catch { /* noop */ }
              return;
            }
            const env = {
              id: "kbd-" + Date.now().toString(36),
              type: "deepseek:command",
              command: cmd,
              payload: null,
              t: Date.now(),
              via: "shortcut",
            };
            (window as any).__deepseekApp?.runCommand?.(env);
            return;
          }
        }
      },
      true // capture
    );
  }

  /** Inject the highlight.js github-dark stylesheet at runtime. The Office
   *  webview sometimes drops webpack-injected CSS for sibling packages, so
   *  we attach a <link> manually after the document is ready. */
  private injectHljsTheme(): void {
    const id = "hljs-theme-github-dark";
    if (document.getElementById(id)) return;
    const link = document.createElement("link");
    link.id = id;
    link.rel = "stylesheet";
    // Resolve relative to the served taskpane.html (webpack dev server or dist/).
    link.href = "https://cdn.jsdelivr.net/npm/highlight.js@11.9.0/styles/github-dark.min.css";
    document.head.appendChild(link);
  }

  /** Apply the user's saved theme on startup. Defaults to "auto" which
   *  follows Office's body[data-theme] attribute (or light if unset). */
  private applyInitialTheme(): void {
    let theme: "light" | "dark" | "auto" = "auto";
    try {
      const stored = localStorage.getItem("deepseek_excel_theme_v1");
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed === "light" || parsed === "dark" || parsed === "auto") {
          theme = parsed;
        }
      }
    } catch {
      /* noop */
    }
    if (theme === "auto") {
      // Follow Office's theme signal if present.
      try {
        const bodyTheme = document.body.getAttribute("data-theme") || "";
        if (bodyTheme.includes("dark") || bodyTheme.includes("Black")) {
          document.documentElement.setAttribute("data-theme", "dark");
          return;
        }
      } catch {
        /* noop */
      }
      document.documentElement.setAttribute("data-theme", "light");
      return;
    }
    document.documentElement.setAttribute("data-theme", theme);
  }

  /** Wire up postMessage handlers so ribbon buttons can drive the UI.
   *  Belt-and-suspenders: every transport the commands iframe might use
   *  is registered, and a recent-id set dedupes them. */
  private setupRibbonBridge(): void {
    if (this._ribbonBound) return;
    this._ribbonBound = true;

    // Brief toast that confirms a command was received. Helps the user
    // see immediately that the button click reached the taskpane even
    // if the underlying action is async or silent.
    const ack = (label: string) => {
      try {
        this.chat?.toast?.(`▶ ${label}`, "info");
      } catch {
        /* noop */
      }
    };

    // A persistent status strip pinned at the top of the taskpane so the
    // user can SEE commands arrive. Updated on every received envelope.
    this.installCommandBanner();

    // Dedupe so multi-channel delivery only fires once per intent.
    const recentIds = new Set<string>();
    const seenRecently = (id: string): boolean => {
      if (!id) return false;
      if (recentIds.has(id)) return true;
      recentIds.add(id);
      // Keep the cache small.
      if (recentIds.size > 64) {
        const arr = Array.from(recentIds);
        recentIds.clear();
        for (const i of arr.slice(arr.length - 32)) recentIds.add(i);
      }
      return false;
    };

    // Core dispatcher - shared by every transport so the command table
    // lives in one place.
    const handle = (raw: unknown, via: string) => {
      let data: any = raw;
      if (typeof raw === "string") {
        try {
          data = JSON.parse(raw);
        } catch {
          return; // not our envelope
        }
      }
      if (!data || typeof data !== "object" || data.type !== "deepseek:command") return;
      if (seenRecently(data.id)) return;
      const cmd: string = data.command;
      try { console.info(`[DeepSeek] ← ${cmd} via ${via}`); } catch {}
      this.updateCommandBanner(cmd, via);
      try {
        switch (cmd) {
          case "analyzeSelection":
            ack("分析选区");
            this.runAnalyzeSelection();
            recordFeature("analyzeSelection");
            break;
          case "generateFormula":
            ack("生成公式");
            this.chat?.runGenerateFormula();
            recordFeature("generateFormula");
            break;
          case "cleanData":
            ack("数据清洗");
            this.chat?.runCleanData();
            recordFeature("cleanData");
            break;
          case "insertLastReply":
            ack("插入回复");
            this.chat?.insertLastReplyPublic();
            recordFeature("insertLastReply");
            break;
          case "exportCurrentSession":
            ack("导出对话");
            this.chat?.exportCurrentSessionPublic();
            recordFeature("exportCurrentSession");
            break;
          case "clearCurrentChat":
            ack("清空对话");
            this.chat?.clearCurrentChatPublic();
            recordFeature("clearCurrentChat");
            break;
          case "toggleTheme":
            ack("切换主题");
            this.chat?.toggleTheme();
            recordFeature("toggleTheme");
            break;
          case "openSettings":
            ack("打开设置");
            this.chat?.openSettingsPublic();
            recordFeature("openSettings");
            break;
          case "diagnoseFormulas":
            ack("诊断公式");
            this.chat?.runDiagnoseFormulas();
            recordFeature("diagnoseFormulas");
            break;
          case "translateToCode":
            ack("公式转 VBA");
            this.chat?.runFormulaToCode("vba");
            recordFeature("translateToCode");
            break;
          case "insertChart":
            ack("插入图表");
            this.chat?.runInsertChart();
            recordFeature("insertChart");
            break;
          case "maskPII":
            ack("数据脱敏");
            this.chat?.runMaskPII();
            recordFeature("maskPII");
            break;
          case "multiSelectionAnalyze":
            ack("多选区分析");
            this.chat?.runMultiSelectionAnalyze();
            recordFeature("multiSelectionAnalyze");
            break;
          case "openKnowledgeBase":
            ack("知识库");
            this.knowledgeBase?.toggle();
            break;
          case "shareSession":
            ack("分享会话");
            this.toggleShareDialog();
            break;
          case "usageDashboard":
            ack("用量看板");
            this.usageDashboard?.toggle();
            break;
          default:
            console.warn("[DeepSeek] Unknown command:", cmd);
        }
      } catch (err) {
        console.error("[DeepSeek] Command handler error:", err);
      }
    };

    // Transport 1: window.postMessage. Generic, catches both cross-frame
    // targets and BroadcastChannel fan-out in some hosts.
    window.addEventListener("message", (ev) => handle(ev.data, "postMessage"));

    // Transport 2: Office.context.ui messageParent - the official cross
    // frame channel. Office routes ExecuteFunction output here for the
    // taskpane that owns the FunctionFile activation.
    try {
      const ui = (Office as any)?.context?.ui;
      if (ui && typeof ui.addHandlerAsync === "function") {
        ui.addHandlerAsync(
          Office.EventType.DialogParentMessageReceived,
          (arg: { message: string | unknown }) => handle(arg && arg.message, "messageParent")
        );
      }
    } catch (err) {
      console.warn("[DeepSeek] addHandlerAsync unavailable", err);
    }

    // Transport 3: BroadcastChannel. Same-origin iframes on modern
    // WebKit/Chromium can broadcast without involving Office at all.
    try {
      if (typeof BroadcastChannel === "function") {
        const ch = new BroadcastChannel("deepseek-cmds-v1");
        ch.addEventListener("message", (ev) => handle(ev.data, "broadcast"));
      }
    } catch (err) {
      console.warn("[DeepSeek] BroadcastChannel unavailable", err);
    }

    // Transport 4: localStorage 'storage' event. Fires in sibling
    // iframes of the same origin whenever another document writes to
    // localStorage. This is the most cross-host friendly transport -
    // it does not depend on Office, postMessage, or BroadcastChannel.
    window.addEventListener("storage", (ev) => {
      if (ev.key !== "deepseek:command" || !ev.newValue) return;
      try {
        handle(JSON.parse(ev.newValue), "storage");
      } catch {
        /* noop */
      }
    });

    // Transport 5: poll localStorage every 600ms. Some hosts (notably
    // Excel on Mac) do NOT fire the 'storage' event between sibling
    // iframes even though both share the same origin. Polling is
    // ugly but it always sees writes that happened, so we use it as
    // the last-resort transport.
    try {
      let lastValue = "";
      window.setInterval(() => {
        try {
          const raw = localStorage.getItem("deepseek:command");
          if (!raw || raw === lastValue) return;
          lastValue = raw;
          const env = JSON.parse(raw);
          // Clear immediately so we don't fire twice on the next tick.
          try { localStorage.removeItem("deepseek:command"); } catch { /* noop */ }
          handle(env, "poll");
        } catch {
          /* noop */
        }
      }, 600);
    } catch {
      /* noop */
    }

    // Transport 4 (inbound): expose a runCommand hook so a same-window
    // command handler can call directly. Used by some hosts when
    // SharedRuntime is active and the two scripts share a global object.
    (window as any).__deepseekApp = {
      runCommand: (env: unknown) => handle(env, "direct"),
      controller: this,
    };

    // Allow dev-mode "ping from the taskpane" - returns a short summary
    // of registered transports so the user can verify bridge wiring
    // even without clicking a ribbon button.
    (window as any).__deepseekPing = () => ({
      hasMessageListener: true,
      hasBroadcastChannel: typeof BroadcastChannel === "function",
      hasMessageParentHandler: true,
      appReady: !!this.chat,
    });
  }

  /** Build a persistent status strip so the user can see ribbon commands
   *  arrive. Updated every time the dispatcher fires. */
  private installCommandBanner(): void {
    if (document.getElementById("deepseek-command-banner")) return;
    const banner = document.createElement("div");
    banner.id = "deepseek-command-banner";
    banner.style.cssText = [
      "position:sticky",
      "top:0",
      "z-index:9999",
      "padding:6px 10px",
      "font:600 11px/1.3 system-ui,-apple-system,Segoe UI,sans-serif",
      "background:#0d6efd",
      "color:#fff",
      "border-radius:6px",
      "margin:6px 8px",
      "box-shadow:0 2px 6px rgba(0,0,0,.15)",
      "transition:background .25s ease",
      "letter-spacing:.3px",
    ].join(";");
    banner.textContent = "🔌 命令桥待命中 — 点击功能区按钮即可触发（点此展开调试）";
    banner.style.cursor = "pointer";
    banner.addEventListener("click", () => panel.classList.toggle("open"));
    document.body.insertBefore(banner, document.body.firstChild);

    // ---- Debug panel ------------------------------------------------------
    const panel = document.createElement("div");
    panel.id = "deepseek-debug-panel";
    panel.style.cssText = [
      "position:sticky",
      "top:38px",
      "z-index:9998",
      "margin:0 8px 8px",
      "padding:8px",
      "background:#1f2937",
      "color:#e5e7eb",
      "border-radius:8px",
      "box-shadow:0 4px 14px rgba(0,0,0,.2)",
      "display:none",
      "font:12px/1.4 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace",
    ].join(";");
    const cmds: Array<[string, string]> = [
      ["分析选区", "analyzeSelection"],
      ["生成公式", "generateFormula"],
      ["数据清洗", "cleanData"],
      ["插入回复", "insertLastReply"],
      ["导出对话", "exportCurrentSession"],
      ["清空对话", "clearCurrentChat"],
      ["切换主题", "toggleTheme"],
      ["打开设置", "openSettings"],
    ];
    const btnRow = document.createElement("div");
    btnRow.style.cssText = "display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin-bottom:6px;";
    for (const [label, cmd] of cmds) {
      const b = document.createElement("button");
      b.textContent = label;
      b.style.cssText = "padding:6px 4px;border:0;border-radius:4px;background:#374151;color:#fff;cursor:pointer;font-size:11px;";
      b.addEventListener("click", () => {
        // Simulate the commands iframe by going through the same handle()
        // entry the ribbon button would hit.
        const app = (window as any).__deepseekApp;
        if (app?.runCommand) {
          app.runCommand({
            id: "manual-" + Date.now() + "-" + Math.random().toString(36).slice(2, 6),
            type: "deepseek:command",
            command: cmd,
            payload: null,
            t: Date.now(),
            via: "manual",
          });
        }
      });
      btnRow.appendChild(b);
    }
    panel.appendChild(btnRow);

    const log = document.createElement("div");
    log.id = "deepseek-debug-log";
    log.style.cssText = "max-height:120px;overflow:auto;background:#111827;padding:4px 6px;border-radius:4px;font-size:11px;line-height:1.45;";
    panel.appendChild(log);
    document.body.insertBefore(panel, banner.nextSibling);

    // Toggle via class so we can show/hide cleanly.
    const styleEl = document.createElement("style");
    styleEl.textContent = "#deepseek-debug-panel.open{display:block !important;}";
    document.head.appendChild(styleEl);

    // Promote the status text. Called from the dispatcher.
    (this as any).__updateBanner = (cmd: string, via: string) => {
      const stamp = new Date().toLocaleTimeString("zh-CN", { hour12: false });
      banner.style.background = "#198754"; // green flash on success
      banner.textContent = `✅ ${stamp}  ${cmd}  ←  ${via}`;
      const line = document.createElement("div");
      line.textContent = `${stamp}  ${cmd}  ←  ${via}`;
      log.prepend(line);
      while (log.childElementCount > 30) log.removeChild(log.lastChild!);
      // Fade back to blue after a short while.
      window.setTimeout(() => { banner.style.background = "#0d6efd"; }, 1200);
    };
  }

  /** Update the banner with a fresh command event. */
  private updateCommandBanner(cmd: string, via: string): void {
    const fn = (this as any).__updateBanner as ((c: string, v: string) => void) | undefined;
    if (fn) fn(cmd, via);
  }

  /** Apply a new configuration - persist and re-render. */
  public applyConfig(config: DeepSeekConfig, reset: boolean): void {
    this.config = { ...config };
    saveConfig(this.config);

    // System prompt changes apply to all future requests. We don't need
    // to retroactively rewrite historical messages.

    if (this.settings) this.settings.setConfig(this.config);
    // Update the model badge in the context bar.
    this.chat?.refreshModelBadge();

    if (reset) {
      this.chat?.toast("已恢复默认设置", "success");
    } else {
      this.chat?.toast("设置已保存", "success");
    }
  }

  /** Replace the entire sessions array (after a bulk delete). */
  public setSessions(sessions: ChatSession[]): void {
    this.sessions = sessions.slice(0, MAX_SESSIONS);
    saveSessions(this.sessions);
    this.refreshHistory();
  }

  /** Insert-or-update one session in place. */
  public upsertSession(session: ChatSession): void {
    const idx = this.sessions.findIndex((s) => s.id === session.id);
    if (idx >= 0) {
      this.sessions[idx] = session;
    } else {
      this.sessions.unshift(session);
    }
    // Enforce the cap - drop oldest when over.
    this.sessions = this.sessions.slice(0, MAX_SESSIONS);
    saveSessions(this.sessions);
    this.activeSessionId = session.id;
    saveActiveSessionId(session.id);
    this.refreshHistory();
    this.chat?.refreshTokenBadge();
  }

  /** Switch the active session. */
  public selectSession(id: string): void {
    this.activeSessionId = id;
    saveActiveSessionId(id);
    const session = this.sessions.find((s) => s.id === id) || null;
    this.chat?.renderSession(session);
    this.shareDialog?.setSession(session);
    this.refreshHistory();
    // Close the sidebar on selection so the chat is visible.
    this.history?.close();
  }

  /** If the current URL has a #share=... fragment, decode it and ask
   *  the user if they want to import it as a new session. */
  private maybeRestoreFromShareUrl(): void {
    const shared = readShareFromUrl();
    if (!shared) return;
    const ok = confirm(
      `检测到一个分享的会话「${shared.title}」(包含 ${shared.messages.length} 条消息)。\n\n是否恢复到当前加载项？`
    );
    if (ok) {
      this.importSharedSession(shared);
    }
    clearShareFromUrl();
  }

  /** Add a new session populated from a ShareableSession payload. */
  private importSharedSession(payload: ShareableSession): void {
    const id = `s-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const session: ChatSession = {
      id,
      title: payload.title ? `${payload.title} (分享)` : "分享的对话",
      createdAt: payload.createdAt || Date.now(),
      updatedAt: Date.now(),
      messages: payload.messages.map((m) => ({
        ...m,
        id: m.id || `${id}-m-${Math.random().toString(36).slice(2, 8)}`,
      })),
    };
    this.sessions.unshift(session);
    saveSessions(this.sessions);
    this.selectSession(id);
    this.chat?.toast(`已恢复分享会话「${session.title}」`, "success");
  }

  /** Toggle the share dialog. */
  public toggleShareDialog(): void {
    const id = this.activeSessionId;
    const session = id ? this.sessions.find((s) => s.id === id) || null : null;
    this.shareDialog?.setSession(session);
    if (this.shareDialog?.element.hidden) {
      this.shareDialog?.show();
    } else {
      this.shareDialog?.hide();
    }
  }

  /** Create a new empty session and select it. */
  public createNewSession(): void {
    const session: ChatSession = {
      id: `s-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      title: "新对话",
      createdAt: Date.now(),
      updatedAt: Date.now(),
      messages: [],
    };
    this.sessions.unshift(session);
    this.sessions = this.sessions.slice(0, MAX_SESSIONS);
    saveSessions(this.sessions);
    this.activeSessionId = session.id;
    saveActiveSessionId(session.id);
    this.refreshHistory();
    this.chat?.clearView();
    this.history?.close();
    this.chat?.toast("已创建新对话", "info");
  }

  /** Rename one session. */
  public renameSession(id: string, newTitle: string): void {
    const s = this.sessions.find((x) => x.id === id);
    if (!s) return;
    s.title = newTitle;
    s.updatedAt = Date.now();
    saveSessions(this.sessions);
    this.refreshHistory();
  }

  /** Delete one or more sessions and clean up the active pointer. */
  public deleteSessions(ids: string[]): void {
    const set = new Set(ids);
    this.sessions = this.sessions.filter((s) => !set.has(s.id));
    saveSessions(this.sessions);

    if (this.activeSessionId && set.has(this.activeSessionId)) {
      this.activeSessionId = this.sessions[0]?.id || null;
      saveActiveSessionId(this.activeSessionId);
      const next = this.activeSessionId
        ? this.sessions.find((s) => s.id === this.activeSessionId) || null
        : null;
      this.chat?.renderSession(next);
      this.shareDialog?.setSession(next);
    }
    this.refreshHistory();
    this.chat?.toast(`已删除 ${ids.length} 个会话`, "success");
  }

  /** Trigger a Markdown download for one session. */
  public exportSession(id: string): void {
    const s = this.sessions.find((x) => x.id === id);
    if (!s) return;
    try {
      downloadSessionAsMarkdown(s);
      this.chat?.toast("已导出 Markdown", "success");
    } catch (err: any) {
      this.chat?.toast("导出失败：" + (err?.message || ""), "error");
    }
  }

  /* ---------------- UI plumbing ---------------- */

  public toggleSettings(): void {
    if (!this.settings) return;
    if (this.history?.isOpen()) this.history.close();
    this.settings.toggle();
  }

  public toggleHistory(): void {
    if (!this.history) return;
    if (this.settings?.isVisible()) this.settings.hide();
    this.history.toggle();
  }

  public refreshHistory(): void {
    this.history?.setSessions(this.sessions, this.activeSessionId);
  }

  /** Public entry point used by the ribbon "Analyze Selection" command. */
  public async runAnalyzeSelection(): Promise<void> {
    if (!this.chat) return;
    try {
      // Prefer the cached selection so we don't trigger an extra Excel round
      // trip on every ribbon click. Fall back to a fresh fetch.
      let selection = this.chat.getCachedSelection();
      if (!selection) {
        selection = await getSelectedData();
      }
      if (!selection) {
        this.chat.toast("请先在 Excel 中选中要分析的数据", "error");
        return;
      }
      const prompt =
        `请分析以下 Excel 选区并给出洞察：\n\n` +
        `工作表：${selection.sheetName}\n` +
        `选区地址：${selection.address}\n` +
        `数据规模：${selection.rowCount} 行 × ${selection.columnCount} 列\n\n` +
        "数据预览：\n" +
        `${excelValuesToMarkdown(selection.values).slice(0, 2000)}\n\n` +
        "请从趋势、异常、改进建议等维度给出结论。";
      await this.chat.sendUserMessage(prompt);
      this.chat.toast("已发送选区给 AI", "success");
    } catch (err: any) {
      const msg = describeApiError(err) || err?.message || "无法获取选区";
      this.chat.toast(msg, "error");
    }
  }
}

/** Singleton instance referenced from commands.ts. */
let app: AppController | null = null;

/** Office.js entry point - called by host when add-in is ready. */
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    // The manifest only declares Excel support, but defensive guard
    // keeps the code from crashing if sideloaded somewhere unexpected.
    console.warn(`[DeepSeek] Unsupported host: ${info.host}`);
  }

  // Hide the sideload splash, reveal the app shell.
  const splash = document.getElementById("sideload-msg");
  const shell = document.getElementById("app-body");
  if (splash) splash.style.display = "none";
  if (shell) (shell as HTMLElement).style.display = "flex";

  // Apply Office theme if available. Office injects data-theme on <html>.
  try {
    const theme = document.body.getAttribute("data-theme") || "default";
    if (theme.includes("dark")) {
      document.documentElement.setAttribute("data-theme", "dark");
    }
  } catch {
    /* noop */
  }

  // Show app version in the console for diagnostics.
  console.info(`[${APP_TITLE}] v${APP_VERSION} loaded. Default prompt length: ${DEFAULT_SYSTEM_PROMPT.length}`);

  // Boot the app.
  app = new AppController();
  app.start();

  // Expose for debugging from the webview devtools.
  (window as any).deepseekApp = app;
});

/** Exported so commands.ts can call into the chat window from the ribbon. */
export function analyzeSelection(): void {
  if (app) {
    app.runAnalyzeSelection();
  } else {
    console.warn("[DeepSeek] App not ready");
  }
}