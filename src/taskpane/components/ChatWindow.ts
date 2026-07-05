/**
 * ============================================================================
 * ChatWindow
 * ----------------------------------------------------------------------------
 * The central orchestrator. Owns:
 *   - The list of ChatMessageView instances (one per message bubble)
 *   - The current ChatSession and the lifecycle of streaming responses
 *   - The AbortController used to cancel an in-flight stream
 *
 * It does NOT own persistence or Excel I/O - those concerns are injected
 * via the `deps` constructor argument. This makes the component testable
 * and decouples it from the rest of the app.
 * ============================================================================
 */

import { ChatMessageView } from "./ChatMessage";
import {
  buildActionPrompt,
  mountQuickActions,
  type QuickActionKind,
} from "./QuickActions";
import { markdownToPlainText } from "../utils/markdownText";
import { chatCompletionStream, describeApiError, diagnoseFormulasStream, translateToScript, recommendChartStream, type FormulaDiagnosis, type ScriptFlavor } from "../services/deepseek";
import { getSelectedData, getRangeFormula, insertTextToCell, writeFormula, writeRange, detectFormula, scanFormulaErrors, fixFormulaAt, collectFormulasForCode, getSelectedRangeInfo, insertChart, scanSelectionForPII, batchReplaceCells, readMultiSelection, type FormulaScanResult } from "../services/excel";
import { excelEvents } from "../services/excel-events";
import { ContextBar } from "./ContextBar";
import { FormulaLibrary } from "./FormulaLibrary";
import { FormulaDiagnosticsView, type FormulaDiagnosticCardEventDetail } from "./FormulaDiagnostics";
import { confirmDialog, promptDialog } from "./Dialog";
import { PromptMenuView, type PromptMenuSelectDetail } from "./PromptMenu";
import { CodeGenPanelView, type CodeGenCopyDetail, type CodeGenRetryDetail } from "./CodeGenPanel";
import { ChartPickerView, type ChartPickerInsertDetail } from "./ChartPicker";
import { PiiMaskerView, type PiiMaskerApplyDetail } from "./PiiMasker";
// Type-only imports — the 5 panel classes live in their own modules and are
// loaded on-demand through phase4Panels. We never reference them at runtime
// from this file, so `import type` keeps them out of the main bundle while
// preserving strong typing for the Phase4Panels record below.
import type { CorrelationMatrixPanel } from "./CorrelationMatrixPanel";
import type { OutlierPanel } from "./OutlierPanel";
import type { PivotBuilderPanel } from "./PivotBuilderPanel";
import type { ReportBuilderPanel } from "./ReportBuilderPanel";
import type { ColumnTypePanel } from "./ColumnTypePanel";
import type { Phase4Panels } from "./phase4Panels";
import { RibbonBanner } from "./RibbonBanner";
import { RibbonFocusController, panelForCommand, type PanelName } from "./RibbonFocusController";
import { excelValuesToMarkdown, generateId, deriveSessionTitle, estimateTokens, formatTokens, escapeHtml } from "../utils/helpers";
import { BranchController } from "../controllers/BranchController";
import { ToolCallController, type ToolCallDispatcher } from "../controllers/ToolCallController";
import { KnowledgeInjector } from "../controllers/KnowledgeInjector";
import { UsageRecorder } from "../controllers/UsageRecorder";
import type { ChatControllerHub } from "../controllers/ChatControllerHub";
// PRD-10: KB retrieval is owned by KnowledgeInjector — see controllers/.
import { recordFeature } from "../services/usage";
import { FORMULA_DIAGNOSIS_LIMIT, FORMULA_DIAGNOSIS_WARN, QUICK_TEMPLATES, type SlashCommand } from "../utils/constants";
import type {
  ChatMessage,
  ChatSession,
  DeepSeekConfig,
  ExcelChartType,
  ExcelSelection,
} from "../types";

export interface ChatWindowDeps {
  /** Current configuration. */
  getConfig: () => DeepSeekConfig;
  /** All persisted sessions. */
  getSessions: () => ChatSession[];
  /** Replace all persisted sessions. */
  setSessions: (sessions: ChatSession[]) => void;
  /** Save just one session back into the active list. */
  upsertSession: (session: ChatSession) => void;
  /** Resolve the active session id. */
  getActiveSessionId: () => string | null;
}

export class ChatWindow implements ChatControllerHub {
  private root: HTMLElement;
  private deps: ChatWindowDeps;

  /** Container for ChatMessageView elements. */
  private listEl!: HTMLElement;
  /** Input area + send/stop button. */
  private inputEl!: HTMLTextAreaElement;
  private sendBtn!: HTMLButtonElement;
  private quickEl!: HTMLElement;
  private templatesEl!: HTMLElement;
  private historyBtn!: HTMLButtonElement;
  private settingsBtn!: HTMLButtonElement;
  private newChatBtn!: HTMLButtonElement;
  private statusBar!: HTMLElement;
  private wordCountEl!: HTMLElement;

  /** Slash-command menu mounted below the input (PRD-02). */
  private promptMenu!: PromptMenuView;
  /** Last query typed after the leading "/" - used to filter the menu. */
  private slashQuery = "";

  /** Context bar shown under the header - displays Excel selection state. */
  private contextBar!: ContextBar;

  /** Latest known selection - kept for quick access by templates/actions. */
  private currentSelection: ExcelSelection | null = null;

  /** Latest known sheet name. */
  private currentSheetName: string | null = null;

  /** Excel event unsubscribers. */
  private excelUnsubscribers: Array<() => void> = [];

  /** Map of message id -> view, for fast streaming updates. */
  private views: Map<string, ChatMessageView> = new Map();

  /** Formula library panel - shows quick-insert cards. */
  private formulaLibrary!: FormulaLibrary;

  /** Code-gen panel (PRD-04) - converts formulas in the selection to
   *  VBA / Office Scripts. */
  private codeGenPanel!: CodeGenPanelView;
  /** In-flight translateToScript abort controller. */
  private codeGenAbort: AbortController | null = null;

  /** Top-of-taskpane status banner shown while a ribbon command is in flight. */
  public ribbonBanner: RibbonBanner = new RibbonBanner();

  /** Scroll/highlight controller for ribbon commands. */
  public ribbonFocus: RibbonFocusController = new RibbonFocusController();

  /** Abort for the chart recommender stream. */
  private chartAbort: AbortController | null = null;

  /** Chart picker panel (PRD-05) - renders 3 AI-recommended chart cards. */
  private chartPicker!: ChartPickerView;
  /** Source range address captured at runInsertChart time. */
  private chartSourceAddress = "";

  /** PII masker panel (PRD-06) - shows detected sensitive values. */
  private piiMasker!: PiiMaskerView;

  /** Lazy-loaded bundle of 5 advanced-analysis panels (Phase 4). Fetched on
   *  first Phase-4 ribbon click; webpack splits this into its own async
   *  chunk so the 5 panel classes + their mount logic stay out of main. */
  private _phase4Panels: Phase4Panels | null = null;
  private _phase4PanelsPromise: Promise<Phase4Panels> | null = null;
  private async ensurePhase4Panels(): Promise<Phase4Panels> {
    if (this._phase4Panels) return this._phase4Panels;
    if (!this._phase4PanelsPromise) {
      this._phase4PanelsPromise = (async () => {
        const mod = await import("./phase4Panels");
        const panels = mod.mountPhase4Panels(this.root, {
          onCorrelationClick: (ev) => this.handleCorrelationClick(ev),
          onOutlierClick: (ev) => this.handleOutlierClick(ev),
          onPivotClick: (ev) => this.handlePivotClick(ev),
          onReportClick: (ev) => this.handleReportClick(ev),
          onColumnTypeClick: (ev) => this.handleColumnTypeClick(ev),
        });
        this._phase4Panels = panels;
        return panels;
      })();
    }
    return this._phase4PanelsPromise;
  }

  /** Lazy-loaded bundle of the 5 advanced ribbon actions (5 streams + 5
   *  Apply functions). Fetched on first advanced-ribbon click; webpack
   *  splits this into its own async chunk so the main bundle stays small. */
  private ribbonActionsPromise: Promise<typeof import("../ribbonActions")> | null = null;
  private async loadRibbonActions(): Promise<typeof import("../ribbonActions")> {
    if (!this.ribbonActionsPromise) {
      this.ribbonActionsPromise = import("../ribbonActions");
    }
    return this.ribbonActionsPromise;
  }

  /** Per-feature controllers (Wave 3 refactor). */
  private branchController!: BranchController;
  private toolCallController!: ToolCallController;
  private knowledgeInjector!: KnowledgeInjector;
  private usageRecorder!: UsageRecorder;
  /** One dispatcher per stream turn so recursion counter resets cleanly. */
  private currentToolCallDispatcher: ToolCallDispatcher | null = null;

  /** Token counter shown in the footer. */
  private tokenBadge!: HTMLElement;

  /** In-flight streaming state. */
  private currentController: AbortController | null = null;
  private currentStreamingMessageId: string | null = null;

  /** Sidebar toggle. */
  private historyOpen: boolean = false;
  /** Settings panel toggle callback. */
  private settingsToggleHandler: () => void = () => undefined;
  private historyToggleHandler: () => void = () => undefined;
  private sessionCreatedHandler: (sessionId: string) => void = () => undefined;

  constructor(parent: HTMLElement, deps: ChatWindowDeps) {
    this.deps = deps;
    this.root = document.createElement("section");
    this.root.classList.add("chat-window");
    this.root.innerHTML = this.renderShell();
    parent.appendChild(this.root);
    this.collectRefs();
    this.bindEvents();
    this.bindQuickActions();
    this.bindMessageActions();
    this.bindHistory();
    this.bindFormulaLibrary();
  }

  /* ---------------- public ---------------- */

  /** Set callbacks the parent wants to observe. */
  public onSettingsToggle(handler: () => void): void {
    this.settingsToggleHandler = handler;
  }

  public onHistoryToggle(handler: () => void): void {
    this.historyToggleHandler = handler;
  }

  public onSessionCreated(handler: (id: string) => void): void {
    this.sessionCreatedHandler = handler;
  }

  /** Replace the visible messages with a session's content. */
  public renderSession(session: ChatSession | null): void {
    this.listEl.innerHTML = "";
    this.views.clear();

    if (!session || session.messages.length === 0) {
      this.showEmptyState();
      return;
    }

    // System messages are not rendered as bubbles (they're context only).
    const visible = session.messages.filter((m) => m.role !== "system");
    if (visible.length === 0) {
      this.showEmptyState();
      return;
    }

    for (const m of visible) {
      this.appendMessageView(m);
    }
    this.scrollToBottom();
  }

  /** Append a message bubble to the chat list. */
  public appendMessage(message: ChatMessage): void {
    if (message.role === "system") return;
    this.appendMessageView(message);
    this.scrollToBottom();
  }

  /** Append a user bubble and immediately trigger a streaming reply.
   *  Returns the promise so callers can await completion (optional). */
  public async sendUserMessage(text: string): Promise<void> {
    if (!text || !text.trim()) return;
    if (this.isStreaming()) return;

    const config = this.deps.getConfig();
    if (!config.apiKey) {
      this.showError("请先在设置中填写 API Key");
      return;
    }

    let session = this.ensureActiveSession();

    // 1. Persist user message
    const userMsg: ChatMessage = {
      id: generateId(),
      role: "user",
      content: text.trim(),
      timestamp: Date.now(),
    };
    session.messages.push(userMsg);
    if (session.messages.filter((m) => m.role !== "system").length === 1) {
      session.title = deriveSessionTitle(text);
    }
    session.updatedAt = Date.now();
    this.deps.upsertSession(session);
    this.appendMessage(userMsg);
    this.updateTokenBadge();

    // 2. Trigger the assistant reply
    await this.streamAssistantReply(session, userMsg);
  }

  /** Inject an action-triggered message (selection-based prompt). */
  public async triggerAction(
    kind: QuickActionKind,
    selection?: ExcelSelection,
    userInput?: string
  ): Promise<void> {
    const prompt = buildActionPrompt(kind, selection, userInput);
    if (kind === "insert") {
      await this.insertLastReplyIntoCell();
      return;
    }
    if (!prompt) return;
    await this.sendUserMessage(prompt);
  }

  /** Cancel an in-flight streaming response. */
  public cancelStreaming(): void {
    if (this.currentController) {
      this.currentController.abort();
      this.currentController = null;
    }
    if (this.currentStreamingMessageId) {
      const view = this.views.get(this.currentStreamingMessageId);
      if (view) {
        const m = view.getMessage();
        view.update({ ...m, content: m.content + (m.content.endsWith("\n") ? "" : "") }, false);
      }
      this.currentStreamingMessageId = null;
    }
    this.setStreamingState(false);
    this.setStatus("");
  }

  /** Are we currently streaming? */
  public isStreaming(): boolean {
    return this.currentController !== null;
  }

  /** Empty the visible chat and prepare for a new conversation. */
  public clearView(): void {
    this.listEl.innerHTML = "";
    this.views.clear();
    this.showEmptyState();
  }

  /** Display a one-shot toast. */
  public toast(message: string, kind: "info" | "error" | "success" = "info"): void {
    const t = document.createElement("div");
    t.classList.add("toast", `toast-${kind}`);
    t.textContent = message;
    document.body.appendChild(t);
    requestAnimationFrame(() => t.classList.add("toast-show"));
    setTimeout(() => {
      t.classList.remove("toast-show");
      setTimeout(() => t.remove(), 300);
    }, 2200);
  }

  /* ---------------- rendering ---------------- */

  private renderShell(): string {
    return `
      <header class="chat-header">
        <button type="button" class="chat-header-btn" data-action="toggleHistory" title="历史">📚</button>
        <h1 class="chat-title">DeepSeek Excel Assistant</h1>
        <button type="button" class="chat-header-btn" data-action="toggleTheme" title="切换主题">🌓</button>
        <button type="button" class="chat-header-btn" data-action="newChat" title="新建对话">➕</button>
        <button type="button" class="chat-header-btn" data-action="toggleSettings" title="设置">⚙️</button>
      </header>

      <div data-ref="contextBar" data-panel="contextBar" class="chat-context-mount"></div>

      <div class="chat-list" data-ref="list" data-panel="chat">
        <div class="chat-empty">
          <div class="chat-empty-icon">💬</div>
          <h3>开始你的第一次对话</h3>
          <p>选中 Excel 数据后，点击快捷操作让 AI 帮你分析，或直接在下方输入问题。</p>
          <div class="chat-empty-tips">
            <p>💡 <strong>提示：</strong></p>
            <ul>
              <li>选中 Excel 区域 → 点击"📊 分析选区"</li>
              <li>Shift+Enter 换行，Enter 直接发送</li>
              <li>点击下方模板可一键发送常用问题</li>
            </ul>
          </div>
        </div>
      </div>

      <div class="chat-status" data-ref="status"></div>

      <div class="chat-templates" data-ref="templates">
        <div class="chat-templates-label">⚡ 快捷模板</div>
        <div class="chat-templates-list"></div>
      </div>

      <div class="chat-quickactions" data-ref="quick"></div>

      <div class="chat-input">
        <textarea
          class="chat-input-field"
          data-ref="input"
          placeholder="输入 / 触发命令菜单 · Enter 发送 · Shift+Enter 换行"
          rows="2"
        ></textarea>
        <button type="button" class="chat-send-btn" data-ref="send">发送</button>
      </div>

      <div data-ref="promptMenu" class="prompt-menu-mount"></div>

      <div data-ref="formulaLibrary" data-panel="formulaLibrary" class="chat-formula-mount"></div>

      <div data-ref="codeGen" data-panel="codeGen" class="chat-codegen-mount"></div>

      <div data-ref="chartPicker" data-panel="chartPicker" class="chat-chartpicker-mount"></div>

      <div data-ref="piiMasker" data-panel="piiMasker" class="chat-pii-mount"></div>

      <div data-ref="correlationPanel" data-panel="correlationMatrix" class="chat-corrpanel-mount"></div>

      <div data-ref="outlierPanel" data-panel="outlierPanel" class="chat-outlierpanel-mount"></div>

      <div data-ref="pivotPanel" data-panel="pivotBuilder" class="chat-pivotpanel-mount"></div>

      <div data-ref="reportPanel" data-panel="reportBuilder" class="chat-reportpanel-mount"></div>

      <div data-ref="colTypePanel" data-panel="columnTypes" class="chat-coltypepanel-mount"></div>

      <div class="chat-footer">
        <span class="chat-word-count" data-ref="wordCount">0 字</span>
        <span class="token-estimate" data-ref="tokenBadge" title="当前会话的输入估算 token 数">0 tokens</span>
        <span class="chat-session-name"></span>
      </div>
    `;
  }

  private collectRefs(): void {
    this.listEl = this.root.querySelector<HTMLElement>('[data-ref="list"]')!;
    this.inputEl = this.root.querySelector<HTMLTextAreaElement>('[data-ref="input"]')!;
    this.sendBtn = this.root.querySelector<HTMLButtonElement>('[data-ref="send"]')!;
    this.quickEl = this.root.querySelector<HTMLElement>('[data-ref="quick"]')!;
    this.templatesEl = this.root.querySelector<HTMLElement>('[data-ref="templates"]')!;
    this.historyBtn = this.root.querySelector<HTMLButtonElement>('[data-action="toggleHistory"]')!;
    this.settingsBtn = this.root.querySelector<HTMLButtonElement>('[data-action="toggleSettings"]')!;
    this.newChatBtn = this.root.querySelector<HTMLButtonElement>('[data-action="newChat"]')!;
    this.statusBar = this.root.querySelector<HTMLElement>('[data-ref="status"]')!;
    this.wordCountEl = this.root.querySelector<HTMLElement>('[data-ref="wordCount"]')!;
    this.tokenBadge = this.root.querySelector<HTMLElement>('[data-ref="tokenBadge"]')!;

    // Mount the context bar inside its dedicated slot.
    const contextMount = this.root.querySelector<HTMLElement>('[data-ref="contextBar"]')!;
    this.contextBar = new ContextBar(contextMount);
    this.contextBar.setModel(this.deps.getConfig().model);

    // Mount the formula library inside its dedicated slot.
    const formulaMount = this.root.querySelector<HTMLElement>('[data-ref="formulaLibrary"]')!;
    this.formulaLibrary = new FormulaLibrary(formulaMount);

    // Mount the code-gen panel (PRD-04).
    const codeGenMount = this.root.querySelector<HTMLElement>('[data-ref="codeGen"]')!;
    this.codeGenPanel = new CodeGenPanelView();
    codeGenMount.appendChild(this.codeGenPanel.element);

    // Hook code-gen copy / retry.
    this.codeGenPanel.element.addEventListener("code-gen-copy", (ev) => {
      const e = ev as CustomEvent<CodeGenCopyDetail>;
      void this.copyGeneratedCode(e.detail.code, e.detail.flavor);
    });
    this.codeGenPanel.element.addEventListener("code-gen-retry", (ev) => {
      const e = ev as CustomEvent<CodeGenRetryDetail>;
      void this.runFormulaToCode(e.detail.flavor);
    });

    // Mount chart picker (PRD-05).
    const chartMount = this.root.querySelector<HTMLElement>('[data-ref="chartPicker"]')!;
    this.chartPicker = new ChartPickerView();
    chartMount.appendChild(this.chartPicker.element);
    this.chartPicker.element.addEventListener("chart-picker-insert", (ev) => {
      const e = ev as CustomEvent<ChartPickerInsertDetail>;
      void this.insertSelectedChart(e.detail.type, e.detail.title);
    });

    // Mount PII masker (PRD-06).
    const piiMount = this.root.querySelector<HTMLElement>('[data-ref="piiMasker"]')!;
    this.piiMasker = new PiiMaskerView();
    piiMount.appendChild(this.piiMasker.element);
    this.piiMasker.element.addEventListener("pii-masker-apply", async (ev) => {
      const e = ev as CustomEvent<PiiMaskerApplyDetail>;
      await this.applyPiiReplacements(e.detail.updates);
    });

    // The 5 advanced-analysis panels (Phase 4) are mounted lazily on the
    // first ribbon click via ensurePhase4Panels() — see phase4Panels.ts.
    // Their mount points are reserved in renderShell() (data-ref="...").

    // Per-feature controllers (Wave 3 refactor). Each takes `this` as the
    // hub dependency; they don't reach back into ChatWindow internals.
    this.branchController = new BranchController(this);
    this.toolCallController = new ToolCallController(this);
    this.knowledgeInjector = new KnowledgeInjector(this);
    this.usageRecorder = new UsageRecorder();

    // Mount the slash-command menu below the input (PRD-02).
    const promptMount = this.root.querySelector<HTMLElement>('[data-ref="promptMenu"]')!;
    this.promptMenu = new PromptMenuView();
    promptMount.appendChild(this.promptMenu.element);
  }

  /* ---------------- event binding ---------------- */

  private bindEvents(): void {
    // Send on Enter; newline on Shift+Enter. Slash-menu keyboard handling
    // is interleaved: ↑/↓ navigate, Enter selects, Esc closes.
    this.inputEl.addEventListener("keydown", (ev) => {
      if (this.promptMenu.visible) {
        if (ev.key === "ArrowDown") { ev.preventDefault(); this.promptMenu.move(1); return; }
        if (ev.key === "ArrowUp")   { ev.preventDefault(); this.promptMenu.move(-1); return; }
        if (ev.key === "Enter" && !ev.shiftKey && !ev.isComposing) {
          const cmd = this.promptMenu.getActive();
          if (cmd) { ev.preventDefault(); this.applySlashCommand(cmd); return; }
        }
        if (ev.key === "Escape") { ev.preventDefault(); this.closeSlashMenu(); return; }
      }
      if (ev.key === "Enter" && !ev.shiftKey && !ev.isComposing) {
        ev.preventDefault();
        this.onSendClicked();
      }
    });

    // Auto-grow textarea + update word count. Also updates the slash-menu
    // visibility as the user types.
    this.inputEl.addEventListener("input", () => {
      this.inputEl.style.height = "auto";
      this.inputEl.style.height = Math.min(this.inputEl.scrollHeight, 180) + "px";
      this.updateWordCount();
      this.refreshSlashMenu();
    });

    // Close the menu when the textarea loses focus (mousedown elsewhere).
    this.inputEl.addEventListener("blur", () => {
      // Delay so a menu click (mousedown) can still register.
      setTimeout(() => this.closeSlashMenu(), 150);
    });

    // Click / Enter from the floating menu itself.
    this.promptMenu.element.addEventListener("prompt-menu-select", (ev) => {
      const e = ev as CustomEvent<PromptMenuSelectDetail>;
      this.applySlashCommand(e.detail.command);
    });

    this.sendBtn.addEventListener("click", () => this.onSendClicked());
    this.settingsBtn.addEventListener("click", () => this.settingsToggleHandler());
    this.historyBtn.addEventListener("click", () => this.historyToggleHandler());
    this.newChatBtn.addEventListener("click", () => {
      this.sessionCreatedHandler("");
      // The parent decides whether to clear current messages and create
      // a new session. We just clear the view immediately for responsiveness.
      this.clearView();
    });

    // Theme toggle button in the header.
    this.root
      .querySelector<HTMLButtonElement>('[data-action="toggleTheme"]')
      ?.addEventListener("click", () => this.toggleTheme());

    // Build the quick-template chips.
    this.renderTemplates();

    // Subscribe to live Excel events.
    this.bindExcelEvents();
  }

  /** Wire up selection-changed / sheet-activated listeners. */
  private bindExcelEvents(): void {
    const unsubSel = excelEvents.onSelectionChange((detail) => {
      this.currentSelection = detail.selection;
      this.contextBar.setSelection(detail.selection);
      // If selection collapsed or moved, surface a tiny status hint so the
      // user understands why quick actions may behave differently.
      if (detail.selection && detail.previousAddress && detail.previousAddress !== detail.selection.address) {
        // No-op for now; keep this here so future UX can hook in.
      }
    });

    const unsubSheet = excelEvents.onSheetChange((detail) => {
      this.currentSheetName = detail.sheetName;
      this.contextBar.setSheet(detail.sheetName);
    });

    this.excelUnsubscribers.push(unsubSel, unsubSheet);

    // Best-effort initial fetch so the bar isn't blank on first load.
    try {
      getSelectedData().then((sel) => {
        this.currentSelection = sel;
        this.contextBar.setSelection(sel);
      }).catch(() => undefined);
    } catch {
      /* noop */
    }
  }

  /** Tear down Excel event listeners. */
  public destroy(): void {
    for (const u of this.excelUnsubscribers) {
      try {
        u();
      } catch {
        /* noop */
      }
    }
    this.excelUnsubscribers = [];
    excelEvents.dispose();
  }

  private renderTemplates(): void {
    const list = this.templatesEl.querySelector<HTMLElement>(".chat-templates-list");
    if (!list) return;
    list.innerHTML = "";
    for (const t of QUICK_TEMPLATES) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.classList.add("chat-template-chip");
      chip.dataset.templateId = t.id;
      chip.innerHTML = `<span class="chip-icon">${t.icon}</span><span class="chip-label">${t.label}</span>`;
      chip.title = t.prompt.slice(0, 80) + "...";
      chip.addEventListener("click", () => this.runTemplate(t.id));
      list.appendChild(chip);
    }
  }

  /** Run a one-click template by id. Pulls selection / formula context
   *  as needed and dispatches the request through the normal flow. */
  public async runTemplate(templateId: string): Promise<void> {
    const tpl = QUICK_TEMPLATES.find((t) => t.id === templateId);
    if (!tpl) return;
    try {
      let selection: ExcelSelection | undefined;
      let formula = "";
      if (tpl.requiresSelection) {
        try {
          selection = await getSelectedData();
        } catch {
          selection = undefined;
        }
        if (templateId === "explain-formula" || templateId === "translate-formula") {
          try {
            if (selection?.address) {
              const formulas = await getRangeFormula(selection.address);
              formula = formulas
                .map((row) => row.filter((c) => c).join(", "))
                .filter((s) => s)
                .join("\n");
            }
          } catch {
            formula = "";
          }
        }
      }

      const ctx =
        selection && selection.values && selection.values.length > 0
          ? `选区地址：${selection.address}\n\n数据：\n${excelValuesToMarkdown(selection.values)}`
          : "(当前未选中数据)";

      let userInput = "";
      // For templates that need additional free-form text, fall back to
      // whatever is currently in the input box.
      if (tpl.prompt.includes("{USER_INPUT}")) {
        userInput = this.inputEl.value || "请根据默认场景填写";
        this.inputEl.value = "";
        this.inputEl.style.height = "auto";
        this.updateWordCount();
      }

      const filled = tpl.prompt
        .replace("{CONTEXT}", ctx)
        .replace("{FORMULA}", formula || "(选区中没有公式)")
        .replace("{USER_INPUT}", userInput);

      await this.sendUserMessage(filled);
    } catch (err: any) {
      this.showError(err?.message || "模板执行失败");
    }
  }

  /** Toggle light/dark theme by toggling the data-theme attribute on
   *  <html>. Persists the user's choice. */
  public toggleTheme(): void {
    const root = document.documentElement;
    const current = root.getAttribute("data-theme") === "dark" ? "dark" : "light";
    const next = current === "dark" ? "light" : "dark";
    root.setAttribute("data-theme", next);
    try {
      localStorage.setItem("deepseek_excel_theme_v1", JSON.stringify(next));
    } catch {
      /* noop */
    }
    this.toast(`已切换到${next === "dark" ? "深色" : "浅色"}主题`, "info");
  }

  /** Apply a theme explicitly - called by the taskpane bootstrap. */
  public applyTheme(theme: "light" | "dark"): void {
    document.documentElement.setAttribute("data-theme", theme);
  }

  /** Insert the last assistant reply into the active cell. Public so
   *  the ribbon command can call it. */
  public async insertLastReplyPublic(): Promise<void> {
    await this.insertLastReplyIntoCell();
  }

  /** Clear the current chat - public for the ribbon command. Asks the
   *  user to confirm first so an accidental ribbon click doesn't lose
   *  the conversation. */
  public clearCurrentChatPublic(): void {
    if (!this.hasMessages()) {
      this.clearView();
      this.toast("当前没有可清空的对话", "info");
      return;
    }
    confirmDialog({
      title: "清空对话",
      message: "确认清空当前对话的消息列表？此操作不影响历史记录。",
      confirmText: "清空",
      cancelText: "取消",
      variant: "danger",
    }).then((ok) => {
      if (ok) {
        this.clearView();
        this.toast("已清空当前对话", "success");
      }
    });
  }

  /** Are there any non-system messages currently shown? */
  private hasMessages(): boolean {
    if (this.views.size > 0) return true;
    const session = this.getActiveSession();
    if (!session) return false;
    return session.messages.some((m) => m.role !== "system");
  }

  /** Export the current session as Markdown. Public for the ribbon. */
  public exportCurrentSessionPublic(): void {
    const sessions = this.deps.getSessions();
    const id = this.deps.getActiveSessionId();
    const session = id ? sessions.find((s) => s.id === id) : null;
    if (!session) {
      this.toast("当前没有可导出的会话", "error");
      return;
    }
    try {
      // Lazy-import to avoid loading download helpers until needed.
      import("../utils/helpers").then(({ downloadSessionAsMarkdown }) => {
        downloadSessionAsMarkdown(session);
        this.toast("已导出为 Markdown", "success");
      });
    } catch (err: any) {
      this.showError(err?.message || "导出失败");
    }
  }

  /** Open the settings panel - public for the ribbon command. */
  public openSettingsPublic(): void {
    this.settingsToggleHandler();
  }

  /** Update the model badge in the context bar - called when the user
   *  switches model in the settings panel. */
  public refreshModelBadge(): void {
    const model = this.deps.getConfig().model;
    this.contextBar?.setModel(model);
  }

  /** Recompute the token badge from the active session. Public so the
   *  parent (taskpane) can call it whenever sessions are mutated. */
  public refreshTokenBadge(): void {
    this.updateTokenBadge();
  }

  /** Public accessor for the latest cached selection. */
  public getCachedSelection(): ExcelSelection | null {
    return this.currentSelection;
  }

  /**
   * Pre-fill the chat input with the given text and focus it. Used by
   * ribbon commands that want to show the user a draft prompt before
   * sending (e.g. "选中分析" - Q3=A: chat prompt already filled).
   */
  public prefillInput(text: string): void {
    this.inputEl.value = text;
    this.updateWordCount();
    try {
      this.inputEl.focus({ preventScroll: true });
      // Move cursor to end of text so the user can edit naturally.
      const len = text.length;
      try {
        this.inputEl.setSelectionRange(len, len);
      } catch {
        /* noop */
      }
      // Smooth-scroll input into view in case the taskpane is tall.
      try {
        this.inputEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
      } catch {
        /* noop */
      }
    } catch {
      /* noop */
    }
  }

  /** Triggered by the ribbon "generate formula" command. Opens the
   *  same prompt the quick-action bar uses. */
  public async runGenerateFormula(): Promise<void> {
    let userText = (this.inputEl.value || "").trim();
    if (!userText) {
      const answer = await promptDialog({
        title: "生成公式",
        message: "请描述你需要的公式（VLOOKUP、SUMIF、IF 等）：",
        placeholder: "例如：根据 A 列查找 B 列对应值并求和",
        confirmText: "生成",
        cancelText: "取消",
        multiline: true,
        rows: 3,
      });
      if (answer === null) return;
      userText = answer;
    }
    if (!userText.trim()) return;
    try {
      const selection = this.currentSelection || (await getSelectedData().catch(() => undefined));
      await this.triggerAction("formula", selection || undefined, userText);
    } catch (err: any) {
      this.showError(err?.message || "生成公式失败");
    }
  }

  /** Triggered by the ribbon "clean data" command. */
  public async runCleanData(): Promise<void> {
    try {
      const selection = this.currentSelection || (await getSelectedData().catch(() => undefined));
      await this.triggerAction("clean", selection || undefined);
    } catch (err: any) {
      this.showError(err?.message || "数据清洗失败");
    }
  }

  /** Triggered by the ribbon "diagnose formulas" command (PRD-01). Scans
   *  the active selection for #REF! / #DIV/0! / etc., sends the list to
   *  the AI, and renders an inline diagnostic report. */
  public async runDiagnoseFormulas(): Promise<void> {
    if (this.isStreaming) {
      this.toast("当前正在生成，请稍后再试", "info");
      return;
    }
    let scan: FormulaScanResult;
    try {
      scan = await scanFormulaErrors();
    } catch (err: any) {
      this.showError(err?.message || "扫描公式失败");
      return;
    }
    if (scan.totalFormulas === 0) {
      this.toast("选区中未发现公式", "info");
      return;
    }
    if (scan.cells.length === 0) {
      this.toast(
        `扫描了 ${scan.totalFormulas} 个公式，全部健康，无需修复`,
        "success"
      );
      return;
    }
    if (scan.totalFormulas > FORMULA_DIAGNOSIS_LIMIT) {
      const ok = await confirmDialog({
        title: "选区较大",
        message: `你选中了 ${scan.totalFormulas} 个公式，AI 诊断可能较慢且消耗较多 token。是否继续？`,
        confirmText: "继续诊断",
        cancelText: "取消",
        variant: "default",
      });
      if (!ok) return;
    } else if (scan.totalFormulas > FORMULA_DIAGNOSIS_WARN) {
      this.toast(`正在扫描 ${scan.cells.length} 个错误公式，预计 5-15 秒…`, "info");
    }

    // Build the host view: empty card, populated when the stream finishes.
    const host = document.createElement("div");
    host.className = "chat-message chat-message-assistant";
    const body = document.createElement("div");
    body.className = "chat-message-body";
    body.innerHTML =
      `<div class="fd-pending">⏳ 正在请求 AI 诊断 ${scan.cells.length} 个错误公式…</div>`;
    host.appendChild(body);
    this.messagesEl.appendChild(host);
    this.scrollToBottom();

    const onAction = async (ev: Event) => {
      const e = ev as CustomEvent<FormulaDiagnosticCardEventDetail>;
      if (e.detail.kind === "apply-fix") {
        const d = e.detail.diagnosis;
        if (!d.fixedFormula) return;
        const confirmed = await confirmDialog({
          title: "应用修复",
          message: `确定将 ${d.address} 的公式覆盖为：\n\n${d.fixedFormula}`,
          confirmText: "应用",
          cancelText: "取消",
          variant: "default",
        });
        if (!confirmed) return;
        try {
          const cell = scan.cells.find((c) => c.address === d.address);
          if (!cell) throw new Error(`找不到单元格 ${d.address}`);
          await fixFormulaAt(cell.fullAddress, d.fixedFormula);
          this.toast(`已修复 ${d.address}`, "success");
        } catch (err: any) {
          this.showError(err?.message || "修复失败");
        }
      } else if (e.detail.kind === "retry") {
        this.toast("重新诊断功能尚未实装", "info");
      }
    };
    host.addEventListener("formula-diagnostic-action", onAction);

    try {
      const ctrl = await diagnoseFormulasStream(
        {
          sheetName: scan.sheetName,
          rangeAddress: scan.rangeAddress,
          errors: scan.cells.map((c) => ({
            address: c.address,
            fullAddress: c.fullAddress,
            formula: c.formula,
            value: c.value,
            error: c.error,
          })),
        },
        this.deps.getConfig(),
        {
          onPartial: (raw) => {
            // Surface a "streaming" hint; the actual JSON is parsed on done.
            const pending = body.querySelector(".fd-pending");
            if (pending) pending.textContent = `⏳ AI 思考中… 已接收 ${raw.length} 字符`;
          },
          onDone: (diagnoses) => {
            body.innerHTML = "";
            const view = new FormulaDiagnosticsView(body, scan, diagnoses);
            this.scrollToBottom();
            // Keep the action handler alive: re-bind on the view's root.
            view["root"].addEventListener("formula-diagnostic-action", onAction);
            this.toast(`诊断完成：${diagnoses.length} 条结果`, "success");
          },
          onError: (err) => {
            body.innerHTML = `<div class="fd-error">诊断失败：${escapeHtml(
              describeApiError(err) || err?.message || "未知错误"
            )}</div>`;
            this.showError(err.message);
          },
        }
      );
      this.currentAbort = ctrl;
    } catch (err: any) {
      body.innerHTML = `<div class="fd-error">启动诊断失败：${escapeHtml(err?.message || "未知错误")}</div>`;
      this.showError(err?.message || "启动诊断失败");
    }
  }

  /** Triggered by the ribbon "公式转 VBA" command (PRD-04). Collects
   *  every formula in the selection, asks the AI to translate them into
   *  a VBA Sub (or Office Script on demand), and renders the resulting
   *  code in `codeGenPanel`. Streams tokens as they arrive. */
  public async runFormulaToCode(flavor: ScriptFlavor = "vba"): Promise<void> {
    if (this.isStreaming && !(this.codeGenPanel as any)?.element) {
      this.toast("当前正在生成，请稍后再试", "info");
      return;
    }

    let collected;
    try {
      collected = await collectFormulasForCode();
    } catch (err: any) {
      this.showError(err?.message || "无法读取选区公式");
      return;
    }
    if (!collected || collected.formulas.length === 0) {
      this.toast("选区中未发现公式", "info");
      return;
    }
    if (collected.formulas.length > 200) {
      const ok = await confirmDialog({
        title: "公式较多",
        message: `你选中了 ${collected.formulas.length} 个公式。AI 转写脚本可能非常长。继续？`,
        confirmText: "继续",
        cancelText: "取消",
        variant: "default",
      });
      if (!ok) return;
    }
    if (collected.formulas.some((f) => f.hasExternalRef)) {
      this.toast("已检测到跨表 / 外部引用，VBA 会保留 Worksheets(...) 引用", "info");
    }

    this.codeGenAbort?.abort();
    this.codeGenPanel.show(flavor, "");

    const t0 = Date.now();
    let aborted = false;
    void (async () => {
      const ctrl = await translateToScript(
        { sheet: collected.sheet, block: collected.block },
        flavor,
        this.deps.getConfig(),
        {
          onPartial: (text) => {
            if (aborted) return;
            this.codeGenPanel.setCode(text);
          },
          onDone: () => {
            if (aborted) return;
            this.codeGenPanel.setStatus(
              "done",
              `✓ 完成 · ${Math.max(1, Math.round((Date.now() - t0) / 1000))}s`
            );
            this.toast(
              `${flavor === "vba" ? "VBA" : "Office Scripts"} 代码已生成`,
              "success"
            );
          },
          onError: (err) => {
            if (aborted) return;
            this.codeGenPanel.setStatus("error", "生成失败");
            this.showError(err?.message || "生成失败");
          },
        }
      );
      if (aborted) {
        try { ctrl.abort(); } catch { /* noop */ }
        return;
      }
      this.codeGenAbort = ctrl;
    })();
  }

  /** Copy the panel's code to the system clipboard. For VBA we also
   *  pop a tiny dialog reminding the user about Alt+F11 per PRD-04 §7. */
  private async copyGeneratedCode(code: string, flavor: ScriptFlavor): Promise<void> {
    if (!code || !code.trim()) {
      this.toast("没有可复制的代码", "info");
      return;
    }
    try {
      if (navigator?.clipboard?.writeText) {
        await navigator.clipboard.writeText(code);
      } else {
        // Fallback for hosts without clipboard API.
        const ta = document.createElement("textarea");
        ta.value = code;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
      }
      this.toast(`已复制 ${flavor === "vba" ? "VBA" : "Office Scripts"} 代码`, "success");
      if (flavor === "vba") {
        const labels: Record<ScriptFlavor, string> = { vba: "VBA", "office-scripts": "Office Scripts" };
        void labels;
        // Per PRD-04 §7 v1: open a dialog guiding the user through Alt+F11.
        await confirmDialog({
          title: "已复制 VBA 代码",
          message:
            "代码已复制到剪贴板。\n\n" +
            "下一步：按 Alt+F11 打开 VBA 编辑器 → 插入 → 模块 → 粘贴代码 → 运行。",
          confirmText: "知道了",
          cancelText: "关闭",
          variant: "default",
        });
      }
    } catch (err: any) {
      this.showError(err?.message || "复制失败");
    }
  }

  /** Triggered by the ribbon "插入图表" command (PRD-05). Reads the
   *  selected range, asks the AI for up to 3 chart recommendations, and
   *  shows them as clickable cards. Clicking a card inserts the chart. */
  public async runInsertChart(): Promise<void> {
    this.chartPicker.hide();

    let info;
    try {
      info = await getSelectedRangeInfo();
    } catch (err: any) {
      this.showError(err?.message || "无法读取选区");
      return;
    }
    if (!info || info.rowCount < 2) {
      this.toast("至少需要 2 行数据才能生成图表", "info");
      return;
    }
    this.chartSourceAddress = info.address;

    this.chartPicker.show("AI 推荐中…");

    try {
      const ctrl = await recommendChartStream(this.deps.getConfig(), info, {
        onPartial: (raw) => {
          const lines = raw.split("\n").slice(-2).join(" ");
          this.chartPicker.setStatus(`生成中 · ${lines.length} 字符`);
        },
        onDone: (result) => {
          this.chartPicker.setRecommendations(
            result.recommendations,
            result.usedFallback
          );
          if (result.usedFallback) {
            this.toast("AI 未响应，已用本地推荐", "info");
          } else {
            this.toast(`已生成 ${result.recommendations.length} 个图表推荐`, "success");
          }
        },
        onError: (err) => {
          this.chartPicker.hide();
          this.showError(err?.message || "推荐失败");
        },
      });
      this.chartAbort = ctrl;
    } catch (err: any) {
      this.chartPicker.hide();
      this.showError(err?.message || "启动失败");
    }
  }

  /** Insert a chart of the given type into Excel using the captured
   *  source address. Shows a toast on success / failure. */
  private async insertSelectedChart(type: ExcelChartType, title: string): Promise<void> {
    if (!this.chartSourceAddress) {
      this.toast("选区信息已过期,请重新点击 ribbon 按钮", "info");
      return;
    }
    this.chartPicker.setStatus("正在插入…");
    try {
      const msg = await insertChart(this.chartSourceAddress, type, title);
      this.chartPicker.hide();
      this.toast(msg || "已插入图表", "success");
    } catch (err: any) {
      this.chartPicker.setStatus("插入失败");
      this.showError(err?.message || "插入失败");
    }
  }

  /** Triggered by the ribbon "多选区分析" command (PRD-07). Reads
   *  multiple disjoint regions of the selection (Ctrl+click) and asks
   *  the AI to analyze them as one cohesive request. */
  public async runMultiSelectionAnalyze(): Promise<void> {
    let multi;
    try {
      multi = await readMultiSelection();
    } catch (err: any) {
      this.showError(err?.message || "无法读取多选区");
      return;
    }
    if (!multi || multi.parts.length === 0) {
      this.toast("未检测到选区", "info");
      return;
    }
    if (multi.parts.length === 1) {
      this.toast("当前仅有一个区域，请使用普通「分析选区」", "info");
      // Still send the single region.
    }
    const prompt =
      `我选中了 ${multi.parts.length} 个区域（按住 Ctrl 选中），请综合分析：\n\n` +
      multi.block +
      `\n\n请说明：1) 各区域的数据意义；2) 它们如何关联；3) 关键发现。`;
    this.inputEl.value = prompt;
    this.inputEl.dispatchEvent(new Event("input"));
    this.toast(
      `已合并 ${multi.parts.length} 个区域 · 共 ${multi.totalCells} 单元格，请点发送`,
      "info"
    );
  }

  /** Triggered by the ribbon "数据脱敏" command (PRD-06). Scans the
   *  selection for sensitive values locally (no AI needed - regex is
   *  enough for the well-known formats) and shows the user a checklist
   *  of replacements before applying. */
  public async runMaskPII(): Promise<void> {
    this.piiMasker.hide();
    let scan;
    try {
      scan = await scanSelectionForPII();
    } catch (err: any) {
      this.showError(err?.message || "扫描失败");
      return;
    }
    if (scan.hits.length === 0) {
      this.toast("未在选区中发现疑似敏感数据", "success");
      return;
    }
    this.piiMasker.show(scan.sheet, scan.hits);
    this.toast(`扫描完成：${scan.hits.length} 个疑似敏感值`, "info");
  }

  /** Apply the chosen batch of fake-value replacements. */
  private async applyPiiReplacements(updates: Array<{ address: string; value: string }>): Promise<void> {
    if (updates.length === 0) return;
    try {
      const written = await batchReplaceCells(updates);
      this.toast(`已替换 ${written} 个单元格`, "success");
      this.piiMasker.hide();
    } catch (err: any) {
      this.showError(err?.message || "替换失败");
    }
  }

  /* =====================================================================
   * Phase 4 — Advanced ribbon handlers
   * Each handler shows its dedicated panel and seeds it with a placeholder
   * payload. The AI completion path is stubbed for now (the panel still
   * renders correctly so verifyRibbon can confirm the wiring).
   * ===================================================================== */

  /** Ribbon "相关性矩阵" — reads numeric columns, shows the heat-map panel. */
  public async runCorrelationMatrix(): Promise<void> {
    const sel = await this.fetchSelectionOrToast();
    if (!sel) return;
    const headers = sel.headers.length ? sel.headers : sel.values[0]?.map((_, i) => `Col${i + 1}`) || [];
    const samples = this.buildSampleRecords(sel.values, headers);
    const panels = await this.ensurePhase4Panels();
    panels.correlation.show({
      selection: sel.address,
      reasoning: "AI 计算中…",
    });
    try {
      const { correlationStream } = await this.loadRibbonActions();
      await correlationStream(
        { selection: sel.address, headers, samples },
        this.deps.getConfig(),
        {
          onPartial: () => this.ribbonBanner.update?.("相关性矩阵 · AI 计算中…"),
          onDone: (r) => {
            const cells = r.cells || [];
            const labels = r.labels?.length ? r.labels : headers;
            panels.correlation.show({
              selection: sel.address,
              matrix: { labels, cells },
              reasoning: r.reasoning || "",
            });
            this.toast(`已计算 ${labels.length}×${labels.length} 相关性矩阵`, "success");
          },
          onError: (err) => this.showError(err?.message || "AI 失败"),
        }
      );
    } catch (err: any) {
      this.showError(err?.message || "启动失败");
    }
  }

  /** Ribbon "异常值" — scans selected numeric range, shows outlier panel. */
  public async runDetectOutliers(): Promise<void> {
    const sel = await this.fetchSelectionOrToast();
    if (!sel) return;
    const headers = sel.headers.length ? sel.headers : [];
    const samples = this.buildSampleRecords(sel.values, headers);
    const panels = await this.ensurePhase4Panels();
    panels.outlier.show({
      selection: sel.address,
      method: "zscore",
      outliers: [],
      summary: "AI 扫描中…",
    });
    try {
      const { outlierStream } = await this.loadRibbonActions();
      await outlierStream(
        { selection: sel.address, headers, samples, method: "zscore" },
        this.deps.getConfig(),
        {
          onPartial: () => this.ribbonBanner.update?.("异常值 · AI 扫描中…"),
          onDone: (r) => {
            panels.outlier.show({
              selection: sel.address,
              method: r.method || "zscore",
              outliers: r.outliers || [],
              summary: r.summary || "",
            });
            this.toast(`检测到 ${r.outliers?.length || 0} 个异常值`, "info");
          },
          onError: (err) => this.showError(err?.message || "AI 失败"),
        }
      );
    } catch (err: any) {
      this.showError(err?.message || "启动失败");
    }
  }

  /** Ribbon "AI 透视表" — proposes a pivot spec, shows preview panel. */
  public async runCreatePivot(): Promise<void> {
    const sel = await this.fetchSelectionOrToast();
    if (!sel) return;
    const headers = sel.headers.length ? sel.headers : [];
    const samples = this.buildSampleRecords(sel.values, headers);
    const panels = await this.ensurePhase4Panels();
    panels.pivot.show({
      selection: sel.address,
      rationale: "AI 推荐中…",
    });
    try {
      const { pivotSpecStream } = await this.loadRibbonActions();
      await pivotSpecStream(
        { selection: sel.address, headers, samples },
        this.deps.getConfig(),
        {
          onPartial: () => this.ribbonBanner.update?.("透视表 · AI 推荐中…"),
          onDone: (r) => {
            panels.pivot.show({
              selection: sel.address,
              spec: {
                rows: r.rows || [],
                columns: r.columns || [],
                values: r.values || [],
              },
              rationale: r.rationale || "",
            });
            this.toast("AI 已推荐透视表布局,可在面板中确认", "success");
          },
          onError: (err) => this.showError(err?.message || "AI 失败"),
        }
      );
    } catch (err: any) {
      this.showError(err?.message || "启动失败");
    }
  }

  /** Ribbon "快速报告" — generates a structured report, shows panel. */
  public async runQuickReport(): Promise<void> {
    const sel = await this.fetchSelectionOrToast();
    if (!sel) return;
    const headers = sel.headers.length ? sel.headers : [];
    const samples = this.buildSampleRecords(sel.values, headers);
    const panels = await this.ensurePhase4Panels();
    panels.report.show({
      selection: sel.address,
      title: "数据报告",
      summary: "AI 生成中…",
    });
    try {
      const { reportStream } = await this.loadRibbonActions();
      await reportStream(
        {
          selection: sel.address,
          headers,
          rowCount: sel.rowCount,
          columnCount: sel.columnCount,
          samples,
        },
        this.deps.getConfig(),
        {
          onPartial: () => this.ribbonBanner.update?.("报告 · AI 生成中…"),
          onDone: (r) => {
            panels.report.show({
              selection: sel.address,
              title: r.title,
              summary: r.summary,
              sections: r.sections || [],
              recommendations: r.recommendations || [],
            });
            this.toast("报告草稿已生成", "success");
          },
          onError: (err) => this.showError(err?.message || "AI 失败"),
        }
      );
    } catch (err: any) {
      this.showError(err?.message || "启动失败");
    }
  }

  /** Ribbon "列类型" — infers types per column, shows panel. */
  public async runInferColumnTypes(): Promise<void> {
    const sel = await this.fetchSelectionOrToast();
    if (!sel) return;
    const headers = sel.headers.length ? sel.headers : [];
    const samples = this.buildSampleRecords(sel.values, headers);
    const panels = await this.ensurePhase4Panels();
    panels.colType.show({
      selection: sel.address,
      rows: [],
      summary: "AI 推断中…",
    });
    try {
      const { columnTypeStream } = await this.loadRibbonActions();
      await columnTypeStream(
        { selection: sel.address, headers, samples },
        this.deps.getConfig(),
        {
          onPartial: () => this.ribbonBanner.update?.("列类型 · AI 推断中…"),
          onDone: (r) => {
            const rows = (r.rows || []).map((row) => ({
              column: row.column,
              proposedType: row.proposedType,
              confidence: row.confidence,
              format: row.format,
              namedRange: row.namedRange || `col_${row.column.replace(/[^A-Za-z0-9_]/g, "_")}`,
              reason: row.reason,
            }));
            panels.colType.show({
              selection: sel.address,
              rows,
              summary: r.summary || `检测到 ${rows.length} 列`,
            });
            this.toast(`已推断 ${rows.length} 列类型`, "success");
          },
          onError: (err) => this.showError(err?.message || "AI 失败"),
        }
      );
    } catch (err: any) {
      this.showError(err?.message || "启动失败");
    }
  }

  /** Read the current selection (headers + values) or toast an error and
   *  return null. Shared by all 5 advanced-panel ribbon handlers. */
  private async fetchSelectionOrToast(): Promise<{
    address: string;
    headers: string[];
    values: any[][];
    rowCount: number;
    columnCount: number;
  } | null> {
    let info;
    try {
      info = await getSelectedRangeInfo();
    } catch (err: any) {
      this.showError(err?.message || "无法读取选区");
      return null;
    }
    if (!info) {
      this.toast("请先选中数据", "info");
      return null;
    }
    if (info.columnCount < 2) {
      this.toast("至少需要 2 列数据", "info");
      return null;
    }
    let values: any[][] = [];
    try {
      const sel = await getSelectedData();
      values = (sel && sel.values) || [];
    } catch {
      values = [];
    }
    return {
      address: info.address,
      headers: info.headers || [],
      values,
      rowCount: info.rowCount,
      columnCount: info.columnCount,
    };
  }

  /** Convert a 2D array of values + a header row into an array of
   *  objects keyed by header name (the shape our AI requests expect). */
  private buildSampleRecords(
    values: any[][],
    headers: string[]
  ): Array<Record<string, string | number | null>> {
    if (!values || values.length === 0) return [];
    const hasHeaderRow =
      headers.length > 0 && headers.length === values[0].length;
    const rows = hasHeaderRow ? values.slice(1) : values;
    const out: Array<Record<string, string | number | null>> = [];
    for (const row of rows.slice(0, 30)) {
      const obj: Record<string, string | number | null> = {};
      for (let i = 0; i < row.length; i++) {
        const k = hasHeaderRow ? headers[i] : `c${i}`;
        const v = row[i];
        obj[k] = typeof v === "number" || typeof v === "string" ? v : v == null ? null : String(v);
      }
      out.push(obj);
    }
    return out;
  }

  /* =====================================================================
   * Phase 4 — Apply/Copy click handlers
   * Each method is wired to the corresponding panel's root element via
   * delegated click listeners mounted in collectRefs.
   * ===================================================================== */

  private async handleCorrelationClick(ev: MouseEvent): Promise<void> {
    const target = ev.target as HTMLElement;
    if (!target.classList.contains("corr-panel__btn")) return;
    const panels = await this.ensurePhase4Panels();
    const payload = panels.correlation.getCurrentPayload();
    if (!payload || !payload.matrix) return;
    if (target.classList.contains("corr-panel__btn--apply")) {
      try {
        // Default: write to the cell two columns to the right of the source.
        const start = this.computeNextFreeAddress(payload.selection) || "H1";
        const { insertCorrelationMatrix } = await this.loadRibbonActions();
        const msg = await insertCorrelationMatrix(
          payload.matrix.labels,
          payload.matrix.cells,
          start
        );
        this.toast(msg, "success");
      } catch (err: any) {
        this.showError(err?.message || "插入失败");
      }
    } else if (target.classList.contains("corr-panel__btn--copy")) {
      const text = JSON.stringify(payload.matrix, null, 2);
      await navigator.clipboard?.writeText(text).catch(() => undefined);
      this.toast("已复制相关性矩阵 JSON 到剪贴板", "info");
    }
  }

  private async handleOutlierClick(ev: MouseEvent): Promise<void> {
    const target = ev.target as HTMLElement;
    if (!target.classList.contains("outlier-panel__btn")) return;
    const panels = await this.ensurePhase4Panels();
    const payload = panels.outlier.getCurrentPayload();
    if (!payload || !payload.outliers) return;
    if (target.classList.contains("outlier-panel__btn--highlight")) {
      try {
        const refs = payload.outliers.map((o) => ({ rowIndex: o.rowIndex, column: o.column }));
        const base = this.parseAddress(payload.selection);
        const { highlightOutliers } = await this.loadRibbonActions();
        const n = await highlightOutliers(refs, base ?? undefined);
        this.toast(`已高亮 ${n} 个异常单元格`, "success");
      } catch (err: any) {
        this.showError(err?.message || "高亮失败");
      }
    } else if (target.classList.contains("outlier-panel__btn--delete")) {
      this.toast("删除行功能:待接入(优先高亮)", "info");
    }
  }

  private async handlePivotClick(ev: MouseEvent): Promise<void> {
    const target = ev.target as HTMLElement;
    if (!target.classList.contains("pivot-panel__btn")) return;
    const panels = await this.ensurePhase4Panels();
    const payload = panels.pivot.getCurrentPayload();
    if (!payload || !payload.spec || !payload.selection) return;
    if (target.classList.contains("pivot-panel__btn--apply")) {
      try {
        const sheetName = `透视_${Date.now().toString(36).slice(-4)}`;
        const { createPivotTable } = await this.loadRibbonActions();
        const msg = await createPivotTable(payload.selection, sheetName, payload.spec);
        this.toast(msg, "success");
      } catch (err: any) {
        this.showError(err?.message || "创建透视表失败");
      }
    } else if (target.classList.contains("pivot-panel__btn--copy")) {
      const text = JSON.stringify(payload.spec, null, 2);
      await navigator.clipboard?.writeText(text).catch(() => undefined);
      this.toast("已复制透视表配置 JSON", "info");
    }
  }

  private async handleReportClick(ev: MouseEvent): Promise<void> {
    const target = ev.target as HTMLElement;
    if (!target.classList.contains("report-panel__btn")) return;
    const panels = await this.ensurePhase4Panels();
    const payload = panels.report.getCurrentPayload();
    if (!payload) return;
    if (target.classList.contains("report-panel__btn--apply")) {
      try {
        const { writeReportSheet } = await this.loadRibbonActions();
        const msg = await writeReportSheet({
          title: payload.title || "数据报告",
          summary: payload.summary || "",
          sections: payload.sections || [],
          recommendations: payload.recommendations || [],
        });
        this.toast(msg, "success");
      } catch (err: any) {
        this.showError(err?.message || "写入报告失败");
      }
    } else if (target.classList.contains("report-panel__btn--copy")) {
      const md = this.renderReportMarkdown(payload);
      await navigator.clipboard?.writeText(md).catch(() => undefined);
      this.toast("已复制 Markdown 到剪贴板", "info");
    }
  }

  private async handleColumnTypeClick(ev: MouseEvent): Promise<void> {
    const target = ev.target as HTMLElement;
    if (!target.classList.contains("coltype-panel__btn")) return;
    const panels = await this.ensurePhase4Panels();
    const payload = panels.colType.getCurrentPayload();
    if (!payload || !payload.rows) return;
    if (target.classList.contains("coltype-panel__btn--apply")) {
      try {
        const base = this.parseAddress(payload.selection);
        if (!base) {
          this.toast("无法解析选区,跳过格式应用", "info");
          return;
        }
        const { applyColumnFormatting } = await this.loadRibbonActions();
        const msg = await applyColumnFormatting(
          base,
          payload.rows.length,
          payload.rows.map((r) => ({
            column: r.column,
            proposedType: r.proposedType,
            confidence: r.confidence,
            format: r.format,
            namedRange: r.namedRange,
          }))
        );
        this.toast(msg, "success");
      } catch (err: any) {
        this.showError(err?.message || "应用格式失败");
      }
    } else if (target.classList.contains("coltype-panel__btn--copy")) {
      const text = JSON.stringify(payload.rows, null, 2);
      await navigator.clipboard?.writeText(text).catch(() => undefined);
      this.toast("已复制列类型配置 JSON", "info");
    }
  }

  /** Parse "Sheet!A1:D10" or "A1:D10" into a base {sheet, rowStart, columnStart}. */
  private parseAddress(addr?: string): { sheet: string; rowStart: number; columnStart: number } | null {
    if (!addr) return null;
    const bang = addr.indexOf("!");
    const sheet = bang >= 0 ? addr.slice(0, bang) : "";
    const a = (bang >= 0 ? addr.slice(bang + 1) : addr).split(":")[0];
    const m = /([A-Z]+)(\d+)/.exec(a);
    if (!m) return null;
    const colStart = this.letterToColIndex(m[1]);
    return { sheet, rowStart: parseInt(m[2], 10), columnStart: colStart };
  }

  /** Convert column letter to 1-based column index (A→1, B→2, ..., Z→26, AA→27). */
  private letterToColIndex(letters: string): number {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  /** Compute a free address 2 columns to the right of the source range
   *  end. Used as the default insert target for the correlation matrix. */
  private computeNextFreeAddress(addr?: string): string | null {
    if (!addr) return null;
    const bang = addr.indexOf("!");
    const clean = bang >= 0 ? addr.slice(bang + 1) : addr;
    const [start, end] = clean.split(":");
    const last = end || start;
    const m = /([A-Z]+)(\d+)/.exec(last);
    if (!m) return null;
    // Shift right by 2 columns from the last column of the source range.
    const newCol = this.letterToColIndex(m[1]) + 2;
    const newColLetters = this.colIndexToLetters(newCol);
    const sheet = bang >= 0 ? addr.slice(0, bang + 1) : "";
    return `${sheet}${newColLetters}1`;
  }

  private colIndexToLetters(idx: number): string {
    let n = idx;
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  private renderReportMarkdown(p: { title?: string; summary?: string; sections?: Array<{ heading: string; body: string }>; recommendations?: string[] }): string {
    const lines: string[] = [];
    lines.push(`# ${p.title || "数据报告"}`);
    if (p.summary) lines.push("", p.summary);
    for (const s of p.sections || []) {
      lines.push("", `## ${s.heading}`, "", s.body);
    }
    if (p.recommendations && p.recommendations.length) {
      lines.push("", "## 建议");
      for (const r of p.recommendations) lines.push(`- ${r}`);
    }
    return lines.join("\n");
  }

  private updateWordCount(): void {
    const v = this.inputEl.value || "";
    // Use Array.from to count CJK chars as one each, not per byte.
    const len = Array.from(v.trim()).length;
    this.wordCountEl.textContent = `${len} 字`;
    this.updateTokenBadge();
  }

  /** Recompute the token estimate for the active session. We sum the
   *  system prompt + every message's content (rough but useful). */
  private updateTokenBadge(): void {
    if (!this.tokenBadge) return;
    const config = this.deps.getConfig();
    let total = 0;
    if (config.systemPrompt) total += estimateTokens(config.systemPrompt);
    const session = this.getActiveSession();
    if (session) {
      for (const m of session.messages) {
        total += estimateTokens(m.content || "");
      }
    }
    this.tokenBadge.textContent = `~${formatTokens(total)} tokens`;
    // Warn at 80% of a typical 8K context window.
    if (total > 6400) this.tokenBadge.classList.add("token-estimate-warn");
    else this.tokenBadge.classList.remove("token-estimate-warn");
    this.tokenBadge.title = `估算 token 数：${total}\n（系统提示 + 当前会话全部消息）`;
  }

  private bindQuickActions(): void {
    mountQuickActions(this.quickEl);

    this.quickEl.addEventListener("quick-action", async (ev) => {
      const e = ev as CustomEvent;
      const { kind, template } = e.detail;

      if (kind === "stop") {
        this.cancelStreaming();
        return;
      }

      // All non-stop actions may need selection context.
      try {
        if (kind === "insert") {
          await this.triggerAction("insert");
          return;
        }
        // For analyze/clean we always need the selection.
        // For formula we ask the user but still try to attach selection if any.
        let selection: ExcelSelection | undefined = this.currentSelection || undefined;
        if (kind === "analyze" || kind === "clean" || kind === "formula") {
          // Use cached selection if it's recent; otherwise re-fetch.
          try {
            const fresh = await getSelectedData();
            selection = fresh;
            this.currentSelection = fresh;
          } catch {
            // Use the cached value if a fresh fetch fails.
            selection = this.currentSelection || undefined;
          }
        }

        // For formula, prefer a focused prompt: replace the placeholder
        // and let the user confirm in the input box.
        if (kind === "formula") {
          let userText = (this.inputEl.value || "").trim();
          if (!userText) {
            const answer = await promptDialog({
              title: "生成公式",
              message: "请描述你需要的公式：",
              placeholder: "例如：按月汇总 A 列的销售数据",
              confirmText: "生成",
              cancelText: "取消",
              multiline: true,
              rows: 3,
            });
            if (answer === null || !answer.trim()) {
              this.toast("请先输入要生成的公式描述", "info");
              return;
            }
            userText = answer;
          } else {
            this.inputEl.value = "";
            this.updateWordCount();
          }
          await this.triggerAction("formula", selection, userText);
          return;
        }

        // Otherwise dispatch with the template + selection.
        await this.triggerAction(kind, selection);
      } catch (err: any) {
        this.showError(err?.message || "操作失败");
      }
    });
  }

  private bindMessageActions(): void {
    this.listEl.addEventListener("message-action", async (ev) => {
      const e = ev as CustomEvent;
      const { action, messageId, rawText, plainText } = e.detail || {};

      if (action === "copy") {
        const ok = await navigator.clipboard?.writeText(rawText || "").then(
          () => true,
          () => false
        );
        this.toast(ok ? "已复制" : "复制失败", ok ? "success" : "error");
        return;
      }

      if (action === "insert") {
        try {
          // Smart insert: if the AI's reply contains a formula, insert
          // as a formula; otherwise insert as plain text.
          const formula = detectFormula(rawText || "");
          if (formula) {
            await writeFormula(formula);
            this.toast("已识别为公式，插入为活动公式", "success");
          } else {
            await insertTextToCell(plainText || rawText || "");
            this.toast("已插入到当前单元格", "success");
          }
        } catch (err: any) {
          this.showError(err?.message || "插入失败");
        }
        return;
      }

      if (action === "regenerate") {
        const session = this.getActiveSession();
        if (!session) return;
        // Find the last user message before this assistant message.
        const idx = session.messages.findIndex((m) => m.id === messageId);
        if (idx < 0) return;
        let userIdx = -1;
        for (let i = idx - 1; i >= 0; i--) {
          if (session.messages[i].role === "user") {
            userIdx = i;
            break;
          }
        }
        if (userIdx < 0) {
          this.toast("找不到原始用户消息", "error");
          return;
        }
        // Drop everything from this assistant message onward (in memory).
        // The persisted copy is rewritten on the next upsert.
        session.messages = session.messages.slice(0, idx);
        const view = this.views.get(messageId);
        if (view) view.destroy();
        this.views.delete(messageId);
        await this.streamAssistantReply(session, session.messages[userIdx]);
        return;
      }

      if (action === "branch") {
        await this.runBranchFromMessage(messageId);
        return;
      }
    });
  }

  /** PRD-08: create a sibling assistant reply for the same user turn so
   *  users can keep multiple AI answers in parallel without losing the
   *  original. The old assistant bubble is hidden (not destroyed) so the
   *  branch hierarchy is visible when the user scrolls back.
   *
   *  Body now lives in `BranchController` — the controller owns the
   *  workflow (count siblings, hide originals, render divider, tag
   *  branch metadata) and depends only on the hub. This wrapper exists
   *  purely so the existing click-handler dispatch (`runBranchFromMessage`)
   *  still routes through the controller. */
  private async runBranchFromMessage(assistantMessageId: string): Promise<void> {
    await this.branchController.runBranchFromMessage(assistantMessageId);
  }

  private bindHistory(): void {
    // No-op here - the HistoryPanel is rendered by the parent.
  }

  /** Listen for formula-pick events from the FormulaLibrary component.
   *  When the user clicks a card we drop the pre-filled prompt into
   *  the chat input and focus it so they can edit the placeholders. */
  private bindFormulaLibrary(): void {
    this.root.addEventListener("formula-pick", (ev) => {
      const e = ev as CustomEvent;
      const detail = e.detail || {};
      const prompt: string = detail.prompt || "";
      if (!prompt) return;
      this.inputEl.value = prompt;
      this.inputEl.dispatchEvent(new Event("input"));
      this.inputEl.focus();
      // Place cursor at the first {PLACEHOLDER} token.
      const m = /\{[^}]+\}/.exec(prompt);
      if (m) {
        this.inputEl.setSelectionRange(m.index, m.index + m[0].length);
      }
      this.toast(`已填入「${detail.name}」公式模板，请替换占位符`, "info");
    });
  }

  /* ---------------- helpers ---------------- */

  /* ---- slash menu (PRD-02) ---- */

  /** Recompute menu visibility based on the current input value.
   *  Menu shows when the only content (trimmed) is "/" or starts with "/". */
  private refreshSlashMenu(): void {
    const v = this.inputEl.value;
    const caret = this.inputEl.selectionStart ?? v.length;
    // Only the very first token (left of caret, with no whitespace yet).
    const before = v.slice(0, caret);
    if (!before.startsWith("/")) {
      this.closeSlashMenu();
      return;
    }
    // If the user has typed whitespace before completing the command, close.
    if (/\s/.test(before)) {
      this.closeSlashMenu();
      return;
    }
    this.slashQuery = before.slice(1); // strip leading "/"
    this.promptMenu.show(this.slashQuery);
  }

  private closeSlashMenu(): void {
    if (this.promptMenu) this.promptMenu.hide();
  }

  /** Insert the chosen command's template into the input, replacing the
   *  "/" trigger the user typed, and place the cursor at {USER_INPUT}. */
  private async applySlashCommand(cmd: SlashCommand): Promise<void> {
    this.closeSlashMenu();

    if (cmd.requiresSelection) {
      const sel = await getSelectedData();
      if (!sel || !sel.address) {
        this.toast("此命令需要先在 Excel 中选中一个区域", "info");
        return;
      }
    }

    let filled = cmd.template;
    try {
      filled = await this.fillSlashTemplate(cmd.template);
    } catch (err) {
      try { console.warn("[DeepSeek] template fill failed", err); } catch {}
    }

    this.inputEl.value = filled;
    this.inputEl.dispatchEvent(new Event("input"));

    // Place cursor at the first {USER_INPUT} token, else at end.
    const userPos = filled.indexOf("{USER_INPUT}");
    if (userPos >= 0) {
      this.inputEl.focus();
      this.inputEl.setSelectionRange(userPos, userPos + "{USER_INPUT}".length);
    } else {
      this.inputEl.focus();
      const end = this.inputEl.value.length;
      this.inputEl.setSelectionRange(end, end);
    }
    this.toast(`已填入 ${cmd.trigger} 模板`, "info");
  }

  /** Replace placeholders in a slash-command template with live data. */
  private async fillSlashTemplate(tpl: string): Promise<string> {
    if (!tpl.includes("{CONTEXT}") && !tpl.includes("{FORMULA}")) return tpl;
    let out = tpl;

    if (out.includes("{CONTEXT}")) {
      try {
        const sel = await getSelectedData();
        const md = sel && sel.values ? excelValuesToMarkdown(sel.values as any[][]) : "(无选区)";
        out = out.split("{CONTEXT}").join(md);
      } catch {
        out = out.split("{CONTEXT}").join("(获取选区失败)");
      }
    }

    if (out.includes("{FORMULA}")) {
      try {
        // Read the formula under the active cell, if any.
        const sel = await getSelectedData();
        const addr = sel && sel.values && sel.values.length > 0 ? sel.address : "";
        let formulaText = "";
        if (addr) {
          const arr = await getRangeFormula(addr);
          formulaText = Array.isArray(arr) && arr.length > 0 ? String(arr[0][0] ?? "") : "";
        }
        out = out.split("{FORMULA}").join(formulaText || "(活动单元格无公式)");
      } catch {
        out = out.split("{FORMULA}").join("(获取公式失败)");
      }
    }

    return out;
  }

  private showEmptyState(): void {
    if (this.listEl.querySelector(".chat-empty")) return;
    this.listEl.innerHTML = `
      <div class="chat-empty">
        <div class="chat-empty-icon">💬</div>
        <h3>开始你的第一次对话</h3>
        <p>选中 Excel 数据后，点击快捷操作让 AI 帮你分析，或直接在下方输入问题。</p>
      </div>
    `;
  }

  public appendMessageView(message: ChatMessage): ChatMessageView {
    // If the empty state is currently displayed, drop it.
    const empty = this.listEl.querySelector(".chat-empty");
    if (empty) empty.remove();

    const view = new ChatMessageView(message, false);
    this.listEl.appendChild(view.element);
    this.views.set(message.id, view);
    return view;
  }

  /** Hub method: retrieve an existing rendered view so controllers can
   *  patch its content without duplicating the DOM node. */
  public getMessageView(messageId: string): ChatMessageView | undefined {
    return this.views.get(messageId);
  }

  /** Hub method: dim a previously-rendered message bubble (used by
   *  BranchController when hiding originals in favor of a new branch). */
  public hideMessageView(messageId: string): void {
    const v = this.views.get(messageId);
    if (v) v.element.classList.add("msg-hidden-branch");
  }

  /** Hub method: append an arbitrary element (e.g. branch divider) to the
   *  chat list without registering it as a message view. */
  public appendListChild(el: HTMLElement): void {
    const empty = this.listEl.querySelector(".chat-empty");
    if (empty) empty.remove();
    this.listEl.appendChild(el);
  }

  /* ---------------- ChatControllerHub contract ----------------
   * These thin wrappers expose the hub's required surface area. Each
   * delegates to either an existing private method or the deps bag. */

  public getConfig(): DeepSeekConfig {
    return this.deps.getConfig();
  }

  public upsertSession(session: ChatSession): void {
    this.deps.upsertSession(session);
  }

  public scrollToBottom(): void {
    requestAnimationFrame(() => {
      this.listEl.scrollTop = this.listEl.scrollHeight;
    });
  }

  public setStatus(text: string): void {
    this.statusBar.textContent = text;
    this.statusBar.style.display = text ? "block" : "none";
  }

  private showError(message: string): void {
    this.setStatus("⚠ " + message);
    this.toast(message, "error");
  }

  private setStreamingState(streaming: boolean): void {
    this.sendBtn.disabled = streaming;
    this.inputEl.disabled = streaming;
    // Hide quick actions except stop while streaming
    const quickBtns = this.quickEl.querySelectorAll<HTMLButtonElement>(".quick-action-btn");
    quickBtns.forEach((b) => {
      if (b.dataset.kind === "stop") {
        b.style.display = streaming ? "inline-flex" : "none";
      } else {
        b.style.display = streaming ? "none" : "inline-flex";
      }
    });
    // Update context bar so the user knows work is in progress.
    this.contextBar?.setStreaming(streaming);
  }

  private async onSendClicked(): Promise<void> {
    if (this.isStreaming()) return;
    const text = this.inputEl.value;
    if (!text.trim()) return;
    this.inputEl.value = "";
    this.inputEl.style.height = "auto";
    await this.sendUserMessage(text);
  }

  /* ---------------- session management ---------------- */

  private ensureActiveSession(): ChatSession {
    let id = this.deps.getActiveSessionId();
    let sessions = this.deps.getSessions();
    let session = id ? sessions.find((s) => s.id === id) : undefined;

    if (!session) {
      session = {
        id: generateId(),
        title: "新对话",
        createdAt: Date.now(),
        updatedAt: Date.now(),
        messages: [],
      };
      sessions.unshift(session);
      this.deps.setSessions(sessions);
      this.sessionCreatedHandler(session.id);
    }
    return session;
  }

  private getActiveSession(): ChatSession | null {
    const id = this.deps.getActiveSessionId();
    if (!id) return null;
    return this.deps.getSessions().find((s) => s.id === id) || null;
  }

  /* ---------------- streaming ---------------- */

  /** Hub contract: BranchController (and future controllers) need to kick
   *  off a fresh assistant stream against a (possibly truncated) session.
   *  Keep the implementation here — this is the actual streaming engine
   *  that talks to DeepSeek. The controllers sit on top. */
  public async streamAssistantReply(
    session: ChatSession,
    lastUserMessage: ChatMessage
  ): Promise<void> {
    const config = this.deps.getConfig();

    // Build messages: system (from config) + history (excluding the
    // current user message - the API will see it as the last user turn).
    const apiMessages: ChatMessage[] = [];
    // PRD-10: KnowledgeInjector owns the KB retrieval + size-cap + warning.
    // If it returns a system message we prepend it; otherwise we just push
    // the conversation history straight to the model.
    const systemMessage = this.knowledgeInjector.inject(
      lastUserMessage.content,
      config.systemPrompt || ""
    );
    if (systemMessage) {
      apiMessages.push(systemMessage);
    }
    // Push the session's messages up to and including the last user.
    const lastUserIdx = session.messages.findIndex((m) => m.id === lastUserMessage.id);
    for (let i = 0; i <= lastUserIdx; i++) {
      const m = session.messages[i];
      if (m.role === "system") continue;
      apiMessages.push(m);
    }
    // Snapshot the prompt text so UsageRecorder can tokenize it after the
    // stream finishes. We avoid holding onto the full apiMessages array
    // — the recorder only needs the joined content.
    accumulatedPromptChars = apiMessages.reduce(
      (n, m) => n + (m.content?.length || 0),
      0
    );
    promptText = apiMessages.map((m) => m.content || "").join("\n");

    // Pre-create the assistant bubble.
    const assistantMsg: ChatMessage = {
      id: generateId(),
      role: "assistant",
      content: "",
      timestamp: Date.now(),
    };
    session.messages.push(assistantMsg);
    this.deps.upsertSession(session);
    const view = this.appendMessageView(assistantMsg);
    this.currentStreamingMessageId = assistantMsg.id;
    this.setStreamingState(true);
    this.setStatus("AI 正在思考…");

    // Fresh dispatcher per stream turn so the round counter resets to 0.
    // Recursion into streamAssistantReply (from the dispatcher's tail) keeps
    // the dispatcher alive via closure.
    this.currentToolCallDispatcher = this.toolCallController.buildDispatcher();

    let accumulated = "";
    let accumulatedPromptChars = 0;
    let promptText = "";

    try {
      this.currentController = await chatCompletionStream(
        apiMessages,
        config,
        (_chunk, full) => {
          accumulated = full;
          view.update({ ...assistantMsg, content: full }, true);
          session.messages[session.messages.length - 1] = { ...assistantMsg, content: full };
          this.scrollToBottom();
          this.setStatus(`正在接收回复… (${full.length} 字)`);
        },
        async (full) => {
          const finalText = full || accumulated;
          assistantMsg.content = finalText;
          view.update({ ...assistantMsg, content: finalText }, false);
          session.messages[session.messages.length - 1] = { ...assistantMsg, content: finalText };
          session.updatedAt = Date.now();
          this.deps.upsertSession(session);
          this.setStatus("");
          // PRD-12: record successful request via UsageRecorder — it owns the
          // CJK/Latin token-estimate heuristic.
          this.usageRecorder.recordApiCall({
            model: config.model,
            promptText,
            finalText,
          });
        },
        async (calls) => {
          // PRD-09: model requested tool invocations. Delegate to
          // ToolCallController — it owns the execution loop, recursion
          // guard, usage recording, and follow-up stream kickoff.
          await this.currentToolCallDispatcher!.handle(calls, assistantMsg, session, lastUserMessage);
        },
        { withTools: true },
        (err) => {
          const msg = describeApiError(err);
          assistantMsg.content = `⚠️ ${msg}\n\n请检查网络、API Key 或模型配置。`;
          view.update({ ...assistantMsg }, false);
          session.messages[session.messages.length - 1] = { ...assistantMsg };
          this.deps.upsertSession(session);
          this.setStatus("");
          this.toast(msg, "error");
          this.usageRecorder.recordApiCall({
            model: config.model,
            promptText,
            finalText: "",
            errored: true,
          });
        }
      );
    } catch (err: any) {
      this.showError(err?.message || "请求失败");
    } finally {
      this.currentController = null;
      this.currentStreamingMessageId = null;
      this.setStreamingState(false);
    }
  }

  private async insertLastReplyIntoCell(): Promise<void> {
    // The last assistant message is the most recent reply.
    const sessions = this.deps.getSessions();
    const id = this.deps.getActiveSessionId();
    const session = id ? sessions.find((s) => s.id === id) : null;
    if (!session) {
      this.toast("当前没有活跃对话", "error");
      return;
    }
    const lastAssistant = [...session.messages].reverse().find((m) => m.role === "assistant");
    if (!lastAssistant) {
      this.toast("当前对话还没有 AI 回复", "error");
      return;
    }
    // Smart insert: detect formula and insert as formula if present.
    const formula = detectFormula(lastAssistant.content);
    try {
      if (formula) {
        await writeFormula(formula);
        this.toast("已识别为公式，插入为活动公式", "success");
      } else {
        const text = markdownToPlainText(lastAssistant.content);
        await insertTextToCell(text);
        this.toast("已插入到当前单元格", "success");
      }
    } catch (err: any) {
      this.showError(err?.message || "插入失败");
    }
  }
}