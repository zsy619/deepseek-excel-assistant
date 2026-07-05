/**
 * ============================================================================
 * RibbonFocusController
 * ----------------------------------------------------------------------------
 * When a ribbon button fires, the user expects the taskpane to visually
 * *do something*. This module handles two of those signals:
 *
 *   1. scrollToPanel(name)
 *      Smoothly scrolls the named panel into view. The chat list is the
 *      default for commands that operate on chat (e.g. analyze, generate
 *      formula). Specialized panels (ChartPicker, PiiMasker, CodeGenPanel,
 *      FormulaDiagnostics) are scrolled to when their command fires.
 *
 *   2. highlightPanel(name)
 *      Adds a temporary accent ring around the panel for 1.5s. Helps the
 *      user locate the panel if it's currently off-screen or small.
 *
 * Panel registry (data-panel attribute on the mount host):
 *   chat          → chat list (default fallback)
 *   chat-input    → chat input textarea (focused instead of scrolled)
 *   chartPicker   → ChartPickerView mount
 *   piiMasker     → PiiMaskerView mount
 *   codeGen       → CodeGenPanel mount
 *   formulaLibrary→ FormulaLibrary mount
 *   contextBar    → ContextBar (for selection-update commands)
 *
 * If a panel name doesn't resolve, this is a no-op (we don't throw).
 * ============================================================================
 */

export type PanelName =
  | "chat"
  | "chat-input"
  | "chartPicker"
  | "piiMasker"
  | "codeGen"
  | "formulaLibrary"
  | "formulaDiagnostics"
  | "contextBar"
  | "commandPalette"
  | "knowledgeBase"
  | "correlationMatrix"
  | "outlierPanel"
  | "pivotBuilder"
  | "reportBuilder"
  | "columnTypes";

const SELECTOR = "[data-panel]";
const HIGHLIGHT_MS = 1500;

/** Map a ribbon command id to the panel that should scroll/highlight. */
export function panelForCommand(commandId: string): PanelName | null {
  switch (commandId) {
    case "showTaskpane":
      return null; // No scroll - just opens taskpane.
    case "analyzeSelection":
      return "chat-input"; // Focus the input (Q3: prompt pre-filled).
    case "generateFormula":
      return "chat-input";
    case "cleanData":
      return "chat-input";
    case "diagnoseFormulas":
      return "formulaDiagnostics";
    case "insertChart":
      return "chartPicker";
    case "translateToCode":
      return "codeGen";
    case "maskPII":
      return "piiMasker";
    case "openKnowledgeBase":
      return "knowledgeBase";
    case "openSettings":
      return null; // Settings is a modal overlay - no scroll.
    case "shareSession":
      return null;
    case "usageDashboard":
      return null;
    case "commandPalette":
      return "commandPalette";
    case "multiSelectionAnalyze":
      return "chat-input";
    case "exportCurrentSession":
    case "clearCurrentChat":
    case "insertLastReply":
    case "toggleTheme":
      return "chat";
    case "correlationMatrix":
      return "correlationMatrix";
    case "detectOutliers":
      return "outlierPanel";
    case "createPivot":
      return "pivotBuilder";
    case "quickReport":
      return "reportBuilder";
    case "inferColumnTypes":
      return "columnTypes";
    default:
      return "chat";
  }
}

export class RibbonFocusController {
  /** Resolve a panel name to its DOM element. Returns null when not found. */
  public findPanel(name: PanelName): HTMLElement | null {
    if (name === "chat-input") {
      const ta = document.querySelector<HTMLTextAreaElement>(
        ".chat-input-field"
      );
      return ta;
    }
    const sel = `${SELECTOR}="${name}"`;
    return document.querySelector<HTMLElement>(sel);
  }

  /** Smooth-scroll a panel into view. No-op if panel not found. */
  public scrollToPanel(name: PanelName): void {
    if (name === "chat-input") {
      const ta = this.findPanel("chat-input");
      if (ta) {
        try {
          ta.focus({ preventScroll: true });
        } catch {
          /* noop */
        }
      }
      return;
    }
    const el = this.findPanel(name);
    if (!el) return;
    try {
      el.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
    } catch {
      /* old webview - fall back to instant */
      try {
        el.scrollIntoView();
      } catch {
        /* noop */
      }
    }
  }

  /** Add an accent highlight ring for HIGHLIGHT_MS then remove. */
  public highlightPanel(name: PanelName, ms: number = HIGHLIGHT_MS): void {
    const el = this.findPanel(name);
    if (!el) return;
    el.classList.add("ribbon-focus-ring");
    window.setTimeout(() => {
      el.classList.remove("ribbon-focus-ring");
    }, ms);
  }

  /** Convenience: scroll + highlight together. */
  public focusPanel(name: PanelName): void {
    this.scrollToPanel(name);
    // Slight delay so the highlight feels like a "ping" after the scroll.
    window.setTimeout(() => {
      this.highlightPanel(name);
    }, 120);
  }
}