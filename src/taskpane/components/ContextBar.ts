/**
 * ============================================================================
 * ContextBar
 * ----------------------------------------------------------------------------
 * Slim bar shown directly under the chat header. It surfaces the current
 * Excel selection so the user always knows what context the panel sees:
 *
 *   - Sheet name        (Sheet1)
 *   - Selection address (A1:D10)
 *   - Row/column count  (4 × 10)
 *   - Cell formula      (only when exactly one cell is selected and it has
 *                        a formula - useful for "explain formula")
 *   - Active model      (small badge)
 *
 * The bar updates live via the excel-events service.
 * ============================================================================
 */

import type { ExcelSelection } from "../types";
import type { DeepSeekModel } from "../types";

export interface ContextBarState {
  sheetName: string | null;
  selection: ExcelSelection | null;
  model: DeepSeekModel | null;
  streaming: boolean;
}

export class ContextBar {
  private root: HTMLElement;
  private state: ContextBarState = {
    sheetName: null,
    selection: null,
    model: null,
    streaming: false,
  };

  constructor(parent: HTMLElement) {
    this.root = document.createElement("div");
    this.root.classList.add("context-bar");
    this.root.innerHTML = this.renderShell();
    parent.appendChild(this.root);
    this.repaint();
  }

  /** Replace the entire state object and re-render. */
  public setState(state: Partial<ContextBarState>): void {
    this.state = { ...this.state, ...state };
    this.repaint();
  }

  /** Convenience setters for the common fields. */
  public setSheet(name: string | null): void {
    this.setState({ sheetName: name });
  }

  public setSelection(sel: ExcelSelection | null): void {
    this.setState({ selection: sel });
  }

  public setModel(model: DeepSeekModel | null): void {
    this.setState({ model });
  }

  public setStreaming(streaming: boolean): void {
    this.setState({ streaming });
  }

  /* ---------------- rendering ---------------- */

  private renderShell(): string {
    return `
      <div class="context-bar-inner">
        <span class="ctx-item ctx-sheet" data-ref="sheet">
          <span class="ctx-icon">📄</span>
          <span class="ctx-text">—</span>
        </span>
        <span class="ctx-divider"></span>
        <span class="ctx-item ctx-selection" data-ref="selection">
          <span class="ctx-icon">🎯</span>
          <span class="ctx-text">未选择</span>
        </span>
        <span class="ctx-divider"></span>
        <span class="ctx-item ctx-formula" data-ref="formula" style="display:none">
          <span class="ctx-icon">ƒ</span>
          <span class="ctx-text"></span>
        </span>
        <span class="ctx-divider ctx-divider-formula" style="display:none"></span>
        <span class="ctx-item ctx-model" data-ref="model">
          <span class="ctx-badge"></span>
        </span>
      </div>
    `;
  }

  private repaint(): void {
    const { sheetName, selection, model, streaming } = this.state;

    // Sheet
    const sheetEl = this.root.querySelector<HTMLElement>('[data-ref="sheet"] .ctx-text');
    if (sheetEl) sheetEl.textContent = sheetName || "—";

    // Selection
    const selEl = this.root.querySelector<HTMLElement>('[data-ref="selection"]');
    const selText = selEl?.querySelector(".ctx-text");
    if (selText) {
      if (!selection) {
        selText.textContent = "未选择";
      } else {
        const dims =
          selection.rowCount === 1 && selection.columnCount === 1
            ? ""
            : ` (${selection.rowCount}×${selection.columnCount})`;
        selText.textContent = `${selection.address}${dims}`;
      }
    }
    if (selEl) {
      selEl.classList.toggle("ctx-active", !!selection);
    }

    // Formula (only show when single cell with a non-empty formula)
    const formulaEl = this.root.querySelector<HTMLElement>('[data-ref="formula"]');
    const divider = this.root.querySelector<HTMLElement>(".ctx-divider-formula");
    if (formulaEl && divider) {
      let formula: string | null = null;
      if (
        selection &&
        selection.formulas &&
        selection.rowCount === 1 &&
        selection.columnCount === 1
      ) {
        const f = selection.formulas[0]?.[0];
        if (f && f.startsWith("=")) formula = f;
      }
      if (formula) {
        formulaEl.style.display = "";
        divider.style.display = "";
        formulaEl.classList.add("ctx-active");
        const text = formulaEl.querySelector(".ctx-text");
        if (text) text.textContent = formula.length > 30 ? formula.slice(0, 30) + "…" : formula;
        formulaEl.title = formula;
      } else {
        formulaEl.style.display = "none";
        divider.style.display = "none";
      }
    }

    // Model badge
    const modelEl = this.root.querySelector<HTMLElement>('[data-ref="model"] .ctx-badge');
    if (modelEl) {
      if (streaming) {
        modelEl.textContent = "⚡ 生成中…";
        modelEl.className = "ctx-badge ctx-badge-streaming";
      } else if (model) {
        const labels: Record<DeepSeekModel, string> = {
          "deepseek-chat": "🤖 Chat",
          "deepseek-reasoner": "🧠 Reasoner",
        };
        modelEl.textContent = labels[model] || model;
        modelEl.className = "ctx-badge";
      } else {
        modelEl.textContent = "未配置";
        modelEl.className = "ctx-badge ctx-badge-warn";
      }
    }
  }
}