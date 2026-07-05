/**
 * ============================================================================
 * OutlierPanel
 * ----------------------------------------------------------------------------
 * Two-step pattern for the "异常值" ribbon button.
 *   1. AI/IQR scan produces a list of outlier rows.
 *   2. User reviews the table and clicks one of two Apply actions:
 *        - 高亮  → mark cells red
 *        - 删除  → delete the rows
 *
 * Empty-state is shown until a payload arrives so the ribbon → panel
 * transition is visually grounded.
 * ============================================================================
 */

export interface OutlierRow {
  rowIndex: number; // 1-based Excel row index, for addressing.
  column: string;
  value: unknown;
  zScore?: number;
  reason: string;
}

export interface OutlierPayload {
  selection?: string;
  method?: "zscore" | "iqr" | "ai";
  outliers?: OutlierRow[];
  summary?: string;
}

export class OutlierPanel {
  public root: HTMLElement;
  private currentPayload: OutlierPayload | null = null;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "outlier-panel";
    this.root.dataset.panelKind = "outlier";
    this.renderEmpty();
  }

  public getCurrentPayload(): OutlierPayload | null {
    return this.currentPayload;
  }

  private renderEmpty(): void {
    this.root.innerHTML =
      '<div class="outlier-panel__empty">' +
      '  <div class="outlier-panel__title">🔍 异常值检测</div>' +
      '  <div class="outlier-panel__hint">选中数字列,此面板会用 Z-score / IQR 检测异常值并显示清单。</div>' +
      '</div>';
  }

  public show(payload: OutlierPayload): void {
    this.currentPayload = payload;
    if (!payload.outliers || payload.outliers.length === 0) {
      this.renderEmpty();
      return;
    }
    let html =
      '<div class="outlier-panel__header">' +
      `  <span class="outlier-panel__title">🔍 异常值检测</span>` +
      `  <span class="outlier-panel__badge">${payload.outliers.length} 个</span>` +
      (payload.method
        ? ` <span class="outlier-panel__meta">方法: ${this.methodLabel(payload.method)}</span>`
        : "") +
      '</div>';
    if (payload.summary) {
      html += `<div class="outlier-panel__summary">${this.escape(payload.summary)}</div>`;
    }
    html += '<div class="outlier-panel__rows">';
    for (const o of payload.outliers) {
      html +=
        '<div class="outlier-row">' +
        `  <span class="outlier-row__addr">行 ${o.rowIndex}</span>` +
        `  <span class="outlier-row__col">${this.escape(o.column)}</span>` +
        `  <span class="outlier-row__val">${this.escape(String(o.value))}</span>` +
        (typeof o.zScore === "number"
          ? `  <span class="outlier-row__z">z=${o.zScore.toFixed(2)}</span>`
          : "") +
        `  <span class="outlier-row__reason">${this.escape(o.reason)}</span>` +
        '</div>';
    }
    html += "</div>";
    html +=
      '<div class="outlier-panel__footer">' +
      '  <button class="outlier-panel__btn outlier-panel__btn--highlight" type="button">✓ 高亮(两步)</button>' +
      '  <button class="outlier-panel__btn outlier-panel__btn--delete" type="button">删除行</button>' +
      '</div>';
    this.root.innerHTML = html;
  }

  public hide(): void {
    this.root.style.display = "none";
  }

  public reveal(): void {
    this.root.style.display = "";
  }

  private methodLabel(m: "zscore" | "iqr" | "ai"): string {
    return m === "zscore" ? "Z-score" : m === "iqr" ? "IQR" : "AI";
  }

  private escape(s: string): string {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c] as string);
  }
}