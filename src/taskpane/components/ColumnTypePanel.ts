/**
 * ============================================================================
 * ColumnTypePanel
 * ----------------------------------------------------------------------------
 * Two-step pattern for the "列类型" ribbon button.
 *   1. AI inspects the first ~50 rows of each column and proposes a type
 *      (number / date / categorical / text / boolean / currency / pct).
 *   2. User reviews the proposals and clicks "Apply" to:
 *        - apply header bold + autofilter
 *        - tag each column with a type-formatted number format
 *        - create named ranges per column
 *
 * Empty-state is shown until the payload arrives.
 * ============================================================================
 */

export type ColumnType =
  | "number"
  | "currency"
  | "percent"
  | "date"
  | "boolean"
  | "categorical"
  | "text"
  | "unknown";

export interface ColumnTypeRow {
  column: string;
  proposedType: ColumnType;
  confidence: number; // 0..1
  format?: string; // Excel number-format code
  namedRange?: string;
  reason?: string;
}

export interface ColumnTypePayload {
  selection?: string;
  rows?: ColumnTypeRow[];
  summary?: string;
}

export class ColumnTypePanel {
  public root: HTMLElement;
  private currentPayload: ColumnTypePayload | null = null;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "coltype-panel";
    this.root.dataset.panelKind = "coltype";
    this.renderEmpty();
  }

  public getCurrentPayload(): ColumnTypePayload | null {
    return this.currentPayload;
  }

  private renderEmpty(): void {
    this.root.innerHTML =
      '<div class="coltype-panel__empty">' +
      '  <div class="coltype-panel__title">🏷️ 列类型推断</div>' +
      '  <div class="coltype-panel__hint">选中多列,AI 会根据样例数据推断每列类型(数字/日期/分类等),Apply 后加表头样式与命名范围。</div>' +
      '</div>';
  }

  public show(payload: ColumnTypePayload): void {
    this.currentPayload = payload;
    if (!payload.rows || payload.rows.length === 0) {
      this.renderEmpty();
      return;
    }
    let html =
      '<div class="coltype-panel__header">' +
      `  <span class="coltype-panel__title">🏷️ 列类型推断</span>` +
      `  <span class="coltype-panel__badge">${payload.rows.length} 列</span>` +
      (payload.selection
        ? ` <span class="coltype-panel__meta">${this.escape(payload.selection)}</span>`
        : "") +
      '</div>';
    if (payload.summary) {
      html += `<div class="coltype-panel__summary">${this.escape(payload.summary)}</div>`;
    }
    html += '<div class="coltype-panel__rows">';
    for (const r of payload.rows) {
      const conf = Math.round(Math.max(0, Math.min(1, r.confidence)) * 100);
      html +=
        '<div class="coltype-row">' +
        `  <span class="coltype-row__name">${this.escape(r.column)}</span>` +
        `  <span class="coltype-row__type coltype-row__type--${r.proposedType}">${this.typeLabel(r.proposedType)}</span>` +
        `  <span class="coltype-row__conf" title="置信度">${conf}%</span>` +
        (r.reason
          ? `  <span class="coltype-row__reason">${this.escape(r.reason)}</span>`
          : "") +
        '</div>';
    }
    html += "</div>";
    html +=
      '<div class="coltype-panel__footer">' +
      '  <button class="coltype-panel__btn coltype-panel__btn--apply" type="button">✓ 应用格式 + 命名范围</button>' +
      '  <button class="coltype-panel__btn coltype-panel__btn--copy" type="button">复制配置 JSON</button>' +
      '</div>';
    this.root.innerHTML = html;
  }

  public hide(): void {
    this.root.style.display = "none";
  }

  public reveal(): void {
    this.root.style.display = "";
  }

  private typeLabel(t: ColumnType): string {
    const map: Record<ColumnType, string> = {
      number: "数字",
      currency: "货币",
      percent: "百分比",
      date: "日期",
      boolean: "布尔",
      categorical: "分类",
      text: "文本",
      unknown: "未知",
    };
    return map[t] || t;
  }

  private escape(s: string): string {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c] as string);
  }
}