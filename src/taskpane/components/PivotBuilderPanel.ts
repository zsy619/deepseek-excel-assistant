/**
 * ============================================================================
 * PivotBuilderPanel
 * ----------------------------------------------------------------------------
 * Two-step pattern for the "AI 透视表" ribbon button.
 *   1. AI inspects the selection headers and proposes a pivot layout:
 *        - rows:  string[] (categorical columns)
 *        - columns: string[] (categorical columns)
 *        - values: { column, aggregation: 'sum'|'count'|'avg' }[]
 *   2. User reviews the preview (rendered as a small mock table) and clicks
 *      "Apply" to call Office's PivotTable API and create the pivot.
 *
 * The panel is empty until a payload arrives.
 * ============================================================================
 */

export interface PivotSpec {
  rows: string[];
  columns: string[];
  values: Array<{ column: string; aggregation: "sum" | "count" | "avg" }>;
}

export interface PivotPayload {
  selection?: string;
  spec?: PivotSpec;
  rationale?: string;
}

export class PivotBuilderPanel {
  public root: HTMLElement;
  private currentPayload: PivotPayload | null = null;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "pivot-panel";
    this.root.dataset.panelKind = "pivot";
    this.renderEmpty();
  }

  public getCurrentPayload(): PivotPayload | null {
    return this.currentPayload;
  }

  private renderEmpty(): void {
    this.root.innerHTML =
      '<div class="pivot-panel__empty">' +
      '  <div class="pivot-panel__title">🧩 AI 透视表</div>' +
      '  <div class="pivot-panel__hint">选中包含表头的数据,AI 会推荐行 / 列 / 值字段并预览。</div>' +
      '</div>';
  }

  public show(payload: PivotPayload): void {
    this.currentPayload = payload;
    const spec = payload.spec;
    if (!spec) {
      this.renderEmpty();
      return;
    }
    let html =
      '<div class="pivot-panel__header">' +
      `  <span class="pivot-panel__title">🧩 AI 透视表</span>` +
      (payload.selection
        ? ` <span class="pivot-panel__meta">${this.escape(payload.selection)}</span>`
        : "") +
      '</div>';
    if (payload.rationale) {
      html += `<div class="pivot-panel__rationale">${this.escape(payload.rationale)}</div>`;
    }
    // Fields preview
    html +=
      '<div class="pivot-panel__fields">' +
      `  <div class="pivot-field pivot-field--rows"><span class="pivot-field__label">行</span><span class="pivot-field__value">${this.list(spec.rows)}</span></div>` +
      `  <div class="pivot-field pivot-field--cols"><span class="pivot-field__label">列</span><span class="pivot-field__value">${this.list(spec.columns)}</span></div>` +
      `  <div class="pivot-field pivot-field--vals"><span class="pivot-field__label">值</span><span class="pivot-field__value">${this.values(spec.values)}</span></div>` +
      '</div>';
    // Mock preview
    html += '<div class="pivot-panel__preview-title">预览 (模拟布局)</div>';
    html += this.renderMockPreview(spec);
    html +=
      '<div class="pivot-panel__footer">' +
      '  <button class="pivot-panel__btn pivot-panel__btn--apply" type="button">✓ 创建透视表</button>' +
      '  <button class="pivot-panel__btn pivot-panel__btn--copy" type="button">复制配置 JSON</button>' +
      '</div>';
    this.root.innerHTML = html;
  }

  private renderMockPreview(spec: PivotSpec): string {
    const rows = spec.rows.length ? spec.rows : ["(全部)"];
    const cols = spec.columns.length ? spec.columns : ["汇总"];
    const vals = spec.values.length ? spec.values : [{ column: "(计数)", aggregation: "count" as const }];
    let h = '<div class="pivot-preview"><table><thead><tr><th></th>';
    for (const c of cols) h += `<th>${this.escape(c)}</th>`;
    h += "</tr></thead><tbody>";
    for (const r of rows.slice(0, 6)) {
      h += `<tr><th>${this.escape(r)}</th>`;
      for (const _c of cols) {
        const v = vals[0];
        h += `<td>${v.aggregation.toUpperCase()}(${this.escape(v.column)})</td>`;
      }
      h += "</tr>";
    }
    h += "</tbody></table></div>";
    return h;
  }

  public hide(): void {
    this.root.style.display = "none";
  }

  public reveal(): void {
    this.root.style.display = "";
  }

  private list(arr: string[]): string {
    return arr.length === 0 ? "<em>(无)</em>" : arr.map((s) => this.escape(s)).join(" · ");
  }

  private values(arr: PivotSpec["values"]): string {
    return arr.length === 0
      ? "<em>(无)</em>"
      : arr.map((v) => `${v.aggregation.toUpperCase()}(${this.escape(v.column)})`).join(" · ");
  }

  private escape(s: string): string {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c] as string);
  }
}