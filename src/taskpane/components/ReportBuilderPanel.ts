/**
 * ============================================================================
 * ReportBuilderPanel
 * ----------------------------------------------------------------------------
 * Two-step pattern for the "快速报告" ribbon button.
 *   1. AI summarises the selection: header, key findings, suggestions.
 *   2. User clicks "Apply" to write the report into a NEW worksheet.
 *
 * The panel renders a markdown-style preview so the user sees what will be
 * inserted before any cell is touched.
 * ============================================================================
 */

export interface ReportSection {
  heading: string;
  body: string;
}

export interface ReportPayload {
  selection?: string;
  title?: string;
  summary?: string;
  sections?: ReportSection[];
  recommendations?: string[];
}

export class ReportBuilderPanel {
  public root: HTMLElement;
  private currentPayload: ReportPayload | null = null;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "report-panel";
    this.root.dataset.panelKind = "report";
    this.renderEmpty();
  }

  public getCurrentPayload(): ReportPayload | null {
    return this.currentPayload;
  }

  private renderEmpty(): void {
    this.root.innerHTML =
      '<div class="report-panel__empty">' +
      '  <div class="report-panel__title">📝 快速报告</div>' +
      '  <div class="report-panel__hint">选中数据,AI 会生成摘要 + 关键发现 + 建议,点击 Apply 后写入新 sheet。</div>' +
      '</div>';
  }

  public show(payload: ReportPayload): void {
    this.currentPayload = payload;
    if (!payload.title && !payload.summary && (!payload.sections || payload.sections.length === 0)) {
      this.renderEmpty();
      return;
    }
    let html =
      '<div class="report-panel__header">' +
      `  <span class="report-panel__title">📝 ${this.escape(payload.title || "数据报告")}</span>` +
      (payload.selection
        ? ` <span class="report-panel__meta">来源: ${this.escape(payload.selection)}</span>`
        : "") +
      '</div>';
    if (payload.summary) {
      html += `<div class="report-panel__summary">${this.escape(payload.summary)}</div>`;
    }
    if (payload.sections && payload.sections.length) {
      html += '<div class="report-panel__sections">';
      for (const s of payload.sections) {
        html += `<div class="report-section"><h4>${this.escape(s.heading)}</h4><div class="report-section__body">${this.escape(s.body)}</div></div>`;
      }
      html += "</div>";
    }
    if (payload.recommendations && payload.recommendations.length) {
      html += '<div class="report-panel__recs"><h4>建议</h4><ul>';
      for (const r of payload.recommendations) {
        html += `<li>${this.escape(r)}</li>`;
      }
      html += "</ul></div>";
    }
    html +=
      '<div class="report-panel__footer">' +
      '  <button class="report-panel__btn report-panel__btn--apply" type="button">✓ 写入新工作表</button>' +
      '  <button class="report-panel__btn report-panel__btn--copy" type="button">复制 Markdown</button>' +
      '</div>';
    this.root.innerHTML = html;
  }

  public hide(): void {
    this.root.style.display = "none";
  }

  public reveal(): void {
    this.root.style.display = "";
  }

  private escape(s: string): string {
    return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c] as string);
  }
}