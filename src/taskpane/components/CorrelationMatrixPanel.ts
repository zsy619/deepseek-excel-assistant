/**
 * ============================================================================
 * CorrelationMatrixPanel
 * ----------------------------------------------------------------------------
 * Renders an NxN correlation heat-map for the selected numeric columns.
 *
 * Two-step pattern:
 *   1. AI computes Pearson r for every column pair (client-side via
 *      services/correlation.ts OR a streamed completion that returns JSON).
 *   2. User reviews the heat-map and clicks "Apply" to insert the matrix
 *      + correlation table starting at the next free cell.
 *
 * When no selection is supplied we still show an empty-state with a hint
 * so the ribbon → panel flow looks intentional.
 * ============================================================================
 */

export interface CorrelationCell {
  row: number;
  col: number;
  value: number; // -1 .. +1
  rowLabel?: string;
  colLabel?: string;
}

export interface CorrelationMatrix {
  /** Square matrix, length === labels.length */
  cells: CorrelationCell[];
  labels: string[];
}

export interface CorrelationMatrixPayload {
  /** Optional selection summary for the user-visible header. */
  selection?: string;
  /** Computed matrix; undefined → empty state. */
  matrix?: CorrelationMatrix;
  /** Free-form AI comment shown above the matrix. */
  reasoning?: string;
}

export class CorrelationMatrixPanel {
  public root: HTMLElement;
  /** Last payload passed to show(); used by the Apply/Copy handlers. */
  private currentPayload: CorrelationMatrixPayload | null = null;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "corr-panel";
    this.root.dataset.panelKind = "correlation";
    this.renderEmpty();
  }

  /** Return the last payload (for the Apply/Copy wiring). */
  public getCurrentPayload(): CorrelationMatrixPayload | null {
    return this.currentPayload;
  }

  private renderEmpty(): void {
    this.root.innerHTML =
      '<div class="corr-panel__empty">' +
      '  <div class="corr-panel__title">📊 相关性矩阵</div>' +
      '  <div class="corr-panel__hint">选中 ≥ 2 列数字数据后,此面板会显示 AI 计算的 Pearson 相关系数热力图。</div>' +
      '</div>';
  }

  public show(payload: CorrelationMatrixPayload): void {
    this.currentPayload = payload;
    const matrix = payload.matrix;
    if (!matrix || matrix.cells.length === 0 || matrix.labels.length === 0) {
      this.renderEmpty();
      return;
    }
    const labels = matrix.labels;
    const n = labels.length;
    // Header row + N data rows.
    let html =
      '<div class="corr-panel__header">' +
      `  <span class="corr-panel__title">📊 相关性矩阵</span>` +
      (payload.selection
        ? ` <span class="corr-panel__meta">${this.escape(payload.selection)}</span>`
        : "") +
      '</div>';
    if (payload.reasoning) {
      html += `<div class="corr-panel__reasoning">${this.escape(payload.reasoning)}</div>`;
    }
    html += '<div class="corr-panel__matrix" style="--cols:' + n + '">';
    // header
    html += '<div class="corr-cell corr-cell--head"></div>';
    for (const label of labels) {
      html += `<div class="corr-cell corr-cell--head">${this.escape(label)}</div>`;
    }
    // rows
    for (let r = 0; r < n; r++) {
      html += `<div class="corr-cell corr-cell--head">${this.escape(labels[r])}</div>`;
      for (let c = 0; c < n; c++) {
        const cell = matrix.cells.find((x) => x.row === r && x.col === c);
        const v = cell ? cell.value : 0;
        const intensity = Math.abs(v);
        const hue = v >= 0 ? "120" : "0"; // green vs red
        const bg = `hsla(${hue}, 70%, 50%, ${0.10 + intensity * 0.55})`;
        const text = intensity > 0.55 ? "#fff" : "#1f2937";
        html += `<div class="corr-cell" style="background:${bg};color:${text}">${v.toFixed(2)}</div>`;
      }
    }
    html += "</div>";
    html +=
      '<div class="corr-panel__footer">' +
      '  <button class="corr-panel__btn corr-panel__btn--apply" type="button">✓ 插入矩阵</button>' +
      '  <button class="corr-panel__btn corr-panel__btn--copy" type="button">复制到剪贴板</button>' +
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