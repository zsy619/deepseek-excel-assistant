/**
 * ============================================================================
 * ChartPicker (PRD-05)
 * ----------------------------------------------------------------------------
 * Renders up to 3 chart recommendation cards. Click a card to fire
 * `chart-picker-insert` with the chosen chart type + title. Emits
 * `chart-picker-error` if the AI returned nothing usable and the host
 * decides to fall back to local heuristics.
 * ============================================================================
 */

import type { ChartRecommendation } from "../services/deepseek";
import { CHART_TYPE_INFO, type ExcelChartType } from "../types";

export interface ChartPickerInsertDetail {
  type: ExcelChartType;
  title: string;
}

export class ChartPickerView {
  private root: HTMLDivElement;
  private listEl!: HTMLDivElement;
  private statusEl!: HTMLElement;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "chart-picker";
    this.root.style.display = "none";
    this.render();
  }

  get element(): HTMLDivElement {
    return this.root;
  }

  show(status = "推荐中…"): void {
    this.root.style.display = "block";
    this.setStatus(status);
  }

  hide(): void {
    this.root.style.display = "none";
  }

  setRecommendations(list: ChartRecommendation[], usedFallback: boolean): void {
    if (!list || list.length === 0) {
      this.listEl.innerHTML =
        '<div class="chart-picker__empty">没有可用的图表推荐</div>';
      return;
    }
    this.listEl.innerHTML = list
      .map((r, idx) => {
        const meta = CHART_TYPE_INFO[r.type] ?? { label: r.type, icon: "📊" };
        return (
          `<button type="button" class="chart-picker__card" data-index="${idx}" data-type="${escapeAttr(r.type)}">` +
            `<div class="chart-picker__icon">${meta.icon}</div>` +
            `<div class="chart-picker__body">` +
              `<div class="chart-picker__type">${escapeHtml(meta.label)}</div>` +
              `<div class="chart-picker__title">${escapeHtml(r.title)}</div>` +
              `<div class="chart-picker__reason">${escapeHtml(r.reason)}</div>` +
            `</div>` +
            `<div class="chart-picker__cta">插入 →</div>` +
          `</button>`
        );
      })
      .join("");

    this.listEl.querySelectorAll<HTMLButtonElement>(".chart-picker__card").forEach((el) => {
      el.addEventListener("click", () => {
        const idx = Number(el.dataset.index ?? "0");
        const item = list[idx];
        if (!item) return;
        const detail: ChartPickerInsertDetail = {
          type: item.type,
          title: item.title,
        };
        this.root.dispatchEvent(
          new CustomEvent("chart-picker-insert", { detail, bubbles: true })
        );
      });
    });

    this.setStatus(usedFallback ? "本地推荐（AI 未响应）" : "就绪");
  }

  setStatus(text: string): void {
    if (this.statusEl) this.statusEl.textContent = text;
  }

  private render(): void {
    this.root.innerHTML = `
      <div class="chart-picker__header">
        <span>📊 AI 推荐的图表</span>
        <span class="chart-picker__status" data-ref="status"></span>
      </div>
      <div class="chart-picker__list" data-ref="list"></div>
    `;
    this.listEl = this.root.querySelector<HTMLDivElement>("[data-ref=list]")!;
    this.statusEl = this.root.querySelector<HTMLElement>("[data-ref=status]")!;
  }
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
function escapeAttr(s: string): string {
  return escapeHtml(s);
}