/**
 * ============================================================================
 * UsageDashboard component (PRD-12)
 * ----------------------------------------------------------------------------
 * Modal panel summarizing recent usage:
 *   - today's totals (tokens, requests, cost)
 *   - bar chart of last 7 days
 *   - top tools / slash commands / features
 *
 * Emits no events; this is read-only summary data sourced from
 * services/usage.ts.
 * ============================================================================
 */

import { escapeHtml, formatRelativeTime } from "../utils/helpers";
import {
  getAllDays,
  getTodayStats,
  topEntries,
  estimateCost,
  totalCostAcrossDays,
  clearAllUsage,
  PRICE_TABLE,
  type UsageDay,
} from "../services/usage";

export class UsageDashboardView {
  public readonly element: HTMLElement;

  constructor() {
    this.element = document.createElement("div");
    this.element.classList.add("usage-panel");
    this.element.dataset.ref = "usageDashboard";
    this.element.hidden = true;
    this.render();
  }

  show(): void {
    this.element.hidden = false;
    this.refresh();
  }

  hide(): void {
    this.element.hidden = true;
  }

  toggle(): void {
    if (this.element.hidden) this.show();
    else this.hide();
  }

  private render(): void {
    this.element.innerHTML = `
      <div class="usage-panel__backdrop" data-action="close"></div>
      <div class="usage-panel__modal" role="dialog" aria-labelledby="usage-title">
        <header class="usage-panel__head">
          <h2 id="usage-title">📊 用量看板</h2>
          <button type="button" class="usage-panel__close" data-action="close" title="关闭">✕</button>
        </header>
        <div class="usage-panel__body">
          <section class="usage-panel__cards" data-ref="cards"></section>
          <section class="usage-panel__chart-section">
            <h3>最近 7 天</h3>
            <div class="usage-panel__chart" data-ref="chart"></div>
          </section>
          <section class="usage-panel__tables">
            <div class="usage-panel__table-block">
              <h3>工具调用 Top 5</h3>
              <div class="usage-panel__table" data-ref="tools"></div>
            </div>
            <div class="usage-panel__table-block">
              <h3>斜杠命令 Top 5</h3>
              <div class="usage-panel__table" data-ref="slashes"></div>
            </div>
            <div class="usage-panel__table-block">
              <h3>功能使用 Top 5</h3>
              <div class="usage-panel__table" data-ref="features"></div>
            </div>
          </section>
          <section class="usage-panel__pricing">
            <h3>计费参考 (USD / 1M tokens)</h3>
            <table class="usage-panel__price-table">
              <thead>
                <tr><th>模型</th><th>输入</th><th>输出</th></tr>
              </thead>
              <tbody data-ref="prices"></tbody>
            </table>
            <p class="usage-panel__hint">价格仅供参考；以 DeepSeek 官方计费页面为准。</p>
          </section>
          <footer class="usage-panel__footer">
            <button type="button" class="usage-panel__btn usage-panel__btn-danger" data-action="reset">🗑️ 重置全部用量记录</button>
          </footer>
        </div>
      </div>
    `;
    this.bind();
    this.refresh();
  }

  private bind(): void {
    this.element.querySelectorAll<HTMLElement>("[data-action='close']").forEach((el) => {
      el.onclick = () => this.hide();
    });
    this.element.querySelector<HTMLButtonElement>("[data-action='reset']")!.onclick = () => {
      if (!confirm("确定清空全部用量记录？此操作不可撤销。")) return;
      clearAllUsage();
      this.toast("已重置", "success");
      this.refresh();
    };
    this.element.addEventListener("keydown", (ev) => {
      if ((ev as KeyboardEvent).key === "Escape") this.hide();
    });
  }

  private refresh(): void {
    const today = getTodayStats();
    const days = getAllDays().slice(-7);

    // Top cards
    const totalCost = totalCostAcrossDays();
    const cards = [
      { label: "今日请求", value: String(today.requests) },
      { label: "今日 Tokens", value: today.totalTokens.toLocaleString("en-US") },
      { label: "今日预估成本", value: `$${estimateCost("deepseek-chat", today.promptTokens, today.completionTokens).toFixed(4)}` },
      { label: "累计成本 (30d)", value: `$${totalCost.toFixed(4)}` },
      { label: "今日错误数", value: String(today.errors) },
      { label: "知识库 / 文档数", value: "—" },
    ];
    this.element.querySelector<HTMLElement>("[data-ref='cards']")!.innerHTML = cards
      .map(
        (c) => `
        <div class="usage-card">
          <div class="usage-card__label">${escapeHtml(c.label)}</div>
          <div class="usage-card__value">${escapeHtml(c.value)}</div>
        </div>`
      )
      .join("");

    // Bar chart of last 7 days (tokens)
    const maxTokens = Math.max(1, ...days.map((d) => d.totalTokens));
    const chartEl = this.element.querySelector<HTMLElement>("[data-ref='chart']")!;
    if (!days.length) {
      chartEl.innerHTML = `<div class="usage-panel__empty">还没有用量数据，开始一次对话就会显示。</div>`;
    } else {
      chartEl.innerHTML = days
        .map((d) => {
          const pct = Math.max(2, Math.round((d.totalTokens / maxTokens) * 100));
          return `
          <div class="usage-bar-row">
            <div class="usage-bar-row__date">${escapeHtml(d.date.slice(5))}</div>
            <div class="usage-bar-row__track">
              <div class="usage-bar-row__fill" style="width:${pct}%"></div>
            </div>
            <div class="usage-bar-row__value">${d.totalTokens.toLocaleString("en-US")}</div>
          </div>`;
        })
        .join("");
    }

    // Top tables
    const renderTable = (ref: string, entries: Array<{ key: string; count: number }>, emptyMsg: string) => {
      const el = this.element.querySelector<HTMLElement>(`[data-ref='${ref}']`)!;
      if (!entries.length) {
        el.innerHTML = `<div class="usage-panel__empty">${escapeHtml(emptyMsg)}</div>`;
        return;
      }
      el.innerHTML = entries
        .map(
          (e) => `
        <div class="usage-row">
          <div class="usage-row__key">${escapeHtml(e.key)}</div>
          <div class="usage-row__count">${e.count}</div>
        </div>`
        )
        .join("");
    };
    renderTable("tools", topEntries(today.tools), "今日还未触发任何工具调用");
    renderTable("slashes", topEntries(today.slashCommands), "今日还未使用任何斜杠命令");
    renderTable("features", topEntries(today.features), "今日还未使用快捷功能");

    // Pricing table
    const priceRows = Object.entries(PRICE_TABLE)
      .map(
        ([model, p]) => `
      <tr>
        <td>${escapeHtml(model)}</td>
        <td>$${p.prompt.toFixed(2)}</td>
        <td>$${p.completion.toFixed(2)}</td>
      </tr>`
      )
      .join("");
    this.element.querySelector<HTMLElement>("[data-ref='prices']")!.innerHTML = priceRows;
  }

  private toast(msg: string, kind: "info" | "success" | "error"): void {
    this.element.dispatchEvent(
      new CustomEvent("usage-toast", { detail: { msg, kind }, bubbles: true })
    );
  }
}