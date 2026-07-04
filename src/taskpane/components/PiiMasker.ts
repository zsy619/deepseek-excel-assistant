/**
 * ============================================================================
 * PiiMasker (PRD-06)
 * ----------------------------------------------------------------------------
 * Card list for "selected cells that look sensitive". Each row shows the
 * kind label, original snippet, fake replacement, and a checkbox. The host
 * aggregates the chosen replacements and calls `pii-masker-apply`.
 * ============================================================================
 */

import type { PiiHit, PiiKind } from "../services/excel";
import { fakeFor } from "../services/excel";

const KIND_LABELS: Record<PiiKind, string> = {
  phone_cn: "手机号",
  phone_intl: "国际电话",
  email: "邮箱",
  id_card_cn: "身份证",
  bank_card: "银行卡",
  ip: "IP地址",
  name_cn: "中文姓名",
  address_cn: "地址",
  credit_card: "信用卡",
};

export interface PiiMaskerApplyDetail {
  updates: Array<{ address: string; value: string }>;
}

export class PiiMaskerView {
  private root: HTMLDivElement;
  private listEl!: HTMLDivElement;
  private countEl!: HTMLElement;
  private applyBtn!: HTMLButtonElement;
  private selectAllBtn!: HTMLButtonElement;
  private hits: PiiHit[] = [];

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "pii-masker";
    this.root.style.display = "none";
    this.render();
  }

  get element(): HTMLDivElement {
    return this.root;
  }

  show(sheet: string, hits: PiiHit[]): void {
    this.hits = hits;
    this.root.style.display = "block";
    this.countEl.textContent = `${sheet} · 共 ${hits.length} 个疑似敏感值`;
    this.applyBtn.disabled = hits.length === 0;
    this.renderList();
  }

  hide(): void {
    this.root.style.display = "none";
  }

  private render(): void {
    this.root.innerHTML = `
      <div class="pii-masker__header">
        <span class="pii-masker__title">🛡️ 数据脱敏</span>
        <span class="pii-masker__count" data-ref="count"></span>
      </div>
      <div class="pii-masker__actions">
        <button type="button" class="pii-masker__btn pii-masker__btn--ghost" data-ref="selectAll">全选</button>
        <button type="button" class="pii-masker__btn" data-ref="apply">应用所选</button>
      </div>
      <div class="pii-masker__list" data-ref="list"></div>
    `;
    this.listEl = this.root.querySelector<HTMLDivElement>("[data-ref=list]")!;
    this.countEl = this.root.querySelector<HTMLElement>("[data-ref=count]")!;
    this.applyBtn = this.root.querySelector<HTMLButtonElement>("[data-ref=apply]")!;
    this.selectAllBtn = this.root.querySelector<HTMLButtonElement>("[data-ref=selectAll]")!;

    this.applyBtn.addEventListener("click", () => this.fireApply());
    this.selectAllBtn.addEventListener("click", () => {
      const checks = this.listEl.querySelectorAll<HTMLInputElement>("input[type=checkbox]");
      const allChecked = Array.from(checks).every((c) => c.checked);
      checks.forEach((c) => (c.checked = !allChecked));
      this.refreshApplyState();
    });
  }

  private renderList(): void {
    if (this.hits.length === 0) {
      this.listEl.innerHTML =
        '<div class="pii-masker__empty">未在选区中发现疑似敏感数据 🎉</div>';
      return;
    }
    this.listEl.innerHTML = this.hits
      .map((h, idx) => {
        const fake = fakeFor(h.kind, h.original, idx + 1);
        return (
          `<label class="pii-masker__row" data-idx="${idx}">` +
            `<input type="checkbox" checked data-idx="${idx}" />` +
            `<span class="pii-masker__addr">${escapeHtml(h.address)}</span>` +
            `<span class="pii-masker__kind">${escapeHtml(KIND_LABELS[h.kind] || h.kind)}</span>` +
            `<span class="pii-masker__from">${escapeHtml(truncate(h.original, 24))}</span>` +
            `<span class="pii-masker__arrow">→</span>` +
            `<span class="pii-masker__to">${escapeHtml(fake)}</span>` +
          `</label>`
        );
      })
      .join("");

    this.listEl.querySelectorAll<HTMLInputElement>("input[type=checkbox]").forEach((cb) => {
      cb.addEventListener("change", () => this.refreshApplyState());
    });
    this.refreshApplyState();
  }

  private refreshApplyState(): void {
    const checks = this.listEl.querySelectorAll<HTMLInputElement>("input[type=checkbox]");
    const count = Array.from(checks).filter((c) => c.checked).length;
    this.applyBtn.textContent = `应用所选 (${count})`;
    this.applyBtn.disabled = count === 0;
  }

  private fireApply(): void {
    const checks = this.listEl.querySelectorAll<HTMLInputElement>("input[type=checkbox]");
    const updates: Array<{ address: string; value: string }> = [];
    checks.forEach((cb) => {
      if (!cb.checked) return;
      const idx = Number(cb.dataset.idx ?? "0");
      const hit = this.hits[idx];
      if (!hit) return;
      updates.push({
        address: hit.address,
        value: fakeFor(hit.kind, hit.original, idx + 1),
      });
    });
    const detail: PiiMaskerApplyDetail = { updates };
    this.root.dispatchEvent(
      new CustomEvent("pii-masker-apply", { detail, bubbles: true })
    );
  }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function truncate(s: string, n: number): string {
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}