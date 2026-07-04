/**
 * ============================================================================
 * ShortcutHelp (PRD-03)
 * ----------------------------------------------------------------------------
 * Modal dialog listing all SHORTCUTS. Triggered by ⌘+/ (or by clicking the
 * "shortcuts" item in the command palette). Emits no events - just a
 * passive display.
 * ============================================================================
 */

import { SHORTCUTS } from "../utils/constants";

export class ShortcutHelpView {
  private root: HTMLDivElement;
  private onClose: () => void;

  constructor(onClose: () => void = () => {}) {
    this.onClose = onClose;
    this.root = document.createElement("div");
    this.root.className = "shortcut-help";
    this.root.setAttribute("role", "dialog");
    this.root.setAttribute("aria-label", "快捷键帮助");
    this.root.style.display = "none";
    this.render();
  }

  get element(): HTMLDivElement {
    return this.root;
  }

  show(): void {
    this.root.style.display = "flex";
  }

  hide(): void {
    this.root.style.display = "none";
  }

  private render(): void {
    const globalRows = SHORTCUTS.filter((s) => s.scope === "global")
      .map(rowHtml).join("");
    const inputRows = SHORTCUTS.filter((s) => s.scope === "input")
      .map(rowHtml).join("");

    this.root.innerHTML = `
      <div class="shortcut-help__backdrop" data-ref="backdrop"></div>
      <div class="shortcut-help__panel" role="document">
        <div class="shortcut-help__header">
          <h3>快捷键</h3>
          <button type="button" class="shortcut-help__close" data-ref="close" aria-label="关闭">✕</button>
        </div>
        <div class="shortcut-help__body">
          <section>
            <h4>全局</h4>
            ${globalRows}
          </section>
          <section>
            <h4>输入框</h4>
            ${inputRows}
          </section>
        </div>
        <div class="shortcut-help__footer">按 Esc 关闭</div>
      </div>
    `;

    this.root.querySelector<HTMLDivElement>('[data-ref="backdrop"]')!
      .addEventListener("click", () => { this.hide(); this.onClose(); });
    this.root.querySelector<HTMLButtonElement>('[data-ref="close"]')!
      .addEventListener("click", () => { this.hide(); this.onClose(); });
  }
}

function rowHtml(s: { combo: string; label: string }): string {
  return (
    `<div class="shortcut-help__row">` +
      `<div class="shortcut-help__keys">${escapeHtml(s.combo)}</div>` +
      `<div class="shortcut-help__label">${escapeHtml(s.label)}</div>` +
    `</div>`
  );
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}