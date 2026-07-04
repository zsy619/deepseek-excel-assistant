/**
 * ============================================================================
 * PromptMenu (PRD-02)
 * ----------------------------------------------------------------------------
 * Floating menu that appears below the chat input when the user types "/".
 * Shows the list of SLASH_COMMANDS with keyboard navigation, mouse hover
 * and click-to-select. Emits `prompt-menu-select` with the chosen command
 * detail; ChatWindow handles the rest (insert template, fill placeholders).
 * ============================================================================
 */

import { SLASH_COMMANDS, type SlashCommand } from "../utils/constants";

export interface PromptMenuSelectDetail {
  command: SlashCommand;
  /** The raw "/analyze" the user had typed at insert time. */
  query: string;
}

/** Build the floating menu DOM. The owner is responsible for positioning
 *  it (absolute, below the input). */
export class PromptMenuView {
  private root: HTMLDivElement;
  private commands: SlashCommand[];
  private activeIndex = 0;
  private filterText = "";

  constructor(commands: SlashCommand[] = SLASH_COMMANDS) {
    this.commands = commands;
    this.root = document.createElement("div");
    this.root.className = "prompt-menu";
    this.root.setAttribute("role", "listbox");
    this.root.setAttribute("aria-label", "斜杠命令菜单");
    this.root.style.display = "none";
    this.render();
  }

  /** Returns the root element so the caller can append it to the DOM. */
  get element(): HTMLDivElement {
    return this.root;
  }

  /** Show the menu and (optionally) pre-filter by what the user has typed
   *  after the leading "/". */
  show(query = ""): void {
    this.filterText = query.toLowerCase();
    this.activeIndex = 0;
    this.render();
    this.root.style.display = "block";
  }

  hide(): void {
    this.root.style.display = "none";
  }

  get visible(): boolean {
    return this.root.style.display !== "none";
  }

  /** Visible (filtered) command list. */
  get filtered(): SlashCommand[] {
    const q = this.filterText;
    if (!q) return this.commands;
    return this.commands.filter((c) =>
      c.id.toLowerCase().includes(q) ||
      c.label.toLowerCase().includes(q) ||
      c.trigger.toLowerCase().includes(q)
    );
  }

  /** Move the highlight by `delta` (+1 / -1). Wraps around. */
  move(delta: number): void {
    const list = this.filtered;
    if (list.length === 0) return;
    this.activeIndex = (this.activeIndex + delta + list.length) % list.length;
    this.render();
  }

  /** Returns the currently-highlighted command, or null if the filtered
   *  list is empty. */
  getActive(): SlashCommand | null {
    const list = this.filtered;
    return list[this.activeIndex] ?? null;
  }

  /** Set highlight by index (used for mouse hover). */
  setActive(index: number): void {
    const list = this.filtered;
    if (index < 0 || index >= list.length) return;
    this.activeIndex = index;
    this.render();
  }

  /* ---------------------------------------------------------------- */
  /* Internal                                                          */
  /* ---------------------------------------------------------------- */

  private render(): void {
    const list = this.filtered;
    if (list.length === 0) {
      this.root.innerHTML =
        '<div class="prompt-menu__empty">没有匹配的命令</div>';
      return;
    }

    const rows = list
      .map((cmd, idx) => {
        const cls = idx === this.activeIndex ? "prompt-menu__item is-active" : "prompt-menu__item";
        const req = cmd.requiresSelection
          ? '<span class="prompt-menu__tag" title="需要选区">选区</span>'
          : "";
        return (
          `<div class="${cls}" role="option" data-index="${idx}" data-id="${cmd.id}">` +
            `<div class="prompt-menu__icon">${escapeHtml(cmd.icon)}</div>` +
            `<div class="prompt-menu__body">` +
              `<div class="prompt-menu__title">${escapeHtml(cmd.label)}${req}</div>` +
              `<div class="prompt-menu__hint">${escapeHtml(cmd.hint)}</div>` +
            `</div>` +
            `<div class="prompt-menu__trigger">${escapeHtml(cmd.trigger)}</div>` +
          `</div>`
        );
      })
      .join("");

    const header =
      '<div class="prompt-menu__header">' +
        '<span>提示词命令</span>' +
        '<span class="prompt-menu__hint-small">↑↓ 选择 · Enter 确认 · Esc 关闭</span>' +
      '</div>';

    this.root.innerHTML = header + rows;
    this.bindEvents();
  }

  private bindEvents(): void {
    const items = this.root.querySelectorAll<HTMLDivElement>(".prompt-menu__item");
    items.forEach((el) => {
      el.addEventListener("mouseenter", () => {
        const idx = Number(el.dataset.index ?? "0");
        this.setActive(idx);
      });
      el.addEventListener("mousedown", (ev) => {
        // mousedown (not click) so we beat the textarea's blur.
        ev.preventDefault();
        const idx = Number(el.dataset.index ?? "0");
        this.setActive(idx);
        const cmd = this.getActive();
        if (cmd) this.fireSelect(cmd);
      });
    });
  }

  private fireSelect(cmd: SlashCommand): void {
    const detail: PromptMenuSelectDetail = { command: cmd, query: this.filterText };
    this.root.dispatchEvent(
      new CustomEvent("prompt-menu-select", { detail, bubbles: true })
    );
  }
}

/* ----------------------------------------------------------------- */
/* Helpers                                                            */
/* ----------------------------------------------------------------- */

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}