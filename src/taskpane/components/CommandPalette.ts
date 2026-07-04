/**
 * ============================================================================
 * CommandPalette (PRD-03)
 * ----------------------------------------------------------------------------
 * VS Code / Linear style "command bar". Triggered by ⌘+K (or Ctrl+K on
 * Windows). Typing a query fuzzy-filters the COMMAND_PALETTE_ITEMS list;
 * ↑/↓ navigates, Enter fires the action, Esc closes.
 *
 *  - The palette does NOT know how to execute each action; it emits a
 *    `command-palette-select` CustomEvent with the chosen item's id and
 *    the host wires it up. This keeps the palette decoupled from
 *    ChatWindow's internals.
 *
 *  - Typing "?" (when the query is exactly "?") switches the body to a
 *    shortcut reference (SHORTCUTS array) per PRD-03 §3 acceptance.
 * ============================================================================
 */

import { COMMAND_PALETTE_ITEMS, SHORTCUTS, type CommandPaletteItem } from "../utils/constants";

export interface CommandPaletteSelectDetail {
  item: CommandPaletteItem;
  /** The query the user had typed at the moment of selection. */
  query: string;
}

export class CommandPaletteView {
  private root: HTMLDivElement;
  private inputEl!: HTMLInputElement;
  private listEl!: HTMLDivElement;
  private activeIndex = 0;
  private query = "";

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "command-palette";
    this.root.setAttribute("role", "dialog");
    this.root.setAttribute("aria-label", "命令面板");
    this.root.style.display = "none";
    this.render();
  }

  get element(): HTMLDivElement {
    return this.root;
  }

  get visible(): boolean {
    return this.root.style.display !== "none";
  }

  show(): void {
    this.query = "";
    this.activeIndex = 0;
    this.root.style.display = "flex";
    requestAnimationFrame(() => {
      this.inputEl.value = "";
      this.inputEl.focus();
      this.renderList();
    });
  }

  hide(): void {
    this.root.style.display = "none";
    // Drop focus from the input so subsequent keys go to ChatWindow.
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
  }

  /* ---------------------------------------------------------------- */
  /* Internal                                                          */
  /* ---------------------------------------------------------------- */

  private render(): void {
    this.root.innerHTML = `
      <div class="command-palette__backdrop" data-ref="backdrop"></div>
      <div class="command-palette__panel" role="document">
        <div class="command-palette__inputwrap">
          <span class="command-palette__prompt">⌘K</span>
          <input
            type="text"
            class="command-palette__input"
            data-ref="q"
            placeholder="输入命令或搜索（输入 ? 查看所有快捷键）"
            spellcheck="false"
            autocomplete="off"
          />
        </div>
        <div class="command-palette__list" data-ref="list" role="listbox"></div>
        <div class="command-palette__footer">
          <span>↑↓ 选择</span>
          <span>Enter 执行</span>
          <span>Esc 关闭</span>
        </div>
      </div>
    `;
    this.inputEl = this.root.querySelector<HTMLInputElement>('[data-ref="q"]')!;
    this.listEl = this.root.querySelector<HTMLDivElement>('[data-ref="list"]')!;

    this.bindEvents();
  }

  private bindEvents(): void {
    this.inputEl.addEventListener("input", () => {
      this.query = this.inputEl.value;
      this.activeIndex = 0;
      this.renderList();
    });

    this.inputEl.addEventListener("keydown", (ev) => {
      if (ev.key === "ArrowDown") { ev.preventDefault(); this.move(1); }
      else if (ev.key === "ArrowUp") { ev.preventDefault(); this.move(-1); }
      else if (ev.key === "Enter" && !ev.isComposing) {
        ev.preventDefault();
        this.fireActive();
      } else if (ev.key === "Escape") {
        ev.preventDefault();
        this.hide();
      }
    });

    const backdrop = this.root.querySelector<HTMLDivElement>('[data-ref="backdrop"]')!;
    backdrop.addEventListener("click", () => this.hide());
  }

  private filteredItems(): CommandPaletteItem[] {
    const q = this.query.trim().toLowerCase();
    if (!q) return COMMAND_PALETTE_ITEMS;
    if (q === "?") return [];
    return COMMAND_PALETTE_ITEMS.filter((c) =>
      c.label.toLowerCase().includes(q) ||
      c.hint.toLowerCase().includes(q) ||
      c.id.toLowerCase().includes(q) ||
      c.action.toLowerCase().includes(q)
    );
  }

  private renderList(): void {
    if (this.query.trim() === "?") {
      this.renderShortcuts();
      return;
    }
    const items = this.filteredItems();
    if (items.length === 0) {
      this.listEl.innerHTML =
        '<div class="command-palette__empty">没有匹配的命令</div>';
      return;
    }
    this.listEl.innerHTML = items
      .map((it, idx) => {
        const cls = idx === this.activeIndex ? "command-palette__item is-active" : "command-palette__item";
        return (
          `<div class="${cls}" role="option" data-index="${idx}" data-id="${it.id}">` +
            `<div class="command-palette__icon">${escapeHtml(it.icon)}</div>` +
            `<div class="command-palette__body">` +
              `<div class="command-palette__label">${escapeHtml(it.label)}</div>` +
              `<div class="command-palette__hint">${escapeHtml(it.hint)}</div>` +
            `</div>` +
            `<div class="command-palette__keys">${escapeHtml(it.shortcut)}</div>` +
          `</div>`
        );
      })
      .join("");

    this.listEl.querySelectorAll<HTMLDivElement>(".command-palette__item").forEach((el) => {
      el.addEventListener("mousedown", (ev) => {
        ev.preventDefault();
        const idx = Number(el.dataset.index ?? "0");
        this.activeIndex = idx;
        this.fireActive();
      });
      el.addEventListener("mouseenter", () => {
        const idx = Number(el.dataset.index ?? "0");
        this.activeIndex = idx;
        this.highlightActive();
      });
    });
    this.highlightActive();
  }

  private renderShortcuts(): void {
    const rows = SHORTCUTS
      .map(
        (s) =>
          `<div class="command-palette__shortcut">` +
            `<div class="command-palette__keys">${escapeHtml(s.combo)}</div>` +
            `<div class="command-palette__shortcut-label">${escapeHtml(s.label)}</div>` +
            `<div class="command-palette__scope">${s.scope === "input" ? "输入框" : "全局"}</div>` +
          `</div>`
      )
      .join("");
    this.listEl.innerHTML =
      '<div class="command-palette__shortcuts-title">所有快捷键</div>' + rows;
  }

  private highlightActive(): void {
    const items = this.listEl.querySelectorAll<HTMLDivElement>(".command-palette__item");
    items.forEach((el, idx) => {
      if (idx === this.activeIndex) {
        el.classList.add("is-active");
        el.scrollIntoView({ block: "nearest" });
      } else {
        el.classList.remove("is-active");
      }
    });
  }

  private move(delta: number): void {
    const items = this.filteredItems();
    if (items.length === 0) return;
    this.activeIndex = (this.activeIndex + delta + items.length) % items.length;
    this.highlightActive();
  }

  private fireActive(): void {
    const items = this.filteredItems();
    const item = items[this.activeIndex];
    if (!item) return;
    this.hide();
    const detail: CommandPaletteSelectDetail = { item, query: this.query };
    this.root.dispatchEvent(
      new CustomEvent("command-palette-select", { detail, bubbles: true })
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