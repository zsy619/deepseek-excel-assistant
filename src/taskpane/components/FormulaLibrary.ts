/**
 * ============================================================================
 * FormulaLibrary
 * ----------------------------------------------------------------------------
 * A collapsible panel that lists the most common Excel formula patterns
 * (VLOOKUP, XLOOKUP, SUMIFS, ...). Clicking a card drops a pre-filled
 * prompt into the chat input box so the user just has to add their own
 * values and hit send.
 *
 * The component is intentionally thin - it owns DOM, the rest of the
 * app (ChatWindow) listens for the `formula-pick` CustomEvent.
 * ============================================================================
 */

import { FORMULA_CARDS, type FormulaCard } from "../utils/constants";

export class FormulaLibrary {
  private root: HTMLElement;
  private collapsed: boolean = false;
  /** Persist the collapsed state across renders. */
  private static readonly STORAGE_KEY = "deepseek_formula_library_collapsed";

  constructor(parent: HTMLElement) {
    this.root = document.createElement("div");
    this.root.classList.add("formula-library");
    try {
      this.collapsed = localStorage.getItem(FormulaLibrary.STORAGE_KEY) === "1";
    } catch {
      /* noop */
    }
    if (this.collapsed) this.root.classList.add("collapsed");
    parent.appendChild(this.root);
    this.render();
  }

  public destroy(): void {
    this.root.remove();
  }

  private render(): void {
    this.root.innerHTML = `
      <div class="formula-library-header">
        <span>📚 公式库（${FORMULA_CARDS.length}）</span>
        <button type="button" class="formula-library-toggle" data-action="toggle">
          ${this.collapsed ? "展开" : "收起"}
        </button>
      </div>
      <div class="formula-library-grid"></div>
    `;

    const grid = this.root.querySelector<HTMLElement>(".formula-library-grid");
    if (!grid) return;

    for (const card of FORMULA_CARDS) {
      grid.appendChild(this.renderCard(card));
    }

    this.root.querySelector<HTMLButtonElement>('[data-action="toggle"]')?.addEventListener(
      "click",
      () => this.toggleCollapse()
    );
  }

  private renderCard(card: FormulaCard): HTMLElement {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.classList.add("formula-card");
    btn.dataset.id = card.id;
    btn.title = card.prompt;
    btn.innerHTML = `
      <span class="formula-card-name">${card.name}</span>
      <span class="formula-card-desc">${card.desc}</span>
    `;
    btn.addEventListener("click", () => this.pickCard(card));
    return btn;
  }

  private toggleCollapse(): void {
    this.collapsed = !this.collapsed;
    this.root.classList.toggle("collapsed", this.collapsed);
    try {
      localStorage.setItem(FormulaLibrary.STORAGE_KEY, this.collapsed ? "1" : "0");
    } catch {
      /* noop */
    }
    const btn = this.root.querySelector<HTMLButtonElement>('[data-action="toggle"]');
    if (btn) btn.textContent = this.collapsed ? "展开" : "收起";
  }

  private pickCard(card: FormulaCard): void {
    this.root.dispatchEvent(
      new CustomEvent("formula-pick", {
        detail: { id: card.id, name: card.name, prompt: card.prompt },
        bubbles: true,
      })
    );
  }
}
