/**
 * ============================================================================
 * CodeGenPanel (PRD-04)
 * ----------------------------------------------------------------------------
 * Card UI for "Formula -> VBA / Office Scripts". Renders:
 *   - flavor tabs (VBA / Office Scripts)
 *   - a streaming code block
 *   - a copy-to-clipboard button
 *   - a copy-and-guide button that opens a dialog explaining Alt+F11
 *
 * The panel itself does not call Excel - the host wires streaming and
 * copy to clipboard on its events. The panel only emits:
 *   - `code-gen-copy`    when the user clicks "复制"
 *   - `code-gen-retry`   when the user switches flavor
 * ============================================================================
 */

import type { ScriptFlavor } from "../services/deepseek";

export interface CodeGenCopyDetail {
  flavor: ScriptFlavor;
  code: string;
}

export interface CodeGenRetryDetail {
  flavor: ScriptFlavor;
}

export class CodeGenPanelView {
  private root: HTMLDivElement;
  private codeEl!: HTMLPreElement;
  private bodyEl!: HTMLDivElement;
  private statusEl!: HTMLElement;
  private tabVba!: HTMLButtonElement;
  private tabTs!: HTMLButtonElement;
  private copyBtn!: HTMLButtonElement;
  private guideBtn!: HTMLButtonElement;
  private flavor: ScriptFlavor = "vba";
  private code = "";

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "code-gen-panel";
    this.root.style.display = "none";
    this.render();
  }

  get element(): HTMLDivElement {
    return this.root;
  }

  show(initialFlavor: ScriptFlavor = "vba", initialCode = ""): void {
    this.flavor = initialFlavor;
    this.code = initialCode;
    this.root.style.display = "block";
    this.renderTabs();
    this.renderCode();
    this.renderStatus("ready");
  }

  hide(): void {
    this.root.style.display = "none";
  }

  /** Append a token to the live code block. */
  appendToken(token: string): void {
    this.code += token;
    if (this.bodyEl.querySelector(".code-gen-panel__code")) {
      this.codeEl.textContent = this.code;
      this.codeEl.scrollTop = this.codeEl.scrollHeight;
    } else {
      this.renderCode();
    }
  }

  /** Replace code wholesale (e.g. on flavor switch). */
  setCode(code: string): void {
    this.code = code;
    this.renderCode();
  }

  setStatus(state: "ready" | "streaming" | "done" | "error", text?: string): void {
    this.renderStatus(state, text);
  }

  setFlavor(flavor: ScriptFlavor): void {
    this.flavor = flavor;
    this.renderTabs();
  }

  get currentFlavor(): ScriptFlavor {
    return this.flavor;
  }

  /* ---------------------------------------------------------------- */
  /* Internal                                                          */
  /* ---------------------------------------------------------------- */

  private render(): void {
    this.root.innerHTML = `
      <div class="code-gen-panel__header">
        <div class="code-gen-panel__tabs">
          <button type="button" class="code-gen-panel__tab" data-ref="tabVba">VBA 脚本</button>
          <button type="button" class="code-gen-panel__tab" data-ref="tabTs">Office Scripts</button>
        </div>
        <div class="code-gen-panel__status" data-ref="status"></div>
      </div>
      <div class="code-gen-panel__body" data-ref="body">
        <pre class="code-gen-panel__code" data-ref="code"></pre>
      </div>
      <div class="code-gen-panel__footer">
        <button type="button" class="code-gen-panel__btn" data-ref="copy">复制到剪贴板</button>
        <button type="button" class="code-gen-panel__btn code-gen-panel__btn--ghost" data-ref="guide">复制并打开 VBA 编辑器</button>
      </div>
    `;
    this.codeEl = this.root.querySelector<HTMLPreElement>("[data-ref=code]")!;
    this.bodyEl = this.root.querySelector<HTMLDivElement>("[data-ref=body]")!;
    this.statusEl = this.root.querySelector<HTMLElement>("[data-ref=status]")!;
    this.tabVba = this.root.querySelector<HTMLButtonElement>("[data-ref=tabVba]")!;
    this.tabTs = this.root.querySelector<HTMLButtonElement>("[data-ref=tabTs]")!;
    this.copyBtn = this.root.querySelector<HTMLButtonElement>("[data-ref=copy]")!;
    this.guideBtn = this.root.querySelector<HTMLButtonElement>("[data-ref=guide]")!;

    this.tabVba.addEventListener("click", () => this.fireRetry("vba"));
    this.tabTs.addEventListener("click", () => this.fireRetry("office-scripts"));
    this.copyBtn.addEventListener("click", () => this.fireCopy());
    this.guideBtn.addEventListener("click", () => this.fireCopy());
    this.guideBtn.style.display = "none";
    // Show the guide button only for VBA flavor.
    const observer = () => {
      this.guideBtn.style.display = this.flavor === "vba" ? "inline-block" : "none";
    };
    this.tabVba.addEventListener("click", observer);
    this.tabTs.addEventListener("click", observer);

    this.renderTabs();
  }

  private renderTabs(): void {
    this.tabVba.classList.toggle("is-active", this.flavor === "vba");
    this.tabTs.classList.toggle("is-active", this.flavor === "office-scripts");
  }

  private renderCode(): void {
    this.codeEl.textContent = this.code || `// 等待 AI 生成 ${this.flavor === "vba" ? "VBA" : "Office Scripts"} 代码...`;
    // Best-effort: keep highlighted section simple (no hljs to keep
    // dependencies small; user can copy and paste into VS Code for full
    // syntax colors).
    if (this.code) this.codeEl.scrollTop = this.codeEl.scrollHeight;
  }

  private renderStatus(state: "ready" | "streaming" | "done" | "error", text?: string): void {
    const map: Record<string, string> = {
      ready: "就绪",
      streaming: "生成中…",
      done: "完成",
      error: "出错",
    };
    this.statusEl.textContent = text || map[state] || "";
    this.statusEl.dataset.state = state;
  }

  private fireCopy(): void {
    const detail: CodeGenCopyDetail = { flavor: this.flavor, code: this.code };
    this.root.dispatchEvent(
      new CustomEvent("code-gen-copy", { detail, bubbles: true })
    );
  }

  private fireRetry(flavor: ScriptFlavor): void {
    this.flavor = flavor;
    this.code = "";
    this.renderTabs();
    this.renderCode();
    const detail: CodeGenRetryDetail = { flavor };
    this.root.dispatchEvent(
      new CustomEvent("code-gen-retry", { detail, bubbles: true })
    );
  }
}