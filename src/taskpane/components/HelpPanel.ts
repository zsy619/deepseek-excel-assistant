/**
 * ============================================================================
 * HelpPanel — floating draggable/resizable help window
 * ----------------------------------------------------------------------------
 * Loads pre-rendered HTML pages from docs/{lang}/*.html and renders them
 * inside a positioned div with:
 *   - title-bar drag (mouse + touch)
 *   - bottom-right resize handle
 *   - left TOC sidebar (per-language)
 *   - top: search input (fuse.js fuzzy) + language switcher + close button
 *   - bottom: prev/next buttons + breadcrumb
 *
 * Search is lazy: fuse.js (~10K minified) is only fetched when the user
 * types a query. The search index is docs/search-index.json, regenerated
 * by scripts/build-help-index.js at webpack build time.
 * ============================================================================
 */

import {
  HELP_TREE,
  nodesForLang,
  findByPath,
  getSiblings,
  type HelpNode,
} from "../helpTree";

type Lang = "zh" | "en";

interface SearchEntry {
  path: string;
  lang: string;
  section: string;
  body: string;
}

export class HelpPanel {
  public root: HTMLElement;
  private titleBar!: HTMLElement;
  private closeBtn!: HTMLButtonElement;
  private langSwitcher!: HTMLSelectElement;
  private searchInput!: HTMLInputElement;
  private searchResults!: HTMLElement;
  private tocEl!: HTMLElement;
  private contentEl!: HTMLElement;
  private breadcrumbEl!: HTMLElement;
  private prevBtn!: HTMLButtonElement;
  private nextBtn!: HTMLButtonElement;
  private resizeHandle!: HTMLElement;

  private lang: Lang = "zh";
  private currentPath = "zh/index.html";
  private fuse: any = null;
  private fusePromise: Promise<any> | null = null;
  private indexData: SearchEntry[] | null = null;

  // Drag state.
  private dragOffsetX = 0;
  private dragOffsetY = 0;
  private dragging = false;

  // Resize state.
  private resizing = false;
  private resizeStartX = 0;
  private resizeStartY = 0;
  private resizeStartW = 0;
  private resizeStartH = 0;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "help-panel";
    this.root.dataset.lang = this.lang;
    this.root.style.display = "none";
    this.root.innerHTML = this.renderShell();
    this.collectRefs();
    this.bindDrag();
    this.bindResize();
    this.bindEvents();
    this.renderTOC();
    this.loadDoc(this.currentPath);
  }

  public show(): void {
    this.root.style.display = "";
    this.positionDefault();
  }

  public hide(): void {
    this.root.style.display = "none";
  }

  public toggle(): void {
    if (this.root.style.display === "none") this.show();
    else this.hide();
  }

  /* ---------------------------------------------------------------- *
   * DOM shell
   * ---------------------------------------------------------------- */

  private renderShell(): string {
    return (
      '<div class="help-panel__titlebar" data-ref="titleBar">' +
      '  <span class="help-panel__brand">📖 DeepSeek 帮助</span>' +
      '  <select class="help-panel__lang" data-ref="langSwitcher" aria-label="语言">' +
      '    <option value="zh">中文</option>' +
      '    <option value="en">English</option>' +
      '  </select>' +
      '  <input type="search" class="help-panel__search" data-ref="searchInput" placeholder="搜索… (按 Enter 跳转)" />' +
      '  <button type="button" class="help-panel__close" data-ref="closeBtn" aria-label="关闭">✕</button>' +
      '</div>' +
      '<div class="help-panel__body">' +
      '  <aside class="help-panel__toc" data-ref="toc"></aside>' +
      '  <main class="help-panel__content" data-ref="content"></main>' +
      '</div>' +
      '<div class="help-panel__results" data-ref="results" hidden></div>' +
      '<div class="help-panel__footer">' +
      '  <button type="button" class="help-panel__nav" data-ref="prevBtn">← 上一页</button>' +
      '  <nav class="help-panel__breadcrumb" data-ref="breadcrumb"></nav>' +
      '  <button type="button" class="help-panel__nav" data-ref="nextBtn">下一页 →</button>' +
      '</div>' +
      '<div class="help-panel__resize" data-ref="resizeHandle" aria-label="调整大小">⇲</div>'
    );
  }

  private collectRefs(): void {
    const find = <T extends HTMLElement>(k: string): T => {
      const el = this.root.querySelector<T>(`[data-ref="${k}"]`);
      if (!el) throw new Error(`HelpPanel: missing data-ref="${k}"`);
      return el;
    };
    this.titleBar = find("titleBar");
    this.closeBtn = find<HTMLButtonElement>("closeBtn");
    this.langSwitcher = find<HTMLSelectElement>("langSwitcher");
    this.searchInput = find<HTMLInputElement>("searchInput");
    this.searchResults = find("results");
    this.tocEl = find("toc");
    this.contentEl = find("content");
    this.breadcrumbEl = find("breadcrumb");
    this.prevBtn = find<HTMLButtonElement>("prevBtn");
    this.nextBtn = find<HTMLButtonElement>("nextBtn");
    this.resizeHandle = find("resizeHandle");
  }

  private bindEvents(): void {
    this.closeBtn.addEventListener("click", () => this.hide());
    this.langSwitcher.value = this.lang;
    this.langSwitcher.addEventListener("change", () => {
      this.lang = this.langSwitcher.value as Lang;
      this.root.dataset.lang = this.lang;
      // Jump to the index of the new language.
      const target = nodesForLang(this.lang)[0];
      this.loadDoc(target.path);
      this.renderTOC();
    });
    this.searchInput.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        void this.runSearch(this.searchInput.value);
      }
    });
    this.searchInput.addEventListener("input", () => {
      // Live results as user types (debounced lightly via rAF).
      if (!this.searchInput.value.trim()) {
        this.searchResults.hidden = true;
        return;
      }
      if (this.searchRaf) cancelAnimationFrame(this.searchRaf);
      this.searchRaf = requestAnimationFrame(() => {
        void this.runSearch(this.searchInput.value);
      });
    });
    this.prevBtn.addEventListener("click", () => {
      const { prev } = getSiblings(this.currentPath);
      if (prev) this.loadDoc(prev.path);
    });
    this.nextBtn.addEventListener("click", () => {
      const { next } = getSiblings(this.currentPath);
      if (next) this.loadDoc(next.path);
    });
  }

  private searchRaf: number | null = null;

  /* ---------------------------------------------------------------- *
   * Drag & resize
   * ---------------------------------------------------------------- */

  private bindDrag(): void {
    this.titleBar.addEventListener("mousedown", (ev) => {
      // Ignore clicks on buttons / inputs / select inside the title bar.
      const target = ev.target as HTMLElement;
      if (target.closest("button, input, select")) return;
      ev.preventDefault();
      this.dragging = true;
      const rect = this.root.getBoundingClientRect();
      this.dragOffsetX = ev.clientX - rect.left;
      this.dragOffsetY = ev.clientY - rect.top;
      document.addEventListener("mousemove", this.onDragMove);
      document.addEventListener("mouseup", this.onDragEnd);
    });
  }

  private onDragMove = (ev: MouseEvent): void => {
    if (!this.dragging) return;
    const x = Math.max(0, Math.min(window.innerWidth - 100, ev.clientX - this.dragOffsetX));
    const y = Math.max(0, Math.min(window.innerHeight - 40, ev.clientY - this.dragOffsetY));
    this.root.style.left = `${x}px`;
    this.root.style.top = `${y}px`;
    this.root.style.right = "auto";
    this.root.style.bottom = "auto";
  };

  private onDragEnd = (): void => {
    this.dragging = false;
    document.removeEventListener("mousemove", this.onDragMove);
    document.removeEventListener("mouseup", this.onDragEnd);
  };

  private bindResize(): void {
    this.resizeHandle.addEventListener("mousedown", (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      this.resizing = true;
      const rect = this.root.getBoundingClientRect();
      this.resizeStartX = ev.clientX;
      this.resizeStartY = ev.clientY;
      this.resizeStartW = rect.width;
      this.resizeStartH = rect.height;
      document.addEventListener("mousemove", this.onResizeMove);
      document.addEventListener("mouseup", this.onResizeEnd);
    });
  }

  private onResizeMove = (ev: MouseEvent): void => {
    if (!this.resizing) return;
    const dx = ev.clientX - this.resizeStartX;
    const dy = ev.clientY - this.resizeStartY;
    const w = Math.max(360, Math.min(window.innerWidth - 20, this.resizeStartW + dx));
    const h = Math.max(280, Math.min(window.innerHeight - 40, this.resizeStartH + dy));
    this.root.style.width = `${w}px`;
    this.root.style.height = `${h}px`;
  };

  private onResizeEnd = (): void => {
    this.resizing = false;
    document.removeEventListener("mousemove", this.onResizeMove);
    document.removeEventListener("mouseup", this.onResizeEnd);
  };

  private positionDefault(): void {
    if (this.root.style.left && this.root.style.top) return;
    const w = 720;
    const h = 520;
    this.root.style.width = `${w}px`;
    this.root.style.height = `${h}px`;
    this.root.style.left = `${Math.max(20, window.innerWidth - w - 40)}px`;
    this.root.style.top = `${Math.max(20, window.innerHeight - h - 40)}px`;
    this.root.style.right = "auto";
    this.root.style.bottom = "auto";
  }

  /* ---------------------------------------------------------------- *
   * TOC, content loading, navigation
   * ---------------------------------------------------------------- */

  private renderTOC(): void {
    const nodes = nodesForLang(this.lang);
    this.tocEl.innerHTML =
      '<div class="help-panel__toc-title">' + (this.lang === "zh" ? "目录" : "Contents") + "</div>" +
      nodes
        .map(
          (n) =>
            `<a class="help-panel__toc-link" data-path="${n.path}" href="#">${escapeHtml(n.title)}</a>`
        )
        .join("");
    this.tocEl.querySelectorAll<HTMLAnchorElement>(".help-panel__toc-link").forEach((a) => {
      a.addEventListener("click", (ev) => {
        ev.preventDefault();
        const path = a.dataset.path || "";
        if (path) this.loadDoc(path);
      });
    });
    this.highlightActiveTOC();
  }

  private highlightActiveTOC(): void {
    this.tocEl.querySelectorAll<HTMLAnchorElement>(".help-panel__toc-link").forEach((a) => {
      if (a.dataset.path === this.currentPath) a.classList.add("is-active");
      else a.classList.remove("is-active");
    });
  }

  private async loadDoc(path: string): Promise<void> {
    this.currentPath = path;
    const node = findByPath(path);
    if (node?.lang) {
      this.lang = node.lang;
      this.langSwitcher.value = this.lang;
      this.root.dataset.lang = this.lang;
    }
    this.searchResults.hidden = true;
    this.searchInput.value = "";
    this.contentEl.innerHTML = '<div class="help-panel__loading">加载中…</div>';
    try {
      const res = await fetch(`docs/${path}`, { cache: "no-cache" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const html = await res.text();
      // Strip the <html>/<head> and inject only the body content so our
      // panel styles take over. shared.css is fetched and inlined into the
      // shadow DOM via the .html file already, so styles apply.
      const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
      const body = bodyMatch ? bodyMatch[1] : html;
      this.contentEl.innerHTML = body;
      // Make in-page links relative to the docs/ folder so prev/next anchors
      // resolve via loadDoc() rather than full-page navigation.
      this.contentEl.querySelectorAll<HTMLAnchorElement>("a[href]").forEach((a) => {
        const href = a.getAttribute("href") || "";
        if (
          href.startsWith("http") ||
          href.startsWith("#") ||
          href.startsWith("mailto:")
        ) {
          // External: open in browser.
          if (href.startsWith("http")) {
            a.setAttribute("target", "_blank");
            a.setAttribute("rel", "noopener noreferrer");
          }
          return;
        }
        // Internal: treat as a doc path and intercept the click.
        a.addEventListener("click", (ev) => {
          ev.preventDefault();
          this.loadDoc(href);
        });
      });
      this.contentEl.scrollTop = 0;
      this.highlightActiveTOC();
      this.updateFooter();
    } catch (err: any) {
      this.contentEl.innerHTML =
        '<div class="help-panel__error">加载失败:' + escapeHtml(err?.message || String(err)) + "</div>";
    }
  }

  private updateFooter(): void {
    const node = findByPath(this.currentPath);
    const { prev, next } = getSiblings(this.currentPath);
    this.breadcrumbEl.innerHTML =
      (node?.title || "") +
      (node?.lang === "zh" ? "  ·  中文" : "  ·  English");
    this.prevBtn.disabled = !prev;
    this.nextBtn.disabled = !next;
  }

  /* ---------------------------------------------------------------- *
   * Search (lazy fuse.js + prebuilt local index)
   * ---------------------------------------------------------------- */

  private async loadFuse(): Promise<any> {
    if (this.fuse) return this.fuse;
    if (!this.fusePromise) {
      this.fusePromise = (async () => {
        // fuse.js v7 ships ESM + UMD. Use the UMD browser bundle.
        // @ts-ignore — fuse.js has no types; we treat it as any.
        const FuseMod: any = await import(/* webpackIgnore: true */ "https://cdn.jsdelivr.net/npm/fuse.js@7.0.0/dist/fuse.min.js");
        const Fuse = FuseMod.default || FuseMod;
        if (!this.indexData) {
          const res = await fetch("docs/search-index.json", { cache: "no-cache" });
          this.indexData = await res.json();
        }
        this.fuse = new Fuse(this.indexData, {
          keys: [
            { name: "section", weight: 0.5 },
            { name: "body", weight: 0.5 },
          ],
          threshold: 0.4,
          includeScore: true,
          minMatchCharLength: 2,
        });
        return this.fuse;
      })();
    }
    return this.fusePromise;
  }

  private async runSearch(query: string): Promise<void> {
    const q = query.trim();
    if (!q) {
      this.searchResults.hidden = true;
      return;
    }
    try {
      const fuse = await this.loadFuse();
      const raw = fuse.search(q) as Array<{ item: SearchEntry; score?: number }>;
      const filtered = raw.filter((r) => r.item.lang === this.lang).slice(0, 12);
      if (!filtered.length) {
        this.searchResults.innerHTML =
          '<div class="help-panel__results-empty">没有匹配项</div>';
        this.searchResults.hidden = false;
        return;
      }
      this.searchResults.innerHTML =
        '<div class="help-panel__results-title">匹配结果(点击跳转)</div>' +
        filtered
          .map((r) => {
            const snippet = makeSnippet(r.item.body, q);
            return (
              '<a class="help-panel__result" data-path="' +
              r.item.path +
              '" href="#">' +
              '<div class="help-panel__result-section">' +
              escapeHtml(r.item.section) +
              "</div>" +
              '<div class="help-panel__result-snippet">' +
              escapeHtml(snippet) +
              "</div>" +
              "</a>"
            );
          })
          .join("");
      this.searchResults.querySelectorAll<HTMLAnchorElement>(".help-panel__result").forEach((a) => {
        a.addEventListener("click", (ev) => {
          ev.preventDefault();
          const path = a.dataset.path || "";
          if (path) this.loadDoc(path);
        });
      });
      this.searchResults.hidden = false;
    } catch (err: any) {
      this.searchResults.innerHTML =
        '<div class="help-panel__results-empty">搜索失败:' +
        escapeHtml(err?.message || String(err)) +
        "</div>";
      this.searchResults.hidden = false;
    }
  }
}

/** Trim a snippet around the first occurrence of `query` (case-insensitive). */
function makeSnippet(body: string, query: string, radius = 60): string {
  if (!body) return "";
  const lower = body.toLowerCase();
  const idx = lower.indexOf(query.toLowerCase());
  if (idx === -1) return body.slice(0, radius * 2) + (body.length > radius * 2 ? "…" : "");
  const start = Math.max(0, idx - radius);
  const end = Math.min(body.length, idx + query.length + radius);
  const prefix = start > 0 ? "…" : "";
  const suffix = end < body.length ? "…" : "";
  return prefix + body.slice(start, end) + suffix;
}

function escapeHtml(s: string): string {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}