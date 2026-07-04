/**
 * ============================================================================
 * KnowledgeBase component (PRD-10)
 * ----------------------------------------------------------------------------
 * Modal panel that lets the user:
 *   - upload .txt / .md / .csv files into the RAG store
 *   - paste raw text directly
 *   - browse existing docs and delete them
 *   - test retrieval against a sample query
 *
 * Emits no events; this is a self-contained management UI. Retrieval itself
 * is invoked transparently from ChatWindow before each AI call.
 * ============================================================================
 */

import { escapeHtml, formatRelativeTime, copyToClipboard } from "../utils/helpers";
import {
  addDocument,
  deleteDocument,
  loadAllDocs,
  retrieve,
  totalStats,
  type KnowledgeDoc,
  type RetrievedChunk,
} from "../services/rag";

export class KnowledgeBaseView {
  public readonly element: HTMLElement;

  constructor() {
    this.element = document.createElement("div");
    this.element.classList.add("kb-panel");
    this.element.dataset.ref = "knowledgeBase";
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
      <div class="kb-panel__backdrop" data-action="close"></div>
      <div class="kb-panel__modal" role="dialog" aria-labelledby="kb-title">
        <header class="kb-panel__head">
          <h2 id="kb-title">📚 知识库</h2>
          <button type="button" class="kb-panel__close" data-action="close" title="关闭">✕</button>
        </header>
        <section class="kb-panel__body">
          <div class="kb-panel__upload">
            <h3>添加文档</h3>
            <div class="kb-panel__upload-row">
              <input type="text" class="kb-panel__name" placeholder="文档名 (例如 财务手册)" />
              <input type="file" class="kb-panel__file" accept=".txt,.md,.csv,.json,.log" />
            </div>
            <textarea class="kb-panel__paste" placeholder="或者直接粘贴文本内容…"></textarea>
            <div class="kb-panel__upload-actions">
              <button type="button" class="kb-panel__btn kb-panel__btn-primary" data-action="addText">📥 添加文本</button>
              <span class="kb-panel__hint">支持 .txt / .md / .csv / .json，单文件 ≤ 5MB</span>
            </div>
          </div>
          <div class="kb-panel__list-section">
            <div class="kb-panel__list-head">
              <h3>已收录文档 <span class="kb-panel__count" data-ref="count">0</span></h3>
              <button type="button" class="kb-panel__btn kb-panel__btn-danger" data-action="clearAll" title="清空所有文档">🗑️ 全部清空</button>
            </div>
            <div class="kb-panel__list" data-ref="list"></div>
          </div>
          <div class="kb-panel__test">
            <h3>检索测试</h3>
            <input type="text" class="kb-panel__query" placeholder="输入查询，验证是否能检索到相关内容" />
            <div class="kb-panel__results" data-ref="results"></div>
          </div>
        </section>
      </div>
    `;
    this.attach();
    this.refresh();
  }

  private attach(): void {
    // Close handlers
    this.element.querySelectorAll<HTMLElement>("[data-action='close']").forEach((el) => {
      el.onclick = () => this.hide();
    });

    // File upload
    const fileInput = this.element.querySelector<HTMLInputElement>(".kb-panel__file")!;
    const nameInput = this.element.querySelector<HTMLInputElement>(".kb-panel__name")!;
    fileInput.onchange = async () => {
      const file = fileInput.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const name = nameInput.value.trim() || file.name;
        const doc = addDocument(name, text);
        this.toast(`已添加「${doc.name}」(${doc.chunks.length} 个片段)`, "success");
        nameInput.value = "";
        fileInput.value = "";
        this.refresh();
      } catch (err: any) {
        this.toast(`读取失败: ${err?.message || err}`, "error");
      }
    };

    // Text paste
    const pasteArea = this.element.querySelector<HTMLTextAreaElement>(".kb-panel__paste")!;
    this.element.querySelector<HTMLButtonElement>("[data-action='addText']")!.onclick = () => {
      const text = pasteArea.value.trim();
      if (!text) {
        this.toast("请输入或粘贴文本内容", "info");
        return;
      }
      const name = nameInput.value.trim() || `文本片段 ${new Date().toLocaleString("zh-CN")}`;
      const doc = addDocument(name, text);
      this.toast(`已添加「${doc.name}」(${doc.chunks.length} 个片段)`, "success");
      pasteArea.value = "";
      nameInput.value = "";
      this.refresh();
    };

    // Clear all
    this.element.querySelector<HTMLButtonElement>("[data-action='clearAll']")!.onclick = () => {
      const stats = totalStats();
      if (stats.docs === 0) {
        this.toast("知识库已为空", "info");
        return;
      }
      if (!confirm(`确定清空全部 ${stats.docs} 个文档？此操作不可撤销。`)) return;
      const { clearAllDocs } = require("../services/rag");
      clearAllDocs();
      this.toast("知识库已清空", "success");
      this.refresh();
    };

    // Live retrieval test
    const queryInput = this.element.querySelector<HTMLInputElement>(".kb-panel__query")!;
    let timer: number | undefined;
    queryInput.oninput = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => this.runTestQuery(), 220);
    };

    // ESC closes panel
    this.element.addEventListener("keydown", (ev) => {
      if ((ev as KeyboardEvent).key === "Escape") this.hide();
    });
  }

  private refresh(): void {
    const docs = loadAllDocs();
    const stats = totalStats();
    const countEl = this.element.querySelector<HTMLElement>("[data-ref='count']")!;
    countEl.textContent = String(stats.docs);

    const list = this.element.querySelector<HTMLElement>("[data-ref='list']")!;
    if (!docs.length) {
      list.innerHTML = `<div class="kb-panel__empty">还没有文档。上传或粘贴内容后，AI 回答时会自动引用。</div>`;
      return;
    }
    list.innerHTML = docs
      .map(
        (d) => `
        <div class="kb-panel__doc" data-id="${escapeHtml(d.id)}">
          <div class="kb-panel__doc-main">
            <div class="kb-panel__doc-name">${escapeHtml(d.name)}</div>
            <div class="kb-panel__doc-meta">
              ${d.chunks.length} 个片段 · ${(d.bytes / 1024).toFixed(1)}KB · ${formatRelativeTime(d.createdAt)}
            </div>
          </div>
          <div class="kb-panel__doc-actions">
            <button type="button" class="kb-panel__btn" data-action="preview" data-id="${escapeHtml(d.id)}">👁️ 预览</button>
            <button type="button" class="kb-panel__btn kb-panel__btn-danger" data-action="delete" data-id="${escapeHtml(d.id)}">删除</button>
          </div>
        </div>`
      )
      .join("");

    list.querySelectorAll<HTMLButtonElement>("[data-action='delete']").forEach((btn) => {
      btn.onclick = () => {
        const id = btn.dataset.id!;
        const doc = docs.find((d) => d.id === id);
        if (!doc) return;
        if (!confirm(`删除「${doc.name}」?`)) return;
        deleteDocument(id);
        this.toast("已删除", "success");
        this.refresh();
      };
    });
    list.querySelectorAll<HTMLButtonElement>("[data-action='preview']").forEach((btn) => {
      btn.onclick = () => {
        const id = btn.dataset.id!;
        const doc = docs.find((d) => d.id === id);
        if (!doc) return;
        const preview = doc.content.length > 2000 ? doc.content.slice(0, 2000) + "…" : doc.content;
        alert(`【${doc.name}】\n\n${preview}`);
      };
    });
  }

  private runTestQuery(): void {
    const q = (this.element.querySelector<HTMLInputElement>(".kb-panel__query")?.value || "").trim();
    const out = this.element.querySelector<HTMLElement>("[data-ref='results']")!;
    if (!q) {
      out.innerHTML = `<div class="kb-panel__empty">输入查询词以验证检索效果</div>`;
      return;
    }
    const results = retrieve(q, 5);
    if (!results.length) {
      out.innerHTML = `<div class="kb-panel__empty">未检索到相关片段</div>`;
      return;
    }
    out.innerHTML = results
      .map(
        (r: RetrievedChunk) => `
        <div class="kb-panel__result">
          <div class="kb-panel__result-head">
            <span class="kb-panel__result-doc">${escapeHtml(r.docName)}</span>
            <span class="kb-panel__result-score">score ${r.score.toFixed(2)}</span>
            <button type="button" class="kb-panel__btn kb-panel__btn-tiny" data-copy="${escapeHtml(r.text)}">复制</button>
          </div>
          <div class="kb-panel__result-text">${escapeHtml(r.text)}</div>
        </div>`
      )
      .join("");
    out.querySelectorAll<HTMLButtonElement>("[data-copy]").forEach((btn) => {
      btn.onclick = async () => {
        const ok = await copyToClipboard(btn.dataset.copy || "");
        this.toast(ok ? "已复制片段" : "复制失败", ok ? "success" : "error");
      };
    });
  }

  private toast(msg: string, kind: "info" | "success" | "error"): void {
    // The host taskpane listens for a `toast` event; we just emit one.
    this.element.dispatchEvent(
      new CustomEvent("kb-toast", { detail: { msg, kind }, bubbles: true })
    );
  }
}