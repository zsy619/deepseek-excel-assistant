/**
 * ============================================================================
 * HistoryPanel
 * ----------------------------------------------------------------------------
 * Sidebar that lists every persisted chat session. Supports:
 *   - Search by title/content
 *   - Switch active session on click
 *   - Per-session menu: rename / export / delete
 *   - Bulk select / delete
 *
 * The panel is a thin view layer; all persistence lives in storage.ts and
 * is driven by the parent (ChatWindow) which holds the canonical state.
 * ============================================================================
 */

import { downloadSessionAsMarkdown, escapeHtml, formatRelativeTime, generateId } from "../utils/helpers";
import { confirmDialog, promptDialog } from "./Dialog";
import type { ChatSession } from "../types";

export interface HistoryPanelEvents {
  /** User picked a session to switch into. */
  select: (sessionId: string) => void;
  /** User asked to create a new session. */
  create: () => void;
  /** User renamed a session. */
  rename: (sessionId: string, newTitle: string) => void;
  /** User asked to delete one or more sessions. */
  delete: (sessionIds: string[]) => void;
  /** User asked to export a session as Markdown. */
  export: (sessionId: string) => void;
}

export class HistoryPanel {
  private root: HTMLElement;
  private sessions: ChatSession[] = [];
  private activeId: string | null = null;
  private query: string = "";
  private bulkMode: boolean = false;
  private selectedIds: Set<string> = new Set();
  private events: HistoryPanelEvents;

  constructor(parent: HTMLElement, events: HistoryPanelEvents) {
    this.events = events;
    this.root = document.createElement("aside");
    this.root.classList.add("history-panel");
    this.root.innerHTML = this.renderShell();
    parent.appendChild(this.root);
    this.bindEvents();
    this.repaint();
  }

  /* ---------------- public ---------------- */

  public setSessions(sessions: ChatSession[], activeId: string | null): void {
    this.sessions = sessions;
    this.activeId = activeId;
    // Drop selections that no longer exist.
    for (const id of Array.from(this.selectedIds)) {
      if (!sessions.find((s) => s.id === id)) this.selectedIds.delete(id);
    }
    this.repaint();
  }

  public setActive(id: string | null): void {
    this.activeId = id;
    this.repaint();
  }

  public toggle(): void {
    this.root.classList.toggle("history-panel-open");
  }

  public close(): void {
    this.root.classList.remove("history-panel-open");
  }

  public isOpen(): boolean {
    return this.root.classList.contains("history-panel-open");
  }

  /* ---------------- rendering ---------------- */

  private renderShell(): string {
    return `
      <div class="history-header">
        <h3>📚 历史对话</h3>
        <button type="button" class="history-close-btn" data-action="close" title="收起">◀</button>
      </div>

      <div class="history-toolbar">
        <button type="button" class="history-btn history-btn-primary" data-action="create">+ 新建</button>
        <button type="button" class="history-btn" data-action="bulk">☑ 多选</button>
      </div>

      <div class="history-search">
        <input
          type="search"
          class="history-search-input"
          placeholder="🔍 搜索会话"
          data-field="search"
        />
      </div>

      <div class="history-bulk-bar" style="display:none">
        <span class="history-bulk-count" data-ref="bulkCount">已选 0</span>
        <button type="button" class="history-btn history-btn-danger" data-action="bulkDelete">🗑 删除所选</button>
        <button type="button" class="history-btn" data-action="bulkCancel">取消</button>
      </div>

      <div class="history-list" data-ref="list"></div>
    `;
  }

  private repaint(): void {
    const list = this.root.querySelector<HTMLElement>('[data-ref="list"]');
    if (!list) return;
    list.innerHTML = "";

    const filtered = this.applyFilter(this.sessions);
    if (filtered.length === 0) {
      list.innerHTML = `<div class="history-empty">${this.query ? "没有匹配的会话" : "暂无会话，点击「+ 新建」开始"}</div>`;
    } else {
      for (const s of filtered) {
        list.appendChild(this.renderItem(s));
      }
    }

    // Bulk mode toggle
    const bulkBar = this.root.querySelector<HTMLElement>(".history-bulk-bar")!;
    bulkBar.style.display = this.bulkMode ? "flex" : "none";
    const counter = this.root.querySelector<HTMLElement>('[data-ref="bulkCount"]')!;
    counter.textContent = `已选 ${this.selectedIds.size}`;
  }

  private renderItem(session: ChatSession): HTMLElement {
    const li = document.createElement("div");
    li.classList.add("history-item");
    if (session.id === this.activeId) li.classList.add("history-item-active");
    if (this.selectedIds.has(session.id)) li.classList.add("history-item-selected");
    li.dataset.id = session.id;

    const msgCount = session.messages.filter((m) => m.role !== "system").length;
    li.innerHTML = `
      <div class="history-item-main">
        ${
          this.bulkMode
            ? `<input type="checkbox" class="history-checkbox" ${this.selectedIds.has(session.id) ? "checked" : ""}/>`
            : ""
        }
        <div class="history-item-text">
          <div class="history-item-title">${escapeHtml(session.title)}</div>
          <div class="history-item-meta">
            <span>${formatRelativeTime(session.updatedAt)}</span>
            <span>•</span>
            <span>${msgCount} 条消息</span>
          </div>
        </div>
      </div>
      ${
        this.bulkMode
          ? ""
          : `<div class="history-item-actions">
              <button type="button" class="history-item-btn" data-action="rename" title="重命名">✏️</button>
              <button type="button" class="history-item-btn" data-action="export" title="导出 Markdown">⬇️</button>
              <button type="button" class="history-item-btn" data-action="delete" title="删除">🗑</button>
            </div>`
      }
    `;
    return li;
  }

  /* ---------------- filtering ---------------- */

  private applyFilter(items: ChatSession[]): ChatSession[] {
    if (!this.query) return items;
    const q = this.query.toLowerCase();
    return items.filter((s) => {
      if (s.title.toLowerCase().includes(q)) return true;
      return s.messages.some((m) => m.content.toLowerCase().includes(q));
    });
  }

  /* ---------------- events ---------------- */

  private bindEvents(): void {
    // Toolbar actions
    this.root.addEventListener("click", (ev) => {
      const target = ev.target as HTMLElement;
      const btn = target.closest<HTMLElement>("[data-action]");
      const item = target.closest<HTMLElement>(".history-item");

      // Toolbar buttons
      if (btn && !item) {
        const action = btn.dataset.action;
        if (action === "close") this.close();
        if (action === "create") this.events.create();
        if (action === "bulk") {
          this.bulkMode = true;
          this.repaint();
        }
        if (action === "bulkCancel") {
          this.bulkMode = false;
          this.selectedIds.clear();
          this.repaint();
        }
        if (action === "bulkDelete") {
          if (this.selectedIds.size === 0) return;
          confirmDialog({
            title: "删除会话",
            message: `确认删除 ${this.selectedIds.size} 个会话？此操作不可撤销。`,
            confirmText: "删除",
            cancelText: "取消",
            variant: "danger",
          }).then((ok) => {
            if (ok) {
              const ids = Array.from(this.selectedIds);
              this.events.delete(ids);
              this.selectedIds.clear();
              this.bulkMode = false;
            }
          });
        }
        return;
      }

      // Item-level actions
      if (item) {
        const id = item.dataset.id!;
        const action = btn?.dataset.action;

        if (this.bulkMode) {
          if (this.selectedIds.has(id)) this.selectedIds.delete(id);
          else this.selectedIds.add(id);
          this.repaint();
          return;
        }

        if (!action) {
          this.events.select(id);
          return;
        }

        if (action === "rename") {
          const session = this.sessions.find((s) => s.id === id);
          if (!session) return;
          promptDialog({
            title: "重命名会话",
            message: "为这个会话起一个新名字：",
            defaultValue: session.title,
            confirmText: "保存",
            cancelText: "取消",
          }).then((next) => {
            if (next !== null && next.trim()) {
              this.events.rename(id, next.trim());
            }
          });
        }

        if (action === "export") {
          this.events.export(id);
        }

        if (action === "delete") {
          const title = this.sessions.find((s) => s.id === id)?.title || "此会话";
          confirmDialog({
            title: "删除会话",
            message: `确认删除会话「${title}」？此操作不可撤销。`,
            confirmText: "删除",
            cancelText: "取消",
            variant: "danger",
          }).then((ok) => {
            if (ok) this.events.delete([id]);
          });
        }
      }
    });

    // Search input
    const searchInput = this.root.querySelector<HTMLInputElement>('[data-field="search"]');
    if (searchInput) {
      searchInput.addEventListener("input", () => {
        this.query = searchInput.value.trim();
        this.repaint();
      });
    }
  }
}

/** Re-exported helpers used by ChatWindow when handling the export event. */
export { downloadSessionAsMarkdown };