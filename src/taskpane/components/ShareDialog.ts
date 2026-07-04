/**
 * ============================================================================
 * ShareDialog component (PRD-11)
 * ----------------------------------------------------------------------------
 * Modal that lets the user:
 *   - copy a shareable URL (base64-encoded session fragment)
 *   - download the session as a .json file
 *   - import a session from a .json file
 *
 * Emits `share-share` with the decoded ShareableSession when the user
 * accepts an imported URL/file. The host restores the session.
 * ============================================================================
 */

import { escapeHtml, copyToClipboard } from "../utils/helpers";
import {
  encodeSession,
  buildShareUrl,
  sessionToShareableFile,
  fileToShareable,
  downloadSessionFile,
  type ShareableSession,
} from "../services/share";
import type { ChatSession } from "../types";

export interface ShareDialogDetail {
  /** Decoded session payload (only set when the user imports / accepts). */
  imported?: ShareableSession;
}

export class ShareDialogView {
  public readonly element: HTMLElement;
  private session: ChatSession | null = null;

  constructor() {
    this.element = document.createElement("div");
    this.element.classList.add("share-panel");
    this.element.dataset.ref = "shareDialog";
    this.element.hidden = true;
  }

  setSession(session: ChatSession | null): void {
    this.session = session;
    if (!this.element.hidden) this.refresh();
  }

  show(): void {
    this.element.hidden = false;
    this.refresh();
  }

  hide(): void {
    this.element.hidden = true;
  }

  private refresh(): void {
    if (!this.session) {
      this.element.innerHTML = `
        <div class="share-panel__backdrop" data-action="close"></div>
        <div class="share-panel__modal" role="dialog" aria-labelledby="share-title">
          <header class="share-panel__head">
            <h2 id="share-title">🔗 分享会话</h2>
            <button type="button" class="share-panel__close" data-action="close" title="关闭">✕</button>
          </header>
          <div class="share-panel__body">
            <div class="share-panel__empty">当前没有活跃对话，无法分享。</div>
          </div>
        </div>
      `;
      this.bindClose();
      return;
    }

    const messageCount = this.session.messages.length;
    const summary = `${this.session.title} · ${messageCount} 条消息`;

    this.element.innerHTML = `
      <div class="share-panel__backdrop" data-action="close"></div>
      <div class="share-panel__modal" role="dialog" aria-labelledby="share-title">
        <header class="share-panel__head">
          <h2 id="share-title">🔗 分享会话</h2>
          <button type="button" class="share-panel__close" data-action="close" title="关闭">✕</button>
        </header>
        <div class="share-panel__body">
          <div class="share-panel__summary">${escapeHtml(summary)}</div>

          <div class="share-panel__section">
            <h3>方式 1 · 复制分享链接</h3>
            <p class="share-panel__hint">链接内嵌入了完整会话内容，对方在同一个加载项里打开即可恢复。</p>
            <div class="share-panel__row">
              <input type="text" class="share-panel__link" readonly value="" />
              <button type="button" class="share-panel__btn share-panel__btn-primary" data-action="copyLink">📋 复制链接</button>
            </div>
          </div>

          <div class="share-panel__section">
            <h3>方式 2 · 导出 JSON 文件</h3>
            <p class="share-panel__hint">适合作为邮件附件长期归档，对方通过「导入」按钮加载。</p>
            <button type="button" class="share-panel__btn" data-action="downloadJson">⬇️ 下载 JSON</button>
          </div>

          <div class="share-panel__section">
            <h3>方式 3 · 导入 JSON 文件</h3>
            <p class="share-panel__hint">从别人分享的 JSON 文件恢复一个会话到本加载项。</p>
            <input type="file" class="share-panel__file" accept=".json,application/json" />
          </div>
        </div>
      </div>
    `;

    this.bindClose();
    this.bindActions();
  }

  private bindClose(): void {
    this.element.querySelectorAll<HTMLElement>("[data-action='close']").forEach((el) => {
      el.onclick = () => this.hide();
    });
  }

  private bindActions(): void {
    // Render the link into the readonly input.
    if (this.session) {
      const link = encodeSession(this.session);
      const url = buildShareUrl(link);
      const input = this.element.querySelector<HTMLInputElement>(".share-panel__link");
      if (input) input.value = url;
    }

    this.element.querySelector<HTMLButtonElement>("[data-action='copyLink']")!.onclick = async () => {
      const input = this.element.querySelector<HTMLInputElement>(".share-panel__link");
      if (!input) return;
      const ok = await copyToClipboard(input.value);
      this.toast(ok ? "链接已复制，发给同事即可" : "复制失败，请手动选择", ok ? "success" : "error");
    };

    this.element.querySelector<HTMLButtonElement>("[data-action='downloadJson']")!.onclick = () => {
      if (!this.session) return;
      downloadSessionFile(this.session);
      this.toast("已开始下载 JSON 文件", "success");
    };

    const fileInput = this.element.querySelector<HTMLInputElement>(".share-panel__file")!;
    fileInput.onchange = async () => {
      const file = fileInput.files?.[0];
      if (!file) return;
      try {
        const parsed = await fileToShareable(file);
        this.emitImported(parsed);
        this.toast(`已导入「${parsed.title}」`, "success");
        this.hide();
      } catch (err: any) {
        this.toast(err?.message || "导入失败", "error");
      } finally {
        fileInput.value = "";
      }
    };
  }

  /** Emit a custom event so the host can decide to import + render. */
  private emitImported(payload: ShareableSession): void {
    this.element.dispatchEvent(
      new CustomEvent<ShareDialogDetail>("share-import", { detail: { imported: payload }, bubbles: true })
    );
  }

  private toast(msg: string, kind: "info" | "success" | "error"): void {
    this.element.dispatchEvent(
      new CustomEvent("share-toast", { detail: { msg, kind }, bubbles: true })
    );
  }
}