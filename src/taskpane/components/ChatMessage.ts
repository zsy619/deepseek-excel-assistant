/**
 * ============================================================================
 * ChatMessage
 * ----------------------------------------------------------------------------
 * Renders a single chat bubble. User and assistant messages get distinct
 * styling; the assistant bubble runs through MarkdownRenderer. A copy
 * button is shown on every assistant message.
 *
 * The component is a class that returns an HTMLElement so the chat window
 * can mount/unmount it without recreating it on every render.
 * ============================================================================
 */

import type { ChatMessage as ChatMessageType } from "../types";
import { renderMarkdown, markdownToPlainText, renderMermaidDiagrams } from "./MarkdownRenderer";
import { copyToClipboard, escapeHtml, formatRelativeTime } from "../utils/helpers";

export class ChatMessageView {
  /** Root element of this message bubble. */
  public readonly element: HTMLElement;

  /** The model object we are rendering. */
  private message: ChatMessageType;

  /** True while the assistant is still streaming this message. */
  private streaming: boolean;

  constructor(message: ChatMessageType, streaming: boolean = false) {
    this.message = message;
    this.streaming = streaming;
    this.element = document.createElement("div");
    this.element.classList.add("chat-message");
    this.element.dataset.role = message.role;
    this.element.dataset.id = message.id;
    this.render();
  }

  /** Patch the message in-place. Used during streaming to avoid
   *  re-creating the DOM node for every token. */
  public update(message: ChatMessageType, streaming: boolean): void {
    this.message = message;
    this.streaming = streaming;
    this.element.dataset.role = message.role;
    this.render();
  }

  /** Return the current underlying model. */
  public getMessage(): ChatMessageType {
    return this.message;
  }

  /** Tear down listeners. Currently a no-op but kept for API stability. */
  public destroy(): void {
    this.element.remove();
  }

  /* ---------------- rendering ---------------- */

  private render(): void {
    const role = this.message.role;
    const isUser = role === "user";
    const isAssistant = role === "assistant";
    const isSystem = role === "system";
    const isTool = role === "tool";

    const avatarChar = isUser ? "我" : isAssistant ? "DS" : isTool ? "🔧" : "SYS";
    const whoLabel = isUser
      ? "你"
      : isAssistant
      ? "DeepSeek"
      : isTool
      ? "工具调用结果"
      : "System";

    // Body content differs by role.
    let bodyHtml: string;
    if (isAssistant) {
      bodyHtml = renderMarkdown(this.message.content || "");
      if (this.streaming) {
        bodyHtml += '<span class="md-cursor">▍</span>';
      }
    } else if (isSystem) {
      bodyHtml = `<pre class="system-prompt">${escapeHtml(this.message.content)}</pre>`;
    } else if (isTool) {
      const name = this.message.toolName || "tool";
      bodyHtml = `<div class="tool-result">
        <div class="tool-result__head">🔧 工具: ${escapeHtml(name)}</div>
        <pre class="tool-result__body">${escapeHtml(this.message.content)}</pre>
      </div>`;
    } else {
      // User messages: keep newlines but escape everything.
      bodyHtml = `<div class="user-text">${escapeHtml(this.message.content).replace(/\n/g, "<br>")}</div>`;
    }

    const timestamp = formatRelativeTime(this.message.timestamp);
    const tokensBadge =
      typeof this.message.tokens === "number"
        ? `<span class="msg-tokens">${this.message.tokens} tokens</span>`
        : "";

    // PRD-09: tool-call badges on assistant messages.
    const toolCallBadges = isAssistant && this.message.toolCalls?.length
      ? `<div class="msg-tool-calls">${this.message.toolCalls
          .map(
            (tc) => `<span class="msg-tool-badge" title="${escapeHtml(tc.name)}: ${escapeHtml(tc.arguments)}">🔧 ${escapeHtml(tc.name)}${tc.executed ? " ✓" : ""}</span>`
          )
          .join("")}</div>`
      : "";

    this.element.innerHTML = `
      ${isUser ? "" : `<div class="msg-avatar ${isTool ? "msg-avatar-tool" : "msg-avatar-bot"}">${avatarChar}</div>`}
      <div class="msg-body">
        <div class="msg-header">
          <span class="msg-author">${whoLabel}</span>
          <span class="msg-time">${timestamp}</span>
          ${tokensBadge}
        </div>
        ${toolCallBadges}
        <div class="msg-content">${bodyHtml}</div>
        ${isAssistant ? this.renderActions() : ""}
      </div>
      ${isUser ? `<div class="msg-avatar msg-avatar-user">${avatarChar}</div>` : ""}
    `;

    this.attachHandlers();

    // Mermaid blocks are placeholders until the library renders them
    // asynchronously. Kick that off after mount; errors are surfaced
    // inline so the user still sees the source.
    if (isAssistant) {
      void renderMermaidDiagrams(this.element);
    }
  }

  private renderActions(): string {
    return `
      <div class="msg-actions">
        <button type="button" class="msg-action-btn" data-action="copy" title="复制消息内容">📋 复制</button>
        <button type="button" class="msg-action-btn" data-action="insert" title="插入到选中单元格">📌 插入</button>
        <button type="button" class="msg-action-btn" data-action="regenerate" title="重新生成回答">🔄 重新生成</button>
        <button type="button" class="msg-action-btn" data-action="branch" title="从这条消息分出新分支，AI 重新回答">🔀 分支</button>
      </div>
    `;
  }

  private attachHandlers(): void {
    // Copy code button inside markdown code blocks
    const copyBtns = this.element.querySelectorAll<HTMLButtonElement>(".md-copy-btn");
    copyBtns.forEach((btn) => {
      btn.onclick = async (ev) => {
        ev.stopPropagation();
        const code = btn.dataset.code || "";
        const decoded = (() => {
          try {
            return decodeURIComponent(code);
          } catch {
            return code;
          }
        })();
        const ok = await copyToClipboard(decoded);
        flashBtn(btn, ok ? "已复制" : "失败");
      };
    });

    // Top-level actions on assistant messages
    const actionBtns = this.element.querySelectorAll<HTMLButtonElement>(".msg-action-btn");
    actionBtns.forEach((btn) => {
      btn.onclick = (ev) => {
        ev.stopPropagation();
        const action = btn.dataset.action;
        this.dispatchAction(action || "");
      };
    });
  }

  /** Emit a CustomEvent so the parent ChatWindow can react without us
   *  importing a shared event bus. */
  private dispatchAction(action: string): void {
    const detail = {
      action,
      messageId: this.message.id,
      rawText: this.message.content,
      plainText: markdownToPlainText(this.message.content),
    };
    this.element.dispatchEvent(
      new CustomEvent("message-action", { detail, bubbles: true })
    );
  }
}

/** Briefly replace the button text to give the user feedback. */
function flashBtn(btn: HTMLButtonElement, text: string): void {
  const orig = btn.textContent;
  btn.textContent = text;
  btn.disabled = true;
  setTimeout(() => {
    btn.textContent = orig;
    btn.disabled = false;
  }, 1200);
}