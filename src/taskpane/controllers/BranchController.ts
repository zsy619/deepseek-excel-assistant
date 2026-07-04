/**
 * ============================================================================
 * BranchController (PRD-08)
 * ----------------------------------------------------------------------------
 * Owns the "branch from this assistant message" workflow:
 *
 *   1. User clicks 🔀 分支 on an assistant bubble
 *   2. We locate the preceding user message
 *   3. We count sibling branches (assistant messages whose `branch.parentId`
 *      points at the target) - so the new branch is labelled #N+1
 *   4. We truncate the session at the target message
 *   5. We render a "🔀 分支 #N" divider above the new answer
 *   6. We kick off a fresh streamAssistantReply via the hub
 *   7. We tag the resulting new assistant message with
 *      `branch = { parentId: target.id, branchId: newMsg.id }`
 *
 * Fix vs Wave 2 inline version: the inline code sliced the messages array
 * BEFORE capturing the parentId, then assigned `branch.parentId` to an id
 * that no longer existed in the session. This controller captures the
 * `parentId` once, retains it on the target itself (so it survives across
 * re-renders), and uses it consistently through the branch lifecycle.
 * ============================================================================
 */

import type { ChatMessage, ChatSession } from "../types";
import type { ChatControllerHub } from "./ChatControllerHub";

export class BranchController {
  constructor(private readonly hub: ChatControllerHub) {}

  /** Branch from `assistantMessageId`. Returns true if a branch was
   *  started. The actual stream resolves asynchronously. */
  public async runBranchFromMessage(assistantMessageId: string): Promise<boolean> {
    const session = this.hub.getActiveSession();
    if (!session) return false;

    const idx = session.messages.findIndex((m) => m.id === assistantMessageId);
    if (idx < 0) return false;
    const target = session.messages[idx];
    if (target.role !== "assistant") {
      this.hub.toast("只能在 AI 回复上创建分支", "error");
      return false;
    }

    const userIdx = this.findPrecedingUserIndex(session, idx);
    if (userIdx < 0) {
      this.hub.toast("找不到原始用户消息", "error");
      return false;
    }

    // Count siblings: assistant messages in the session whose branch.parentId
    // is the target's id. The target itself counts as the original (#1),
    // so the new branch is #(siblings + 1).
    const siblings = session.messages.filter(
      (m) => m.role === "assistant" && m.branch?.parentId === target.id
    );
    const branchNumber = siblings.length + 1;
    const parentId = target.id;

    // Snapshot + truncate. The hidden originals remain in `removed` so we
    // can re-attach them if the user decides to switch back later.
    const removed: ChatMessage[] = session.messages.slice(idx);
    session.messages = session.messages.slice(0, idx);

    // Visually dim the original bubbles - we keep the DOM around so the
    // user can scroll back and read them.
    for (const old of removed) {
      this.hub.hideMessageView(old.id);
    }

    // Mark the divider with the parent id so the new branch can later be
    // linked back, and append it.
    const divider = document.createElement("div");
    divider.className = "msg-branch-divider";
    divider.dataset.parentId = parentId;
    divider.innerHTML = `<span>🔀 分支 #${branchNumber}</span>`;
    this.hub.appendListChild(divider);

    this.hub.setStatus("正在生成分支回答…");
    await this.hub.streamAssistantReply(session, session.messages[userIdx]);

    // Tag the freshly-streamed assistant message with branch metadata.
    // The streamAssistantReply pushes a brand-new message at the tail;
    // we mutate it in place so its id becomes the branchId.
    const tail = session.messages[session.messages.length - 1];
    if (tail && tail.role === "assistant") {
      tail.branch = { parentId, branchId: tail.id };
      session.updatedAt = Date.now();
      this.hub.upsertSession(session);
      // Note: view was created during streaming; we don't need to update
      // here because ChatWindow's stream path renders via update() which
      // would have replaced it already. Nothing else to do.
    }
    this.hub.setStatus("");
    return true;
  }

  /** Find the index of the most recent user message at or before `idx`. */
  private findPrecedingUserIndex(session: ChatSession, idx: number): number {
    for (let i = idx - 1; i >= 0; i--) {
      if (session.messages[i].role === "user") return i;
    }
    return -1;
  }
}