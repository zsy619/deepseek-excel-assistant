/**
 * ============================================================================
 * KnowledgeInjector (PRD-10)
 * ----------------------------------------------------------------------------
 * Owns the "given a user question, what context from the local knowledge
 * base should we hand to the model" step. Today this is pure BM25 over
 * localStorage; tomorrow it could swap to embeddings, hybrid search, etc.
 *
 * Flow:
 *   1. ChatWindow's streamAssistantReply calls `inject()` with the user's
 *      question and the configured system prompt.
 *   2. We retrieve top-k chunks via rag.ts.
 *   3. We format them into a `[参考资料] ... [/参考资料]` block.
 *   4. We return a single `system`-role ChatMessage ready to prepend to the
 *      API message list — or undefined if there's nothing to inject.
 *
 * DoS / quota guards (Wave 3 fix):
 *   Wave 2 inline version had no size cap on the injected context. A user
 *   with a 5MB KB could trigger massive prompts and token costs. We cap
 *   the chunk count and total char budget; the rest is silently dropped
 *   and surfaced via toast so users know their KB is contributing.
 *
 * Hub-only dependency: this controller never imports ChatWindow.
 * ============================================================================
 */

import type { ChatMessage } from "../types";
import {
  retrieve as ragRetrieve,
  formatRetrievedAsContext,
  totalStats as ragTotalStats,
} from "../services/rag";
import type { ChatControllerHub } from "./ChatControllerHub";

/** Top-k chunks to surface per query. Modest on purpose — context bloat
 *  hurts latency and cost more than it helps relevance. */
const TOP_K = 3;

/** Hard cap on the rendered `[参考资料]` block size (chars). BM25 may
 *  return more chunks than this represents; we truncate the block. */
const MAX_CONTEXT_CHARS = 4000;

/** How often to remind the user the KB is being used. The first time
 *  per stream we surface a brief toast; subsequent rounds in the same
 *  recursion are silent so we don't spam. */
let lastWarnAt = 0;

export class KnowledgeInjector {
  constructor(private readonly hub: ChatControllerHub) {}

  /**
   * Return a single `role: "system"` ChatMessage containing the user's
   * configured system prompt plus the retrieved KB context, or undefined
   * if there's nothing meaningful to inject.
   *
   * Pure function over the hub config — no side effects beyond an
   * occasional status toast when the KB is large.
   */
  public inject(userQuestion: string, systemPrompt: string): ChatMessage | undefined {
    const stats = ragTotalStats();
    if (stats.docs === 0) {
      // No KB loaded - just return the user's system prompt verbatim.
      return this.systemOnly(systemPrompt);
    }

    const hits = ragRetrieve(userQuestion, TOP_K);
    const ragBlock = formatRetrievedAsContext(hits).trim();
    const truncated = this.truncate(ragBlock);

    if (truncated.warned) {
      const now = Date.now();
      if (now - lastWarnAt > 60_000) {
        this.hub.toast("知识库内容较多,只截取前 4000 字符", "info");
        lastWarnAt = now;
      }
    }

    const parts = [systemPrompt || "", truncated.text].filter(Boolean);
    const combined = parts.join("\n\n").trim();
    if (!combined) return undefined;

    return {
      id: "system",
      role: "system",
      content: combined,
      timestamp: Date.now(),
    };
  }

  /** Build a system message containing only the user's prompt (no KB). */
  private systemOnly(systemPrompt: string): ChatMessage | undefined {
    const text = (systemPrompt || "").trim();
    if (!text) return undefined;
    return {
      id: "system",
      role: "system",
      content: text,
      timestamp: Date.now(),
    };
  }

  /** Cap the block at MAX_CONTEXT_CHARS, returning whether truncation
   *  happened so the caller can surface a warning. */
  private truncate(block: string): { text: string; warned: boolean } {
    if (block.length <= MAX_CONTEXT_CHARS) return { text: block, warned: false };
    return { text: block.slice(0, MAX_CONTEXT_CHARS) + "\n…", warned: true };
  }
}