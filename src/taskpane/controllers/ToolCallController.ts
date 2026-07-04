/**
 * ============================================================================
 * ToolCallController (PRD-09)
 * ----------------------------------------------------------------------------
 * Owns the "model emitted tool_calls → execute locally → feed results back
 * → recurse for follow-up reply" loop.
 *
 * Flow:
 *   1. DeepSeek stream emits `onToolCalls(calls)` with one or more ToolCall.
 *   2. We attach them to the current assistant message, persist.
 *   3. We execute each call via `applyToolCall(call)` (in services/excel).
 *   4. We record usage per invocation.
 *   5. We append a `role: "tool"` ChatMessage per result.
 *   6. We kick off a fresh `streamAssistantReply` so the model can react to
 *      the tool outputs.
 *
 * Recursion guard (Wave 3 fix):
 *   Wave 2 inline version had NO cap — a model that loops "call tool →
 *   call another tool → …" would blow the budget. We cap at
 *   MAX_TOOL_ROUNDS. Beyond the cap we drop further tool calls and surface
 *   a toast so the user knows we stopped.
 *
 * Hub-only dependency: this controller never imports ChatWindow.
 * ============================================================================
 */

import type { ChatMessage, ChatSession, ToolCall } from "../types";
import { applyToolCall } from "../services/excel";
import { generateId } from "../utils/helpers";
import { recordToolInvocation } from "../services/usage";
import type { ChatControllerHub } from "./ChatControllerHub";

/** Hard cap on consecutive tool-call rounds. Three is enough for ~all
 *  real Excel workflows (e.g. "scan PII → mask → write back"); beyond
 *  that the model is almost certainly looping. */
export const MAX_TOOL_ROUNDS = 3;

interface ExecuteOptions {
  /** Round index — increments each time we recurse into the stream. */
  round: number;
  /** The original user turn. Recursive streams reuse this so the API
   *  sees the same question context with growing tool results. */
  lastUserMessage: ChatMessage;
}

/** Handle returned to ChatWindow so it can ask the controller to start a
 *  brand-new tool round at round=0. */
export interface ToolCallDispatcher {
  /** Called when a fresh stream emits tool_calls. Owns the execution +
   *  recursion. */
  handle(calls: ToolCall[], assistantMsg: ChatMessage, session: ChatSession, lastUserMessage: ChatMessage): Promise<void>;
}

export class ToolCallController {
  constructor(private readonly hub: ChatControllerHub) {}

  /** Build a dispatcher bound to a particular session + assistant turn.
   *  The dispatcher increments its own round counter on each recursion so
   *  multiple concurrent branches don't trample each other's counters. */
  public buildDispatcher(): ToolCallDispatcher {
    let round = 0;
    const self = this;
    return {
      async handle(calls, assistantMsg, session, lastUserMessage) {
        await self.handleToolCalls(calls, assistantMsg, session, lastUserMessage, {
          round,
          lastUserMessage,
        });
        round = Math.min(round + 1, MAX_TOOL_ROUNDS);
      },
    };
  }

  /** Execute `calls`, append their results to the session as tool messages,
   *  then ask the hub to stream a follow-up reply. */
  public async handleToolCalls(
    calls: ToolCall[],
    assistantMsg: ChatMessage,
    session: ChatSession,
    lastUserMessage: ChatMessage,
    opts: ExecuteOptions
  ): Promise<void> {
    if (opts.round >= MAX_TOOL_ROUNDS) {
      this.hub.toast(
        `工具调用已超过 ${MAX_TOOL_ROUNDS} 轮，已停止以避免无限循环`,
        "error"
      );
      this.hub.setStatus("");
      return;
    }

    // Attach the calls to the assistant message so they render with the
    // "🔧 已调用工具 N 个" badge and persist on reload. Patch the existing
    // streaming view in place rather than creating a duplicate DOM node.
    assistantMsg.toolCalls = calls;
    session.messages[session.messages.length - 1] = { ...assistantMsg, toolCalls: calls };
    this.hub.upsertSession(session);
    const existing = this.hub.getMessageView?.(assistantMsg.id);
    if (existing) existing.update({ ...assistantMsg }, false);

    this.hub.setStatus(`正在调用工具 (${calls.length})…`);

    for (const call of calls) {
      const resultText = await applyToolCall(call);
      call.result = resultText;
      call.executed = true;
      recordToolInvocation(call.name);
      const toolMsg: ChatMessage = {
        id: generateId(),
        role: "tool",
        content: resultText,
        timestamp: Date.now(),
        toolCallId: call.id,
        toolName: call.name,
      };
      session.messages.push(toolMsg);
      this.hub.appendMessageView(toolMsg);
      this.hub.upsertSession(session);
    }

    this.hub.setStatus("正在根据工具结果生成回复…");
    await this.hub.streamAssistantReply(session, opts.lastUserMessage);
  }
}