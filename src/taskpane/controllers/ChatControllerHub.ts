/**
 * ============================================================================
 * ChatControllerHub
 * ----------------------------------------------------------------------------
 * Single dependency surface shared by all per-feature controllers.
 *
 * Each controller (BranchController, ToolCallController, KnowledgeInjector,
 * UsageRecorder, ...) depends ONLY on this hub, never on ChatWindow
 * directly. That keeps the controllers unit-testable and stops the
 * "ChatWindow god-object" creep that Wave 2 left behind.
 *
 * Hub responsibilities:
 *   - hold the active session + config (read-only accessors)
 *   - persist session mutations
 *   - append / patch rendered message views
 *   - surface user-visible feedback (toast / status)
 *   - kick off another round of streaming
 *
 * Anything that doesn't fit these buckets stays in ChatWindow.
 * ============================================================================
 */

import type { ChatMessage, ChatSession, DeepSeekConfig } from "../types";

/** Minimal handle the hub exposes to controllers for stream kickoff. */
export interface StreamRequest {
  /** Caller wants a fresh assistant reply on `lastUserMessage`. */
  resume(lastUserMessage: ChatMessage): Promise<void>;
}

export interface ChatControllerHub {
  /* ---- session ---- */
  getActiveSession(): ChatSession | null;
  getConfig(): DeepSeekConfig;
  upsertSession(session: ChatSession): void;

  /* ---- rendering ---- */
  appendMessageView(message: ChatMessage): { update(msg: ChatMessage, streaming: boolean): void };
  /** Look up an already-rendered message view by id (returns undefined if
   *  it was never appended — e.g. for messages only present in storage). */
  getMessageView?(messageId: string): { update(msg: ChatMessage, streaming: boolean): void } | undefined;
  hideMessageView(messageId: string): void;
  appendListChild(el: HTMLElement): void;
  scrollToBottom(): void;

  /* ---- feedback ---- */
  toast(message: string, kind?: "info" | "error" | "success"): void;
  setStatus(text: string): void;

  /* ---- streaming ---- */
  streamAssistantReply(
    session: ChatSession,
    lastUserMessage: ChatMessage
  ): Promise<void>;
}