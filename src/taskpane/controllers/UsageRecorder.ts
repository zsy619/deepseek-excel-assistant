/**
 * ============================================================================
 * UsageRecorder (PRD-12)
 * ----------------------------------------------------------------------------
 * Thin wrapper over the four usage-tracking entry points in services/usage.
 * Lives behind the hub so callers don't import the service directly and so
 * the heuristics live in ONE place.
 *
 * Today the controller exposes one new method, `recordApiCall()`, that
 * ChatWindow calls in two spots: on a successful stream and on an errored
 * one. It computes token estimates from the prompt + completion text using
 * a blended CJK/Latin heuristic — much more accurate than the old
 * chars/3 shortcut, especially for Chinese Excel workbooks where the model
 * is reading/writing lots of CJK cell labels.
 *
 * Token-estimate heuristic (Wave 3 fix):
 *   - Latin runs (a-zA-Z0-9, whitespace, punctuation): 1 token ≈ 4 chars
 *   - CJK runs: 1 token ≈ 1.5 chars (DeepSeek's BPE is denser per char)
 *   - Blend by simple regex pass; cheap, no tokenizer dependency.
 *
 * Hub-only dependency: this controller never imports ChatWindow.
 * ============================================================================
 */

import { recordRequest } from "../services/usage";

export interface ApiCallRecord {
  /** Model id (e.g. "deepseek-chat" or "deepseek-reasoner") — drives the
   *  cost-per-1M lookup in services/usage. */
  model: string;
  /** Concatenated prompt text we sent to the model. */
  promptText: string;
  /** Final completion text we received. */
  finalText: string;
  /** Set when the request errored — record zero tokens + errored flag. */
  errored?: boolean;
}

export class UsageRecorder {
  /** Single entry point used by ChatWindow's stream lifecycle. */
  public recordApiCall(record: ApiCallRecord): void {
    if (record.errored) {
      recordRequest({
        model: record.model,
        promptTokens: 0,
        completionTokens: 0,
        errored: true,
      });
      return;
    }
    const promptTokens = Math.max(1, this.estimate(record.promptText));
    const completionTokens = Math.max(1, this.estimate(record.finalText));
    recordRequest({
      model: record.model,
      promptTokens,
      completionTokens,
    });
  }

  /**
   * Estimate token count for `text` using a simple per-character heuristic
   * blended across Latin and CJK runs. No tokenizer import — that would
   * bloat the bundle by ~10MB and we only need ballpark accuracy for the
   * usage dashboard.
   */
  public estimate(text: string): number {
    if (!text) return 0;
    let cjk = 0;
    let other = 0;
    for (const ch of text) {
      // CJK Unified Ideographs + Hiragana + Katakana + CJK Symbols.
      // Roughly matches what DeepSeek's BPE ends up tokenizing as ≥2 bytes.
      const code = ch.codePointAt(0) ?? 0;
      if (
        (code >= 0x4e00 && code <= 0x9fff) ||
        (code >= 0x3040 && code <= 0x30ff) ||
        (code >= 0x3400 && code <= 0x4dbf)
      ) {
        cjk++;
      } else {
        other++;
      }
    }
    // CJK ≈ 1 token / 1.5 chars; Latin ≈ 1 token / 4 chars.
    const cjkTokens = cjk / 1.5;
    const otherTokens = other / 4;
    return Math.ceil(cjkTokens + otherTokens);
  }

  /** Helper to concatenate all message contents into one string for
   *  token estimation. Public so ChatWindow can hand us the prompt without
   *  re-implementing the join. */
  public concatContents(parts: Array<string | undefined | null>): string {
    return parts.filter((p): p is string => typeof p === "string" && p.length > 0).join("\n");
  }
}