/**
 * ============================================================================
 * markdownText — lightweight markdown-to-plain-text utility
 * ----------------------------------------------------------------------------
 * Pure-regex converter with no external deps (no `marked`, no `highlight.js`).
 * Stays in the main bundle so the context bar, message actions, and
 * "insert to cell" can strip formatting without pulling in the heavy
 * MarkdownRenderer (which lazy-loads marked + hljs + 9 language modules).
 * ============================================================================
 */

/** Strip markdown syntax and return plain text. Used when "insert to cell"
 *  should drop formatting markers rather than carry them into Excel. */
export function markdownToPlainText(input: string): string {
  if (!input) return "";
  return input
    // fenced code blocks -> content
    .replace(/```[a-zA-Z]*\n([\s\S]*?)```/g, "$1")
    // inline code -> content
    .replace(/`([^`]+)`/g, "$1")
    // bold / italic
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/_([^_]+)_/g, "$1")
    // links -> visible text
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    // headings / blockquotes / lists markers
    .replace(/^#+\s+/gm, "")
    .replace(/^>\s?/gm, "")
    .replace(/^[-*+]\s+/gm, "")
    .replace(/^\d+\.\s+/gm, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}