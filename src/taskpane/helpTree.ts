/**
 * ============================================================================
 * helpTree — declarative structure of the help documentation
 * ----------------------------------------------------------------------------
 * Lists every HTML page in the docs/ bundle with its title, prev/next
 * siblings, and language. The HelpPanel uses this to render the TOC sidebar
 * and prev/next buttons without hard-coding paths in the component.
 *
 * Order matters: prev/next follows the order declared here.
 * ============================================================================
 */

export interface HelpNode {
  /** Path relative to docs/, e.g. "zh/guide.html". */
  path: string;
  /** Title shown in the TOC and the page header. */
  title: string;
  /** Language tag. */
  lang: "zh" | "en";
}

export const HELP_TREE: HelpNode[] = [
  { path: "zh/index.html",             title: "📖 目录",            lang: "zh" },
  { path: "zh/guide.html",             title: "🚀 用户指南",         lang: "zh" },
  { path: "zh/buttons.html",           title: "🎛️ 按钮参考",         lang: "zh" },
  { path: "zh/faq.html",               title: "❓ 常见问题",          lang: "zh" },
  { path: "zh/troubleshooting.html",   title: "🔧 疑难排查",          lang: "zh" },

  { path: "en/index.html",             title: "📖 Index",            lang: "en" },
  { path: "en/guide.html",             title: "🚀 User Guide",       lang: "en" },
  { path: "en/buttons.html",           title: "🎛️ Button Reference", lang: "en" },
  { path: "en/faq.html",               title: "❓ FAQ",               lang: "en" },
  { path: "en/troubleshooting.html",   title: "🔧 Troubleshooting",  lang: "en" },
];

/** Return the prev/next HelpNode siblings of `path`, or null at edges. */
export function getSiblings(path: string): { prev: HelpNode | null; next: HelpNode | null } {
  const idx = HELP_TREE.findIndex((n) => n.path === path);
  if (idx === -1) return { prev: null, next: null };
  return {
    prev: idx > 0 ? HELP_TREE[idx - 1] : null,
    next: idx < HELP_TREE.length - 1 ? HELP_TREE[idx + 1] : null,
  };
}

/** Find the entry that matches the current path (used to set the default page). */
export function findByPath(path: string): HelpNode | undefined {
  return HELP_TREE.find((n) => n.path === path);
}

/** All nodes for a single language (drives the language-specific TOC). */
export function nodesForLang(lang: "zh" | "en"): HelpNode[] {
  return HELP_TREE.filter((n) => n.lang === lang);
}