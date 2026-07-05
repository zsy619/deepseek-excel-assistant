/**
 * ============================================================================
 * MarkdownRenderer
 * ----------------------------------------------------------------------------
 * Stateless utility that turns a markdown string into safe HTML, applies
 * syntax highlighting to fenced code blocks, and decorates every code block
 * with a copy button.
 *
 * Important: we sanitize user-supplied content. Although the model's output
 * is already markdown, system prompts and user inputs can contain raw HTML
 * which we never want to inject into the DOM as-is.
 * ============================================================================
 */

import { marked } from "marked";
import hljs from "highlight.js/lib/core";

// Register the languages Excel users actually need. Adding more here
// increases the bundle size by ~5-10KB each - keep it lean.
import javascript from "highlight.js/lib/languages/javascript";
import typescript from "highlight.js/lib/languages/typescript";
import python from "highlight.js/lib/languages/python";
import json from "highlight.js/lib/languages/json";
import xml from "highlight.js/lib/languages/xml";
import css from "highlight.js/lib/languages/css";
import bash from "highlight.js/lib/languages/bash";
import sql from "highlight.js/lib/languages/sql";
import excel from "highlight.js/lib/languages/excel";

hljs.registerLanguage("javascript", javascript);
hljs.registerLanguage("js", javascript);
hljs.registerLanguage("typescript", typescript);
hljs.registerLanguage("ts", typescript);
hljs.registerLanguage("python", python);
hljs.registerLanguage("py", python);
hljs.registerLanguage("json", json);
hljs.registerLanguage("xml", xml);
hljs.registerLanguage("html", xml);
hljs.registerLanguage("css", css);
hljs.registerLanguage("bash", bash);
hljs.registerLanguage("sh", bash);
hljs.registerLanguage("sql", sql);
hljs.registerLanguage("excel", excel);

/** Custom renderer that hooks into marked to:
 *  - add language classes for hljs
 *  - inject a copy button into each code block
 *  - rewrite links to open in a new tab
 *  - escape raw HTML by default
 *
 * marked v12 passes primitives (strings, booleans) into renderer
 * methods, NOT token objects, so we cast arguments carefully. */
const renderer = new marked.Renderer();

renderer.code = function (code: string, infostring: string | undefined): string {
  const language = (infostring || "").trim().split(/\s+/)[0] || "plaintext";

  // Mermaid blocks get a placeholder; the actual SVG is rendered by
  // renderMermaidDiagrams() once the message is mounted in the DOM.
  if (language === "mermaid") {
    return (
      `<div class="md-mermaid">` +
      `<div class="md-mermaid-header">` +
      `<span class="md-code-lang">mermaid</span>` +
      `<button type="button" class="md-copy-btn" data-code="${encodeAttr(code)}" title="复制源码">复制源码</button>` +
      `</div>` +
      `<div class="md-mermaid-source" data-source="${encodeAttr(code)}">` +
      `<pre class="mermaid-fallback">${escapeHtml(code)}</pre>` +
      `</div>` +
      `</div>`
    );
  }

  let highlighted: string;
  try {
    if (language && hljs.getLanguage(language)) {
      highlighted = hljs.highlight(code, { language, ignoreIllegals: true }).value;
    } else {
      highlighted = escapeHtml(code);
    }
  } catch {
    highlighted = escapeHtml(code);
  }
  const safeLang = escapeAttr(language);
  return (
    `<div class="md-code-block">` +
    `<div class="md-code-header">` +
    `<span class="md-code-lang">${safeLang}</span>` +
    `<button type="button" class="md-copy-btn" data-code="${encodeAttr(code)}" title="复制代码">复制</button>` +
    `</div>` +
    `<pre class="hljs"><code class="language-${safeLang}">${highlighted}</code></pre>` +
    `</div>`
  );
};

renderer.link = function (href: string, title: string | null | undefined, text: string): string {
  const safeHref = escapeAttr(href || "#");
  const titleAttr = title ? ` title="${escapeAttr(title)}"` : "";
  return `<a href="${safeHref}"${titleAttr} target="_blank" rel="noopener noreferrer">${text}</a>`;
};

renderer.image = function (href: string, title: string | null, text: string): string {
  const safeHref = escapeAttr(href || "");
  const titleAttr = title ? ` title="${escapeAttr(title)}"` : "";
  const alt = escapeAttr(text || "");
  return `<img src="${safeHref}"${titleAttr} alt="${alt}" loading="lazy" />`;
};

renderer.table = function (header: string, body: string): string {
  return `<div class="md-table-wrap"><table class="md-table"><thead>${header}</thead><tbody>${body}</tbody></table></div>`;
};

// Configure marked once at module load.
marked.setOptions({
  renderer,
  gfm: true,
  breaks: true,
  pedantic: false,
});

/** Render a markdown string to a safe HTML string. */
export function renderMarkdown(input: string): string {
  if (!input) return "";
  try {
    // marked v12 returns string when async:false (default).
    const html = marked.parse(input, { async: false }) as string;
    return html;
  } catch (err) {
    // If parsing fails for any reason fall back to escaped text.
    return `<pre>${escapeHtml(input)}</pre>`;
  }
}

/** Strip markdown syntax and return plain text. Useful when "insert to cell"
 *  should drop formatting markers rather than carry them into Excel.
 *
 *  NOTE: this pure-regex version now lives in utils/markdownText.ts so the
 *  context bar / insert-to-cell flow doesn't have to pull in marked + hljs.
 *  It's re-exported here for callers that already import from this module. */
export { markdownToPlainText } from "../utils/markdownText";

/** Highlight.js instance for direct use when needed (e.g. raw text rendering). */
export function highlightCode(text: string, language?: string): string {
  try {
    if (language && hljs.getLanguage(language)) {
      return hljs.highlight(text, { language, ignoreIllegals: true }).value;
    }
  } catch {
    /* fall through */
  }
  return escapeHtml(text);
}

/* ---------------------------------------------------------------- *
 * Mermaid diagram rendering                                         *
 * ---------------------------------------------------------------- */

/** Track the loading state so we only fetch the library once. */
let mermaidLoadPromise: Promise<any> | null = null;
let mermaidIdCounter = 0;

/**
 * Load the Mermaid library from a CDN. The Office webview sometimes
 * drops webpack-injected scripts for sibling packages, so we attach
 * the <script> tag manually after the document is ready.
 */
function loadMermaid(): Promise<any> {
  if (mermaidLoadPromise) return mermaidLoadPromise;
  mermaidLoadPromise = new Promise((resolve, reject) => {
    const existing = (window as any).mermaid;
    if (existing && typeof existing.render === "function") {
      resolve(existing);
      return;
    }
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/mermaid@10.9.1/dist/mermaid.min.js";
    script.async = true;
    script.onload = () => {
      const m = (window as any).mermaid;
      if (m && typeof m.initialize === "function") {
        try {
          const theme =
            document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "default";
          m.initialize({
            startOnLoad: false,
            theme,
            securityLevel: "strict",
            fontFamily: "inherit",
          });
        } catch {
          /* noop */
        }
      }
      resolve(m);
    };
    script.onerror = () => reject(new Error("Mermaid CDN load failed"));
    document.head.appendChild(script);
  });
  return mermaidLoadPromise;
}

/**
 * Find every .md-mermaid-source placeholder inside `container` and
 * replace it with the rendered SVG. Safe to call repeatedly; already
 * rendered diagrams are skipped.
 */
export async function renderMermaidDiagrams(container: HTMLElement): Promise<void> {
  if (!container) return;
  const placeholders = Array.from(
    container.querySelectorAll<HTMLElement>(".md-mermaid-source:not([data-rendered])")
  );
  if (placeholders.length === 0) return;

  let mermaid: any;
  try {
    mermaid = await loadMermaid();
  } catch {
    placeholders.forEach((p) => {
      const note = document.createElement("div");
      note.className = "md-mermaid-error";
      note.textContent = "图表渲染失败：无法加载 Mermaid 库";
      p.appendChild(note);
    });
    return;
  }

  for (const el of placeholders) {
    const source = el.dataset.source || "";
    if (!source.trim()) continue;
    el.dataset.rendered = "1";
    const id = `mermaid-${Date.now()}-${++mermaidIdCounter}`;
    try {
      const { svg } = await mermaid.render(id, source);
      el.innerHTML = svg;
      el.classList.add("md-mermaid-rendered");
    } catch (err: any) {
      el.innerHTML = "";
      const note = document.createElement("div");
      note.className = "md-mermaid-error";
      note.textContent = `图表语法错误：${err?.message || "未知错误"}`;
      el.appendChild(note);
    }
  }
}

/* ---------------------------------------------------------------- *
 * HTML safety helpers                                             *
 * ---------------------------------------------------------------- */

function escapeHtml(s: string): string {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttr(s: string): string {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/** Encode for use in data-* attribute (preserves quotes inside the value). */
function encodeAttr(s: string): string {
  return encodeURIComponent(s);
}