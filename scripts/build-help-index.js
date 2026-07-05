#!/usr/bin/env node
/**
 * ============================================================================
 * build-help-index
 * ----------------------------------------------------------------------------
 * Walks docs HTML files, strips tags, writes search-index.json.
 * The HelpPanel loads this file on demand and feeds it to fuse.js for
 * client-side fuzzy search.
 *
 * Run via:  node scripts/build-help-index.js
 * Or wired into npm run build before webpack.
 * ============================================================================
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const DOCS_DIR = path.join(ROOT, "docs");
const OUT = path.join(DOCS_DIR, "search-index.json");

/** Strip HTML tags, decode entities, collapse whitespace. */
function stripHtml(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<\/?[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function extractTitle(html) {
  const m = html.match(/<title>([\s\S]*?)<\/title>/i);
  return m ? m[1].trim() : "";
}

function extractH1(html) {
  const m = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  return m ? stripHtml(m[1]) : "";
}

function extractSection(html) {
  // Use the first heading to label the doc in search results.
  return extractH1(html) || extractTitle(html);
}

function walk(dir) {
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...walk(full));
    } else if (entry.isFile() && entry.name.endsWith(".html")) {
      out.push(full);
    }
  }
  return out;
}

function build() {
  if (!fs.existsSync(DOCS_DIR)) {
    console.error(`[build-help-index] docs/ not found: ${DOCS_DIR}`);
    process.exit(1);
  }
  const files = walk(DOCS_DIR);
  const entries = [];
  for (const file of files) {
    const rel = path.relative(DOCS_DIR, file).replace(/\\/g, "/");
    const html = fs.readFileSync(file, "utf8");
    const section = extractSection(html);
    const body = stripHtml(html);
    // Keep body under ~8KB per doc to bound the index size.
    entries.push({
      path: rel,
      lang: rel.startsWith("zh/") ? "zh" : rel.startsWith("en/") ? "en" : "",
      section,
      body: body.slice(0, 8000),
    });
  }
  // Sort for deterministic output.
  entries.sort((a, b) => a.path.localeCompare(b.path));
  fs.writeFileSync(OUT, JSON.stringify(entries, null, 2), "utf8");
  console.log(
    `[build-help-index] wrote ${entries.length} entries → ${path.relative(ROOT, OUT)}`
  );
}

build();