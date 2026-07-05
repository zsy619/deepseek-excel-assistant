/**
 * ============================================================================
 * ApplySuggestionModal
 * ----------------------------------------------------------------------------
 * Two-step ("AI suggest → user reviews → user applies") modal shared by:
 *   - generateFormula  (writes a formula to one cell)
 *   - diagnoseFormulas (rewrites one or more formula cells)
 *   - cleanData        (writes cleaned values back to a range)
 *   - insertChart      (inserts a chart object)
 *   - maskPII          (writes masked values back to a range)
 *
 * Why this exists:
 *   Without it, every AI run that returns a write action would mutate the
 *   workbook immediately on the user pressing the ribbon button - bad UX,
 *   especially on shared workbooks. Pattern is:
 *
 *     user clicks ribbon
 *       → ribbon fires FunctionFile → cross-frame message
 *       → taskpane receives command + selected range + AI suggestion
 *       → ApplySuggestionModal.show({before, after, ...})
 *         shows diff + 3 buttons:
 *           [取消]               → discard suggestion
 *           [复制]               → copy suggestion to clipboard
 *           [✓ 应用]             → Excel.run() writes back, returns Promise<true>
 *
 *   The Apply button is the only path that touches the workbook.
 *
 * Two rendering modes:
 *   1. Cell-diff mode (formula / diagnose / clean) — renders an
 *      Office-Fluent table with "Before" / "After" columns side-by-side
 *      and red/green tinting on changed cells.
 *   2. Action mode (chart / pii) — renders a static description + a
 *      preview area + the same 3 buttons.
 *
 * Returns a Promise that resolves when the modal is dismissed:
 *   - "apply"  : user approved, caller should run the writer
 *   - "copy"   : user copied to clipboard
 *   - "cancel" : user dismissed (incl. Escape / click-outside)
 * ============================================================================
 */

export interface DiffCell {
  before: string;
  after: string;
  changed: boolean;
}
export interface CellDiffPayload {
  kind: "cell-diff";
  /** Cells to highlight. Rectangular range, top-left first. */
  cells: DiffCell[][];
  /** Optional message shown above the table (e.g. "AI will replace X with Y"). */
  summary?: string;
}

export interface ActionPayload {
  kind: "action";
  /** Title row above the preview. */
  title: string;
  /** Short, plain-text preview of what AI will write — kept short so it
   *  fits inside the modal without scrolling. Multi-line is fine. */
  preview: string;
  /** Optional larger rendered preview (HTML string, will be rendered as-is,
   *  for code listings — already escaped by caller). */
  previewHtml?: string;
  /** True for write operations that mutate the workbook. False for
   *  read-only or copy-only flows where Apply may be hidden. */
  requiresWrite?: boolean;
}

export type ApplySuggestionPayload = CellDiffPayload | ActionPayload;

export interface ShowOptions {
  title: string;
  /** The AI's reasoning for the change, shown as a small subtitle. */
  reasoning?: string;
}

export type ApplySuggestionResult = "apply" | "copy" | "cancel";

const HOST_ID = "deepseek-apply-modal-host";
const MODAL_TITLE = "deepseek-apply-modal";
const ESCAPE_KEY = "Escape";

/** Find or create the singleton host. */
function ensureHost(): HTMLElement {
  let host = document.getElementById(HOST_ID);
  if (!host) {
    host = document.createElement("div");
    host.id = HOST_ID;
    host.classList.add("apply-modal-host");
    document.body.appendChild(host);
  }
  return host;
}

/** Escape any user-controlled string before we inject it into the DOM. */
function esc(s: unknown): string {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    c === "&"
      ? "&amp;"
      : c === "<"
        ? "&lt;"
        : c === ">"
          ? "&gt;"
          : c === '"'
            ? "&quot;"
            : "&#39;"
  );
}

function buildFooter(
  requiresWrite: boolean,
  onApply: () => void,
  onCopy: () => void,
  onCancel: () => void
): HTMLElement {
  const footer = document.createElement("div");
  footer.className = "apply-modal__footer";

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "apply-modal__btn apply-modal__btn--ghost";
  cancelBtn.textContent = "取消";
  cancelBtn.addEventListener("click", onCancel);

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "apply-modal__btn apply-modal__btn--secondary";
  copyBtn.textContent = "复制到剪贴板";
  copyBtn.addEventListener("click", onCopy);

  footer.appendChild(cancelBtn);
  footer.appendChild(copyBtn);

  if (requiresWrite) {
    const applyBtn = document.createElement("button");
    applyBtn.type = "button";
    applyBtn.className = "apply-modal__btn apply-modal__btn--primary";
    applyBtn.textContent = "✓ 应用更改";
    applyBtn.addEventListener("click", onApply);
    footer.appendChild(applyBtn);
  }
  return footer;
}

function renderCellDiff(payload: CellDiffPayload): HTMLElement {
  const wrap = document.createElement("div");
  wrap.className = "apply-modal__diff";
  if (payload.summary) {
    const sum = document.createElement("p");
    sum.className = "apply-modal__summary";
    sum.textContent = payload.summary;
    wrap.appendChild(sum);
  }
  const table = document.createElement("table");
  table.className = "apply-modal__table";
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  for (const head of ["原值 (Before)", "AI 建议 (After)"]) {
    const th = document.createElement("th");
    th.textContent = head;
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  // Flatten across rows so the diff reads top-to-bottom.
  for (const row of payload.cells) {
    for (const cell of row) {
      const tr = document.createElement("tr");
      const td1 = document.createElement("td");
      td1.className = "apply-modal__cell apply-modal__cell--before";
      td1.textContent = cell.before;
      tr.appendChild(td1);
      const td2 = document.createElement("td");
      td2.className =
        "apply-modal__cell " +
        (cell.changed ? "apply-modal__cell--after" : "apply-modal__cell--after-unchanged");
      td2.textContent = cell.after;
      tr.appendChild(td2);
      tbody.appendChild(tr);
    }
  }
  table.appendChild(tbody);
  wrap.appendChild(table);
  return wrap;
}

function renderAction(payload: ActionPayload): HTMLElement {
  const wrap = document.createElement("div");
  wrap.className = "apply-modal__action";
  const title = document.createElement("h4");
  title.className = "apply-modal__action-title";
  title.textContent = payload.title;
  wrap.appendChild(title);
  const preview = document.createElement("pre");
  preview.className = "apply-modal__preview";
  preview.textContent = payload.preview;
  wrap.appendChild(preview);
  if (payload.previewHtml) {
    const extra = document.createElement("div");
    extra.className = "apply-modal__extra";
    extra.innerHTML = payload.previewHtml; // caller escaped already
    wrap.appendChild(extra);
  }
  return wrap;
}

function buildShell(
  body: HTMLElement,
  reasoning: string | undefined,
  footer: HTMLElement
): HTMLElement {
  const overlay = document.createElement("div");
  overlay.className = "apply-modal__overlay";
  overlay.id = MODAL_TITLE;

  const card = document.createElement("div");
  card.className = "apply-modal__card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");

  const head = document.createElement("div");
  head.className = "apply-modal__head";
  const h = document.createElement("h3");
  h.textContent = "AI 建议";
  head.appendChild(h);
  if (reasoning) {
    const p = document.createElement("p");
    p.className = "apply-modal__reasoning";
    p.textContent = reasoning;
    head.appendChild(p);
  }

  card.appendChild(head);
  card.appendChild(body);
  card.appendChild(footer);
  overlay.appendChild(card);
  return overlay;
}

async function copyToClipboard(text: string): Promise<boolean> {
  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    /* noop - fall through to legacy path */
  }
  // Legacy fallback (Office webview sometimes lacks Clipboard API).
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch {
    return false;
  }
}

/**
 * Render the modal and resolve when the user dismisses it. Returns one of:
 *   - "apply"  : the caller should run the writer
 *   - "copy"   : the suggestion has been copied to clipboard
 *   - "cancel" : dismissed without action
 */
export function showApplySuggestion(
  payload: ApplySuggestionPayload,
  options: ShowOptions
): Promise<ApplySuggestionResult> {
  return new Promise((resolve) => {
    const host = ensureHost();
    const body =
      payload.kind === "cell-diff"
        ? renderCellDiff(payload)
        : renderAction(payload);
    const requiresWrite = payload.kind === "cell-diff" || !!payload.requiresWrite;

    let resolved = false;
    const cleanup = () => {
      if (resolved) return;
      resolved = true;
      try {
        overlay.remove();
      } catch {
        /* noop */
      }
      try {
        unbindKey();
      } catch {
        /* noop */
      }
    };
    const finish = (val: ApplySuggestionResult) => {
      cleanup();
      resolve(val);
    };

    const onCopy = () => {
      const text =
        payload.kind === "cell-diff"
          ? payload.cells.map((row) => row.map((c) => c.after).join("\t")).join("\n")
          : payload.preview;
      void copyToClipboard(text).then((ok) => {
        if (!ok) {
          try {
            console.warn("[DeepSeek] clipboard copy failed");
          } catch {
            /* noop */
          }
        }
        finish("copy");
      });
    };
    const onApply = () => finish("apply");
    const onCancel = () => finish("cancel");

    const footer = buildFooter(requiresWrite, onApply, onCopy, onCancel);
    const overlay = buildShell(body, options.reasoning, footer);
    host.appendChild(overlay);

    // Wire escape and click-outside.
    const onKey = (e: KeyboardEvent) => {
      if (e.key === ESCAPE_KEY) {
        e.preventDefault();
        onCancel();
      }
    };
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) onCancel();
    });
    document.addEventListener("keydown", onKey);
    const unbindKey = () => document.removeEventListener("keydown", onKey);

    // Auto-focus the most "destructive" button first to discourage
    // accidental application — but only the safest, Copy.
    setTimeout(() => {
      const copy = overlay.querySelector<HTMLButtonElement>(
        ".apply-modal__btn--secondary"
      );
      copy?.focus();
    }, 50);
  });
}
