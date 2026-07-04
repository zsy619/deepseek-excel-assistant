/**
 * ============================================================================
 * QuickActions
 * ----------------------------------------------------------------------------
 * Renders the row of shortcut buttons shown above the message input. Each
 * button dispatches a CustomEvent that the ChatWindow observes and turns
 * into a concrete action (e.g. fetch selected data, swap into formula mode).
 * ============================================================================
 */

import { QUICK_ACTION_TEMPLATES } from "../utils/constants";
import type { ExcelSelection } from "../types";
import { excelValuesToMarkdown } from "../utils/helpers";

export type QuickActionKind =
  | "analyze"
  | "formula"
  | "clean"
  | "insert"
  | "stop";

export interface QuickActionEventDetail {
  kind: QuickActionKind;
  template: string;
  selection?: ExcelSelection;
}

/**
 * Mount the quick action bar inside `container`. Returns a teardown
 * function that removes event listeners.
 */
export function mountQuickActions(container: HTMLElement): () => void {
  container.innerHTML = "";
  container.classList.add("quick-actions");

  const actions: Array<{
    key: QuickActionKind;
    icon: string;
    label: string;
    primary?: boolean;
  }> = [
    { key: "analyze", icon: "📊", label: "分析选区" },
    { key: "formula", icon: "📝", label: "生成公式" },
    { key: "clean", icon: "🧹", label: "数据清洗" },
    { key: "insert", icon: "📌", label: "插入单元格" },
  ];

  for (const a of actions) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.classList.add("quick-action-btn");
    btn.dataset.kind = a.key;
    btn.innerHTML = `<span class="qa-icon">${a.icon}</span><span class="qa-label">${a.label}</span>`;
    container.appendChild(btn);
  }

  // The "stop" button is hidden by default; ChatWindow shows it during streaming.
  const stopBtn = document.createElement("button");
  stopBtn.type = "button";
  stopBtn.classList.add("quick-action-btn", "quick-action-stop");
  stopBtn.dataset.kind = "stop";
  stopBtn.style.display = "none";
  stopBtn.innerHTML = `<span class="qa-icon">⏹</span><span class="qa-label">停止</span>`;
  container.appendChild(stopBtn);

  /** Dispatch helper - bubble up so ChatWindow can listen once. */
  const fire = (kind: QuickActionKind, template: string, selection?: ExcelSelection) => {
    const detail: QuickActionEventDetail = { kind, template, selection };
    container.dispatchEvent(new CustomEvent("quick-action", { detail, bubbles: true }));
  };

  const handler = async (ev: Event) => {
    const target = ev.target as HTMLElement;
    const btn = target.closest<HTMLButtonElement>(".quick-action-btn");
    if (!btn) return;
    const kind = btn.dataset.kind as QuickActionKind | undefined;
    if (!kind) return;

    if (kind === "stop") {
      fire("stop", "");
      return;
    }

    if (kind === "insert") {
      // No template, just signal the intent.
      fire("insert", "");
      return;
    }

    // All other actions may need selected data; the ChatWindow will pull
    // selection if needed. We just emit the bare template here so the
    // parent can choose how to enrich it.
    const template = QUICK_ACTION_TEMPLATES[kind as "analyze" | "formula" | "clean"] || "";
    fire(kind, template);
  };

  container.addEventListener("click", handler);

  return () => {
    container.removeEventListener("click", handler);
  };
}

/** Build the actual user prompt for an action, given current selection. */
export function buildActionPrompt(
  kind: QuickActionKind,
  selection?: ExcelSelection,
  userInput?: string
): string {
  if (kind === "insert") return "";
  const tpl =
    QUICK_ACTION_TEMPLATES[kind as "analyze" | "formula" | "clean"] || "{USER_INPUT}";

  const ctx =
    selection && selection.values && selection.values.length > 0
      ? `选区地址：${selection.address}\n\n数据：\n${excelValuesToMarkdown(selection.values)}`
      : "(当前未选中数据)";

  return tpl
    .replace("{CONTEXT}", ctx)
    .replace("{USER_INPUT}", userInput || "请帮我生成合适的公式");
}