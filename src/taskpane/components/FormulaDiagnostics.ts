/**
 * ============================================================================
 * FormulaDiagnostics
 * ----------------------------------------------------------------------------
 * Renders the inline diagnostic report inside the chat window. One card
 * per scanned cell, with the original formula, the AI's analysis, and a
 * one-click "Apply Fix" button that overwrites the formula in Excel.
 *
 * The component is pure: it takes the scan + diagnosis result and emits
 * `apply-fix` / `retry` CustomEvents for the ChatWindow to forward into
 * the excel service.
 * ============================================================================
 */

import type { FormulaDiagnosis } from "../services/deepseek";
import type { FormulaScanResult } from "../services/excel";
import { FORMULA_ERROR_INFO } from "../utils/constants";

export interface FormulaDiagnosticCardEventDetail {
  kind: "apply-fix" | "retry";
  diagnosis: FormulaDiagnosis;
  scan: FormulaScanResult;
}

export class FormulaDiagnosticsView {
  private root: HTMLElement;

  constructor(
    parent: HTMLElement,
    private scan: FormulaScanResult,
    private diagnoses: FormulaDiagnosis[]
  ) {
    this.root = document.createElement("div");
    this.root.className = "formula-diagnostics";
    parent.appendChild(this.root);
    this.render();
  }

  /** Update the diagnosis list in place. Called when streaming completes
   *  or the user clicks "Retry" to re-run the AI. */
  public setDiagnoses(diagnoses: FormulaDiagnosis[]): void {
    this.diagnoses = diagnoses;
    this.render();
  }

  public destroy(): void {
    this.root.remove();
  }

  /* ----------------------------------------------------------- */

  private render(): void {
    this.root.innerHTML = "";
    this.root.appendChild(this.renderHeader());
    if (this.diagnoses.length === 0) {
      this.root.appendChild(this.renderEmpty());
      return;
    }
    const list = document.createElement("div");
    list.className = "formula-diagnostics-list";
    for (const d of this.diagnoses) {
      list.appendChild(this.renderCard(d));
    }
    this.root.appendChild(list);
  }

  private renderHeader(): HTMLElement {
    const head = document.createElement("div");
    head.className = "formula-diagnostics-head";
    const summary = document.createElement("div");
    summary.className = "formula-diagnostics-summary";
    const errCount = this.scan.cells.length;
    summary.innerHTML = [
      `<strong>公式诊断报告</strong>`,
      `<span class="fd-meta">${this.scan.sheetName} · ${this.scan.rangeAddress}</span>`,
      `<span class="fd-count">扫描 ${this.scan.totalFormulas} 个公式，发现 <b>${errCount}</b> 个错误</span>`,
    ].join(" · ");
    head.appendChild(summary);
    return head;
  }

  private renderEmpty(): HTMLElement {
    const empty = document.createElement("div");
    empty.className = "formula-diagnostics-empty";
    empty.textContent = "AI 暂时没有给出可解析的诊断结果。可缩小选区后重试。";
    return empty;
  }

  private renderCard(d: FormulaDiagnosis): HTMLElement {
    const card = document.createElement("div");
    card.className = "formula-diagnostics-card";

    const info = FORMULA_ERROR_INFO[d.error] || {
      code: d.error,
      label: d.error || "未知错误",
      reason: "",
    };

    const header = document.createElement("div");
    header.className = "fd-card-header";
    header.innerHTML =
      `<span class="fd-addr">${escapeHtml(d.address || "(未知)")}</span>` +
      `<span class="fd-badge fd-badge-${info.code.toLowerCase()}">${escapeHtml(info.label)}</span>` +
      (typeof d.confidence === "number"
        ? `<span class="fd-conf">置信度 ${(d.confidence * 100).toFixed(0)}%</span>`
        : "");
    card.appendChild(header);

    const body = document.createElement("div");
    body.className = "fd-card-body";

    const formulaSection = document.createElement("div");
    formulaSection.className = "fd-section";
    formulaSection.innerHTML =
      `<div class="fd-label">当前公式</div>` +
      `<pre class="fd-formula">${escapeHtml(this.findOriginalFormula(d.address))}</pre>`;
    body.appendChild(formulaSection);

    if (info.reason) {
      const reasonSection = document.createElement("div");
      reasonSection.className = "fd-section fd-reason";
      reasonSection.innerHTML =
        `<div class="fd-label">常见原因</div><div>${escapeHtml(info.reason)}</div>`;
      body.appendChild(reasonSection);
    }

    if (d.cause) {
      const causeSection = document.createElement("div");
      causeSection.className = "fd-section fd-cause";
      causeSection.innerHTML =
        `<div class="fd-label">AI 诊断</div><div>${escapeHtml(d.cause)}</div>`;
      body.appendChild(causeSection);
    }

    if (d.suggestion) {
      const suggestSection = document.createElement("div");
      suggestSection.className = "fd-section fd-suggestion";
      suggestSection.innerHTML =
        `<div class="fd-label">修复建议</div><div>${escapeHtml(d.suggestion)}</div>`;
      body.appendChild(suggestSection);
    }

    if (d.fixedFormula) {
      const fixedSection = document.createElement("div");
      fixedSection.className = "fd-section fd-fixed";
      fixedSection.innerHTML =
        `<div class="fd-label">建议公式</div>` +
        `<pre class="fd-formula fd-formula-fixed">${escapeHtml(d.fixedFormula)}</pre>`;
      body.appendChild(fixedSection);

      const actions = document.createElement("div");
      actions.className = "fd-actions";
      const apply = document.createElement("button");
      apply.className = "fd-btn fd-btn-primary";
      apply.textContent = "应用修复";
      apply.addEventListener("click", () => this.emit({ kind: "apply-fix" }));
      actions.appendChild(apply);

      const copy = document.createElement("button");
      copy.className = "fd-btn";
      copy.textContent = "复制公式";
      copy.addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(d.fixedFormula);
          apply.textContent = "已复制 ✓";
          setTimeout(() => (apply.textContent = "应用修复"), 1500);
        } catch {
          /* noop */
        }
      });
      actions.appendChild(copy);

      body.appendChild(actions);
    } else {
      const actions = document.createElement("div");
      actions.className = "fd-actions";
      const retry = document.createElement("button");
      retry.className = "fd-btn";
      retry.textContent = "重新诊断";
      retry.addEventListener("click", () => this.emit({ kind: "retry" }));
      actions.appendChild(retry);
      body.appendChild(actions);
    }

    card.appendChild(body);
    return card;
  }

  private findOriginalFormula(address: string): string {
    const cell = this.scan.cells.find((c) => c.address === address);
    return cell ? cell.formula : "(未找到原始公式)";
  }

  private emit(detail: { kind: "apply-fix" | "retry"; diagnosis?: FormulaDiagnosis }): void {
    this.root.dispatchEvent(
      new CustomEvent<FormulaDiagnosticCardEventDetail>("formula-diagnostic-action", {
        detail: { ...detail, diagnosis: detail.diagnosis!, scan: this.scan } as FormulaDiagnosticCardEventDetail,
        bubbles: true,
      })
    );
  }
}

function escapeHtml(s: string): string {
  if (s == null) return "";
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
