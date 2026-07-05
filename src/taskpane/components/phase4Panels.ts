/**
 * ============================================================================
 * phase4Panels — lazy-loaded bundle of the 5 advanced-analysis panels
 * ----------------------------------------------------------------------------
 * Carved into its own webpack chunk so the main taskpane bundle stays lean.
 * Loaded on first click of one of the 5 Phase-4 ribbon buttons (Correlation
 * / Outliers / Pivot / Report / ColumnTypes).
 *
 * The factory wires each panel into the right DOM mount point and attaches
 * the click handlers — the parent ChatWindow only has to call
 * `ensurePhase4Panels()` once, then read `panels.correlation` etc.
 * ============================================================================
 */

import { CorrelationMatrixPanel } from "./CorrelationMatrixPanel";
import { OutlierPanel } from "./OutlierPanel";
import { PivotBuilderPanel } from "./PivotBuilderPanel";
import { ReportBuilderPanel } from "./ReportBuilderPanel";
import { ColumnTypePanel } from "./ColumnTypePanel";

export interface Phase4Panels {
  correlation: CorrelationMatrixPanel;
  outlier: OutlierPanel;
  pivot: PivotBuilderPanel;
  report: ReportBuilderPanel;
  colType: ColumnTypePanel;
}

export interface Phase4PanelHandlers {
  onCorrelationClick: (ev: MouseEvent) => void;
  onOutlierClick: (ev: MouseEvent) => void;
  onPivotClick: (ev: MouseEvent) => void;
  onReportClick: (ev: MouseEvent) => void;
  onColumnTypeClick: (ev: MouseEvent) => void;
}

/** Mount all 5 panels into the DOM mount points and wire delegated clicks.
 *  Returns the panel instances so the caller can drive show/getCurrentPayload. */
export function mountPhase4Panels(root: HTMLElement, handlers: Phase4PanelHandlers): Phase4Panels {
  const correlation = new CorrelationMatrixPanel();
  const outlier = new OutlierPanel();
  const pivot = new PivotBuilderPanel();
  const report = new ReportBuilderPanel();
  const colType = new ColumnTypePanel();

  const findMount = (ref: string): HTMLElement => {
    const el = root.querySelector<HTMLElement>(`[data-ref="${ref}"]`);
    if (!el) throw new Error(`phase4Panels: mount point [data-ref="${ref}"] not found`);
    return el;
  };

  findMount("correlationPanel").appendChild(correlation.root);
  findMount("outlierPanel").appendChild(outlier.root);
  findMount("pivotPanel").appendChild(pivot.root);
  findMount("reportPanel").appendChild(report.root);
  findMount("colTypePanel").appendChild(colType.root);

  correlation.root.addEventListener("click", handlers.onCorrelationClick);
  outlier.root.addEventListener("click", handlers.onOutlierClick);
  pivot.root.addEventListener("click", handlers.onPivotClick);
  report.root.addEventListener("click", handlers.onReportClick);
  colType.root.addEventListener("click", handlers.onColumnTypeClick);

  return { correlation, outlier, pivot, report, colType };
}