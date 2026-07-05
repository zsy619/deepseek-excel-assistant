/**
 * ============================================================================
 * ribbonActions — lazy-loaded barrel for the 5 advanced Ribbon commands
 * ----------------------------------------------------------------------------
 * Loaded on-demand via dynamic `import("./ribbonActions")` from ChatWindow,
 * which lets webpack carve it into its own async chunk and keeps the main
 * taskpane bundle small. Only fetched the first time the user clicks one of
 * the advanced ribbon buttons (Correlation / Outliers / Pivot / Report /
 * ColumnTypes), not on taskpane open.
 * ============================================================================
 */

// 5 AI stream helpers (DeepSeek JSON streaming)
export {
  correlationStream,
  outlierStream,
  pivotSpecStream,
  reportStream,
  columnTypeStream,
} from "./services/deepseek";

// 5 Apply functions (real Excel write operations)
export {
  highlightOutliers,
  createPivotTable,
  writeReportSheet,
  applyColumnFormatting,
  insertCorrelationMatrix,
} from "./services/excel";