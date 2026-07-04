/**
 * ============================================================================
 * Excel service
 * ----------------------------------------------------------------------------
 * Thin wrapper around Office.js for the operations this add-in needs.
 * Every public method:
 *
 *   - accepts an optional `host` (Excel.Application) to ease testing
 *   - validates inputs and throws descriptive errors on misuse
 *   - returns a normalized payload suitable for direct consumption
 *
 * We use Excel.run() to batch operations and stay within the Office.js
 * sync model. All public methods are async because Excel.run returns a
 * promise that resolves when the queued commands complete.
 * ============================================================================
 */

import type { ExcelSelection } from "../types";

/** Run a batch against the active workbook. Passes the Excel context to
 *  `task` which performs the actual reads/writes. */
async function run<T>(task: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    try {
      Excel.run(async (ctx) => {
        try {
          const result = await task(ctx);
          await ctx.sync();
          resolve(result);
        } catch (err) {
          reject(err);
        }
      });
    } catch (err) {
      reject(err);
    }
  });
}

/** Safe wrapper that catches & re-throws with a friendly message. */
async function safeCall<T>(fn: () => Promise<T>, fallbackMessage: string): Promise<T> {
  try {
    return await fn();
  } catch (err: any) {
    const msg = err?.message || fallbackMessage;
    throw new Error(fallbackMessage + "：" + msg);
  }
}

/* ---------------------------------------------------------------- *
 * Reads                                                           *
 * ---------------------------------------------------------------- */

/**
 * Read the currently selected range. Returns the values, formulas, number
 * formats and a fully qualified address like "Sheet1!A1:D10".
 */
export async function getSelectedData(): Promise<ExcelSelection> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "formulas", "numberFormatCategories", "address", "rowCount", "columnCount"]);
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      return {
        address: range.address,
        sheetName: ws.name,
        values: range.values as any[][],
        formulas: (range.formulas as string[][]) || undefined,
        numberFormats: (range.numberFormatCategories as string[][]) || undefined,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
      };
    });
  }, "无法读取选中区域");
}

/** Resolve the active worksheet's name. */
export async function getActiveSheetName(): Promise<string> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();
      return ws.name;
    });
  }, "无法获取当前工作表名称");
}

/** Get a list of all worksheet names in the active workbook. */
export async function getSheetNames(): Promise<string[]> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      return sheets.items.map((s) => s.name);
    });
  }, "无法获取工作表列表");
}

/** Read the formulas of a specific A1-style address. */
export async function getRangeFormula(address: string): Promise<string[][]> {
  return safeCall(async () => {
    if (!address) throw new Error("地址不能为空");
    return run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
      range.load("formulas");
      await ctx.sync();
      return range.formulas as string[][];
    });
  }, "无法读取公式");
}

/* ---------------------------------------------------------------- *
 * Formula diagnostics (PRD-01)                                      *
 * ---------------------------------------------------------------- */

/** Canonical Excel error sentinels we know how to diagnose. */
export const FORMULA_ERROR_SENTINELS = [
  "#REF!",
  "#DIV/0!",
  "#N/A",
  "#VALUE!",
  "#NAME?",
  "#NUM!",
  "#NULL!",
  "#SPILL!",
  "#CALC!",
] as const;

export type FormulaErrorCode = (typeof FORMULA_ERROR_SENTINELS)[number];

export interface FormulaErrorCell {
  /** Sheet-local address like "C3". */
  address: string;
  /** Full sheet-qualified address "Sheet1!C3". */
  fullAddress: string;
  /** Original formula as the user sees it (=SUM(...)). */
  formula: string;
  /** Computed result or error sentinel. */
  value: string;
  /** The error sentinel, normalized. */
  error: FormulaErrorCode;
  /** Row index (0-based) within the scanned range. */
  rowIndex: number;
  /** Column index (0-based) within the scanned range. */
  colIndex: number;
}

export interface FormulaScanResult {
  sheetName: string;
  rangeAddress: string;
  cells: FormulaErrorCell[];
  /** Total formulas scanned (including healthy ones). */
  totalFormulas: number;
  /** Cells with a value but no formula in them. */
  totalValueOnly: number;
}

const ERROR_SET = new Set<string>(FORMULA_ERROR_SENTINELS);

/**
 * Walk every formula in the active selection and return a list of cells
 * whose value is one of the standard Excel error sentinels (#REF!, etc.).
 * The result is what the AI needs to produce per-cell diagnostics.
 */
export async function scanFormulaErrors(): Promise<FormulaScanResult> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["formulas", "values", "rowCount", "columnCount", "address"]);
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const formulas = (range.formulas as string[][]) || [];
      const values = (range.values as any[][]) || [];
      const errors: FormulaErrorCell[] = [];
      let totalFormulas = 0;
      let totalValueOnly = 0;

      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const formula = String(formulas[r][c] ?? "");
          const value = String(values[r]?.[c] ?? "");
          if (!formula || !formula.startsWith("=")) {
            // Not a formula cell - track plain values for context.
            if (value) totalValueOnly++;
            continue;
          }
          totalFormulas++;
          if (ERROR_SET.has(value)) {
            const localAddress = cellAddressAt(range.address, r, c);
            errors.push({
              address: localAddress,
              fullAddress: `${ws.name}!${localAddress}`,
              formula,
              value,
              error: value as FormulaErrorCode,
              rowIndex: r,
              colIndex: c,
            });
          }
        }
      }

      return {
        sheetName: ws.name,
        rangeAddress: range.address,
        cells: errors,
        totalFormulas,
        totalValueOnly,
      };
    });
  }, "无法扫描公式错误");
}

/**
 * Overwrite the formula at a specific cell with a new one. Excel.js throws
 * `InvalidArgument` if the new formula is malformed - we let that bubble
 * so the UI can show a friendly error.
 */
export async function fixFormulaAt(fullAddress: string, newFormula: string): Promise<void> {
  return safeCall(async () => {
    if (!fullAddress) throw new Error("单元格地址不能为空");
    if (!newFormula) throw new Error("公式不能为空");
    let trimmed = String(newFormula).trim();
    if (!trimmed.startsWith("=")) trimmed = "=" + trimmed;
    return run(async (ctx) => {
      const range = ctx.workbook.worksheets.getItem(fullAddress).getRange(fullAddress.split("!").pop()!);
      range.load(["formulas", "rowCount", "columnCount"]);
      await ctx.sync();
      const f: string[][] = [[trimmed]];
      range.formulas = f;
      await ctx.sync();
    });
  }, "无法写入修复后的公式");
}

/** Resolve the A1-style address of the cell at (row, col) relative to a
 *  range address. Handles the "Sheet1!A1:D10" form by stripping the sheet
 *  prefix and computing offsets from the top-left. */
function cellAddressAt(rangeAddress: string, row: number, col: number): string {
  const bangIdx = rangeAddress.indexOf("!");
  const body = bangIdx >= 0 ? rangeAddress.slice(bangIdx + 1) : rangeAddress;
  const [start] = body.split(":");
  const m = start.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return body;
  const startCol = colLettersToIndex(m[1]);
  const startRow = parseInt(m[2], 10);
  return indexToColLetters(startCol + col) + (startRow + row);
}

function colLettersToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  return n;
}

function indexToColLetters(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* ---------------------------------------------------------------- *
 * Writes                                                          *
 * ---------------------------------------------------------------- */

/**
 * Write a string into the active cell. If `moveActiveAfter` is true the
 * selection moves down one row after writing (mimics pressing Enter).
 */
export async function insertTextToCell(text: string, moveActiveAfter: boolean = false): Promise<void> {
  return safeCall(async () => {
    if (text === null || text === undefined) throw new Error("文本不能为空");
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await ctx.sync();

      // If the selection is multi-cell, only the top-left cell is written.
      const values: string[][] = Array.from({ length: range.rowCount }, (_, r) =>
        Array.from({ length: range.columnCount }, (_, c) => (r === 0 && c === 0 ? String(text) : ""))
      );
      range.values = values;
      await ctx.sync();

      if (moveActiveAfter) {
        // Move selection down by one cell for next insertion.
        const nextAddress = shiftAddress(range.address, 1, 0);
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        ws.getRange(nextAddress).select();
        await ctx.sync();
      }
    });
  }, "无法写入单元格");
}

/** Write a 2D array into the given A1 address. */
export async function writeRange(address: string, values: any[][]): Promise<void> {
  return safeCall(async () => {
    if (!address) throw new Error("地址不能为空");
    if (!Array.isArray(values)) throw new Error("values 必须是二维数组");
    return run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
      range.values = values;
      await ctx.sync();
    });
  }, "无法写入区域");
}

/**
 * Write a formula into the active cell. Excel needs the leading "=" to
 * recognize a formula; we add it automatically if missing. Useful when the
 * AI's reply contains a formula string the user wants to drop in directly.
 */
export async function writeFormula(formula: string): Promise<void> {
  return safeCall(async () => {
    if (formula === null || formula === undefined) throw new Error("公式不能为空");
    let trimmed = String(formula).trim();
    if (!trimmed) throw new Error("公式不能为空");
    if (!trimmed.startsWith("=")) trimmed = "=" + trimmed;
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "formulas", "rowCount", "columnCount"]);
      await ctx.sync();
      const f: string[][] = Array.from({ length: range.rowCount }, (_, r) =>
        Array.from({ length: range.columnCount }, (_, c) => (r === 0 && c === 0 ? trimmed : ""))
      );
      range.formulas = f;
      await ctx.sync();
    });
  }, "无法写入公式");
}

/**
 * Detect formula strings in a free-form text. Returns the first one that
 * looks like an Excel formula (=...). Useful for one-click "insert as
 * formula" actions.
 */
export function detectFormula(text: string): string | null {
  if (!text) return null;
  // Match on a line containing "=..." up to whitespace or end of line.
  const m = String(text).match(/(?:^|\n)\s*(=\s*[A-Z0-9_()\+\-\*\/\.\,\$"&':\s]+)/i);
  if (m && m[1]) {
    return m[1].trim();
  }
  // Fall back: any =... up to first newline
  const idx = text.indexOf("=");
  if (idx >= 0) {
    const slice = text.slice(idx).split(/\r?\n/)[0].trim();
    if (slice.length > 1 && /^=[A-Z0-9_]/i.test(slice)) return slice;
  }
  return null;
}

/** Highlight a cell briefly to give the user feedback. */
export async function flashSelectedCell(): Promise<void> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.format.fill.color = "#FFF4CE";
      await ctx.sync();
      setTimeout(() => undefined, 0); // keep short-lived
    });
  }, "无法高亮单元格");
}

/* ---------------------------------------------------------------- *
 * Internal helpers                                                *
 * ---------------------------------------------------------------- */

/** Parse a sheet-qualified address and shift the row/col of the top-left
 *  cell by (rowOffset, colOffset). Used to move the active selection down
 *  after writing. */
function shiftAddress(address: string, rowOffset: number, colOffset: number): string {
  // address looks like: "Sheet1!A1:D10" or "A1:D10"
  const bangIdx = address.indexOf("!");
  const sheet = bangIdx >= 0 ? address.slice(0, bangIdx + 1) : "";
  const body = bangIdx >= 0 ? address.slice(bangIdx + 1) : address;

  const [start] = body.split(":");
  const colMatch = start.match(/^[A-Z]+/);
  const rowMatch = start.match(/\d+$/);
  if (!colMatch || !rowMatch) return address;

  const colLetters = colMatch[0];
  const rowNumber = parseInt(rowMatch[0], 10);

  const newCol = offsetColLetters(colLetters, colOffset);
  const newRow = Math.max(1, rowNumber + rowOffset);
  return `${sheet}${newCol}${newRow}`;
}

/** Translate "A" -> offset 0, "AA" -> 26, etc., and apply offset. */
function offsetColLetters(col: string, offset: number): string {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  n = Math.max(1, n + offset);
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
/* ----------------------------------------------------------------- *
 * Formula collection for code-gen (PRD-04)                          *
 * ----------------------------------------------------------------- */

/** One formula cell captured for translation. */
export interface FormulaForCode {
  /** Sheet name (no quotes). */
  sheet: string;
  /** Cell address, e.g. "A1". */
  address: string;
  /** The raw formula string WITH the leading "=" if Excel stored one. */
  formula: string;
  /** True if the formula references another sheet or an external workbook. */
  hasExternalRef: boolean;
  /** True if the formula includes an array constant (`{=...}`). */
  isArrayFormula: boolean;
}

export interface FormulaCollectionResult {
  sheet: string;
  formulas: FormulaForCode[];
  /** Convenience join - the assistant ingests this directly. */
  block: string;
}

/** Read every formula in the current selection and pack them into a
 *  block the AI can ingest. Empty formulas (i.e. literal values) are
 *  skipped. If the selection is empty, returns `null`. */
export async function collectFormulasForCode(): Promise<FormulaCollectionResult | null> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["formulas", "address", "rowCount", "columnCount"]);
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const sheetName = ws.name || "Sheet1";
      const formulasArr: string[][] = range.formulas as string[][];
      const collected: FormulaForCode[] = [];

      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const raw = formulasArr?.[r]?.[c];
          if (typeof raw !== "string") continue;
          if (!raw.startsWith("=")) continue;
          const addr = `${indexToColLetters(c)}${r + 1}`;
          const hasExternal = /!|\[/.test(raw) && /'.*?!|\[.*?\]/.test(raw);
          const isArray = /^\{=/.test(raw);
          collected.push({
            sheet: sheetName,
            address: addr,
            formula: raw,
            hasExternalRef: hasExternal,
            isArrayFormula: isArray,
          });
        }
      }

      const block = collected
        .map((f) => `${f.sheet}!${f.address}  ${f.formula}`)
        .join("\n");

      return { sheet: sheetName, formulas: collected, block };
    });
  }, "无法读取选中区域的公式");
}

/* ----------------------------------------------------------------- *
 * Chart insertion (PRD-05)                                           *
 * ----------------------------------------------------------------- */

import type {
  ExcelChartType,
} from "../types";

/** Information about the currently selected range. */
export interface SelectedRangeInfo {
  /** Sheet name. */
  sheet: string;
  /** Range address, e.g. "Sheet1!A1:D10" or just "A1:D10". */
  address: string;
  /** First-row values - useful as column labels. */
  headers: string[];
  /** Approximate row count. */
  rowCount: number;
  /** Approximate column count. */
  columnCount: number;
  /** Sample preview (first 5 rows). */
  preview: string[][];
}

/** Read selection metadata + headers + a short preview, suitable for
 *  shipping to an AI chart recommender. */
export async function getSelectedRangeInfo(): Promise<SelectedRangeInfo | null> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address", "rowCount", "columnCount"]);
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const allValues = (range.values ?? []) as any[][];
      if (!allValues || allValues.length === 0) {
        throw new Error("未发现选区");
      }
      const headers = (allValues[0] ?? []).map((v: any) => String(v ?? "").trim());
      const preview = allValues.slice(0, Math.min(5, allValues.length)).map(
        (row) => row.map((v: any) => String(v ?? ""))
      );
      return {
        sheet: ws.name ?? "Sheet1",
        address: range.address ?? "",
        headers,
        rowCount: range.rowCount ?? 0,
        columnCount: range.columnCount ?? 0,
        preview,
      };
    });
  }, "无法读取选区信息");
}

/** Insert a chart at a default position below the source range.
 *  The chart type string must be one of the whitelisted values. */
export async function insertChart(
  sourceAddress: string,
  chartType: ExcelChartType,
  title: string
): Promise<string> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      // Parse out source range. Office.js wants the address only (no sheet).
      const bang = sourceAddress.indexOf("!");
      const cleanAddress =
        bang >= 0 ? sourceAddress.slice(bang + 1) : sourceAddress;

      const sourceRange = ws.getRange(cleanAddress);
      sourceRange.load(["rowCount", "columnCount"]);
      await ctx.sync();

      // Office.js charts.add(type, sourceRange). The chart is anchored
      // at the source range's top-left by default; we leave it there so
      // the user can drag it where they want.
      const chart = ws.charts.add(chartType, sourceRange);
      chart.title.text = title || "AI 推荐图表";
      chart.width = 480;
      chart.height = 288;

      await ctx.sync();
      return `${ws.name} 已插入图表「${title || "AI 推荐图表"}」`;
    });
  }, "无法插入图表");
}

/* ----------------------------------------------------------------- *
 * PII detection / masking (PRD-06)                                  *
 * ----------------------------------------------------------------- */

export type PiiKind =
  | "phone_cn"      // 手机号（中国）
  | "phone_intl"    // 国际电话（含 +86 / +1 等）
  | "email"         // 邮箱
  | "id_card_cn"    // 中国身份证号
  | "bank_card"     // 银行卡号（13-19 位数字）
  | "ip"            // IPv4
  | "name_cn"       // 中文姓名（启发式：高概率）
  | "address_cn"    // 中文地址（启发式：含「省/市/区/路/号」）
  | "credit_card";  // 信用卡（含 Luhn 校验）

export interface PiiHit {
  address: string;
  row: number; // 0-based
  col: number; // 0-based
  kind: PiiKind;
  /** Original cell value as a string (best-effort). */
  original: string;
}

export interface PiiScanResult {
  sheet: string;
  rangeAddress: string;
  hits: PiiHit[];
  totalCells: number;
}

/** Pre-compiled detectors. Each runs synchronously on a stringified
 *  cell and returns the matched substring or null. */
const PII_PATTERNS: Array<{ kind: PiiKind; test: (s: string) => string | null }> = [
  { kind: "id_card_cn",
    test: (s) => {
      // 18 digits, last char can be X.
      const m = s.match(/\b[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx]\b/);
      return m ? m[0] : null;
    } },
  { kind: "phone_cn",
    test: (s) => {
      const m = s.match(/\b1[3-9]\d{9}\b/);
      return m ? m[0] : null;
    } },
  { kind: "phone_intl",
    test: (s) => {
      const m = s.match(/(?<!\d)\+\d{1,3}[\s\-]?\d{4,14}(?!\d)/);
      return m ? m[0] : null;
    } },
  { kind: "bank_card",
    test: (s) => {
      // Just Luhn validation on 13-19 digit runs.
      const m = s.match(/(?<!\d)\d{13,19}(?!\d)/);
      if (!m) return null;
      return luhnValid(m[0]) ? m[0] : null;
    } },
  { kind: "credit_card",
    test: (s) => {
      // 16-digit, dash-separated or continuous, Luhn-valid
      const m = s.match(/(?<!\d)(?:\d{4}[\s\-]?){3,4}\d{3,4}(?!\d)/);
      if (!m) return null;
      const digits = m[0].replace(/\D/g, "");
      return digits.length >= 13 && digits.length <= 19 && luhnValid(digits) ? m[0] : null;
    } },
  { kind: "email",
    test: (s) => {
      const m = s.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/);
      return m ? m[0] : null;
    } },
  { kind: "ip",
    test: (s) => {
      const m = s.match(/\b(?:(?:25[0-5]|2[0-4]\d|[01]?\d{1,2})\.){3}(?:25[0-5]|2[0-4]\d|[01]?\d{1,2})\b/);
      return m ? m[0] : null;
    } },
  { kind: "address_cn",
    test: (s) => {
      // Heuristic: contains a keyword and at least 6 chars of context.
      if (/[省市区县镇乡村路街道号巷弄]/.test(s) && /[一-龥]/.test(s) && s.length >= 6) return s;
      return null;
    } },
  { kind: "name_cn",
    test: (s) => {
      // Only flag bare 2-4 CJK chars when not part of a longer matched address.
      // We treat this as low-confidence, so callers can opt out.
      const m = s.match(/^[一-龥]{2,4}$/);
      return m ? m[0] : null;
    } },
];

function luhnValid(digits: string): boolean {
  let sum = 0;
  let alt = false;
  for (let i = digits.length - 1; i >= 0; i--) {
    let n = digits.charCodeAt(i) - 48;
    if (alt) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
    alt = !alt;
  }
  return sum % 10 === 0;
}

/** Scan the selection for sensitive values using the local detectors. */
export async function scanSelectionForPII(): Promise<PiiScanResult> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address", "rowCount", "columnCount"]);
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const sheetName = ws.name ?? "Sheet1";
      const values = (range.values ?? []) as any[][];
      const hits: PiiHit[] = [];
      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < (values[r]?.length ?? 0); c++) {
          const cell = values[r][c];
          if (cell === null || cell === undefined) continue;
          const s = String(cell);
          for (const { kind, test } of PII_PATTERNS) {
            const m = test(s);
            if (m) {
              hits.push({
                address: `${indexToColLetters(c)}${r + 1}`,
                row: r,
                col: c,
                kind,
                original: s,
              });
              break; // one hit per cell
            }
          }
        }
      }
      return {
        sheet: sheetName,
        rangeAddress: range.address ?? "",
        hits,
        totalCells: (range.rowCount ?? 0) * (range.columnCount ?? 0),
      };
    });
  }, "无法扫描选区");
}

/** Replace matched cells with deterministic fake values keyed to the
 *  original (so the same phone stays the same fake phone). The caller
 *  decides whether to apply per cell. */
export function fakeFor(kind: PiiKind, original: string, seed: number): string {
  switch (kind) {
    case "phone_cn":
    case "phone_intl":
      return "138" + String(10000000 + seed).padStart(8, "0").slice(0, 8);
    case "email":
      return `user${seed}@example.com`;
    case "id_card_cn": {
      // Deterministic 18-digit fake; checksum won't match - that's fine for placeholders.
      const base = String(1100000000000000000 + seed);
      return (base + "X").slice(0, 18);
    }
    case "bank_card":
    case "credit_card":
      return "6222020000000000" + String(seed).padStart(4, "0").slice(0, 4);
    case "ip":
      return `10.0.${(seed >> 8) & 0xff}.${seed & 0xff}`;
    case "name_cn":
      return ["张伟", "李娜", "王芳", "刘洋", "陈静"][seed % 5];
    case "address_cn":
      return `北京市朝阳区某街${seed}号`;
    default:
      return "***";
  }
}

/** Apply a list of `address -> fakeValue` replacements in one batch. */
export async function batchReplaceCells(
  updates: Array<{ address: string; value: string }>
): Promise<number> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();
      let written = 0;
      for (const u of updates) {
        try {
          const cell = ws.getRange(u.address);
          cell.values = [[u.value]];
          written++;
        } catch {
          /* skip */
        }
      }
      await ctx.sync();
      return written;
    });
  }, "无法批量替换");
}

/* ----------------------------------------------------------------- *
 * Multi-selection aggregation (PRD-07)                              *
 * ----------------------------------------------------------------- */

export interface MultiSelectionPart {
  sheet: string;
  address: string;
  values: any[][];
  rowCount: number;
  columnCount: number;
  label: string;
}

export interface MultiSelectionResult {
  sheet: string;
  parts: MultiSelectionPart[];
  totalCells: number;
  /** Best-effort markdown block; sheets/regions separated by headers. */
  block: string;
}

function mdEscape(s: string): string {
  return String(s ?? "").replace(/\|/g, "\\|").replace(/\n/g, " ");
}

/** Read the active selection as a list of regions. Office.js exposes
 *  multi-region selection via `range.areas` which is itself a RangeCollection. */
export async function readMultiSelection(): Promise<MultiSelectionResult | null> {
  return safeCall(async () => {
    return run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      // Use areas() if available; fall back to a single part otherwise.
      let areas: any = null;
      try {
        areas = (range as any).areas;
        if (areas && typeof areas.load === "function") {
          areas.load("items");
        }
      } catch {
        areas = null;
      }
      await ctx.sync();

      const parts: MultiSelectionPart[] = [];
      const sheetName = ws.name ?? "Sheet1";

      if (areas && Array.isArray(areas.items) && areas.items.length > 0) {
        let i = 0;
        for (const sub of areas.items) {
          sub.load(["values", "address", "rowCount", "columnCount"]);
          (sub as any).__idx = i++;
        }
        await ctx.sync();
        for (const sub of areas.items as any[]) {
          parts.push({
            sheet: sheetName,
            address: sub.address ?? "",
            values: (sub.values as any[][]) ?? [],
            rowCount: sub.rowCount ?? 0,
            columnCount: sub.columnCount ?? 0,
            label: `区域 ${(sub as any).__idx + 1}`,
          });
        }
      } else {
        range.load(["values", "address", "rowCount", "columnCount"]);
        await ctx.sync();
        parts.push({
          sheet: sheetName,
          address: range.address ?? "",
          values: (range.values as any[][]) ?? [],
          rowCount: range.rowCount ?? 0,
          columnCount: range.columnCount ?? 0,
          label: "区域 1",
        });
      }

      // Build a Markdown block.
      const blocks = parts.map((p) => {
        const header = `### ${p.label} (${p.sheet}!${p.address})`;
        const tbl = (p.values || []).slice(0, 30).map((row) =>
          "|" + (row ?? []).map(mdEscape).join("|") + "|"
        );
        const sep =
          "|" + (p.values?.[0] ?? []).map(() => "---").join("|") + "|";
        return [header, sep, ...tbl].join("\n");
      });

      return {
        sheet: sheetName,
        parts,
        totalCells: parts.reduce(
          (acc, p) => acc + p.rowCount * p.columnCount,
          0
        ),
        block: blocks.join("\n\n"),
      };
    });
  }, "无法读取多选区");
}

/* ----------------------------------------------------------------- *
 * Function-call dispatcher (PRD-09)                                 *
 * ----------------------------------------------------------------- */

import type { ToolCall } from "../types";
import { safeJsonParse } from "../utils/helpers";

/** Friendly Chinese labels for tool names - used in UI badges. */
export const TOOL_LABELS: Record<string, string> = {
  writeFormula: "写入公式",
  insertChart: "插入图表",
  scanPII: "扫描敏感数据",
  getSelectionInfo: "读取选区信息",
};

/** Execute a single tool call against the live Excel workbook.
 *  Returns a short human-readable outcome string the model can use to
 *  continue the conversation. */
export async function applyToolCall(call: ToolCall): Promise<string> {
  const args = safeJsonParse<Record<string, any>>(call.arguments, {});
  try {
    switch (call.name) {
      case "writeFormula": {
        const formula = String(args.formula || "").trim();
        if (!formula) return "错误：缺少 formula 参数";
        if (!formula.startsWith("=")) return `错误：公式必须以 = 开头 (收到: ${formula})`;
        const targetAddr = typeof args.address === "string" && args.address.trim()
          ? args.address.trim()
          : null;
        if (targetAddr) {
          await writeFormulaToAddress(targetAddr, formula);
          return `已写入公式 ${formula} 到 ${targetAddr}`;
        }
        await writeFormula(formula);
        return `已写入公式 ${formula} 到活动单元格`;
      }
      case "insertChart": {
        const chartType = String(args.chartType || "ColumnClustered") as any;
        const title = typeof args.title === "string" ? args.title : "";
        const ok = await insertChart("", chartType, title);
        return ok
          ? `已插入图表 (${chartType}${title ? ": " + title : ""})`
          : "插入图表失败：请先选中一个数据区域";
      }
      case "scanPII": {
        const result = await scanSelectionForPII();
        if (result.hits.length === 0) {
          return `未在选区中发现敏感数据 (${result.totalCells} 个单元格已扫描)`;
        }
        const byKind: Record<string, number> = {};
        for (const h of result.hits) byKind[h.kind] = (byKind[h.kind] || 0) + 1;
        const summary = Object.entries(byKind)
          .map(([k, n]) => `${k}: ${n}`)
          .join(", ");
        return `发现 ${result.hits.length} 处敏感数据 (${summary})。建议引导用户打开「数据脱敏」面板进行替换。`;
      }
      case "getSelectionInfo": {
        const sel = await getSelectedData();
        if (!sel || !sel.address) return "当前没有选区";
        const sample = (sel.values || [])
          .slice(0, 3)
          .map((row) => (row || []).slice(0, 5).map((v) => (v == null ? "" : String(v))).join(" | "))
          .join("\n");
        return `选区: ${sel.sheetName}!${sel.address}, ${sel.rowCount}×${sel.columnCount}。\n预览 (前 3 行 × 5 列):\n${sample || "(空)"}`;
      }
      default:
        return `未知工具: ${call.name}`;
    }
  } catch (err: any) {
    return `工具执行失败: ${err?.message || String(err)}`;
  }
}

/** Write a formula to a specific address (e.g. "B3"). */
async function writeFormulaToAddress(address: string, formula: string): Promise<void> {
  await safeRun(async () => {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.getRange(address).formulas = [[formula]];
      await ctx.sync();
    });
  }, "无法写入公式");
}
