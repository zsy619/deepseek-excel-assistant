/**
 * ============================================================================
 * Excel event bridge
 * ----------------------------------------------------------------------------
 * Registers handlers for the Office.js "binding" and "worksheet" events so the
 * taskpane can react to changes the user makes inside Excel. Specifically:
 *
 *   - selectionChanged on the active binding  -> notify subscribers
 *   - worksheetActivated on the workbook      -> notify subscribers
 *
 * The bridge is intentionally thin: it just dispatches typed events. Consumers
 * (typically the ChatWindow + a context-bar component) decide what to do with
 * the information.
 * ============================================================================
 */

import type { ExcelSelection } from "../types";
import { getSelectedData, getActiveSheetName } from "./excel";

export interface SelectionChangedDetail {
  selection: ExcelSelection | null;
  /** Address of the previous selection, or null. */
  previousAddress: string | null;
}

export interface SheetChangedDetail {
  sheetName: string;
}

type SelectionListener = (detail: SelectionChangedDetail) => void;
type SheetListener = (detail: SheetChangedDetail) => void;

/**
 * Singleton bridge. Created lazily on first access.
 */
class ExcelEventBridge {
  private selectionListeners: Set<SelectionListener> = new Set();
  private sheetListeners: Set<SheetListener> = new Set();
  private lastAddress: string | null = null;
  private lastSheet: string | null = null;
  private registered: boolean = false;

  /** Register listeners with the Office.js event API. Safe to call repeatedly. */
  public register(): void {
    if (this.registered) return;
    this.registered = true;

    try {
      // Selection change - fires every time the user clicks a different cell.
      // We attach to the workbook's active cell rather than a named binding
      // because we want any selection inside any sheet to be observed.
      const onSelectionChanged = async () => {
        let selection: ExcelSelection | null = null;
        try {
          selection = await getSelectedData();
        } catch {
          selection = null;
        }
        const previousAddress = this.lastAddress;
        this.lastAddress = selection?.address ?? null;
        const detail: SelectionChangedDetail = { selection, previousAddress };
        for (const l of this.selectionListeners) {
          try {
            l(detail);
          } catch (err) {
            console.error("[DeepSeek] selection listener error", err);
          }
        }
      };
      Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getActiveWorksheet();
        ws.onSelectionChanged.add(onSelectionChanged);
        await ctx.sync();
      }).catch((err) => {
        console.warn("[DeepSeek] Failed to register onSelectionChanged", err);
      });
    } catch (err) {
      console.warn("[DeepSeek] onSelectionChanged not available", err);
    }

    try {
      const onSheetActivated = async () => {
        let name = "";
        try {
          name = await getActiveSheetName();
        } catch {
          name = "";
        }
        if (name === this.lastSheet) return;
        this.lastSheet = name;
        const detail: SheetChangedDetail = { sheetName: name };
        for (const l of this.sheetListeners) {
          try {
            l(detail);
          } catch (err) {
            console.error("[DeepSeek] sheet listener error", err);
          }
        }
      };
      Excel.run(async (ctx) => {
        ctx.workbook.onActivated.add(onSheetActivated);
        await ctx.sync();
      }).catch((err) => {
        console.warn("[DeepSeek] Failed to register onActivated", err);
      });
    } catch (err) {
      console.warn("[DeepSeek] onActivated not available", err);
    }
  }

  /** Subscribe to selection changes. Returns an unsubscribe function. */
  public onSelectionChange(listener: SelectionListener): () => void {
    this.selectionListeners.add(listener);
    this.register();
    return () => this.selectionListeners.delete(listener);
  }

  /** Subscribe to sheet activation events. Returns an unsubscribe function. */
  public onSheetChange(listener: SheetListener): () => void {
    this.sheetListeners.add(listener);
    this.register();
    return () => this.sheetListeners.delete(listener);
  }

  /** Tear down all listeners. Currently a no-op; kept for future use. */
  public dispose(): void {
    this.selectionListeners.clear();
    this.sheetListeners.clear();
    this.registered = false;
  }
}

export const excelEvents = new ExcelEventBridge();