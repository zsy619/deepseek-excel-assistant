/**
 * ============================================================================
 * modalExtras — lazy-loaded bundle of 5 modal/dialog components
 * ----------------------------------------------------------------------------
 * Carved into its own webpack chunk so the main taskpane bundle stays lean.
 * Loaded on first ⌘K / ⌘B / ⌘⇧S / ⌘D / ⌘/ keystroke (whichever comes
 * first). Each component is mounted at document.body level so it can overlay
 * the whole app.
 *
 * Returns the mounted instances so the caller can drive show()/hide()
 * through cached references after the one-time load.
 * ============================================================================
 */

import { CommandPaletteView, type CommandPaletteSelectDetail } from "./CommandPalette";
import { ShortcutHelpView } from "./ShortcutHelp";
import { KnowledgeBaseView } from "./KnowledgeBase";
import { ShareDialogView, type ShareDialogDetail } from "./ShareDialog";
import { UsageDashboardView } from "./UsageDashboard";

export interface ModalExtras {
  palette: CommandPaletteView;
  shortcutHelp: ShortcutHelpView;
  knowledgeBase: KnowledgeBaseView;
  shareDialog: ShareDialogView;
  usageDashboard: UsageDashboardView;
}

export interface ModalExtrasHandlers {
  onPaletteSelect: (detail: CommandPaletteSelectDetail) => void;
  onKbToast: (msg: string, kind: string) => void;
  onShareImport: (detail: ShareDialogDetail) => void;
  onShareToast: (msg: string, kind: string) => void;
  onUsageToast: (msg: string, kind: string) => void;
}

/** Mount all 5 modal extras at document.body level and wire their events. */
export function mountModalExtras(handlers: ModalExtrasHandlers): ModalExtras {
  const palette = new CommandPaletteView();
  palette.element.addEventListener("command-palette-select", (ev) => {
    const e = ev as CustomEvent<CommandPaletteSelectDetail>;
    handlers.onPaletteSelect(e.detail);
  });
  document.body.appendChild(palette.element);

  const shortcutHelp = new ShortcutHelpView();
  document.body.appendChild(shortcutHelp.element);

  const knowledgeBase = new KnowledgeBaseView();
  knowledgeBase.element.addEventListener("kb-toast", (ev) => {
    const e = ev as CustomEvent;
    const { msg, kind } = e.detail || {};
    handlers.onKbToast(msg, kind);
  });
  document.body.appendChild(knowledgeBase.element);

  const shareDialog = new ShareDialogView();
  shareDialog.element.addEventListener("share-import", (ev) => {
    const e = ev as CustomEvent<ShareDialogDetail>;
    handlers.onShareImport(e.detail);
  });
  shareDialog.element.addEventListener("share-toast", (ev) => {
    const e = ev as CustomEvent;
    const { msg, kind } = e.detail || {};
    handlers.onShareToast(msg, kind);
  });
  document.body.appendChild(shareDialog.element);

  const usageDashboard = new UsageDashboardView();
  usageDashboard.element.addEventListener("usage-toast", (ev) => {
    const e = ev as CustomEvent;
    const { msg, kind } = e.detail || {};
    handlers.onUsageToast(msg, kind);
  });
  document.body.appendChild(usageDashboard.element);

  return { palette, shortcutHelp, knowledgeBase, shareDialog, usageDashboard };
}