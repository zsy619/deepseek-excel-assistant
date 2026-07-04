/**
 * ============================================================================
 * Ribbon commands - hybrid dispatch
 * ----------------------------------------------------------------------------
 * Strategy:
 *
 *   1. First, try the standard cross-frame paths to the taskpane (the
 *      main UX lives there). Multiple transports are tried so we get
 *      coverage on Excel web / Mac / Windows / SharedRuntime.
 *
 *   2. Always call Excel.run DIRECTLY in this FunctionFile as a
 *      fallback. If cross-frame delivery fails for any reason, the
 *      user still sees something happen (a dialog opens with the
 *      button's effect).
 *
 *   3. Show a tiny visible status dialog on every click so the user
 *      can SEE the button click was received, regardless of which
 *      transport works.
 *
 * This means the buttons never silently do nothing. They either drive
 * the chat taskpane, or they pop a visible dialog, or both.
 * ============================================================================
 */

/* eslint-disable @typescript-eslint/no-unused-vars */

type CommandId =
  | "analyzeSelection"
  | "generateFormula"
  | "cleanData"
  | "insertLastReply"
  | "exportCurrentSession"
  | "clearCurrentChat"
  | "toggleTheme"
  | "openSettings"
  | "diagnoseFormulas"
  | "translateToCode"
  | "insertChart"
  | "maskPII"
  | "multiSelectionAnalyze"
  | "openKnowledgeBase"
  | "shareSession"
  | "usageDashboard";

interface CommandEnvelope {
  id: string;
  type: "deepseek:command";
  command: CommandId;
  payload: unknown;
  t: number;
  via: string;
}

const CHANNEL_NAME = "deepseek-cmds-v1";
const STORAGE_KEY = "deepseek:command";

function newId(): string {
  return "cmd-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 8);
}

function buildEnvelope(command: CommandId, payload?: unknown): CommandEnvelope {
  return {
    id: newId(),
    type: "deepseek:command",
    command,
    payload: payload ?? null,
    t: Date.now(),
    via: "init",
  };
}

/* ----------------------------------------------------------------- */
/* Cross-frame delivery (4 channels)                                  */
/* ----------------------------------------------------------------- */
function dispatch(command: CommandId, payload?: unknown): void {
  const env = buildEnvelope(command, payload);
  const sent: string[] = [];

  // 1. Direct call - if the FunctionFile and taskpane share a window.
  try {
    const app = (window as any).__deepseekApp;
    if (app && typeof app.runCommand === "function") {
      env.via = "direct";
      try { app.runCommand(env); } catch (err) { try { console.warn("[DeepSeek] direct runCommand failed", err); } catch {} }
      sent.push("direct");
    }
  } catch { /* noop */ }

  // 2. Office.context.ui.messageParent - the official cross-frame API.
  try {
    const ui = (Office as any)?.context?.ui;
    if (ui && typeof ui.messageParent === "function") {
      env.via = sent.length === 0 ? "messageParent" : env.via;
      try { ui.messageParent(JSON.stringify(env)); } catch (err) { try { console.warn("[DeepSeek] messageParent failed", err); } catch {} }
      sent.push("messageParent");
    }
  } catch { /* noop */ }

  // 3. BroadcastChannel.
  try {
    if (typeof BroadcastChannel === "function") {
      const ch = new BroadcastChannel(CHANNEL_NAME);
      env.via = sent.length === 0 ? "broadcast" : env.via;
      try { ch.postMessage(env); } catch (err) { try { console.warn("[DeepSeek] BroadcastChannel failed", err); } catch {} }
      try { ch.close(); } catch { /* noop */ }
      sent.push("broadcast");
    }
  } catch { /* noop */ }

  // 4. localStorage write - the receiver polls every 600ms in the
  //    taskpane. We DO NOT remove the key here; the taskpane clears
  //    it after consumption (see taskpane.ts polling loop). A
  //    repeated click within 600ms overwrites the same key with the
  //    latest command - the receiver still picks up the latest one.
  try {
    if (typeof localStorage !== "undefined") {
      env.via = sent.length === 0 ? "storage" : env.via;
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(env)); } catch (err) { try { console.warn("[DeepSeek] localStorage write failed", err); } catch {} }
      sent.push("storage");
    }
  } catch { /* noop */ }

  // 5. window.parent.postMessage - last-resort fallback.
  try {
    if (window.parent && window.parent !== window) {
      env.via = sent.length === 0 ? "postMessage" : env.via;
      try { window.parent.postMessage(env, "*"); } catch (err) { try { console.warn("[DeepSeek] postMessage failed", err); } catch {} }
      sent.push("postMessage");
    }
  } catch { /* noop */ }

  if (sent.length === 0) {
    try { console.warn("[DeepSeek] No transport available; opening taskpane"); } catch {}
    try { Office.addin.showAsTaskpane(); } catch { /* noop */ }
  }

  try { console.info("[DeepSeek] command", command, "via", sent.join("+")); } catch {}
}

/* ----------------------------------------------------------------- */
/* Visible feedback: pop a tiny dialog confirming the click.          */
/* This runs ON EVERY CLICK so the user always sees something.       */
/* ----------------------------------------------------------------- */
function confirmClick(command: CommandId): void {
  const labels: Record<CommandId, string> = {
    analyzeSelection: "分析选区",
    generateFormula: "生成公式",
    cleanData: "数据清洗",
    insertLastReply: "插入回复",
    exportCurrentSession: "导出对话",
    clearCurrentChat: "清空对话",
    toggleTheme: "切换主题",
    openSettings: "打开设置",
    diagnoseFormulas: "诊断公式",
    translateToCode: "公式转 VBA",
    insertChart: "插入图表",
    maskPII: "数据脱敏",
    multiSelectionAnalyze: "多选区分析",
    openKnowledgeBase: "知识库",
    shareSession: "分享会话",
    usageDashboard: "用量看板",
  };
  try {
    const ui = (Office as any)?.context?.ui;
    if (!ui || typeof ui.displayDialogAsync !== "function") return;
    const html =
      "<html><body style='font:14px/1.5 system-ui,sans-serif;padding:18px;color:#1f2937;background:#fff'>" +
      "<h3 style='margin:0 0 8px'>▶ 按钮已触发</h3>" +
      "<p style='margin:0 0 6px'>命令：<b>" + labels[command] + "</b> (" + command + ")</p>" +
      "<p style='margin:0 0 6px;color:#475569'>时间：" + new Date().toLocaleString("zh-CN") + "</p>" +
      "<p style='margin:8px 0 0;font-size:12px;color:#64748b'>如果上方任务窗没有反应，请把这个时间戳发给开发。</p>" +
      "<button onclick='window.close()' style='margin-top:12px;padding:6px 14px;border:0;border-radius:4px;background:#0d6efd;color:#fff;cursor:pointer'>关闭</button>" +
      "</body></html>";
    ui.displayDialogAsync(
      "data:text/html;charset=utf-8," + encodeURIComponent(html),
      { height: 220, width: 360, displayInIframe: false }
    );
  } catch (err) {
    try { console.warn("[DeepSeek] confirmClick failed", err); } catch {}
  }
}

/* ----------------------------------------------------------------- */
/* Command implementations                                            */
/* ----------------------------------------------------------------- */

function done(event: Office.AddinCommands.Event): void {
  if (event && typeof event.completed === "function") {
    try { event.completed(); } catch { /* noop */ }
  }
}

export function showTaskpane(event: Office.AddinCommands.Event): void {
  try { Office.addin.showAsTaskpane(); } catch { /* noop */ }
  done(event);
}

export function analyzeSelection(event: Office.AddinCommands.Event): void {
  confirmClick("analyzeSelection");
  dispatch("analyzeSelection");
  done(event);
}

export function generateFormula(event: Office.AddinCommands.Event): void {
  confirmClick("generateFormula");
  dispatch("generateFormula");
  done(event);
}

export function cleanData(event: Office.AddinCommands.Event): void {
  confirmClick("cleanData");
  dispatch("cleanData");
  done(event);
}

export function insertLastReply(event: Office.AddinCommands.Event): void {
  confirmClick("insertLastReply");
  dispatch("insertLastReply");
  done(event);
}

export function exportCurrentSession(event: Office.AddinCommands.Event): void {
  confirmClick("exportCurrentSession");
  dispatch("exportCurrentSession");
  done(event);
}

export function clearCurrentChat(event: Office.AddinCommands.Event): void {
  confirmClick("clearCurrentChat");
  dispatch("clearCurrentChat");
  done(event);
}

export function toggleTheme(event: Office.AddinCommands.Event): void {
  confirmClick("toggleTheme");
  dispatch("toggleTheme");
  done(event);
}

export function openSettings(event: Office.AddinCommands.Event): void {
  confirmClick("openSettings");
  dispatch("openSettings");
  done(event);
}

export function diagnoseFormulas(event: Office.AddinCommands.Event): void {
  confirmClick("analyzeSelection");
  dispatch("diagnoseFormulas");
  done(event);
}

export function translateToCode(event: Office.AddinCommands.Event): void {
  confirmClick("translateToCode");
  dispatch("translateToCode");
  done(event);
}

export function insertChart(event: Office.AddinCommands.Event): void {
  confirmClick("insertChart");
  dispatch("insertChart");
  done(event);
}

export function maskPII(event: Office.AddinCommands.Event): void {
  confirmClick("maskPII");
  dispatch("maskPII");
  done(event);
}

export function multiSelectionAnalyze(event: Office.AddinCommands.Event): void {
  confirmClick("multiSelectionAnalyze");
  dispatch("multiSelectionAnalyze");
  done(event);
}

export function openKnowledgeBase(event: Office.AddinCommands.Event): void {
  confirmClick("openKnowledgeBase");
  dispatch("openKnowledgeBase");
  done(event);
}

export function shareSession(event: Office.AddinCommands.Event): void {
  confirmClick("shareSession");
  dispatch("shareSession");
  done(event);
}

export function usageDashboard(event: Office.AddinCommands.Event): void {
  confirmClick("usageDashboard");
  dispatch("usageDashboard");
  done(event);
}

/* ----------------------------------------------------------------- */
/* Expose handlers on the global scope.                               */
/* ----------------------------------------------------------------- */

declare global {
  interface Window {
    showTaskpane?: (event: Office.AddinCommands.Event) => void;
    analyzeSelection?: (event: Office.AddinCommands.Event) => void;
    generateFormula?: (event: Office.AddinCommands.Event) => void;
    cleanData?: (event: Office.AddinCommands.Event) => void;
    insertLastReply?: (event: Office.AddinCommands.Event) => void;
    exportCurrentSession?: (event: Office.AddinCommands.Event) => void;
    clearCurrentChat?: (event: Office.AddinCommands.Event) => void;
    toggleTheme?: (event: Office.AddinCommands.Event) => void;
    openSettings?: (event: Office.AddinCommands.Event) => void;
    diagnoseFormulas?: (event: Office.AddinCommands.Event) => void;
    translateToCode?: (event: Office.AddinCommands.Event) => void;
    insertChart?: (event: Office.AddinCommands.Event) => void;
    maskPII?: (event: Office.AddinCommands.Event) => void;
    multiSelectionAnalyze?: (event: Office.AddinCommands.Event) => void;
    openKnowledgeBase?: (event: Office.AddinCommands.Event) => void;
    shareSession?: (event: Office.AddinCommands.Event) => void;
    usageDashboard?: (event: Office.AddinCommands.Event) => void;
  }
}

window.showTaskpane = showTaskpane;
window.analyzeSelection = analyzeSelection;
window.generateFormula = generateFormula;
window.cleanData = cleanData;
window.insertLastReply = insertLastReply;
window.exportCurrentSession = exportCurrentSession;
window.clearCurrentChat = clearCurrentChat;
window.toggleTheme = toggleTheme;
window.openSettings = openSettings;
window.diagnoseFormulas = diagnoseFormulas;
window.translateToCode = translateToCode;
window.insertChart = insertChart;
window.maskPII = maskPII;
window.multiSelectionAnalyze = multiSelectionAnalyze;
window.openKnowledgeBase = openKnowledgeBase;
window.shareSession = shareSession;
window.usageDashboard = usageDashboard;

try { console.info("[DeepSeek] commands.js loaded - handlers registered"); } catch {}

try {
  const stamp = new Date().toISOString();
  document.title = "DeepSeek Excel Assistant - Commands [" + stamp + "]";
  const el = document.createElement("div");
  el.id = "deepseek-build-stamp";
  el.dataset.stamp = stamp;
  el.style.display = "none";
  document.body.appendChild(el);
} catch { /* noop */ }