/**
 * ============================================================================
 * Constants
 * ----------------------------------------------------------------------------
 * Single source of truth for keys, defaults, and limits used throughout the
 * add-in. Anything that could change between deployments lives here so that
 * tuning never requires hunting across files.
 * ============================================================================
 */

import type { DeepSeekConfig, DeepSeekModel } from "../types";

/** localStorage key for the user's persisted configuration. */
export const STORAGE_KEY_CONFIG = "deepseek_excel_config_v1";

/** localStorage key for the array of chat sessions. */
export const STORAGE_KEY_SESSIONS = "deepseek_excel_sessions_v1";

/** localStorage key for the active session id. */
export const STORAGE_KEY_ACTIVE_SESSION = "deepseek_excel_active_session_v1";

/** Hard cap on the number of sessions retained locally. */
export const MAX_SESSIONS = 50;

/** Maximum formulas we attempt to diagnose in a single scan. Larger ranges
 *  require explicit user confirmation. */
export const FORMULA_DIAGNOSIS_LIMIT = 500;

/** Soft warning threshold. Users get a "this might take a moment" toast. */
export const FORMULA_DIAGNOSIS_WARN = 200;

/** Human-readable explanation of each Excel error sentinel. The keys match
 *  the strings Excel itself emits (#REF!, #DIV/0!, etc.). */
export const FORMULA_ERROR_INFO: Record<string, { code: string; label: string; reason: string }> = {
  "#REF!": {
    code: "REF",
    label: "引用无效",
    reason: "公式引用了一个已被删除或不存在的单元格（通常是复制/粘贴时引用断裂）。",
  },
  "#DIV/0!": {
    code: "DIV0",
    label: "除以零",
    reason: "公式尝试除以零或空单元格。请检查分母是否被正确填写。",
  },
  "#N/A": {
    code: "NA",
    label: "值不可用",
    reason: "查找函数（VLOOKUP / XLOOKUP / MATCH 等）找不到匹配项；或公式主动返回 NA()。",
  },
  "#VALUE!": {
    code: "VALUE",
    label: "类型错误",
    reason: "公式中的参数类型不匹配（例如把文本传给了数值函数）。",
  },
  "#NAME?": {
    code: "NAME",
    label: "名称未识别",
    reason: "公式里引用了一个不存在的函数名、命名范围或文本未加引号。",
  },
  "#NUM!": {
    code: "NUM",
    label: "数值无效",
    reason: "公式产生了无效数值（如 SQRT(-1) 或超出范围的迭代结果）。",
  },
  "#NULL!": {
    code: "NULL",
    label: "区域交叉无效",
    reason: "两个不相交的区域之间用了空格运算符，或区域引用语法错误。",
  },
  "#SPILL!": {
    code: "SPILL",
    label: "溢出冲突",
    reason: "动态数组公式的溢出区域被其他数据阻挡。",
  },
  "#CALC!": {
    code: "CALC",
    label: "计算错误",
    reason: "Excel 365 内部计算引擎报告的未分类错误，通常由不兼容的数组运算引起。",
  },
};

/** Network timeout for any single DeepSeek request (ms). */
export const REQUEST_TIMEOUT_MS = 60_000;

/** Default system prompt - shown to the model as the first message. */
export const DEFAULT_SYSTEM_PROMPT = `你是一个专业的 Excel 助手。你能帮助用户：
1. 生成和优化 Excel 公式
2. 分析选中区域的数据并给出洞察
3. 提供数据清洗和格式化建议
4. 解释复杂的 Excel 功能
请用简洁清晰的中文回答，公式请用代码块包裹。`;

/** Models surfaced in the settings dropdown. */
export const AVAILABLE_MODELS: { value: DeepSeekModel; label: string; description: string }[] = [
  { value: "deepseek-chat", label: "deepseek-chat", description: "通用对话（速度快）" },
  { value: "deepseek-reasoner", label: "deepseek-reasoner", description: "深度推理（思考链）" },
];

/** Factory that returns a fresh default config - never share the same object
 *  reference between users to avoid accidental mutation. */
export function createDefaultConfig(): DeepSeekConfig {
  return {
    apiKey: "",
    baseUrl: "https://api.deepseek.com",
    model: "deepseek-chat",
    temperature: 0.7,
    maxTokens: 2048,
    topP: 0.9,
    systemPrompt: DEFAULT_SYSTEM_PROMPT,
  };
}

/** Shortcut templates prepended to the user's text in the input box. */
export const QUICK_ACTION_TEMPLATES = {
  analyze:
    "请分析以下 Excel 数据并给出洞察：\n\n" +
    "{CONTEXT}\n\n" +
    "请从趋势、异常、改进建议等维度给出结论。",
  formula:
    "请根据以下需求生成 Excel 公式：\n\n" +
    "需求描述：{USER_INPUT}\n\n" +
    "{CONTEXT}\n\n" +
    "请给出可直接使用的公式并解释每个参数的含义。",
  clean:
    "请为以下数据提供清洗和格式化建议：\n\n" +
    "{CONTEXT}\n\n" +
    "请指出潜在的问题（缺失值、格式不一致、异常值等）并给出处理方案。",
};

/** One-click templates shown as chips above the input. */
export const QUICK_TEMPLATES: Array<{
  id: string;
  icon: string;
  label: string;
  prompt: string;
  requiresSelection?: boolean;
}> = [
  {
    id: "explain-formula",
    icon: "🔍",
    label: "解释公式",
    prompt: "请解释以下 Excel 公式的含义和每个参数的作用：\n\n{FORMULA}",
    requiresSelection: true,
  },
  {
    id: "translate-formula",
    icon: "🌐",
    label: "公式转中文",
    prompt: "请把以下 Excel 公式的逻辑用中文描述清楚：\n\n{FORMULA}",
    requiresSelection: true,
  },
  {
    id: "vlookup-helper",
    icon: "🔗",
    label: "VLOOKUP",
    prompt:
      "请基于以下数据生成 VLOOKUP / XLOOKUP 公式：\n\n{CONTEXT}\n\n" +
      "查找条件：{USER_INPUT}\n请给出可直接使用的公式并解释。",
    requiresSelection: true,
  },
  {
    id: "pivot-suggestion",
    icon: "📊",
    label: "透视表建议",
    prompt:
      "请基于以下数据建议一个数据透视表的结构：\n\n{CONTEXT}\n\n" +
      "请说明行列字段、值字段以及推荐的可视化方式。",
    requiresSelection: true,
  },
  {
    id: "chart-suggestion",
    icon: "📈",
    label: "图表建议",
    prompt:
      "请基于以下数据推荐合适的 Excel 图表类型：\n\n{CONTEXT}\n\n" +
      "请说明推荐理由及具体操作步骤。",
    requiresSelection: true,
  },
  {
    id: "sql-translate",
    icon: "🗃️",
    label: "SQL 转公式",
    prompt:
      "请把以下 SQL 语句转换为 Excel 公式或 Power Query M 公式：\n\n{CONTEXT}\n\nSQL：{USER_INPUT}",
  },
  {
    id: "vba-snippet",
    icon: "💻",
    label: "VBA 片段",
    prompt: "请根据以下需求生成一段 VBA 宏代码：\n\n需求：{USER_INPUT}\n\n{CONTEXT}",
  },
  {
    id: "summary",
    icon: "📝",
    label: "数据摘要",
    prompt:
      "请为以下数据生成一段简洁的摘要说明（中英文均可）：\n\n{CONTEXT}",
    requiresSelection: true,
  },
];

/** localStorage key for the active theme. */
export const STORAGE_KEY_THEME = "deepseek_excel_theme_v1";

/** Application title shown in the header. */
export const APP_TITLE = "DeepSeek Excel Assistant";

/** Application version surfaced in the about dialog. */
export const APP_VERSION = "1.0.0";

/* ----------------------------------------------------------------- */
/* Slash commands (PRD-02)                                            */
/* ----------------------------------------------------------------- */

/** Definition of a single slash command shown when the user types "/".
 *  Templates use placeholders that get substituted at insert time:
 *    {CONTEXT}     - current Excel selection as Markdown
 *    {FORMULA}     - formula under the active cell (if any)
 *    {USER_INPUT}  - cursor position for follow-up text
 */
export interface SlashCommand {
  /** Command id, also the trigger (e.g. "analyze"). */
  id: string;
  /** Trigger after the slash, e.g. "/analyze". */
  trigger: string;
  /** Emoji icon shown in the menu. */
  icon: string;
  /** Short label (1-2 words). */
  label: string;
  /** One-line description shown in the menu. */
  hint: string;
  /** Template body. The {USER_INPUT} slot is where the cursor lands. */
  template: string;
  /** Requires an active Excel selection. */
  requiresSelection: boolean;
  /** Optional category for grouping in the menu. */
  category: "analyze" | "formula" | "clean" | "explain" | "visualize" | "summary";
}

export const SLASH_COMMANDS: SlashCommand[] = [
  {
    id: "analyze",
    trigger: "/analyze",
    icon: "📊",
    label: "分析选区",
    hint: "让 AI 解释数据趋势 / 异常 / 改进建议",
    template: "请分析以下 Excel 选区并给出洞察：\n\n{CONTEXT}",
    requiresSelection: true,
    category: "analyze",
  },
  {
    id: "formula",
    trigger: "/formula",
    icon: "📝",
    label: "生成公式",
    hint: "基于数据生成 Excel 公式 + 解释",
    template: "请基于以下 Excel 数据生成公式：\n\n{CONTEXT}\n\n需求：{USER_INPUT}",
    requiresSelection: true,
    category: "formula",
  },
  {
    id: "clean",
    trigger: "/clean",
    icon: "🧹",
    label: "清洗数据",
    hint: "给出数据清洗 / 格式化建议",
    template: "请基于以下数据给出清洗建议：\n\n{CONTEXT}",
    requiresSelection: true,
    category: "clean",
  },
  {
    id: "explain",
    trigger: "/explain",
    icon: "🔍",
    label: "解释公式",
    hint: "解释活动单元格的公式含义",
    template: "请解释以下 Excel 公式的含义和每个参数的作用：\n\n{FORMULA}",
    requiresSelection: true,
    category: "explain",
  },
  {
    id: "translate",
    trigger: "/translate",
    icon: "🌐",
    label: "公式转中文",
    hint: "把公式逻辑用中文描述清楚",
    template: "请把以下 Excel 公式的逻辑用中文描述清楚：\n\n{FORMULA}",
    requiresSelection: true,
    category: "explain",
  },
  {
    id: "vlookup",
    trigger: "/vlookup",
    icon: "🔗",
    label: "VLOOKUP",
    hint: "基于数据生成 VLOOKUP / XLOOKUP 公式",
    template:
      "请基于以下数据生成 VLOOKUP / XLOOKUP 公式：\n\n{CONTEXT}\n\n" +
      "查找条件：{USER_INPUT}\n请给出可直接使用的公式并解释。",
    requiresSelection: true,
    category: "formula",
  },
  {
    id: "pivot",
    trigger: "/pivot",
    icon: "📊",
    label: "透视表建议",
    hint: "建议数据透视表结构 + 字段",
    template:
      "请基于以下数据建议一个数据透视表的结构：\n\n{CONTEXT}\n\n" +
      "请说明行列字段、值字段以及推荐的可视化方式。",
    requiresSelection: true,
    category: "visualize",
  },
  {
    id: "chart",
    trigger: "/chart",
    icon: "📈",
    label: "图表推荐",
    hint: "推荐合适的图表类型 + 设计建议",
    template:
      "请基于以下数据推荐最合适的可视化图表，并说明原因：\n\n{CONTEXT}",
    requiresSelection: true,
    category: "visualize",
  },
  {
    id: "summary",
    trigger: "/summary",
    icon: "📋",
    label: "摘要",
    hint: "为选中内容生成 100 字摘要",
    template: "请为以下内容生成 100 字以内的摘要：\n\n{CONTEXT}",
    requiresSelection: false,
    category: "summary",
  },
];

/* ----------------------------------------------------------------- */
/* Command palette (PRD-03)                                           */
/* ----------------------------------------------------------------- */

/** A single action surfaced in the ⌘K command palette. */
export interface CommandPaletteItem {
  /** Stable id used in the keyboard layer and in telemetry. */
  id: string;
  /** Emoji icon shown at the left of each row. */
  icon: string;
  /** Human-readable label (Chinese). */
  label: string;
  /** One-line description shown in dim text. */
  hint: string;
  /** Ribbon / system command id that this maps to. The dispatcher in
   *  taskpane.ts recognizes these strings. */
  action: string;
  /** Optional category for grouping when searching. */
  category: "ribbon" | "session" | "view" | "help";
  /** Key combo for the shortcut hint, e.g. "⌘+2". */
  shortcut: string;
  /** Set true if this command needs an active Excel selection to make sense. */
  requiresSelection?: boolean;
}

export const COMMAND_PALETTE_ITEMS: CommandPaletteItem[] = [
  { id: "analyze", icon: "📊", label: "分析选区", hint: "让 AI 解读当前选区", action: "analyzeSelection", category: "ribbon", shortcut: "⌘+2", requiresSelection: true },
  { id: "formula", icon: "📝", label: "生成公式", hint: "根据上下文生成 Excel 公式", action: "generateFormula", category: "ribbon", shortcut: "⌘+3", requiresSelection: true },
  { id: "clean", icon: "🧹", label: "数据清洗", hint: "给出数据清洗 / 格式化建议", action: "cleanData", category: "ribbon", shortcut: "⌘+4", requiresSelection: true },
  { id: "diagnose", icon: "🩺", label: "诊断公式", hint: "扫描 #REF! 等错误并修复", action: "diagnoseFormulas", category: "ribbon", shortcut: "⌘+0", requiresSelection: true },
  { id: "insert", icon: "📥", label: "插入回复", hint: "把最后一条 AI 回复写回 Excel", action: "insertLastReply", category: "ribbon", shortcut: "⌘+5" },
  { id: "export", icon: "📤", label: "导出对话", hint: "把当前会话保存为 Markdown", action: "exportCurrentSession", category: "ribbon", shortcut: "⌘+6" },
  { id: "clear", icon: "🗑️", label: "清空对话", hint: "清空当前会话所有消息", action: "clearCurrentChat", category: "ribbon", shortcut: "⌘+7" },
  { id: "theme", icon: "🌗", label: "切换主题", hint: "在浅色 / 深色 / 跟随系统间切换", action: "toggleTheme", category: "view", shortcut: "⌘+8" },
  { id: "settings", icon: "⚙️", label: "打开设置", hint: "API Key、模型、温度等", action: "openSettings", category: "view", shortcut: "⌘+9" },
  { id: "history", icon: "🕘", label: "历史记录", hint: "查看过去的会话", action: "toggleHistory", category: "session", shortcut: "⌘+H" },
  { id: "newchat", icon: "➕", label: "新建对话", hint: "开启一个新会话", action: "newSession", category: "session", shortcut: "⌘+N" },
  { id: "shortcuts", icon: "❓", label: "快捷键帮助", hint: "查看所有快捷键", action: "showShortcuts", category: "help", shortcut: "⌘+/" },
  { id: "kb", icon: "📚", label: "知识库", hint: "管理上传的参考资料，AI 会自动检索引用", action: "toggleKnowledgeBase", category: "view", shortcut: "⌘+B" },
  { id: "share", icon: "🔗", label: "分享会话", hint: "生成可分享的链接或导出 JSON", action: "shareSession", category: "view", shortcut: "⌘+⇧+S" },
  { id: "usage", icon: "📊", label: "用量看板", hint: "查看今日 / 累计 Tokens、成本与工具调用统计", action: "usageDashboard", category: "view", shortcut: "⌘+D" },
];

/* ----------------------------------------------------------------- */
/* Shortcut reference (for the help dialog)                           */
/* ----------------------------------------------------------------- */

export interface ShortcutEntry {
  /** Key combo shown to the user, e.g. "⌘+K". */
  combo: string;
  /** What this combo does. */
  label: string;
  /** Where this combo is active: "global" = anywhere in taskpane,
   *  "input" = only when the chat input is focused. */
  scope: "global" | "input";
}

export const SHORTCUTS: ShortcutEntry[] = [
  { combo: "⌘+K", label: "打开 / 关闭命令面板", scope: "global" },
  { combo: "⌘+/", label: "打开快捷键帮助", scope: "global" },
  { combo: "Esc", label: "关闭命令面板 / 弹窗", scope: "global" },
  { combo: "⌘+Enter", label: "发送消息", scope: "input" },
  { combo: "Shift+Enter", label: "在输入框内换行", scope: "input" },
  { combo: "⌘+1", label: "聚焦对话", scope: "global" },
  { combo: "⌘+2", label: "分析选区", scope: "global" },
  { combo: "⌘+3", label: "生成公式", scope: "global" },
  { combo: "⌘+4", label: "数据清洗", scope: "global" },
  { combo: "⌘+5", label: "插入回复到 Excel", scope: "global" },
  { combo: "⌘+6", label: "导出当前对话", scope: "global" },
  { combo: "⌘+7", label: "清空当前对话", scope: "global" },
  { combo: "⌘+8", label: "切换主题", scope: "global" },
  { combo: "⌘+9", label: "打开设置", scope: "global" },
  { combo: "⌘+N", label: "新建对话", scope: "global" },
  { combo: "⌘+H", label: "打开历史记录", scope: "global" },
  { combo: "⌘+B", label: "打开 / 关闭知识库", scope: "global" },
  { combo: "⌘+⇧+S", label: "分享当前会话", scope: "global" },
  { combo: "⌘+D", label: "打开 / 关闭用量看板", scope: "global" },
  { combo: "?", label: "在命令面板输入 ? 显示所有快捷键", scope: "global" },
];

/* ----------------------------------------------------------------- */
/* Formula library - quick-insert cards for common Excel formulas.    */
/* ----------------------------------------------------------------- */

export interface FormulaCard {
  /** Unique id - used for telemetry and as the data attribute. */
  id: string;
  /** Formula name shown on the card. */
  name: string;
  /** Short one-line description. */
  desc: string;
  /** Prompt fragment that gets pre-filled into the chat input. The
   *  user can edit it before sending. */
  prompt: string;
}

export const FORMULA_CARDS: FormulaCard[] = [
  {
    id: "vlookup",
    name: "VLOOKUP",
    desc: "按列垂直查找",
    prompt:
      "请给我一个 VLOOKUP 公式，根据「{KEY_COL}」在「{TABLE}」中查找对应的「{RESULT_COL}」。\n" +
      "匹配模式：FALSE（精确匹配）\n" +
      "如果有数据范围，请使用绝对引用。",
  },
  {
    id: "xlookup",
    name: "XLOOKUP",
    desc: "现代通用查找",
    prompt:
      "请用 XLOOKUP（Excel 365 / 2021+）实现：\n" +
      "查找值在「{KEY_COL}」中，查找范围在「{TABLE}」中，\n" +
      "返回列在「{RESULT_COL}」中，找不到时返回「{NOT_FOUND}」。",
  },
  {
    id: "index-match",
    name: "INDEX+MATCH",
    desc: "灵活的双向查找",
    prompt:
      "请用 INDEX + MATCH 组合实现：\n" +
      "用「{KEY_COL}」中的值在「{TABLE}」中查找行号，\n" +
      "返回「{RESULT_COL}」中对应位置的值。",
  },
  {
    id: "sumifs",
    name: "SUMIFS",
    desc: "多条件求和",
    prompt:
      "请给我一个 SUMIFS 公式：\n" +
      "求和范围：「{SUM_RANGE}」\n" +
      "条件 1：「{CRIT_RANGE_1}」等于「{CRIT_VAL_1}」\n" +
      "条件 2：「{CRIT_RANGE_2}」等于「{CRIT_VAL_2}」\n" +
      "（可继续添加条件）",
  },
  {
    id: "countifs",
    name: "COUNTIFS",
    desc: "多条件计数",
    prompt:
      "请给我一个 COUNTIFS 公式：\n" +
      "条件 1：「{CRIT_RANGE_1}」等于「{CRIT_VAL_1}」\n" +
      "条件 2：「{CRIT_RANGE_2}」大于「{CRIT_VAL_2}」",
  },
  {
    id: "iferror",
    name: "IFERROR",
    desc: "错误值兜底",
    prompt:
      "请把以下公式包一层 IFERROR，让出错时返回指定值：\n" +
      "原公式：{FORMULA}\n" +
      "出错时返回：{FALLBACK}",
  },
  {
    id: "if",
    name: "IF",
    desc: "条件分支",
    prompt:
      "请给我一个 IF 嵌套公式：\n" +
      "条件 1：{CONDITION_1} → 返回 {VALUE_1}\n" +
      "条件 2：{CONDITION_2} → 返回 {VALUE_2}\n" +
      "否则 → 返回 {VALUE_OTHER}",
  },
  {
    id: "text-join",
    name: "TEXTJOIN",
    desc: "文本拼接（去空）",
    prompt:
      "请用 TEXTJOIN 把「{RANGE}」范围内所有非空单元格用分隔符「{DELIM}」拼接起来，\n" +
      "并忽略空白单元格。",
  },
  {
    id: "date-diff",
    name: "DATEDIF",
    desc: "日期差值",
    prompt:
      "请用 DATEDIF 计算两个日期的差值：\n" +
      "开始日期：「{START_CELL}」\n" +
      "结束日期：「{END_CELL}」\n" +
      "需要单位：Y（年）/ M（月）/ D（日）",
  },
  {
    id: "left-mid-right",
    name: "LEFT/MID/RIGHT",
    desc: "文本截取",
    prompt:
      "请用 LEFT / MID / RIGHT 处理以下字符串：\n" +
      "源文本：{SOURCE_CELL}\n" +
      "操作：{OPERATION}（例如：从第 3 位开始取 5 个字符）",
  },
];