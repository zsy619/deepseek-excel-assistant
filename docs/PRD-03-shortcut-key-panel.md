# PRD-03 全局快捷键 + 命令面板

## 1. 背景与目标

**痛点**：
- 用户想用快捷键触发功能，但 Office ribbon 按钮没有键盘绑定
- 用户在任务窗内输入 prompt 时，想快速切换主题、清空、重新生成

**目标**：
- 在 taskpane 内实现统一的快捷键体系
- `⌘+K` 调出命令面板（VS Code 风格）
- 8 个 ribbon 功能都有 ⌘+数字 快捷键
- 任务窗支持 `?` 弹出"所有快捷键"列表

**差异化**：把 IDE 级的交互体验带入 Office 助手。

## 2. 快捷键总表

| 快捷键 | 动作 | 范围 |
|--------|------|------|
| `⌘+K` | 打开/关闭命令面板 | 全局 |
| `⌘+/` | 打开快捷键帮助 | 全局 |
| `⌘+B` | 加粗（输入框内） | 输入框 |
| `⌘+Enter` | 发送消息 | 输入框 |
| `Shift+Enter` | 换行 | 输入框 |
| `Esc` | 关闭命令面板/对话框 | 全局 |
| `⌘+1` | 打开对话 | 全局 |
| `⌘+2` | 分析选区 | 全局 |
| `⌘+3` | 生成公式 | 全局 |
| `⌘+4` | 数据清洗 | 全局 |
| `⌘+5` | 插入回复 | 全局 |
| `⌘+6` | 导出对话 | 全局 |
| `⌘+7` | 清空对话 | 全局 |
| `⌘+8` | 切换主题 | 全局 |
| `⌘+9` | 打开设置 | 全局 |

**注意**：macOS 用 `⌘`，Windows 用 `Ctrl`。代码用 `e.metaKey || e.ctrlKey` 兼容。

## 3. 验收点

- [ ] 在 taskpane 任意位置按 `⌘+K` → 命令面板弹出
- [ ] 命令面板 input 自动 focus
- [ ] 输入命令名（如"清空"）→ 列表过滤
- [ ] `↑/↓` 移动高亮，`Enter` 选中执行，`Esc` 关闭
- [ ] 输入 `?` → 显示所有快捷键列表
- [ ] 输入框内 `⌘+Enter` 发送消息（不管按钮 click）
- [ ] `Shift+Enter` 在输入框内换行
- [ ] 8 个 ribbon 按钮的快捷键在 ribbon 控件上加 `Keytip`（需要在 manifest 加 `keytip` 字段，但 v1 可选）

## 4. 技术方案

| 层 | 改动 |
|----|------|
| `components/CommandPalette.ts` | 新组件，Modal 风格命令面板 |
| `components/ShortcutHelp.ts` | 新组件，快捷键帮助弹窗 |
| `components/ChatWindow.ts` | 注册全局 keydown 监听，分发到 handler |
| `utils/constants.ts` | 定义 `COMMANDS` 数组 + `SHORTCUTS` 数组 |
| `manifest.xml` | ribbon 按钮加 `keytip` 属性（v1 可选） |

## 5. 命令面板命令清单

```ts
[
  { id: "analyze",    label: "分析选区",   shortcut: "⌘+2", action: "analyzeSelection" },
  { id: "formula",    label: "生成公式",   shortcut: "⌘+3", action: "generateFormula" },
  { id: "clean",      label: "数据清洗",   shortcut: "⌘+4", action: "cleanData" },
  { id: "insert",     label: "插入回复",   shortcut: "⌘+5", action: "insertLastReply" },
  { id: "export",     label: "导出对话",   shortcut: "⌘+6", action: "exportCurrentSession" },
  { id: "clear",      label: "清空对话",   shortcut: "⌘+7", action: "clearCurrentChat" },
  { id: "theme",      label: "切换主题",   shortcut: "⌘+8", action: "toggleTheme" },
  { id: "settings",   label: "打开设置",   shortcut: "⌘+9", action: "openSettings" },
  { id: "history",    label: "历史记录",   shortcut: "⌘+H", action: "toggleHistory" },
  { id: "newchat",    label: "新建对话",   shortcut: "⌘+N", action: "newSession" },
  { id: "shortcuts",  label: "快捷键帮助", shortcut: "⌘+/", action: "showShortcuts" },
]
```

## 6. 边界

- Excel 内 WebView 对 `⌘` 键的捕获可能有延迟 → 用 capture phase 监听
- 输入框内只拦截 `⌘+Enter` 和 `Shift+Enter`，其他快捷键不抢
- 命令面板打开时禁用其他快捷键

## 7. 测试用例

1. 按 `⌘+K` → 命令面板出现，输入"清空" → Enter → 清空对话
2. 按 `⌘+/` → 快捷键帮助出现
3. 输入框内按 `⌘+Enter` → 消息发送
4. 输入框内按 `Shift+Enter` → 换行（不发送）
5. 命令面板打开时按 `Esc` → 关闭
6. 命令面板输入 `?` → 显示快捷键列表
