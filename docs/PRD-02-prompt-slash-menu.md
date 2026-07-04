# PRD-02 智能提示词 `/` 菜单

## 1. 背景与目标

**痛点**：用户不知道 AI 能干什么，每次都要想"我该怎么问"。Quick Action 按钮只能点固定的几个场景，灵活性不够。

**目标**：在输入框里输入 `/` 触发命令菜单，列出所有支持的 AI 能力，键盘上下选 + Enter 触发。

**差异化**：把 Excel Copilot 的 `/edit /format /chart` 体验带到 DeepSeek 助手。

## 2. 验收点

- [ ] 输入框 focus 后，输入 `/` → 弹出 9 个候选命令的菜单
- [ ] 菜单每项含：图标、命令名、提示文本、快捷键（如有）
- [ ] 键盘 `↑/↓` 移动高亮，`Enter` 选中，`Esc` 关闭
- [ ] 鼠标点击也能选中
- [ ] 选中后：自动填入 prompt 模板到输入框（光标停在 `{USER_INPUT}` 位置）
- [ ] `/analyze` / `/formula` / `/clean` 等命令的模板定义在 `utils/constants.ts`
- [ ] 输入框空时输入 `/` 也触发菜单
- [ ] 输入框已有文字时（如 `帮我`），按 `/` 不触发（避免歧义）
- [ ] 菜单可滚动，超过 6 项时滚动

## 3. 命令清单（v1）

| 命令 | 标签 | 模板 | 需要选区 |
|------|------|------|----------|
| `/analyze` | 分析选区 | `请分析以下 Excel 选区并给出洞察：\n\n{CONTEXT}` | 是 |
| `/formula` | 生成公式 | `请基于以下数据生成 Excel 公式：\n\n{CONTEXT}\n\n需求：{USER_INPUT}` | 是 |
| `/clean` | 清洗数据 | `请基于以下数据给出清洗建议：\n\n{CONTEXT}` | 是 |
| `/explain` | 解释公式 | `请解释以下公式：\n\n{FORMULA}` | 是 |
| `/translate` | 公式转中文 | `请用中文描述：\n\n{FORMULA}` | 是 |
| `/vlookup` | VLOOKUP | `请基于以下数据生成 VLOOKUP 公式：\n\n{CONTEXT}\n\n查找：{USER_INPUT}` | 是 |
| `/pivot` | 透视表建议 | `请基于以下数据建议透视表：\n\n{CONTEXT}` | 是 |
| `/chart` | 图表推荐 | `请基于以下数据推荐可视化方式：\n\n{CONTEXT}` | 是 |
| `/summary` | 摘要 | `请为以下内容生成 100 字摘要：\n\n{CONTEXT}` | 否 |

## 4. 技术方案

| 层 | 改动 |
|----|------|
| `components/PromptMenu.ts` | 新组件，absolute 定位下拉菜单，键盘/鼠标事件 |
| `utils/constants.ts` | 新增 `SLASH_COMMANDS` 数组 |
| `components/ChatWindow.ts` | 在 input 上挂 keydown 监听器，`/` 触发菜单 |
| CSS | `.prompt-menu` 样式（沿用 Fluent UI 风格） |

## 5. 边界

- 移动端：菜单改为底部 sheet（暂不支持，标记 TODO）
- 国际化：菜单文字先写死中文
- 用户自定义命令：v1 不支持

## 6. 测试用例

1. 聚焦输入框 → 输入 `/` → 菜单弹出
2. 键盘 ↓ 移动 → Enter 选中 → 模板填入
3. 鼠标点击 → 模板填入
4. Esc → 菜单关闭
5. 在"帮我"后输入 `/` → 不触发（避免歧义）
6. 选中命令后按 Backspace 删字符 → 模板被清除
