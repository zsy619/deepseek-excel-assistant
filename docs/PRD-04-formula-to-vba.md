# PRD-04 公式 → VBA / Office Scripts 一键转写

## 1. 背景与目标

**痛点**：复杂场景下用户写了 10 个公式，每个 50 字符，但本质上能压缩成 5 行 VBA / Office Script。手动转写需要专业技能。

**目标**：选区选中含公式的单元格 → AI 识别公式逻辑 → 生成等价的 VBA Sub 或 Office Script → 一键复制到剪贴板。

**差异化**：从"AI 写公式"升级为"AI 写脚本"——**跃迁**。

## 2. 验收点

- [ ] 新增 ribbon 按钮 **"公式转 VBA"**，归到"工具"分组
- [ ] 选中含公式的单元格（>= 1 个）→ 点击 → 弹 toast「正在分析 N 个公式…」
- [ ] AI 分析完成后渲染一张代码卡片：
  - 标题：`VBA 脚本` 或 `Office Script`（用户可切换 tab）
  - 代码块（hljs excel 语法高亮）
  - 「复制到剪贴板」按钮
  - 「插入到模块」按钮（仅 VBA 模式可见，调用 Office API 添加到当前 workbook）
- [ ] 选区无公式 → toast「未发现公式」
- [ ] 公式跨 sheet 引用 → 保留 `Worksheets("Sheet2")` 引用
- [ ] 公式含外部引用 → 标注 warning「含外部引用，VBA 无法直接复现」

## 3. 技术方案

| 层 | 改动 |
|----|------|
| `services/excel.ts` | 新增 `getRangeFormulasAsCode()` —— 收集所有公式 + 单元格地址 |
| `services/deepseek.ts` | 新增 `translateToVBA(formulas): AsyncIterable<Token>` |
| `components/CodeExport.ts` | 新组件，渲染代码块 + 复制/插入按钮 |
| `components/ChatWindow.ts` | 加 `runFormulaToVBA()` |
| `commands/commands.ts` + `manifest.xml` | 加 ribbon 按钮 |
| `assets/generate-icons.js` | 生成 `ribbon-vba-{16,32,80}.png` |

## 4. AI Prompt 模板

```
你是一个 VBA 专家。请将以下 Excel 公式转换为等价的 VBA Sub 函数：

工作表：{SHEET_NAME}
单元格地址：{ADDRESSES}
公式：
{FORMULAS}

要求：
1. 用 Worksheets("Sheet").Range("A1") 引用单元格
2. 用 Application.WorksheetFunction 调 Excel 内置函数
3. 保持公式之间的依赖顺序
4. 输出完整 Sub 函数（包含 Sub ... End Sub）

VBA 代码：
```

## 5. Office Script 模式

按 Microsoft 365 趋势，Office Scripts（TypeScript 语法）正在替代 VBA。同样 prompt 切换到：

```
输出为 Office Script 语法（TypeScript），使用 ExcelScript API。
```

## 6. 边界

- 公式含数组公式 `{=...}`：保留 `FormulaArray` 属性
- 公式含名称管理器引用：标注 warning
- 公式含数据验证/条件格式：不在 v1 处理范围

## 7. "插入到模块" 技术细节

`Office.js` 本身不支持添加 VBA 模块。需要走两个路径：

**路径 A**：调 `Office.context.ui.displayDialogAsync` 打开一个特殊对话框，提示用户手动复制粘贴（最安全）

**路径 B**：用 Power Automate 桥接（v2 实现）

**v1 走路径 A**，按钮文案改为「复制并打开 VBA 编辑器」——复制后用 `Office.addin.showAsTaskpane` + 弹窗引导用户 `Alt+F11` 打开 VBA 编辑器。

## 8. 测试用例

1. 选中 `=SUM(A1:A10)` → 转换为 VBA
2. 选中 `=VLOOKUP(B1, Sheet2!A:D, 3, FALSE)` → 跨 sheet 引用保留
3. 选中文本（非公式）→ 提示"未发现公式"
4. 切换 VBA / Office Script tab → 重新生成
5. 点击复制 → 剪贴板验证（粘贴到任意文本框）
