# PRD-01 公式诊断与一键修复

## 1. 背景与目标

**痛点**：用户最常遇见的 #REF!、#DIV/0!、#N/A、#VALUE!、#NAME? 五大错误公式，目前没有 AI 助手能直接定位并修复。手动排查往往要逐个点击公式、查 Stack Overflow、试错。

**目标**：用户选中一个或多个公式所在的单元格 → 一键让 AI 扫描 → 列出每条错误公式 + 原因 + 修复建议 → 一键覆盖修复。

**差异化**：从"AI 解释"升级为"AI 修复"——**闭环**。

## 2. 验收点

- [ ] 新增 ribbon 按钮 **"诊断公式"**（带专属图标），归到"对话"分组
- [ ] 选中区域（>= 1 个公式单元格）→ 点击 → 弹 toast「正在扫描 N 个公式…」
- [ ] 扫描完成后渲染一个**错误报告卡片**（在对话区），列出：
  - 单元格地址
  - 错误类型
  - 当前公式
  - 错误原因
  - 建议修复公式
  - 「应用修复」按钮（仅当公式非空时显示）
- [ ] 点击「应用修复」→ 弹出自定义 confirm → 确认后用 `writeFormula` 覆盖 → toast「已修复 1/3」
- [ ] 选区无公式 → toast「选区中未发现公式」
- [ ] 选区超过 500 个公式 → 弹出「性能警告」确认是否继续
- [ ] 修复后单元格闪烁一下（用 `flashSelectedCell`）

## 3. 技术方案

| 层 | 改动 |
|----|------|
| `services/excel.ts` | 新增 `scanFormulaErrors(range): Promise<FormulaError[]>` —— 用 `range.getFormulas()` + `range.getValues()` + `range.getSpecialCells` 一次拉取 |
| `services/deepseek.ts` | 新增 `diagnoseFormulas(errors): AsyncIterable<Token>` —— 复用 chatCompletionStream |
| `components/FormulaDiagnostics.ts` | 新组件，渲染错误报告卡片 |
| `components/ChatWindow.ts` | 在 `runAnalyzeSelection` 旁加 `runDiagnoseFormulas()` |
| `commands/commands.ts` + `manifest.xml` | 新增 `diagnoseFormulas` ribbon 按钮 + 生成 `ribbon-diagnose-{16,32,80}.png` |
| `utils/constants.ts` | 加 5 个标准错误码的中文描述映射 |

## 4. 边界 / 异常

- 跨 sheet 选区：递归处理每个 sheet
- 公式含外部引用：跳过，提示"含外部引用，AI 无法诊断"
- AI 建议的公式无法被 Excel 接受：捕获 `InvalidArgument` 异常，toast 提示"建议公式无法执行"

## 5. 性能

- 单 sheet ≤ 200 公式：直接处理
- 200-500：分批，每批 50
- > 500：弹窗确认

## 6. 测试用例

1. 选中含 `#REF!` 的公式 → 修复
2. 选中含 `#DIV/0!` 的公式 → 修复
3. 选中文本（非公式）→ 提示"未发现公式"
4. 选区为空 → 提示"请先选中公式"
5. AI 返回非法公式 → 错误处理 toast
