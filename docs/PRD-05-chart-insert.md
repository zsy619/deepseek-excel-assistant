# PRD-05 一键插入图表

## 1. 背景与目标

**痛点**：用户选中数据 → 不知道选什么图 → 手动从「插入 → 图表 → 选类型 → 配置」要 30 秒。
AI 推荐后再插入应当一气呵成。

**目标**：
- 选中数据 → ribbon 点击 → AI 推荐 3 个最合适的图表（带理由）
- 用户点卡片 → 立即插入到 Excel（默认靠右/下方）
- 支持的图表：折线 / 柱状 / 饼图 / 散点 / 面积 / 雷达

**差异化**：把"图表类型决策 + 执行"两步合并成一步。

## 2. 验收点

- [ ] 新 ribbon 按钮 **"插入图表"**（📈 蓝色）
- [ ] 选区必须 ≥ 2 个非空单元格，否则 toast 提示
- [ ] AI 推荐失败 → 给出本地 fallback：数据有「类别 + 数值」→ 柱状；多列数值 → 折线
- [ ] 一键插入：用 `Office.js` `worksheet.charts.add()`，位置在选区下方 2 行
- [ ] 插入成功后 toast「已插入 XX 图表」+ 给出"撤销"链接（一期：手动撤销）

## 3. 技术方案

| 层 | 改动 |
|----|------|
| `services/excel.ts` | 新增 `getSelectedRangeInfo()` + `insertChart(type, sourceAddress, title?)` |
| `services/deepseek.ts` | 新增 `recommendChartStream()` 流式返回 JSON 数组 |
| `components/ChartPicker.ts` | 新组件，3 张卡片，点击触发 `chart-picker-insert` 事件 |
| `components/ChatWindow.ts` | 新方法 `runInsertChart()` |
| `commands/commands.ts` + `manifest.xml` | 加 ribbon 按钮 |
| `assets/generate-icons.js` | 生成 `ribbon-chart-*.png` |

## 4. 数据流

1. 用户点 ribbon → 扫描选区 → 调 AI stream
2. AI 流式输出：`[{type:"bar", title:"...", reason:"..."}, ...]`
3. 渲染 3 张卡片：图标 + 类型 + 标题 + 理由
4. 用户点击卡片 → 插入到 Excel → toast

## 5. 边界

- 选区无表头（仅 1 列）：退化显示「请先添加表头」
- AI 推荐类型不在白名单（柱/折/饼/散/面/雷）：用 fallback
- 隐藏行 / 合并单元格：尝试正常插入，失败时 toast 给出原因

## 6. 测试用例

1. 选中 `A1:B10`（类别 + 数值）→ 推荐柱状图 → 插入成功
2. 选中 5 列数值 → 推荐折线图
3. 1 行数据 → toast「至少需要 2 行」
4. AI 返回非法类型 → 用 fallback 默认柱状
5. 同位置已有图表 → 覆盖还是偏移？v1 偏移 2 行
