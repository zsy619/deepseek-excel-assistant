# DeepSeek Excel Assistant

> 在 Excel 任务窗格中与 DeepSeek AI 对话，辅助数据分析、公式生成、内容优化。

一个基于 [Yeoman Office Generator](https://github.com/OfficeDev/generator-office) 的 Excel Task Pane 插件，把 DeepSeek 大模型对话能力直接带到 Excel 界面里。

## ✨ 功能特性

| 模块 | 功能 |
| --- | --- |
| 💬 **聊天对话** | 支持 Markdown 渲染 + 代码高亮，流式输出逐字显示，可随时停止 |
| ⚙️ **配置管理** | API Key / 端点 / 模型 / Temperature / Top P / 系统提示词，全部持久化 |
| 📚 **历史会话** | 最多保留 50 个会话，支持搜索、重命名、导出 Markdown、批量删除 |
| 📊 **Excel 上下文** | 一键抓取当前选区数据，自动拼接到 prompt 里 |
| 📌 **插入单元格** | 把 AI 回复一键写回当前单元格 |
| 🎨 **Office 主题** | 自动适配深色/浅色模式，Fluent UI 风格 |
| ⏹ **可中断** | AbortController 实时取消生成中的请求 |

## 📁 项目结构

```
deepseek-excel-assistant/
├── manifest.xml                # Office 插件清单
├── package.json
├── webpack.config.js           # 双 bundle: taskpane + commands
├── tsconfig.json
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html
│   │   ├── taskpane.ts         # 主入口
│   │   ├── taskpane.css        # 全局样式（Fluent UI）
│   │   ├── components/
│   │   │   ├── ChatWindow.ts
│   │   │   ├── ChatMessage.ts
│   │   │   ├── SettingsPanel.ts
│   │   │   ├── HistoryPanel.ts
│   │   │   ├── QuickActions.ts
│   │   │   └── MarkdownRenderer.ts
│   │   ├── services/
│   │   │   ├── deepseek.ts     # API + SSE 流式
│   │   │   ├── storage.ts      # localStorage 封装
│   │   │   └── excel.ts        # Office.js 封装
│   │   ├── types/index.ts
│   │   └── utils/
│   │       ├── constants.ts
│   │       └── helpers.ts
│   └── commands/
│       └── commands.ts         # 功能区按钮命令
└── assets/                     # 图标
```

## 🚀 快速开始

### 前置要求

- Node.js >= 18
- Office 365 / Excel 桌面版（Mac 或 Windows）
- DeepSeek API Key（[申请地址](https://platform.deepseek.com/)）

### 安装

```bash
# 1. 全局安装 Office 脚手架（首次需要）
npm install -g yo generator-office

# 2. 安装项目依赖
cd deepseek-excel-assistant
npm install

# 3. 安装 Office 本地开发证书（首次需要，会修改系统信任设置）
npx office-addin-dev-certs install
```

### 启动开发服务器

```bash
# 终端 1：启动本地 HTTPS 开发服务器（端口 3000）
npm run dev-server

# 终端 2：启动 Excel 并加载插件（自动打开 Excel）
npm start
```

第一次运行时 Excel 会弹出"加载开发插件"的提示，勾选"信任此加载项"即可。

### 在 Excel 中使用

1. 打开 Excel，会看到一个名为 **DeepSeek** 的自定义选项卡
2. 点击 **打开对话** 打开右侧任务窗格
3. 点击任务窗格右上角 ⚙️ 图标，填入 DeepSeek API Key
4. 保存后即可开始对话

### 清除开发环境

```bash
npm run stop             # 停止调试
npx office-addin-dev-certs uninstall  # 卸载证书（可选）
```

## 🏗️ 构建生产版本

```bash
npm run build            # 生成 dist/ 目录
npm run validate         # 校验 manifest.xml
```

`dist/` 目录里的内容可以直接上传到自己的 CDN 或 Web 服务器。

## 🔧 配置项说明

| 项 | 默认值 | 说明 |
| --- | --- | --- |
| **API Key** | 空 | DeepSeek 平台申请，格式 `sk-xxx...`，仅本地存储 |
| **API 端点** | `https://api.deepseek.com` | 兼容 OpenAI 格式，可指向任意代理 |
| **模型** | `deepseek-chat` | `deepseek-chat` 通用对话 / `deepseek-reasoner` 深度推理 |
| **Temperature** | 0.7 | 0~2，越高越发散 |
| **Max Tokens** | 2048 | 256~8192，单次回复上限 |
| **Top P** | 0.9 | 0~1，核采样阈值 |
| **系统提示词** | 中文 Excel 助手 | 控制 AI 的人设与回答风格 |

## 🔌 DeepSeek API 规范

本插件直接对接 DeepSeek 的 OpenAI 兼容端点：

```
POST {baseUrl}/chat/completions
Headers:
  Authorization: Bearer {apiKey}
  Content-Type: application/json
Body:
  {
    "model": "deepseek-chat",
    "messages": [...],
    "temperature": 0.7,
    "max_tokens": 2048,
    "top_p": 0.9,
    "stream": true   // true = SSE 流式
  }
```

支持自定义端点 → 兼容任何暴露 OpenAI 格式的代理（Azure OpenAI、OpenRouter、自建网关等）。

## 📋 快捷操作说明

| 按钮 | 行为 |
| --- | --- |
| 📊 分析选区 | 自动抓取选中数据，AI 输出趋势/异常/建议 |
| 📝 生成公式 | 弹窗输入需求，AI 生成可用的 Excel 公式 |
| 🧹 数据清洗 | 抓取选区，AI 输出清洗方案（缺失值/格式/异常） |
| 📌 插入单元格 | 把 AI 最近一次回答写到当前选中单元格 |
| ⏹ 停止 | 中断流式响应（仅在生成中显示） |

## 🛡️ 隐私与安全

- ✅ API Key 仅保存在浏览器 `localStorage`，**不上传任何服务器**
- ✅ 所有 DeepSeek 请求直接发给 `api.deepseek.com` 或用户自定义端点
- ✅ 不收集任何用户行为数据
- ⚠️ 启用第三方代理时请确认端点的可信度

## 🐛 常见问题

**Q: 任务窗格打不开？**
A: 检查 `npm run dev-server` 是否在 3000 端口正常运行；检查 `npx office-addin-dev-certs install` 是否成功。

**Q: 报 401 / 鉴权失败？**
A: API Key 错误或已过期，到 ⚙️ 设置里重新填入。

**Q: 流式输出卡住？**
A: 检查网络；点 ⏹ 停止重试；某些代理不支持 SSE，可在端点后加上 `?stream=false`（不支持，需改代码）。

**Q: 想在 Excel 网页版使用？**
A: 把 `manifest.xml` 里 `localhost:3000` 替换成可公网访问的地址后上传到 [AppSource](https://learn.microsoft.com/en-us/office/dev/store/)。

## 📜 License

MIT