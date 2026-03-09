# VS Code 风格化写作助手

一款深度集成于 VS Code 的智能写作辅助插件，通过 14 维度分析自动学习写作风格，生成风格高度一致的文章，助您实现“所想即所得”的风格化创作。

---

## 🚀 快速开始

### 1. 安装依赖
```bash
npm install
```

### 2. 编译插件
```bash
npm run compile
```

### 3. 启动调试
1. 在 VS Code 中打开本项目。
2. 按 `F5` 启动调试。
3. 在弹出的新窗口（扩展开发宿主）中即可测试插件功能。

---

## ✨ 核心功能

### 1. 风格学习（14 维度深度分析）
- **维度体系**：涵盖词汇、句式、语气、结构、修辞（核心）、时间、空间、逻辑、密度、情感（深度）、受众、文化、叙事、节奏（语境）。
- **档案管理**：支持多风格档案的创建、切换、重命名与删除。

### 2. 风格化生成
- **灵活输出**：支持生成“完整文章”或“写作提纲”（提纲默认提供 2 套以上思路）。
- **去 AI 味**：集成 Humanizer-zh 处理逻辑，弱化机械模板感。
- **自动落地**：生成结果自动保存至配置的「文集」Markdown 文件中。

### 3. 素材智能管理
- **智能提取**：自动从样本中提取金句、段落及名言。
- **多维检索**：支持关键词、标签、类型过滤，并可一键导出为 Markdown。
- **风格隔离**：素材按当前激活风格独立存储与检索。

---

## 🛠 详细指南

### 常用操作
- **导入样本**：`Cmd+Shift+I` (Mac) / `Ctrl+Shift+I` (Win)
- **生成文章**：`Cmd+Shift+G` (Mac) / `Ctrl+Shift+G` (Win)
- **搜索素材**：`Cmd+Shift+S` (Mac) / `Ctrl+Shift+S` (Win)
- **分析当前风格**：`Cmd+Shift+A` (Mac) / `Ctrl+Shift+A` (Win)

### 配置选项
在 VS Code 设置中搜索 "写作助手"：
- `writingAgent.autoExtract`: 自动提取素材（默认：true）
- `writingAgent.ai.channel`: AI 通道选择（`api` / `ide-chat`）
- `writingAgent.storage.path`: 内容存储根目录（默认：`~/Downloads/writing-angent`）

---

## 🏗 技术架构

- **语言框架**: TypeScript + VS Code Extension API
- **NLP 处理**: 内置轻量分词 + Mammoth (DOCX 解析)
- **UI 实现**: VS Code TreeView + Webview API
- **存储机制**: 外部 JSON 索引 + 文集 Markdown 文件

### 项目结构
```text
writing-agent/
├── src/
│   ├── extension.ts          # 插件入口
│   ├── core/                 # 核心逻辑（风格引擎、素材管理）
│   ├── providers/            # 侧边栏视图驱动
│   └── panels/               # Webview 面板（搜索、报告）
├── resources/                # 静态资源
└── package.json              # 插件配置
```

---

## 📈 开发进度 (v0.0.1 MVP)

### 已实现
- ✅ 14 维度风格分析引擎
- ✅ 素材提取、隔离存储与多维检索
- ✅ 文章生成框架及 Humanizer-zh 集成
- ✅ 多风格档案管理及文集落盘系统
- ✅ **维护工具：自动清理空文件夹 (2026-03-09)**

### 进行中
- 🔄 文章生成算法优化
- 🔄 用户界面体验升级

---

## 🤝 贡献与开源

本项目采用 **[MIT 许可证](LICENSE)** 开源。

### 如何贡献
1. **Fork** 本仓库。
2. 创建特性分支：`git checkout -b feature/AmazingFeature`。
3. 提交更改：`git commit -m 'Add some AmazingFeature'`。
4. 推送到分支：`git push origin feature/AmazingFeature`。
5. 开启 **Pull Request**。

### 维护反思 (2026-03-09)
我们在清理空文件夹时发现项目重构后易产生残留。建议定期执行清理脚本，并考虑在 `package.json` 中集成 `clean` 钩子。

---

## 📞 联系我们
如有任何问题或建议，欢迎提交 [Issue](https://github.com/daydayup2026/writing-agent/issues)。
