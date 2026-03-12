# VS Code 风格化写作助手

一款深度集成于 VS Code 的智能写作辅助插件，通过 14 维度分析框架自动学习写作风格，助您生成文风统一、高质量的风格化文章。

---

## 目录

- [项目简介](#项目简介)
- [快速开始](#快速开始)
- [功能特性](#功能特性)
- [安装配置](#安装配置)
- [使用指南](#使用指南)
- [API 文档](#api-文档)
- [贡献指南](#贡献指南)
- [许可证](#许可证)

---

## 项目简介

**VS Code 风格化写作助手** 旨在解决创作过程中“风格不统一”和“素材碎片化”的痛点。它不仅能深度分析样本文章的遣词造句，还能智能沉淀素材，并在创作新文章时自动应用学习到的风格特征。

### 核心价值
- **精准模仿**：基于 14 维度分析框架，捕捉核心笔触。
- **效率提升**：自动化生成提纲与全文，支持多种输出模式。
- **沉浸体验**：完全融入 VS Code 工作流，无需切换应用。

---

## 快速开始

只需几步，即可在本地启动并体验写作助手：

1. **克隆项目**
   ```bash
   git clone https://github.com/daydayup2026/writing-agent.git
   cd writing-agent
   ```

2. **安装依赖**
   ```bash
   npm install
   ```

3. **编译项目**
   ```bash
   npm run compile
   ```

4. **启动调试**
   - 在 VS Code 中打开项目，按 `F5` 键。
   - 在弹出的 [扩展开发宿主] 窗口中即可使用插件。

---

## 功能特性

### 1. 风格学习与管理
插件通过专业的分析框架，为您的每一份创作建立风格档案：
- **14 维度分析**：涵盖词汇、句式、语气、结构、修辞等核心维度，以及受众、文化、叙事等语境维度。
- **多档案切换**：支持为不同场景（如：技术文档、散文、公文）创建并切换独立的风格档案。

### 2. 智能文章生成
- **多样化输出**：支持生成“完整文章”或“写作提纲”（提纲默认提供 2 套以上思路）。
- **去 AI 味处理**：集成 Humanizer-zh 思路，优化机械化表述，使文字更具人情味。
- **文集自动落盘**：生成的文章自动保存至配置的文集目录中。

### 3. 素材库管理
- **自动提取**：在学习风格的同时，自动沉淀金句、段落及引用。
- **高效检索**：支持关键词、标签、素材类型等多维度过滤。
- **风格隔离**：素材按当前激活的风格进行隔离存储，确保引用精准。

---

## 安装配置

您可以在 VS Code 设置中搜索 `写作助手` 或修改 `settings.json` 进行个性化配置：

| 配置项 | 说明 | 默认值 |
| :--- | :--- | :--- |
| `writingAgent.storage.path` | 插件内容（风格、素材、文集）存储根目录 | `~/Downloads/writing-angent` |
| `writingAgent.ai.channel` | AI 通道（可选 `api` 或 `ide-chat`） | `api` |
| `writingAgent.api.provider` | API 提供商（新增支持 `openrouter`） | `iflow` |
| `writingAgent.api.baseUrl` | 统一 API 地址（OpenAI 兼容） | `https://apis.iflow.cn/v1` |
| `writingAgent.api.model` | 主模型（兼容旧配置） | `Qwen/Qwen3-8B` |
| `writingAgent.api.models` | 多模型列表（按顺序自动轮换） | `[]` |
| `writingAgent.autoExtract` | 是否在导入文章时自动提取素材 | `true` |
| `writingAgent.suggestionMinChars` | 自动补全触发所需的最小字符数 (2~20) | `6` |

> 使用 OpenRouter 时，可直接配置默认模型 `stepfun/step-3.5-flash:free`，并在 `writingAgent.api.models` 中追加备用模型实现故障顺序切换。

---

## 使用指南

### 1. 导入样本并学习
1. 使用快捷键 `Cmd/Ctrl + Shift + I` 或点击侧边栏“导入”按钮。
2. 选择多份代表性文章（支持 `.txt`, `.md`, `.docx`）。
3. 查看风格报告并确认素材提取结果。

### 2. 生成新文章
1. 使用快捷键 `Cmd/Ctrl + Shift + G`。
2. 输入主题，并选择输出类型（文章或提纲）。
3. 生成完成后，插件会自动打开生成的 Markdown 文件。

### 3. 素材检索与应用
1. 使用快捷键 `Cmd/Ctrl + Shift + S` 打开搜索页。
2. 输入关键词，找到心仪素材后点击“复制”或“插入”。

---

## API 文档

插件主要通过命令面板与快捷键对外提供能力。

### 核心命令清单
- `writingAgent.generateArticle`: 生成风格化文章。
- `writingAgent.importArticle`: 导入样本数据进行分析。
- `writingAgent.analyzeStyle`: 分析当前打开文档的风格。
- `writingAgent.searchMaterial`: 开启独立素材搜索页面。

### 快捷键对照表
- `Cmd/Ctrl + Shift + G`: 生成文章
- `Cmd/Ctrl + Shift + I`: 导入样本
- `Cmd/Ctrl + Shift + S`: 搜索素材
- `Cmd/Ctrl + Shift + A`: 分析当前文档

---

## 贡献指南

我们非常欢迎开发者参与贡献！

1. **Fork** 本仓库并创建特性分支。
2. **提交** 您的代码更改，并确保符合 TypeScript 开发规范。
3. **开启** 一个 Pull Request 并详细说明您的改进。

### 维护反思
我们在开发过程中始终坚持代码质量：
- ✅ **自动维护**：内置清理工具，定期移除无效空目录。
- ✅ **结构清晰**：遵循 SOLID 原则，核心逻辑按模块（风格、素材、存储）拆分。

---

## 许可证

本项目基于 [MIT 许可证](LICENSE) 开源。

Copyright (c) 2026 daydayup2026
