# VS Code 风格化写作助手 (VS Code Stylized Writing Assistant)

一款深度集成于 VS Code 的智能写作辅助插件，通过 14 维度分析自动学习您的写作风格，并辅助生成风格一致的高质量文章。

---

## 目录

- [项目简介](#项目简介)
- [快速开始](#快速开始)
- [功能特性](#功能特性)
- [安装配置](#安装配置)
- [使用指南](#使用指南)
- [技术实现](#技术实现)
- [贡献指南](#贡献指南)
- [许可证](#许可证)

---

## 项目简介

在内容创作中，保持一致的文字风格往往需要长期的积累和刻意的练习。**VS Code 风格化写作助手** 旨在通过 AI 技术简化这一过程。它不仅能分析您过往文章的文字特征，还能在您创作新内容时提供风格契合的建议、素材和提纲，是专业作者、博主及文案工作者的理想伴侣。

---

## 快速开始

导语：只需几步，即可在本地开发环境中体验这款强大的写作工具。

### 环境要求

- **操作系统**: macOS 15.4.1 (推荐) / Windows / Linux
- **Node.js**: 16.x 及以上版本
- **VS Code**: 1.70.0 及以上版本

### 启动步骤

1. **安装依赖**：
   ```bash
   npm install
   ```
2. **编译项目**：
   ```bash
   npm run compile
   ```
3. **运行调试**：
   - 在 VS Code 中打开本项目。
   - 按下 `F5` 键启动扩展开发宿主。
   - 在弹出的新窗口中即可开始测试插件功能。

---

## 功能特性

导语：本插件集成了风格学习、文章生成和素材管理等核心功能，提供全方位的创作支持。

### 1. 深度风格学习
- **14 维度分析**：从词汇、句式、语气、结构、修辞等 14 个专业维度深度解构文章风格。
- **档案管理**：支持创建、切换、重命名及删除多个风格档案，满足不同场景的创作需求。

### 2. 风格化内容生成
- **智能生成**：根据指定主题，自动生成与当前风格档案契合的文章。
- **灵活输出**：支持生成“完整文章”或“写作提纲”，提纲模式默认提供两套以上思路。
- **去 AI 味处理**：集成 Humanizer-zh 算法，弱化机械感，使文字更加自然生动。

### 3. 素材智能管理
- **自动提取**：从历史文章中智能识别并保存金句、段落及名言。
- **高效检索**：支持按关键词、标签及类型进行多维度过滤搜索。
- **一键导出**：支持将当前风格的素材库一键导出为 Markdown 文档。

---

## 安装配置

导语：您可以通过 VS Code 的设置界面对插件进行个性化配置，以获得最佳使用体验。

在 VS Code 设置中搜索 `writingAgent` 即可找到相关配置项：

- `writingAgent.autoExtract`: 是否开启自动提取素材功能（默认：`true`）。
- `writingAgent.suggestionDelay`: 建议弹出的延迟时间（默认：`500ms`）。
- `writingAgent.suggestionMinChars`: 自动补全触发的最小字符数（默认：`6`）。
- `writingAgent.storage.path`: 插件数据的存储路径（默认为 `~/Downloads/writing-angent`）。
- `writingAgent.ai.channel`: AI 服务通道，支持 `api` 或 `ide-chat` 模式。

---

## 使用指南

导语：按照以下流程，您可以快速从零开始构建并使用您的风格库。

### 1. 建立风格档案
- 使用命令 `写作助手: 导入样本数据` (快捷键: `Cmd/Ctrl + Shift + I`)。
- 选择多篇历史文章，插件将自动完成风格学习与素材沉淀。

### 2. 进行风格化创作
- 使用命令 `写作助手: 生成风格化文章` (快捷键: `Cmd/Ctrl + Shift + G`)。
- 输入主题并选择输出类型，生成的文章将自动保存至您的“文集”目录。

### 3. 素材搜索与复用
- 使用命令 `写作助手: 搜索素材` (快捷键: `Cmd/Ctrl + Shift + S`)。
- 在搜索面板中查找所需内容，点击即可一键插入到当前文档。

---

## 技术实现

导语：了解项目背后的架构设计与核心技术选型。

### 项目结构

- [extension.ts](file:///Users/caomoumou/Documents/CodeArts%20%E5%8D%8E%E4%B8%BA%E4%BA%91%E7%A0%81%E9%81%93/%E5%86%99%E4%BD%9CAgent/src/extension.ts): 插件入口及命令注册中心。
- [core/styleEngine](file:///Users/caomoumou/Documents/CodeArts%20%E5%8D%8E%E4%B8%BA%E4%BA%91%E7%A0%81%E9%81%93/%E5%86%99%E4%BD%9CAgent/src/core/styleEngine): 负责文章风格的维度分析与建模。
- [core/materialManager](file:///Users/caomoumou/Documents/CodeArts%20%E5%8D%8E%E4%B8%BA%E4%BA%91%E7%A0%81%E9%81%93/%E5%86%99%E4%BD%9CAgent/src/core/materialManager): 处理素材的提取、持久化与搜索逻辑。
- [providers/](file:///Users/caomoumou/Documents/CodeArts%20%E5%8D%8E%E4%B8%BA%E4%BA%91%E7%A0%81%E9%81%93/%E5%86%99%E4%BD%9CAgent/src/providers): 提供侧边栏视图的数据驱动。

### 核心技术栈
- **语言**: TypeScript
- **框架**: VS Code Extension API
- **NLP**: 内置轻量级分词引擎，支持 Mammoth (DOCX 解析)
- **UI**: VS Code TreeView + Webview API

---

## 贡献指南

导语：本项目是一个开源社区项目，我们非常欢迎并感谢您的每一份贡献。

1. **Fork 本仓库**：点击页面右上角的 Fork 按钮。
2. **创建特性分支**：执行 `git checkout -b feature/AmazingFeature`。
3. **提交更改**：执行 `git commit -m 'Add some AmazingFeature'`。
4. **推送到远程**：执行 `git push origin feature/AmazingFeature`。
5. **开启 Pull Request**：在 GitHub 上提交您的合并请求。

---

## 许可证

本项目基于 [MIT 许可证](LICENSE) 开源。您可以自由地使用、修改和分发。

---

### 维护反思 (Internal Only)

在最近的维护中，我们清理了项目中的冗余空文件夹，并重新规范化了 README 文档的结构。建议未来在添加新功能时，同步更新“功能特性”和“使用指南”章节，保持文档的时效性。
