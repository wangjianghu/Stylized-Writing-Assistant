# VS Code 风格化写作助手

风格化写作助手是一个深度集成到 VS Code 的写作插件：导入样本学习风格，自动沉淀素材，并按既定风格生成文章或提纲。

当前文档对应代码版本：`0.0.3`。

## 核心能力

- 风格学习：基于 14 维度分析（词汇、句式、结构、语气、修辞等）构建风格档案。
- 多风格管理：支持创建、切换、重命名、删除风格；素材与文集按风格隔离。
- 主题写作：支持“完整文章”与“写作提纲（默认 2 套以上方案）”。
- 选区改写：支持“优化表达/扩写细节/帮写重构/自定义指令”。
- 文集管理：生成结果自动写入文稿并进入文集视图。
- API 与 IDE 双通道：可走统一 API 直出，也可走 IDE Chat（Cursor/CodeArts/VS Code Chat）。

## 0.0.3 关键更新

- 新增 OpenRouter 默认模型：`stepfun/step-3.5-flash:free`。
- 支持多提供商顺序调用：按 `writingAgent.api.providerOrder` 从上到下依次尝试。
- 支持多模型顺序调用：同一提供商按模型顺序依次回退。
- API 配置支持拖拽排序：
  - `写作助手: 拖拽调整 API 提供商顺序`
  - `写作助手: 拖拽调整 API 模型顺序`
- 配置 API 时，`iflow/openrouter` 支持“拉取模型列表选择”与“手动输入”两种模式。
- 生成文章后移除强提示弹窗，改为直接写入并打开文集。
- 选区改写改为轻提示浮层，浮层展示改写预览文本，并可直接选择替换/插入/取消。
- “显示写作助手”触发后可直接加载目标视图，无需手动切页面。
- Cursor IDE Chat 场景下，禁止自动新建会话，仅向当前 Chat 输入框注入提示词。
- 存储目录增加权限兜底：默认目录不可写时自动回退扩展全局存储目录。
- 启动性能优化：仓库索引懒加载 + 非核心视图默认折叠 + 补全注册延迟初始化。

## 快速开始

1. 克隆仓库

```bash
git clone https://github.com/wangjianghu/Stylized-Writing-Assistant.git
cd Stylized-Writing-Assistant
```

2. 安装依赖

```bash
npm install
```

3. 编译

```bash
npm run compile
```

4. 调试运行

- 在 VS Code 打开项目，按 `F5` 启动扩展开发宿主。

## 安装与打包

- 本地打包 VSIX：

```bash
npm run package:vsix
```

- 产物路径：`dist/writing-agent-<version>.vsix`

## 配置项

在 VS Code 设置中搜索 `写作助手`，或直接修改 `settings.json`。

| 配置项 | 说明 | 默认值 |
| :--- | :--- | :--- |
| `writingAgent.storage.path` | 插件存储根目录（风格/素材/文集）。留空时优先 `~/Downloads/writing-angent`，不可写时自动回退扩展目录。 | `""` |
| `writingAgent.ai.channel` | AI 通道：`api` 或 `ide-chat` | `api` |
| `writingAgent.api.enabled` | 是否启用统一 API 通道 | `true` |
| `writingAgent.api.provider` | 当前提供商（`iflow/openrouter/kimi/deepseek/minimax/qwen/custom`） | `iflow` |
| `writingAgent.api.providerOrder` | 提供商调用顺序（从上到下依次尝试） | `[]` |
| `writingAgent.api.baseUrl` | 当前提供商 Base URL（OpenAI 兼容） | `https://apis.iflow.cn/v1` |
| `writingAgent.api.model` | 当前主模型 | `Qwen/Qwen3-8B` |
| `writingAgent.api.models` | 当前模型链（顺序回退） | `[]` |
| `writingAgent.api.temperature` | 生成温度 | `0.7` |
| `writingAgent.api.maxTokens` | 最大 token | `10000` |
| `writingAgent.api.timeoutMs` | 超时毫秒 | `60000` |
| `writingAgent.suggestionDelay` | 自动补全延迟（毫秒） | `500` |
| `writingAgent.suggestionMinChars` | 自动补全最小触发字符数（2~20） | `6` |

说明：

- `writingAgent.api.providerConfigs` 由命令自动维护，不建议手工编辑。
- 旧的 `writingAgent.iflow.*` 字段仍兼容读取，建议迁移到 `writingAgent.api.*`。

## 主要命令

- `writingAgent.showAssistantViews`：显示写作助手
- `writingAgent.generateArticle`：根据主题生成文章/提纲
- `writingAgent.rewriteSelectionWithIFlow`：AI 编辑选中文本
- `writingAgent.importArticle`：导入样本数据
- `writingAgent.searchMaterial`：搜索素材
- `writingAgent.analyzeStyle`：分析当前文档风格
- `writingAgent.configureApi`：配置 API
- `writingAgent.testApiConnection`：测试 API 配置
- `writingAgent.reorderApiProviders`：拖拽调整 API 提供商顺序
- `writingAgent.reorderApiModels`：拖拽调整 API 模型顺序
- `writingAgent.applyClipboardToDraft`：将剪贴板内容写入当前文稿

## 快捷键

- `Cmd/Ctrl + Shift + G`：生成文章
- `Cmd/Ctrl + Shift + I`：导入样本
- `Cmd/Ctrl + Shift + S`：搜索素材
- `Cmd/Ctrl + Shift + A`：分析当前文档
- `Cmd/Ctrl + Shift + M`：保存选中文本为素材

## 典型流程

1. 导入样本文章，建立风格档案。
2. 配置 API（可配置多个提供商与多模型顺序）。
3. 输入主题生成文章，结果自动追加到文稿并自动打开文集。
4. 在文稿中选中段落执行 AI 改写，通过浮层预览后决定替换或插入。
5. 使用素材库检索与复用高质量片段。

## 贡献

欢迎提交 Issue 或 PR。提交前建议执行：

```bash
npm run compile
bash final-verify.sh
```

## 许可证

本项目基于 [MIT 许可证](LICENSE) 开源。
