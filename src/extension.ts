import * as vscode from 'vscode';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { MaterialType, StyleProfile, WritingMaterial, WritingStyle } from './core/types';
import { StyleEngine } from './core/styleEngine';
import { MaterialExtractor } from './core/materialManager/extractor';
import { MaterialRepository } from './core/materialManager/repository';
import { StyleRepository } from './core/styleManager/repository';
import { ArticleRecord, ArticleRepository } from './core/articleManager/repository';
import { generateStyleAlignedDraft } from './core/articleManager/generator';
import { humanizeZh } from './core/articleManager/humanizerZh';
import {
  generateWithOpenAiCompatibleApi,
  listOpenAiCompatibleModels
} from './core/articleManager/openAiCompatibleClient';
import { MaterialSearchPanel } from './panels/materialSearchPanel';
import { MaterialLibraryProvider } from './providers/materialLibraryProvider';
import { StyleProfileProvider } from './providers/styleProfileProvider';
import { QuickEntryProvider } from './providers/quickEntryProvider';
import { ArticleCollectionProvider } from './providers/articleCollectionProvider';

interface SearchCommandArgs {
  presetTypes?: MaterialType[];
  materialId?: string;
  initialQuery?: string;
}

interface SwitchStyleArgs {
  styleId?: string;
}

type ImportTargetDecision =
  | { mode: 'create'; createName: string }
  | { mode: 'update'; profileId: string };

type ApiProvider = 'iflow' | 'openrouter' | 'kimi' | 'deepseek' | 'minimax' | 'qwen' | 'custom';
type TopicOutputMode = 'article' | 'outline';
type AiExecutionChannel = 'api' | 'ide-chat';

interface ProviderPreset {
  provider: ApiProvider;
  displayName: string;
  defaultBaseUrl: string;
  defaultModel: string;
}

interface RuntimeProviderCandidate {
  provider: ApiProvider;
  apiKey: string;
  baseUrl: string;
  model: string;
  models: string[];
}

interface RuntimeApiSettings {
  enabled: boolean;
  provider: ApiProvider;
  providerOrder: ApiProvider[];
  providerCandidates: RuntimeProviderCandidate[];
  apiKey: string;
  baseUrl: string;
  model: string;
  models: string[];
  temperature: number;
  maxTokens: number;
  timeoutMs: number;
}

interface ApiGenerationResult {
  content: string;
  provider: ApiProvider;
  model: string;
  fallbackUsed: boolean;
  fallbackFromProvider?: ApiProvider;
  fallbackFromModel?: string;
}

interface ApiValidationState {
  fingerprint: string;
  valid: boolean;
  message: string;
  checkedAt: string;
}

interface StorageRootResolutionCache {
  preferredPath: string;
  resolvedPath: string;
}

type RewriteApplyMode = 'replace' | 'insertAfter';

interface PendingRewriteApplyDecision {
  editorUri: string;
  selectionRange: vscode.Range;
  resolve: (mode: RewriteApplyMode | undefined) => void;
  timeoutHandle: NodeJS.Timeout;
}

interface DeleteStyleArticlesResult {
  removedRecords: number;
  deletedFiles: number;
  missingFiles: number;
}

interface MergeStyleArticlesResult {
  updatedRecords: number;
  movedFiles: number;
  renamedFiles: number;
  missingFiles: number;
}

interface MaterialSedimentationResult {
  processedArticles: number;
  addedMaterials: number;
}

interface StyleArticleIndexSyncResult {
  indexedRecords: number;
  scannedFiles: number;
}

type SuggestionSourceType = 'material' | 'article';

interface SuggestionCandidate {
  text: string;
  styleId: string;
  styleName: string;
  sourceType: SuggestionSourceType;
  isCurrentStyle: boolean;
  quality: number;
}

interface SuggestionBucketCache {
  items: SuggestionCandidate[];
  builtAt: number;
}

type MammothModule = {
  extractRawText: (input: { buffer: Buffer }) => Promise<{ value?: string }>;
};

const LEGACY_IFLOW_SECRET_KEY = 'writingAgent.iflow.apiKey';
const API_VALIDATION_STATE_KEY = 'writingAgent.api.validation.v1';
const STORAGE_ROOT_RESOLUTION_CACHE_KEY = 'writingAgent.storage.resolvedRoot.v1';
const COLLECTION_RELATIVE_PREFIX = '文集';
const LEGACY_COLLECTION_RELATIVE_PREFIX = '.writing-agent/文集';
const STORAGE_DATA_RELATIVE_PREFIX = 'data';
const STORAGE_DEFAULT_DIR_NAME = 'writing-angent';
const STATUS_MESSAGE_TIMEOUT_MS = 5000;
const SUGGESTION_CACHE_TTL_MS = 20_000;
const COMPLETION_MAX_ITEMS = 6;
const API_VALIDATION_TIMEOUT_MIN_MS = 8000;
const API_VALIDATION_TIMEOUT_MAX_MS = 30000;
const DEFERRED_BOOTSTRAP_DELAY_MS = 1200;
const REWRITE_APPLY_OVERLAY_TIMEOUT_MS = 45_000;
const COMPLETION_DOCUMENT_SELECTOR: vscode.DocumentSelector = [
  { scheme: 'file', language: 'markdown' },
  { scheme: 'untitled', language: 'markdown' },
  { scheme: 'file', language: 'plaintext' },
  { scheme: 'untitled', language: 'plaintext' }
];
const COMPLETION_TRIGGER_CHARACTERS = [' ', '，', '。', '；', '：', ',', '.', ';', ':', '、', '！', '？'];
let showAssistantViewsInFlight: Promise<void> | null = null;
let cachedMammothModule: MammothModule | null | undefined;
let activeExtensionContext: vscode.ExtensionContext | null = null;
let storageRootPathCache: string | null = null;
let rewriteApplyDecorationType: vscode.TextEditorDecorationType | null = null;
let pendingRewriteApplyDecision: PendingRewriteApplyDecision | null = null;
const PROVIDER_PRESETS: Record<ApiProvider, ProviderPreset> = {
  iflow: {
    provider: 'iflow',
    displayName: '心流',
    defaultBaseUrl: 'https://apis.iflow.cn/v1',
    defaultModel: 'Qwen/Qwen3-8B'
  },
  openrouter: {
    provider: 'openrouter',
    displayName: 'OpenRouter',
    defaultBaseUrl: 'https://openrouter.ai/api/v1',
    defaultModel: 'stepfun/step-3.5-flash:free'
  },
  kimi: {
    provider: 'kimi',
    displayName: 'Kimi',
    defaultBaseUrl: 'https://api.moonshot.cn/v1',
    defaultModel: 'moonshot-v1-8k'
  },
  deepseek: {
    provider: 'deepseek',
    displayName: 'DeepSeek',
    defaultBaseUrl: 'https://api.deepseek.com/v1',
    defaultModel: 'deepseek-chat'
  },
  minimax: {
    provider: 'minimax',
    displayName: 'MiniMax',
    defaultBaseUrl: 'https://api.minimax.chat/v1',
    defaultModel: 'abab6.5s-chat'
  },
  qwen: {
    provider: 'qwen',
    displayName: '千问',
    defaultBaseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    defaultModel: 'qwen-plus'
  },
  custom: {
    provider: 'custom',
    displayName: '自定义',
    defaultBaseUrl: 'https://api.openai.com/v1',
    defaultModel: 'gpt-4o-mini'
  }
};

export function activate(context: vscode.ExtensionContext): void {
  console.log('[writing-agent] 插件已激活');
  console.log(`[writing-agent] vscode.version=${vscode.version}`);
  console.log(`[writing-agent] extensionPath=${context.extensionPath}`);

  activeExtensionContext = context;
  storageRootPathCache = null;
  const storageRootPath = resolveStorageRootPath();
  const styleEngine = new StyleEngine();
  const materialExtractor = new MaterialExtractor();
  const materialRepository = new MaterialRepository(context, storageRootPath);
  const styleRepository = new StyleRepository(context, storageRootPath);
  const articleRepository = new ArticleRepository(context, storageRootPath);
  const quickEntryProvider = new QuickEntryProvider(async () => {
    const settings = await resolveApiSettings(context);
    if (!settings || !settings.enabled) {
      return { configured: false, text: '已关闭', iconId: 'circle-slash' };
    }
    if (hasAnyApiProviderKey(settings)) {
      const providerName = PROVIDER_PRESETS[settings.provider].displayName;
      const modelLabel = formatModelChainLabel(settings.models);
      const fingerprint = buildApiValidationFingerprint(settings);
      const validation = loadApiValidationState(context);
      if (!validation || validation.fingerprint !== fingerprint) {
        return { configured: true, text: `已配置·${providerName}/${modelLabel}·未测试`, iconId: 'warning' };
      }
      if (validation.valid) {
        return { configured: true, text: `已配置·${providerName}/${modelLabel}·有效`, iconId: 'check' };
      }
      const brief = summarizeError(validation.message, 36);
      return { configured: true, text: `已配置·${providerName}/${modelLabel}·无效（${brief}）`, iconId: 'error' };
    }
    return { configured: false, text: '未配置', iconId: 'warning' };
  });
  const materialLibraryProvider = new MaterialLibraryProvider(materialRepository, () => {
    const active = styleRepository.getActiveProfile();
    if (!active) {
      return null;
    }
    return { id: active.id, name: active.name };
  });
  const styleProfileProvider = new StyleProfileProvider(styleRepository);
  const articleCollectionProvider = new ArticleCollectionProvider(
    articleRepository,
    () => {
      const active = styleRepository.getActiveProfile();
      if (!active) {
        return null;
      }
      return { id: active.id, name: active.name };
    },
    relativePath => {
      return resolveArticleUriByRelativePath(relativePath);
    }
  );

  const activeProfile = styleRepository.getActiveProfile();
  if (activeProfile) {
    styleEngine.setCurrentStyle(activeProfile.style);
  } else {
    styleEngine.resetStyle();
  }
  const suggestionBucketCache = new Map<string, SuggestionBucketCache>();
  const suggestionTriggerTimers = new Map<string, NodeJS.Timeout>();
  const clearSuggestionCache = (styleId?: string): void => {
    if (styleId) {
      suggestionBucketCache.delete(styleId);
      return;
    }
    suggestionBucketCache.clear();
  };
  const clearSuggestionTimers = (): void => {
    suggestionTriggerTimers.forEach(timer => clearTimeout(timer));
    suggestionTriggerTimers.clear();
  };

  context.subscriptions.push(
    vscode.window.registerTreeDataProvider('quickActions', quickEntryProvider),
    vscode.window.registerTreeDataProvider('quickActionsExplorer', quickEntryProvider),
    vscode.window.registerTreeDataProvider('articleCollection', articleCollectionProvider),
    vscode.window.registerTreeDataProvider('articleCollectionExplorer', articleCollectionProvider),
    vscode.window.registerTreeDataProvider('materialLibrary', materialLibraryProvider),
    vscode.window.registerTreeDataProvider('styleProfile', styleProfileProvider),
    vscode.window.registerTreeDataProvider('materialLibraryExplorer', materialLibraryProvider),
    vscode.window.registerTreeDataProvider('styleProfileExplorer', styleProfileProvider)
  );
  rewriteApplyDecorationType = vscode.window.createTextEditorDecorationType({
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: 'var(--vscode-focusBorder)',
    backgroundColor: 'var(--vscode-editor-selectionBackground)',
    rangeBehavior: vscode.DecorationRangeBehavior.ClosedClosed
  });
  context.subscriptions.push({
    dispose: () => {
      clearRewriteApplyDecorations();
      rewriteApplyDecorationType?.dispose();
      rewriteApplyDecorationType = null;
      resolvePendingRewriteApplyDecision(undefined);
    }
  });
  const deferredLanguageFeaturesTimer = setTimeout(() => {
    context.subscriptions.push(
      vscode.languages.registerCompletionItemProvider(
        COMPLETION_DOCUMENT_SELECTOR,
        {
          async provideCompletionItems(document, position) {
            const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
            const minTriggerChars = getSuggestionMinChars();
            const query = extractSuggestionQuery(linePrefix);
            if (!query || query.length < minTriggerChars) {
              return [];
            }

            const matched = await collectCompletionSuggestions(
              query,
              styleRepository,
              materialRepository,
              articleRepository,
              suggestionBucketCache
            );
            if (matched.length === 0) {
              return [];
            }

            const replaceRange = resolveSuggestionReplaceRange(document, position, query);
            return matched.map((entry, index) => createCompletionItem(entry, index, replaceRange));
          }
        },
        ...COMPLETION_TRIGGER_CHARACTERS
      ),
      vscode.workspace.onDidChangeTextDocument(event => {
        scheduleSuggestionTrigger(event, suggestionTriggerTimers);
      })
    );
  }, DEFERRED_BOOTSTRAP_DELAY_MS);
  context.subscriptions.push({
    dispose: () => {
      clearTimeout(deferredLanguageFeaturesTimer);
      clearSuggestionTimers();
    }
  });
  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration(event => {
      if (
        event.affectsConfiguration('writingAgent.api')
        || event.affectsConfiguration('writingAgent.iflow')
        || event.affectsConfiguration('writingAgent.storage.path')
      ) {
        quickEntryProvider.refresh();
      }
      if (event.affectsConfiguration('writingAgent.storage.path')) {
        storageRootPathCache = null;
        void ensureStorageDirectories().catch(error => {
          const message = error instanceof Error ? error.message : String(error);
          vscode.window.showErrorMessage(`写作助手存储目录初始化失败：${message}`);
        });
        void vscode.window.showInformationMessage(
          `写作助手存储路径已更新为：${resolveStorageRootPath()}。重载窗口后生效。`,
          '立即重载'
        ).then(action => {
          if (action === '立即重载') {
            void vscode.commands.executeCommand('workbench.action.reloadWindow');
          }
        });
      }
      if (
        event.affectsConfiguration('writingAgent.suggestionDelay')
        || event.affectsConfiguration('writingAgent.suggestionMinChars')
      ) {
        clearSuggestionTimers();
      }
    })
  );
  const bootstrapTimer = setTimeout(() => {
    void (async () => {
      try {
        await ensureStorageDirectories();
        await migrateLegacyCollectionDirectory(articleRepository);
        await syncArticleCollectionWithDisk(articleRepository, articleCollectionProvider);
        articleCollectionProvider.refresh();
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.error('[writing-agent] 初始化文集目录失败', message);
        notifyWarn(`初始化文集目录失败：${message}`);
      }
    })();
  }, DEFERRED_BOOTSTRAP_DELAY_MS);
  context.subscriptions.push({
    dispose: () => clearTimeout(bootstrapTimer)
  });

  registerSafeCommand(context, 'writingAgent.generateArticle', async () => {
    const currentStyle = styleEngine.getCurrentStyle();
    if (!currentStyle) {
      const result = await vscode.window.showWarningMessage(
        '尚未学习任何写作风格，是否先导入样本数据？',
        '导入样本数据',
        '切换风格',
        '取消'
      );
      if (result === '导入样本数据') {
        await vscode.commands.executeCommand('writingAgent.importArticle');
      } else if (result === '切换风格') {
        await vscode.commands.executeCommand('writingAgent.switchStyle');
      }
      return;
    }

    const topic = await vscode.window.showInputBox({
      placeHolder: '输入写作主题，例如：介绍 TypeScript 的类型系统',
      prompt: '请输入要生成的文章主题'
    });
    if (!topic) {
      return;
    }

    const outputMode = await pickTopicOutputMode();
    if (!outputMode) {
      return;
    }

    const active = styleRepository.getActiveProfile();
    const executionChannel = resolveAiExecutionChannel();
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Window,
        title: outputMode === 'outline' ? '正在生成写作提纲...' : '正在生成文章...',
        cancellable: false
      },
      async progress => {
        progress.report({ increment: 0, message: '准备文集路径...' });
        const root = resolveStorageRootUri();
        await ensureStorageDirectories();

        const styleName = active?.name || '默认风格';
        const styleId = active?.id || 'default-style';
        const candidateMaterials = active ? materialRepository.search(active.id, { semantic: topic, limit: 8 }) : [];
        const aiPrompt = buildTopicWritingPrompt(
          topic,
          currentStyle,
          styleName,
          candidateMaterials,
          outputMode
        );
        const styleDraft = generateStyleAlignedDraft(topic, currentStyle, candidateMaterials);
        const outlineDraft = generateLocalTopicOutlines(topic, currentStyle, candidateMaterials);
        const apiSettings = await resolveApiSettings(context);
        let generatedBody = styleDraft.body;
        let generatedConclusion = styleDraft.conclusion;
        let generatedOutline = outlineDraft;
        let sourceLabel = '本地风格草稿';
        let aiCommand = 'local:fallback';
        let shouldPromptApiConfig = false;
        let routedToIdeChat = false;

        progress.report({ increment: 30, message: '写入文集草稿...' });
        const draft = buildArticleDraft(topic, styleName, aiPrompt);
        const articleFile = await createArticleFile(root, styleName, topic, draft);

        await articleRepository.upsert({
          topic,
          title: articleFile.title,
          styleId,
          styleName,
          relativePath: articleFile.relativePath
        });
        clearSuggestionCache(styleId);
        articleCollectionProvider.refresh();

        progress.report({ increment: 65, message: '打开文稿...' });
        const document = await vscode.workspace.openTextDocument(articleFile.uri);
        const editor = await vscode.window.showTextDocument(document, vscode.ViewColumn.One);
        focusArticleBody(editor);

        const routePromptToHostChat = async (): Promise<void> => {
          routedToIdeChat = await dispatchPromptToIdeChat(aiPrompt);
          sourceLabel = routedToIdeChat ? '本地草稿 + IDE AI 提示词' : '本地草稿（提示词已复制）';
          aiCommand = 'ide-chat:prompt';
        };

        if (executionChannel === 'ide-chat') {
          progress.report({ increment: 78, message: '填充 IDE Chat 提示词...' });
          await routePromptToHostChat();
        } else if (apiSettings?.enabled && hasAnyApiProviderKey(apiSettings)) {
          try {
            progress.report({ increment: 78, message: outputMode === 'outline' ? '调用 API 生成提纲...' : '调用 API 生成正文...' });
            const apiResult = await generateWithFallbackModel(apiSettings, aiPrompt);
            if (outputMode === 'outline') {
              generatedOutline = ensureOutlineVariants(
                normalizeOutlineOutput(apiResult.content),
                topic,
                currentStyle,
                candidateMaterials
              );
            } else {
              const parsed = parseAiArticle(apiResult.content);
              generatedBody = parsed.body;
              generatedConclusion = parsed.conclusion;
            }
            const providerName = PROVIDER_PRESETS[apiResult.provider].displayName;
            sourceLabel = `${providerName} API · ${apiResult.model}`;
            aiCommand = `${apiResult.provider}:${apiResult.model}`;
            const fallbackNotice = buildApiFallbackNotice(apiResult);
            if (fallbackNotice) {
              notifyWarn(fallbackNotice);
            }
          } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            console.error('[writing-agent] API 生成失败，回退本地草稿', message);
            notifyWarn(`API 生成失败，已回退本地草稿：${message}`);
            const action = await vscode.window.showWarningMessage(
              'API 生成失败，是否改用 IDE AI 对话并自动填充提示词？',
              '改用 IDE AI 对话',
              '继续本地草稿'
            );
            if (action === '改用 IDE AI 对话') {
              await routePromptToHostChat();
            }
          }
        } else if (apiSettings?.enabled) {
          shouldPromptApiConfig = true;
          const action = await vscode.window.showWarningMessage(
            '未配置 API Key，当前将使用本地草稿。是否改用 IDE AI 对话并填充提示词？',
            '改用 IDE AI 对话',
            '配置 API'
          );
          if (action === '改用 IDE AI 对话') {
            await routePromptToHostChat();
          } else if (action === '配置 API') {
            await vscode.commands.executeCommand('writingAgent.configureApi');
          }
        }

        progress.report({ increment: 92, message: outputMode === 'outline' ? '写入提纲...' : '写入正文...' });
        if (outputMode === 'outline') {
          await appendOutlineToEnd(editor, generatedOutline, sourceLabel);
        } else {
          await appendGeneratedDraftToEnd(
            editor,
            generatedBody,
            generatedConclusion,
            currentStyle,
            topic,
            sourceLabel
          );
        }

        await articleRepository.upsert({
          topic,
          title: articleFile.title,
          styleId,
          styleName,
          relativePath: articleFile.relativePath,
          aiCommand
        });
        clearSuggestionCache(styleId);

        progress.report({ increment: 100, message: '完成' });
        if (shouldPromptApiConfig) {
          notifyWarn('API Key 未配置，本次使用本地草稿；可在“写作助手: 配置 API”完成后再试。');
        }
        if (routedToIdeChat) {
          notifyInfo('已填充 IDE AI 提示词，可执行“写作助手: 写入 AI 结果”落盘。');
        }
        await vscode.commands.executeCommand('writingAgent.showArticleCollection');
      }
    );
  });

  registerSafeCommand(context, 'writingAgent.importArticle', async () => {
    const targetDecision = await selectImportTarget(styleRepository);
    if (!targetDecision) {
      return;
    }

    const selections = await vscode.window.showOpenDialog({
      filters: { 文档文件: ['txt', 'md', 'doc', 'docx'] },
      canSelectFiles: true,
      canSelectFolders: true,
      canSelectMany: true
    });
    if (!selections || selections.length === 0) {
      return;
    }
    const files = await collectImportableDocuments(selections);
    if (files.length === 0) {
      notifyWarn('所选路径下未找到可导入文档（.txt/.md/.doc/.docx）。');
      return;
    }

    const previousActiveProfile = styleRepository.getActiveProfile();
    let targetProfile: StyleProfile;

    if (targetDecision.mode === 'create') {
      targetProfile = await styleRepository.createProfile(targetDecision.createName);
    } else {
      const existing = styleRepository.getProfile(targetDecision.profileId);
      if (!existing) {
        vscode.window.showErrorMessage('目标风格不存在，请重新选择。');
        return;
      }
      targetProfile = existing;
    }

    if (targetDecision.mode === 'update') {
      styleEngine.setCurrentStyle(targetProfile.style);
    } else {
      styleEngine.resetStyle();
    }

    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: `正在导入到风格「${targetProfile.name}」...`,
        cancellable: false
      },
      async progress => {
        let learnedCount = 0;
        let materialCount = 0;

        for (let i = 0; i < files.length; i += 1) {
          const file = files[i];
          const fileName = file.path.split('/').pop() || file.path;
          progress.report({
            increment: ((i + 1) / files.length) * 100,
            message: `处理文件 ${i + 1}/${files.length}: ${fileName}`
          });

          try {
            const text = await readImportDocumentText(file);
            if (!text.trim()) {
              continue;
            }

            await styleEngine.learnFromText(text);
            learnedCount += 1;

            const extracted = await materialExtractor.extractMaterials(text, file.path);
            materialCount += extracted.length;
            await materialRepository.addMany(targetProfile.id, extracted);

            const currentStyle = styleEngine.getCurrentStyle();
            if (currentStyle) {
              targetProfile = await styleRepository.updateProfile(targetProfile.id, currentStyle, file.fsPath || file.path);
            }
          } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            console.error(`处理文件失败: ${file.path}`, message);
          }
        }

        if (learnedCount === 0) {
          if (targetDecision.mode === 'create') {
            await styleRepository.deleteProfile(targetProfile.id);
            clearSuggestionCache(targetProfile.id);
          }
          await restoreActiveStyle(styleRepository, styleEngine, previousActiveProfile?.id || null);
          styleProfileProvider.refresh();
          materialLibraryProvider.refresh();
          articleCollectionProvider.refresh();
          notifyWarn('未成功导入任何样本数据，请检查文件内容与格式后重试。');
          return;
        }

        await styleRepository.setActiveProfile(targetProfile.id);
        const currentProfile = styleRepository.getProfile(targetProfile.id);
        if (currentProfile) {
          styleEngine.setCurrentStyle(currentProfile.style);
        }

        materialLibraryProvider.refresh();
        styleProfileProvider.refresh();
        articleCollectionProvider.refresh();
        clearSuggestionCache(targetProfile.id);

        const currentStyle = styleEngine.getCurrentStyle();
        const avgSentenceLength = currentStyle?.sentenceStructure.avgLength.toFixed(1) || '0.0';
        const ratio = ((currentStyle?.vocabulary.uniqueWordRatio || 0) * 100).toFixed(1);
        notifyInfo(`成功导入 ${learnedCount} 篇样本到风格「${targetProfile.name}」，提取 ${materialCount} 条素材。平均句长 ${avgSentenceLength} 字，词汇丰富度 ${ratio}%。`);
      }
    );
  });

  registerSafeCommand(context, 'writingAgent.analyzeStyle', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      notifyWarn('请先打开一个文档');
      return;
    }

    const text = editor.document.getText();
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: '正在分析风格...',
        cancellable: false
      },
      async progress => {
        progress.report({ increment: 0, message: '提取特征...' });
        const style = await styleEngine.analyzeText(text);

        progress.report({ increment: 50, message: '生成报告...' });
        await delay(250);

        const panel = vscode.window.createWebviewPanel(
          'styleAnalysis',
          '风格分析',
          vscode.ViewColumn.Two,
          { enableScripts: true }
        );
        panel.webview.html = getStyleAnalysisWebview(style);
        progress.report({ increment: 100, message: '完成！' });
      }
    );
  });

  registerSafeCommand(context, 'writingAgent.saveMaterial', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      notifyWarn('请先打开一个文档');
      return;
    }

    const selectedText = editor.document.getText(editor.selection);
    if (!selectedText.trim()) {
      notifyWarn('请先选择要保存的文本');
      return;
    }

    const activeProfile = styleRepository.getActiveProfile();
    if (!activeProfile) {
      const action = await vscode.window.showWarningMessage(
        '当前没有激活风格，无法保存素材。请先导入样本数据或创建风格。',
        '新建风格',
        '导入样本数据',
        '切换风格'
      );
      if (action === '新建风格') {
        await vscode.commands.executeCommand('writingAgent.createStyle');
      } else if (action === '导入样本数据') {
        await vscode.commands.executeCommand('writingAgent.importArticle');
      } else if (action === '切换风格') {
        await vscode.commands.executeCommand('writingAgent.switchStyle');
      }
      return;
    }

    const material = await materialRepository.addOneFromSelection(
      activeProfile.id,
      selectedText,
      editor.document.uri.toString()
    );
    materialLibraryProvider.refresh();
    clearSuggestionCache(activeProfile.id);
    notifyInfo(`已保存素材：${material.name}`);
  });

  registerSafeCommand(context, 'writingAgent.searchMaterial', async (...commandArgs: unknown[]) => {
    const args = commandArgs[0] as SearchCommandArgs | undefined;
    const activeProfile = styleRepository.getActiveProfile();
    if (!activeProfile) {
      const action = await vscode.window.showWarningMessage(
        '当前没有激活风格，请先导入样本数据或创建风格后再搜索素材。',
        '新建风格',
        '导入样本数据',
        '切换风格'
      );
      if (action === '新建风格') {
        await vscode.commands.executeCommand('writingAgent.createStyle');
      } else if (action === '导入样本数据') {
        await vscode.commands.executeCommand('writingAgent.importArticle');
      } else if (action === '切换风格') {
        await vscode.commands.executeCommand('writingAgent.switchStyle');
      }
      return;
    }

    MaterialSearchPanel.show(context, materialRepository, {
      presetTypes: args?.presetTypes,
      materialId: args?.materialId,
      initialQuery: args?.initialQuery,
      styleId: activeProfile.id,
      styleName: activeProfile.name
    });
  });

  registerSafeCommand(context, 'writingAgent.exportMaterials', async () => {
    const activeProfile = styleRepository.getActiveProfile();
    if (!activeProfile) {
      const action = await vscode.window.showWarningMessage(
        '当前没有激活风格，请先导入样本数据或创建风格后再导出素材。',
        '新建风格',
        '导入样本数据',
        '切换风格'
      );
      if (action === '新建风格') {
        await vscode.commands.executeCommand('writingAgent.createStyle');
      } else if (action === '导入样本数据') {
        await vscode.commands.executeCommand('writingAgent.importArticle');
      } else if (action === '切换风格') {
        await vscode.commands.executeCommand('writingAgent.switchStyle');
      }
      return;
    }

    const materials = materialRepository.search(activeProfile.id, { limit: 2000 });
    if (materials.length === 0) {
      notifyWarn('当前风格素材为空，暂无可导出内容。');
      return;
    }

    const fileName = `${sanitizePathSegment(activeProfile.name)}-素材导出-${formatTimestamp(new Date())}.md`;
    const defaultUri = vscode.Uri.joinPath(resolveStorageRootUri(), 'exports', fileName);
    const targetUri = await vscode.window.showSaveDialog({
      saveLabel: '导出素材',
      defaultUri,
      filters: { Markdown: ['md'] }
    });
    if (!targetUri) {
      return;
    }

    await vscode.workspace.fs.createDirectory(getParentUri(targetUri));
    const markdown = buildMaterialExportMarkdown(activeProfile.name, materials);
    await vscode.workspace.fs.writeFile(targetUri, Buffer.from(markdown, 'utf-8'));

    const action = await vscode.window.showInformationMessage(
      `已导出 ${materials.length} 条素材到 ${targetUri.fsPath}`,
      '打开文件'
    );
    if (action === '打开文件') {
      await vscode.commands.executeCommand('vscode.open', targetUri);
    }
  });

  registerSafeCommand(context, 'writingAgent.createStyle', async () => {
    const mode = await vscode.window.showQuickPick(
      [
        { label: '创建空白风格', value: 'empty' },
        { label: '复制当前风格', value: 'copy' }
      ],
      { placeHolder: '选择创建方式' }
    );
    if (!mode) {
      return;
    }

    const inputName = await vscode.window.showInputBox({
      prompt: '请输入新风格名称',
      value: '新风格'
    });
    if (!inputName) {
      return;
    }

    const currentStyle = styleEngine.getCurrentStyle();
    const baseStyle = mode.value === 'copy' ? currentStyle || undefined : undefined;
    const profile = await styleRepository.createProfile(inputName, baseStyle);
    await styleRepository.setActiveProfile(profile.id);
    styleEngine.setCurrentStyle(profile.style);
    styleProfileProvider.refresh();
    materialLibraryProvider.refresh();
    articleCollectionProvider.refresh();
    clearSuggestionCache();
    notifyInfo(`已创建风格：${profile.name}`);
  });

  registerSafeCommand(context, 'writingAgent.switchStyle', async (...commandArgs: unknown[]) => {
    const args = commandArgs[0] as SwitchStyleArgs | undefined;
    let target = args?.styleId ? styleRepository.getProfile(args.styleId) : undefined;

    if (!target) {
      const profiles = styleRepository.listProfiles();
      if (profiles.length === 0) {
        notifyWarn('暂无可切换风格，请先导入样本数据或创建风格。');
        return;
      }

      const picked = await vscode.window.showQuickPick(
        profiles.map(profile => ({
          label: profile.name,
          description: `${profile.articleCount} 篇样本`,
          detail: `最近更新：${profile.updatedAt.toLocaleString()}`,
          profileId: profile.id
        })),
        { placeHolder: '选择要切换的风格' }
      );
      if (!picked) {
        return;
      }
      target = styleRepository.getProfile(picked.profileId);
    }

    if (!target) {
      vscode.window.showErrorMessage('目标风格不存在。');
      return;
    }

    await styleRepository.setActiveProfile(target.id);
    styleEngine.setCurrentStyle(target.style);
    styleProfileProvider.refresh();
    materialLibraryProvider.refresh();
    articleCollectionProvider.refresh();
    clearSuggestionCache();
    notifyInfo(`已切换到风格：${target.name}`);
  });

  registerSafeCommand(context, 'writingAgent.renameStyle', async () => {
    const target = await pickStyleProfile(styleRepository, '选择要重命名的风格');
    if (!target) {
      return;
    }

    const newName = await vscode.window.showInputBox({
      prompt: '输入新的风格名称',
      value: target.name
    });
    if (!newName) {
      return;
    }

    await styleRepository.renameProfile(target.id, newName);
    styleProfileProvider.refresh();
    materialLibraryProvider.refresh();
    articleCollectionProvider.refresh();
    clearSuggestionCache(target.id);
    notifyInfo('风格名称已更新。');
  });

  registerSafeCommand(context, 'writingAgent.deleteStyle', async () => {
    const target = await pickStyleProfile(styleRepository, '选择要删除的风格');
    if (!target) {
      return;
    }

    const allProfiles = styleRepository.listProfiles();
    const candidates = allProfiles.filter(profile => profile.id !== target.id);
    const strategy = await vscode.window.showQuickPick(
      candidates.length > 0
        ? [
          { label: '删除风格及其素材和文章', value: 'delete-all' as const },
          { label: '删除风格并将素材和文章合并到其他风格', value: 'merge-all' as const }
        ]
        : [{ label: '删除风格及其素材和文章', value: 'delete-all' as const }],
      { placeHolder: '请选择删除风格时的数据处理方式' }
    );
    if (!strategy) {
      return;
    }

    let mergeTarget: StyleProfile | undefined;
    if (strategy.value === 'merge-all') {
      const picked = await vscode.window.showQuickPick(
        candidates.map(profile => ({
          label: profile.name,
          description: `${profile.articleCount} 篇样本`,
          profileId: profile.id
        })),
        { placeHolder: `选择用于接收「${target.name}」素材和文章的目标风格` }
      );
      if (!picked) {
        return;
      }
      mergeTarget = styleRepository.getProfile(picked.profileId);
      if (!mergeTarget) {
        vscode.window.showErrorMessage('目标风格不存在，无法执行合并。');
        return;
      }
    }

    const confirmation = await vscode.window.showWarningMessage(
      strategy.value === 'merge-all' && mergeTarget
        ? `确认删除风格「${target.name}」并将素材和文章合并到「${mergeTarget.name}」吗？`
        : `确认删除风格「${target.name}」及其素材和文章吗？此操作不可撤销。`,
      { modal: true },
      '删除'
    );
    if (confirmation !== '删除') {
      return;
    }

    const activeProfileIdBeforeDelete = styleRepository.getActiveProfileId();
    const storageRoot = resolveStorageRootUri();

    let materialResult: { moved: number; merged: number } | undefined;
    let articleDeleteResult: DeleteStyleArticlesResult | undefined;
    let articleMergeResult: MergeStyleArticlesResult | undefined;
    if (strategy.value === 'merge-all' && mergeTarget) {
      materialResult = await materialRepository.moveStyleMaterials(target.id, mergeTarget.id);
      articleMergeResult = await moveArticlesToStyle(target, mergeTarget, articleRepository, storageRoot);
    } else {
      await materialRepository.deleteStyleMaterials(target.id);
      articleDeleteResult = await removeArticlesByStyle(target.id, articleRepository, storageRoot);
    }

    await styleRepository.deleteProfile(target.id);

    if (strategy.value === 'merge-all' && mergeTarget && activeProfileIdBeforeDelete === target.id) {
      await styleRepository.setActiveProfile(mergeTarget.id);
    }

    const activeAfterDelete = styleRepository.getActiveProfile();
    if (activeAfterDelete) {
      styleEngine.setCurrentStyle(activeAfterDelete.style);
    } else {
      styleEngine.resetStyle();
    }
    styleProfileProvider.refresh();
    materialLibraryProvider.refresh();
    articleCollectionProvider.refresh();
    clearSuggestionCache();
    if (materialResult) {
      notifyInfo(
        `已删除风格：${target.name}。已合并到「${mergeTarget?.name}」：素材新增 ${materialResult.moved} 条、去重 ${materialResult.merged} 条；文章迁移 ${articleMergeResult?.movedFiles || 0} 篇，索引更新 ${articleMergeResult?.updatedRecords || 0} 篇。`
      );
    } else {
      notifyInfo(
        `已删除风格：${target.name}（素材已删除，文章索引清理 ${articleDeleteResult?.removedRecords || 0} 篇，文件删除 ${articleDeleteResult?.deletedFiles || 0} 篇）。`
      );
    }
  });

  registerSafeCommand(context, 'writingAgent.refreshMaterialLibrary', async () => {
    const active = styleRepository.getActiveProfile();
    if (!active) {
      materialLibraryProvider.refresh();
      notifyWarn('素材库已刷新。当前未激活风格，未执行文集沉淀。');
      return;
    }

    const indexSync = await syncStyleArticleIndexFromDisk(active, articleRepository);
    const sedimentation = await sedimentCurrentStyleArticlesToMaterials(
      active,
      articleRepository,
      materialExtractor,
      materialRepository
    );
    materialLibraryProvider.refresh();
    clearSuggestionCache(active.id);
    notifyInfo(
      `素材库已刷新：扫描 ${indexSync.scannedFiles} 篇文集文件，同步 ${indexSync.indexedRecords} 条索引，已处理 ${sedimentation.processedArticles} 篇并沉淀新增 ${sedimentation.addedMaterials} 条素材。`
    );
  });

  registerSafeCommand(context, 'writingAgent.refreshArticleCollection', async () => {
    await syncArticleCollectionWithDisk(articleRepository, articleCollectionProvider, { notifyOnCleanup: true });
    articleCollectionProvider.refresh();
    clearSuggestionCache();
  });

  registerSafeCommand(context, 'writingAgent.deleteArticle', async (arg?: unknown) => {
    const root = resolveStorageRootUri();

    const target = await resolveArticleRecordTarget(arg, articleRepository);
    if (!target) {
      return;
    }

    const confirmation = await vscode.window.showWarningMessage(
      `确认删除文集文章《${target.title}》吗？此操作会删除磁盘文件且不可撤销。`,
      { modal: true },
      '删除'
    );
    if (confirmation !== '删除') {
      return;
    }

    const articleUri = await resolveArticleUriForDelete(target.relativePath, root);
    let fileDeleted = false;
    try {
      if (await uriExists(articleUri)) {
        await vscode.workspace.fs.delete(articleUri, { recursive: false, useTrash: false });
        fileDeleted = true;
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error('[writing-agent] 删除文集文件失败', message);
      throw new Error(`文集文件删除失败：${message}`);
    }

    await articleRepository.deleteById(target.id);
    await syncArticleCollectionWithDisk(articleRepository, articleCollectionProvider);
    articleCollectionProvider.refresh();
    clearSuggestionCache(target.styleId);

    if (fileDeleted) {
      notifyInfo(`已删除文集文章：${target.title}`);
    } else {
      notifyWarn(`文集文件不存在，已清理索引：${target.title}`);
    }
  });

  registerSafeCommand(context, 'writingAgent.showArticleCollection', async () => {
    await syncArticleCollectionWithDisk(articleRepository, articleCollectionProvider);
    await revealWritingAgentView({
      primaryFocusCommand: 'workbench.actions.treeView.articleCollection.focus',
      explorerFocusCommand: 'workbench.actions.treeView.articleCollectionExplorer.focus',
      failMessage: '无法打开文集视图'
    });
  });

  registerSafeCommand(context, 'writingAgent.openMaterialLibrary', async () => {
    await revealWritingAgentView({
      primaryFocusCommand: 'workbench.actions.treeView.materialLibrary.focus',
      explorerFocusCommand: 'workbench.actions.treeView.materialLibraryExplorer.focus',
      failMessage: '无法打开素材库视图'
    });
  });

  registerSafeCommand(context, 'writingAgent.openStyleProfile', async () => {
    await revealWritingAgentView({
      primaryFocusCommand: 'workbench.actions.treeView.styleProfile.focus',
      explorerFocusCommand: 'workbench.actions.treeView.styleProfileExplorer.focus',
      failMessage: '无法打开风格配置视图'
    });
    await vscode.commands.executeCommand('writingAgent.switchStyle');
  });

  registerSafeCommand(context, 'writingAgent.useHostAI', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      notifyWarn('请先打开一篇文稿，再调用 IDE AI 续写。');
      return;
    }
    const active = styleRepository.getActiveProfile();
    const style = active?.style || styleEngine.getCurrentStyle();
    if (!style) {
      notifyWarn('当前没有可用风格，无法生成 AI 提示词。');
      return;
    }

    const title = editor.document.fileName.split('/').pop() || '当前文稿';
    const topic = title.replace(/\.md$/i, '');
    const prompt = buildTopicWritingPrompt(
      topic,
      style,
      active?.name || '默认风格',
      active ? materialRepository.search(active.id, { semantic: editor.document.getText().slice(0, 120), limit: 5 }) : [],
      'article'
    );
    const dispatched = await dispatchPromptToIdeChat(prompt);
    if (dispatched) {
      notifyInfo('已将续写提示词填充到 IDE Chat 输入框。');
    } else {
      notifyWarn('未能自动填充到 IDE Chat 输入框，提示词已复制到剪贴板。');
    }
  });

  registerSafeCommand(context, 'writingAgent.rewriteSelectionWithIFlow', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      notifyWarn('请先打开文档并选中文本。');
      return;
    }

    if (editor.selection.isEmpty) {
      notifyWarn('请先选中要优化的文本。');
      return;
    }

    const selectionRange = new vscode.Range(editor.selection.start, editor.selection.end);
    const selectedRaw = editor.document.getText(selectionRange);
    const selectedText = selectedRaw.trim();
    if (!selectedText) {
      notifyWarn('选中文本为空，无法执行改写。');
      return;
    }

    const mode = await vscode.window.showQuickPick(
      [
        { label: '优化表达', instruction: '优化这段文字的表达，使其更清晰自然。' },
        { label: '扩写细节', instruction: '在保持原意基础上扩写这段文字，补充细节与论据。' },
        { label: '帮写重构', instruction: '重构这段文字的结构与节奏，增强可读性。' },
        { label: '自定义指令', instruction: '' }
      ],
      { placeHolder: '选择改写模式' }
    );
    if (!mode) {
      return;
    }

    const customInstruction = await vscode.window.showInputBox({
      prompt: '输入本次改写指令（生成后可选“替换”或“插入到选区后”）',
      value: mode.instruction,
      placeHolder: '例如：保留原意，语气更克制，加入一个具体例子'
    });
    if (!customInstruction) {
      return;
    }

    const active = styleRepository.getActiveProfile();
    const style = active?.style || styleEngine.getCurrentStyle() || undefined;
    const prompt = buildSelectionRewritePrompt(
      selectedText,
      customInstruction.trim(),
      style,
      active?.name
    );
    const executionChannel = resolveAiExecutionChannel();
    if (executionChannel === 'ide-chat') {
      const dispatched = await dispatchPromptToIdeChat(prompt);
      if (dispatched) {
        notifyInfo('已将改写提示词填充到 IDE Chat，请将结果复制后替换选区。');
      } else {
        notifyWarn('无法自动填充 IDE Chat，改写提示词已复制到剪贴板。');
      }
      return;
    }

    const settings = await resolveApiSettings(context);
    if (!settings?.enabled) {
      notifyWarn('API 已关闭，请先开启后再试。');
      return;
    }
    if (!hasAnyApiProviderKey(settings)) {
      const action = await vscode.window.showWarningMessage(
        '未配置 API Key。是否改用 IDE AI 对话，或先配置 API？',
        '改用 IDE AI 对话',
        '配置 API'
      );
      if (action === '改用 IDE AI 对话') {
        const dispatched = await dispatchPromptToIdeChat(prompt);
        if (dispatched) {
          notifyInfo('已将改写提示词填充到 IDE Chat，请将结果复制后替换选区。');
        } else {
          notifyWarn('无法自动填充 IDE Chat，改写提示词已复制到剪贴板。');
        }
        return;
      }
      if (action === '配置 API') {
        await vscode.commands.executeCommand('writingAgent.configureApi');
      }
      return;
    }

    try {
      let rewritten = '';
      let fallbackInfo: ApiGenerationResult | undefined;
      let usedStrictRetry = false;
      const recordFallback = (result: ApiGenerationResult): void => {
        if (result.fallbackUsed) {
          fallbackInfo = result;
        }
      };
      await vscode.window.withProgress(
        {
          location: vscode.ProgressLocation.Window,
          title: '正在改写选中文本...',
          cancellable: false
        },
        async progress => {
          progress.report({ increment: 20, message: '调用 API...' });
          const result = await generateWithFallbackModel(settings, prompt);
          rewritten = normalizeRewriteOutput(result.content);
          recordFallback(result);

          if (!rewritten) {
            throw new Error('API 返回为空，未生成可替换内容');
          }

          if (isRewriteEffectivelyUnchanged(selectedRaw, rewritten)) {
            progress.report({ increment: 45, message: '结果与原文接近，自动强制改写重试...' });
            const strictPrompt = buildStrictRewritePrompt(prompt, selectedRaw);
            const strictResult = await generateWithFallbackModel(settings, strictPrompt);
            rewritten = normalizeRewriteOutput(strictResult.content);
            usedStrictRetry = true;
            recordFallback(strictResult);
          }

          if (!rewritten) {
            throw new Error('API 返回为空，未生成可替换内容');
          }
          if (isRewriteEffectivelyUnchanged(selectedRaw, rewritten)) {
            throw new Error('改写结果与原文几乎一致，未产生有效修改。请更换指令后重试。');
          }

          progress.report({ increment: 100, message: '完成' });
        }
      );

      const outputMode = await pickRewriteApplyModeWithPreview(editor, selectionRange, rewritten);
      if (!outputMode) {
        return;
      }

      const applied = await editor.edit(editBuilder => {
        if (outputMode === 'replace') {
          editBuilder.replace(selectionRange, rewritten);
          return;
        }
        const selectedRaw = editor.document.getText(selectionRange);
        const insertPrefix = selectedRaw.endsWith('\n') ? '\n' : '\n\n';
        editBuilder.insert(selectionRange.end, `${insertPrefix}${rewritten}`);
      });
      if (!applied) {
        throw new Error('改写结果写入失败，请重试。');
      }

      if (fallbackInfo) {
        const fallbackNotice = buildApiFallbackNotice(fallbackInfo);
        if (fallbackNotice) {
          notifyWarn(fallbackNotice);
        }
      }
      if (usedStrictRetry) {
        notifyInfo('检测到首轮改写与原文过于接近，已自动执行二次强制改写。');
      }

      if (outputMode === 'replace') {
        notifyInfo('已按指令替换选中文本。');
      } else {
        notifyInfo('已按指令插入到选区后。');
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      vscode.window.showErrorMessage(`改写失败：${message}`);
      const action = await vscode.window.showWarningMessage(
        'API 改写失败，是否改用 IDE AI 对话并填充提示词？',
        '改用 IDE AI 对话'
      );
      if (action === '改用 IDE AI 对话') {
        const dispatched = await dispatchPromptToIdeChat(prompt);
        if (dispatched) {
          notifyInfo('已将改写提示词填充到 IDE Chat。');
        } else {
          notifyWarn('无法自动填充 IDE Chat，改写提示词已复制到剪贴板。');
        }
      }
    }
  });

  const configureApiHandler = async () => {
    const config = vscode.workspace.getConfiguration('writingAgent');
    const target = getConfigurationTarget();
    const current = await resolveApiSettings(context);
    const currentProvider = current?.provider || 'iflow';
    const provider = await pickApiProvider(currentProvider);
    if (!provider) {
      return;
    }
    const preset = getProviderPreset(provider);

    const action = await vscode.window.showQuickPick(
      [
        { label: '设置/更新 API Key', value: 'setKey' },
        { label: '清除 API Key', value: 'clearKey' },
        { label: '仅修改模型与地址', value: 'configOnly' },
        { label: '调整 API 提供商顺序', value: 'providerOrder' }
      ],
      { placeHolder: `配置 API（${preset.displayName}）` }
    );
    if (!action) {
      return;
    }

    if (action.value === 'providerOrder') {
      const configuredOrder = getConfiguredProviderOrder(config, currentProvider);
      const reordered = await editProviderOrderByDrag('API 提供商顺序', configuredOrder);
      if (!reordered || reordered.length === 0) {
        return;
      }
      await updateProviderOrderConfig(reordered, target);
      notifyInfo('API 提供商顺序已更新，调用将按从上到下依次尝试。');
      quickEntryProvider.refresh();
      return;
    }

    if (action.value === 'setKey') {
      const input = await vscode.window.showInputBox({
        title: `${preset.displayName} API Key`,
        prompt: '输入 API Key（将存入 VS Code SecretStorage）',
        password: true,
        ignoreFocusOut: true
      });
      if (input === undefined) {
        return;
      }
      const key = input.trim();
      if (!key) {
        notifyWarn('API Key 为空，已取消更新。');
        return;
      }
      await context.secrets.store(getProviderSecretKey(provider), key);
      if (provider === 'iflow') {
        await context.secrets.store(LEGACY_IFLOW_SECRET_KEY, key);
      }
    } else if (action.value === 'clearKey') {
      await context.secrets.delete(getProviderSecretKey(provider));
      if (provider === 'iflow') {
        await context.secrets.delete(LEGACY_IFLOW_SECRET_KEY);
      }
    }

    const baseDefault = provider === currentProvider && current?.baseUrl
      ? current.baseUrl
      : preset.defaultBaseUrl;
    const baseUrlInput = await vscode.window.showInputBox({
      prompt: `${preset.displayName} Base URL（OpenAI 兼容）`,
      value: baseDefault
    });
    if (baseUrlInput === undefined) {
      return;
    }

    const defaultModels = provider === currentProvider && current?.models?.length
      ? current.models
      : [preset.defaultModel];
    const apiKey = await resolveProviderApiKey(context, provider);
    let selectedModels: string[] | undefined;
    if (provider === 'iflow' || provider === 'openrouter') {
      const mode = await vscode.window.showQuickPick(
        [
          { label: '从模型列表选择（推荐）', value: 'pick' },
          { label: '手动输入模型名称', value: 'manual' }
        ],
        { placeHolder: `配置 ${preset.displayName} 模型` }
      );
      if (!mode) {
        return;
      }
      if (mode.value === 'pick') {
        selectedModels = await pickProviderModelCandidates(
          preset.displayName,
          apiKey,
          baseUrlInput.trim() || preset.defaultBaseUrl,
          defaultModels,
          15000
        );
        if (!selectedModels) {
          const fallbackAction = await vscode.window.showWarningMessage(
            '未选择模型，是否改为手动输入？',
            '手动输入'
          );
          if (fallbackAction !== '手动输入') {
            return;
          }
          selectedModels = await promptManualModelCandidates(preset.displayName, defaultModels);
        }
      } else {
        selectedModels = await promptManualModelCandidates(preset.displayName, defaultModels);
      }
    } else {
      selectedModels = await promptManualModelCandidates(preset.displayName, defaultModels);
    }
    if (!selectedModels) {
      return;
    }
    const pickedModels = selectedModels.length > 0 ? selectedModels : [preset.defaultModel];
    let normalizedModels = pickedModels;
    if (pickedModels.length > 1) {
      const reorder = await editModelOrderByDrag(`${preset.displayName} 模型顺序`, pickedModels);
      if (reorder) {
        normalizedModels = reorder;
      }
    }

    await config.update('api.enabled', true, target);
    await config.update('api.provider', provider, target);
    const normalizedBaseUrl = baseUrlInput.trim() || preset.defaultBaseUrl;
    await config.update('api.baseUrl', normalizedBaseUrl, target);
    await config.update('api.model', normalizedModels[0], target);
    await config.update('api.models', normalizedModels, target);
    const rawProviderConfigs = config.get<unknown>('api.providerConfigs', {});
    const providerConfigs = rawProviderConfigs && typeof rawProviderConfigs === 'object'
      ? { ...(rawProviderConfigs as Record<string, unknown>) }
      : {};
    providerConfigs[provider] = {
      baseUrl: normalizedBaseUrl,
      model: normalizedModels[0],
      models: normalizedModels
    };
    await config.update('api.providerConfigs', providerConfigs, target);

    const latest = await resolveApiSettings(context);
    if (latest?.enabled && hasAnyApiProviderKey(latest)) {
      const validation = await validateApiConnection(context, latest, true);
      if (validation.valid) {
        notifyInfo(`API 配置已更新并验证通过：${preset.displayName}`);
      } else {
        notifyWarn(`API 配置已保存，但测试失败：${validation.message}`);
      }
    } else {
      clearApiValidationState(context);
      notifyInfo(`API 配置已更新：${preset.displayName}（未测试）`);
    }

    quickEntryProvider.refresh();
  };
  registerSafeCommand(context, 'writingAgent.configureApi', configureApiHandler);
  registerSafeCommand(context, 'writingAgent.configureIFlowApi', configureApiHandler);
  registerSafeCommand(context, 'writingAgent.reorderApiProviders', async () => {
    const config = vscode.workspace.getConfiguration('writingAgent');
    const settings = await resolveApiSettings(context);
    const currentProvider = settings?.provider || 'iflow';
    const configuredOrder = getConfiguredProviderOrder(config, currentProvider);
    const reordered = await editProviderOrderByDrag('API 提供商顺序', configuredOrder);
    if (!reordered || reordered.length === 0) {
      return;
    }
    await updateProviderOrderConfig(reordered, getConfigurationTarget());
    notifyInfo('API 提供商顺序已更新，调用将按从上到下依次尝试。');
    quickEntryProvider.refresh();
  });
  registerSafeCommand(context, 'writingAgent.reorderApiModels', async () => {
    const settings = await resolveApiSettings(context);
    if (!settings?.enabled) {
      notifyWarn('API 已关闭，请先开启后再调整模型顺序。');
      return;
    }
    const models = normalizeModelCandidates(settings.models.length > 0 ? settings.models : [settings.model]);
    if (models.length <= 1) {
      notifyWarn('当前模型数量不足 2 个，无需调整顺序。');
      return;
    }

    const title = `${getProviderPreset(settings.provider).displayName} 模型顺序`;
    const reordered = await editModelOrderByDrag(title, models);
    if (!reordered || reordered.length === 0) {
      return;
    }

    const normalized = normalizeModelCandidates(reordered);
    const config = vscode.workspace.getConfiguration('writingAgent');
    const target = getConfigurationTarget();
    await config.update('api.model', normalized[0], target);
    await config.update('api.models', normalized, target);
    notifyInfo('模型顺序已更新，后续调用将按从上到下顺序依次尝试。');
    quickEntryProvider.refresh();
  });
  registerSafeCommand(context, 'writingAgent.configureSuggestionMinChars', async () => {
    const config = vscode.workspace.getConfiguration('writingAgent');
    const target = getConfigurationTarget();
    const current = getSuggestionMinChars();
    const input = await vscode.window.showInputBox({
      title: '自动补全最小触发字符数',
      prompt: '请输入 2~20 之间的整数',
      value: String(current),
      validateInput: value => {
        const trimmed = value.trim();
        if (!/^\d+$/.test(trimmed)) {
          return '请输入整数';
        }
        const parsed = Number(trimmed);
        if (parsed < 2 || parsed > 20) {
          return '取值范围为 2~20';
        }
        return undefined;
      }
    });
    if (input === undefined) {
      return;
    }
    const parsed = Number(input.trim());
    const nextValue = Math.min(20, Math.max(2, Math.round(parsed)));
    await config.update('suggestionMinChars', nextValue, target);
    notifyInfo(`自动补全最小触发字符数已设置为 ${nextValue}。`);
  });

  registerSafeCommand(context, 'writingAgent.testApiConnection', async () => {
    const settings = await resolveApiSettings(context);
    if (!settings?.enabled) {
      notifyWarn('API 已关闭，请先开启后再测试。');
      return;
    }
    if (!hasAnyApiProviderKey(settings)) {
      const action = await vscode.window.showWarningMessage('未配置 API Key，请先配置后再测试。', '配置 API');
      if (action === '配置 API') {
        await vscode.commands.executeCommand('writingAgent.configureApi');
      }
      return;
    }

    const result = await validateApiConnection(context, settings, true);
    quickEntryProvider.refresh();
    if (result.valid) {
      notifyInfo(`API 测试通过：${result.message}`);
      return;
    }
    notifyWarn(`API 测试失败：${result.message}`);
  });

  registerSafeCommand(context, 'writingAgent.applyClipboardToDraft', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      notifyWarn('请先打开文集中的文稿文件。');
      return;
    }
    const style = styleRepository.getActiveProfile()?.style || styleEngine.getCurrentStyle();
    if (!style) {
      notifyWarn('当前没有可用风格，无法进行风格对齐。');
      return;
    }

    const clipboard = (await vscode.env.clipboard.readText()).trim();
    if (!looksLikeArticleContent(clipboard)) {
      notifyWarn('剪贴板内容不像完整文章，请先在聊天窗口复制完整生成结果。');
      return;
    }

    const topic = editor.document.fileName.split('/').pop()?.replace(/\.md$/i, '') || '当前主题';
    await applyAiResultToDraft(editor, clipboard, style, topic);
    const activeProfileId = styleRepository.getActiveProfileId();
    if (activeProfileId) {
      clearSuggestionCache(activeProfileId);
    }
    notifyInfo('已将剪贴板内容写入当前文稿。');
  });

  registerSafeCommand(context, 'writingAgent.showAssistantViews', async () => {
    if (showAssistantViewsInFlight) {
      await showAssistantViewsInFlight;
      return;
    }
    showAssistantViewsInFlight = (async () => {
      try {
        await vscode.commands.executeCommand('workbench.action.openView', 'quickActions');
        return;
      } catch {
        // 继续降级到容器+聚焦命令
      }
      try {
        await vscode.commands.executeCommand('workbench.view.extension.writing-agent');
        try {
          await vscode.commands.executeCommand('workbench.actions.treeView.quickActions.focus');
        } catch {
          // 某些兼容层不支持 focus 命令，容器已打开即可。
        }
        return;
      } catch {
        // 继续降级
      }
      try {
        await vscode.commands.executeCommand('workbench.action.openView', 'quickActionsExplorer');
        try {
          await vscode.commands.executeCommand('workbench.view.explorer');
        } catch {
          // 打开 Explorer 失败时不阻断，view 已可见即可。
        }
        return;
      } catch {
        // 继续降级
      }
      try {
        await vscode.commands.executeCommand('workbench.view.explorer');
        try {
          await vscode.commands.executeCommand('workbench.actions.treeView.quickActionsExplorer.focus');
        } catch {
          // Explorer 已打开时不再调用 openView，避免触发 view quick-open。
        }
        return;
      } catch (primaryError) {
        const primaryMessage = primaryError instanceof Error ? primaryError.message : String(primaryError);
        try {
          await vscode.commands.executeCommand('workbench.action.openView', 'quickActions');
          return;
        } catch (fallbackError) {
          const fallbackMessage = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
          vscode.window.showErrorMessage(`无法打开写作助手视图。ActivityBar: ${primaryMessage}；Explorer: ${fallbackMessage}`);
        }
      } finally {
        showAssistantViewsInFlight = null;
      }
    })();

    await showAssistantViewsInFlight;
  });
  registerSafeCommand(context, 'writingAgent.rewriteApplyReplace', async () => {
    resolvePendingRewriteApplyDecision('replace');
  });
  registerSafeCommand(context, 'writingAgent.rewriteApplyInsertAfter', async () => {
    resolvePendingRewriteApplyDecision('insertAfter');
  });
  registerSafeCommand(context, 'writingAgent.rewriteApplyCancel', async () => {
    resolvePendingRewriteApplyDecision(undefined);
  });

  registerSafeCommand(context, 'writingAgent.refreshStyleProfile', async () => {
    const active = styleRepository.getActiveProfile();
    if (!active) {
      styleProfileProvider.refresh();
      notifyWarn('当前未激活风格，无法刷新风格特征。');
      return;
    }

    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: `正在刷新风格「${active.name}」...`,
        cancellable: false
      },
      async progress => {
        progress.report({ increment: 5, message: '同步文集索引...' });
        const indexSync = await syncStyleArticleIndexFromDisk(active, articleRepository);

        progress.report({ increment: 15, message: '读取当前风格素材...' });
        const materialTexts = materialRepository.getAll(active.id)
          .map(material => normalizeStyleAnalysisText(material.content))
          .filter(text => isMeaningfulSuggestionText(text, 8));

        progress.report({ increment: 30, message: '读取当前风格文集...' });
        const articleTexts: string[] = [];
        const styleArticles = articleRepository.listByStyle(active.id);
        for (const record of styleArticles) {
          const articleUri = resolveArticleUriByRelativePath(record.relativePath);
          if (!articleUri) {
            continue;
          }
          const articleText = normalizeStyleAnalysisText(await readTextFromUriSafely(articleUri));
          if (!isMeaningfulSuggestionText(articleText, 24)) {
            continue;
          }
          articleTexts.push(articleText);
        }

        const corpusBlocks = [...articleTexts, ...materialTexts];
        if (corpusBlocks.length === 0) {
          throw new Error('当前风格缺少可分析内容，请先导入样本或生成文稿。');
        }

        progress.report({ increment: 35, message: '重算风格维度...' });
        const corpus = corpusBlocks.join('\n\n');
        const refreshedStyle = await styleEngine.analyzeText(corpus);

        progress.report({ increment: 10, message: '写入风格档案...' });
        const updated = await styleRepository.refreshProfileStyle(
          active.id,
          refreshedStyle,
          `refresh:materials(${materialTexts.length})+articles(${articleTexts.length})`
        );
        styleEngine.setCurrentStyle(updated.style);
        styleProfileProvider.refresh();
        materialLibraryProvider.refresh();
        articleCollectionProvider.refresh();
        clearSuggestionCache(active.id);

        progress.report({ increment: 5, message: '完成' });
        notifyInfo(
          `风格已刷新：文集 ${articleTexts.length} 篇、素材 ${materialTexts.length} 条（索引补齐 ${indexSync.indexedRecords} 条）。`
        );
      }
    );
  });
}

function registerSafeCommand(
  context: vscode.ExtensionContext,
  commandId: string,
  handler: (...args: unknown[]) => Promise<void>
): void {
  const disposable = vscode.commands.registerCommand(commandId, async (...args: unknown[]) => {
    try {
      await handler(...args);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      vscode.window.showErrorMessage(`${commandId} 执行失败: ${message}`);
    }
  });
  context.subscriptions.push(disposable);
}

function notifyInfo(message: string): void {
  vscode.window.setStatusBarMessage(`$(info) 写作助手 ${message}`, STATUS_MESSAGE_TIMEOUT_MS);
}

function notifyWarn(message: string): void {
  vscode.window.setStatusBarMessage(`$(warning) 写作助手 ${message}`, STATUS_MESSAGE_TIMEOUT_MS);
}

async function pickTopicOutputMode(): Promise<TopicOutputMode | undefined> {
  const picked = await vscode.window.showQuickPick(
    [
      {
        label: '完整文章（默认）',
        description: '生成并追加正文',
        mode: 'article' as TopicOutputMode
      },
      {
        label: '写作提纲',
        description: '默认输出 2 套以上不同思路',
        mode: 'outline' as TopicOutputMode
      }
    ],
    { placeHolder: '根据主题写作：选择输出类型' }
  );
  return picked?.mode;
}

function resolveAiExecutionChannel(): AiExecutionChannel {
  const config = vscode.workspace.getConfiguration('writingAgent');
  const rawValue = (config.get<string>('ai.channel', 'api') || '').trim().toLowerCase();
  return rawValue === 'ide-chat' ? 'ide-chat' : 'api';
}

function isCursorHost(commandSet: Set<string>): boolean {
  const appName = (vscode.env.appName || '').toLowerCase();
  if (appName.includes('cursor')) {
    return true;
  }
  return (
    commandSet.has('cursor.chat.insert')
    || commandSet.has('cursor.chat.open')
    || commandSet.has('cursor.chat.focusInput')
  );
}

async function executeKnownCommand(
  commandSet: Set<string>,
  command: string,
  args?: unknown
): Promise<boolean> {
  if (!commandSet.has(command)) {
    return false;
  }
  try {
    if (args === undefined) {
      await vscode.commands.executeCommand(command);
    } else {
      await vscode.commands.executeCommand(command, args);
    }
    return true;
  } catch {
    return false;
  }
}

async function dispatchPromptToCursorChat(
  normalizedPrompt: string,
  commandSet: Set<string>
): Promise<boolean> {
  const focusAttempts: Array<{ command: string; args?: unknown }> = [
    { command: 'cursor.chat.focusInput' },
    { command: 'cursor.chat.open' },
    { command: 'workbench.panel.chat.view.focus' },
    { command: 'workbench.action.chat.open' }
  ];
  let focused = false;
  for (const attempt of focusAttempts) {
    if (await executeKnownCommand(commandSet, attempt.command, attempt.args)) {
      focused = true;
      break;
    }
  }

  // Cursor 场景只允许向当前 Chat 输入框注入内容，禁止新建会话。
  const insertAttempts: Array<{ command: string; args: unknown }> = [
    { command: 'cursor.chat.insert', args: { text: normalizedPrompt } },
    { command: 'cursor.chat.insert', args: normalizedPrompt },
    { command: 'workbench.action.chat.insert', args: { text: normalizedPrompt } },
    { command: 'workbench.action.chat.insert', args: normalizedPrompt }
  ];
  for (const attempt of insertAttempts) {
    if (await executeKnownCommand(commandSet, attempt.command, attempt.args)) {
      return true;
    }
  }
  return focused;
}

async function dispatchPromptToIdeChat(prompt: string): Promise<boolean> {
  const normalized = prompt.trim();
  if (!normalized) {
    return false;
  }

  const commandSet = new Set(await vscode.commands.getCommands(true));
  if (isCursorHost(commandSet)) {
    const inserted = await dispatchPromptToCursorChat(normalized, commandSet);
    if (inserted) {
      return true;
    }
    await vscode.env.clipboard.writeText(normalized);
    return false;
  }

  const promptAttempts: Array<{ command: string; args: unknown }> = [
    { command: 'workbench.action.chat.open', args: { query: normalized, isPartialQuery: true } },
    { command: 'workbench.action.chat.open', args: { query: normalized } },
    { command: 'workbench.action.chat.open', args: { inputValue: normalized } },
    { command: 'workbench.action.chat.open', args: { initialQuery: normalized } },
    { command: 'workbench.action.quickchat.open', args: { query: normalized } },
    { command: 'workbench.action.quickchat.open', args: normalized },
    { command: 'vscode.chat.open', args: { query: normalized } },
    { command: 'vscode.chat.open', args: { inputValue: normalized } },
    { command: 'cursor.chat.open', args: { query: normalized } },
    { command: 'codearts.agent.chat.open', args: { query: normalized } },
    { command: 'codearts.agent.chat.open', args: { inputValue: normalized } },
    { command: 'codearts.chat.open', args: { query: normalized } },
    { command: 'codearts.chat.open', args: { inputValue: normalized } }
  ];
  let opened = false;

  for (const attempt of promptAttempts) {
    if (await executeKnownCommand(commandSet, attempt.command, attempt.args)) {
      opened = true;
    }
  }

  const focusAttempts = [
    'workbench.action.chat.focusInput',
    'workbench.panel.chat.view.focus',
    'workbench.action.chat.open',
    'workbench.action.quickchat.open',
    'cursor.chat.focusInput',
    'cursor.chat.open',
    'codearts.agent.chat.focusInput',
    'codearts.agent.chat.open',
    'codearts.chat.focusInput',
    'codearts.chat.open'
  ];
  let focused = false;
  for (const command of focusAttempts) {
    if (await executeKnownCommand(commandSet, command)) {
      focused = true;
      break;
    }
  }

  if (opened || focused) {
    const insertAttempts: Array<{ command: string; args: unknown }> = [
      { command: 'workbench.action.chat.insert', args: { text: normalized } },
      { command: 'workbench.action.chat.insert', args: normalized },
      { command: 'cursor.chat.insert', args: { text: normalized } },
      { command: 'cursor.chat.insert', args: normalized },
      { command: 'codearts.agent.chat.insert', args: { text: normalized } },
      { command: 'codearts.agent.chat.insert', args: normalized },
      { command: 'codearts.chat.insert', args: { text: normalized } },
      { command: 'codearts.chat.insert', args: normalized },
      { command: 'workbench.action.quickchat.open', args: { query: normalized } },
      { command: 'workbench.action.quickchat.open', args: normalized }
    ];
    for (const attempt of insertAttempts) {
      if (await executeKnownCommand(commandSet, attempt.command, attempt.args)) {
        return true;
      }
    }
  }

  await vscode.env.clipboard.writeText(normalized);
  return false;
}

async function migrateLegacyCollectionDirectory(articleRepository: ArticleRepository): Promise<void> {
  const roots = collectArticleSearchRoots();
  let movedRoots = 0;
  for (const root of roots) {
    const legacyDir = joinUriByRelativePath(root, LEGACY_COLLECTION_RELATIVE_PREFIX);
    const currentDir = joinUriByRelativePath(root, COLLECTION_RELATIVE_PREFIX);
    if (!(await uriExists(legacyDir))) {
      continue;
    }

    if (!(await uriExists(currentDir))) {
      await vscode.workspace.fs.rename(legacyDir, currentDir, { overwrite: false });
      movedRoots += 1;
      continue;
    }

    const legacyFiles = await listFilesRecursive(legacyDir);
    for (const source of legacyFiles) {
      const relative = source.path.slice(legacyDir.path.length).replace(/^\/+/, '');
      const target = vscode.Uri.joinPath(currentDir, ...relative.split('/'));
      if (await uriExists(target)) {
        continue;
      }
      await vscode.workspace.fs.createDirectory(getParentUri(target));
      await vscode.workspace.fs.rename(source, target, { overwrite: false });
    }
  }

  const rewritten = await articleRepository.replacePathPrefix(
    `${LEGACY_COLLECTION_RELATIVE_PREFIX}/`,
    `${COLLECTION_RELATIVE_PREFIX}/`
  );
  if (movedRoots > 0 || rewritten > 0) {
    notifyInfo(`已兼容迁移旧版文集目录（目录迁移 ${movedRoots} 个，索引更新 ${rewritten} 条）。`);
  }
}

async function syncArticleCollectionWithDisk(
  articleRepository: ArticleRepository,
  articleCollectionProvider: ArticleCollectionProvider,
  options?: { notifyOnCleanup?: boolean }
): Promise<void> {
  const roots = collectArticleSearchRoots();
  const result = await articleRepository.pruneMissingFiles(roots);
  if (result.removed > 0) {
    articleCollectionProvider.refresh();
    if (options?.notifyOnCleanup) {
      notifyInfo(`文集已清理 ${result.removed} 条失效索引。`);
    }
  }
}

async function sedimentCurrentStyleArticlesToMaterials(
  activeProfile: StyleProfile,
  articleRepository: ArticleRepository,
  materialExtractor: MaterialExtractor,
  materialRepository: MaterialRepository
): Promise<MaterialSedimentationResult> {
  const directRecords = articleRepository.listByStyle(activeProfile.id);
  const fallbackRecords = directRecords.length > 0
    ? []
    : articleRepository.listAll().filter(record => record.styleName === activeProfile.name);
  const records = [...directRecords, ...fallbackRecords];
  const seenPaths = new Set<string>();
  let processedArticles = 0;
  let addedMaterials = 0;

  for (const record of records) {
    const normalizedPath = normalizeRelativePath(record.relativePath);
    if (!normalizedPath || seenPaths.has(normalizedPath)) {
      continue;
    }
    seenPaths.add(normalizedPath);

    const articleUri = resolveArticleUriByRelativePath(record.relativePath);
    if (!articleUri) {
      continue;
    }

    const articleText = await readTextFromUriSafely(articleUri);
    const extractableText = normalizeArticleForMaterialExtraction(articleText);
    if (!isMeaningfulSuggestionText(extractableText, 24)) {
      continue;
    }

    const extracted = await materialExtractor.extractMaterials(
      extractableText,
      articleUri.fsPath || articleUri.toString()
    );
    if (extracted.length === 0) {
      continue;
    }
    processedArticles += 1;
    addedMaterials += await materialRepository.addMany(activeProfile.id, extracted);
  }

  return { processedArticles, addedMaterials };
}

async function syncStyleArticleIndexFromDisk(
  activeProfile: StyleProfile,
  articleRepository: ArticleRepository
): Promise<StyleArticleIndexSyncResult> {
  const styleDirName = sanitizePathSegment(activeProfile.name);
  const roots = collectArticleSearchRoots();
  const allRecords = articleRepository.listAll();
  const recordByPath = new Map<string, ArticleRecord>();
  for (const record of allRecords) {
    const normalized = normalizeRelativePath(toCanonicalCollectionRelativePath(record.relativePath));
    if (!normalized || recordByPath.has(normalized)) {
      continue;
    }
    recordByPath.set(normalized, record);
  }

  let indexedRecords = 0;
  let scannedFiles = 0;
  const visited = new Set<string>();

  for (const root of roots) {
    for (const collectionPrefix of getCollectionRelativePrefixes()) {
      const styleDir = joinUriByRelativePath(root, `${collectionPrefix}/${styleDirName}`);
      const markdownFiles = await listMarkdownFilesRecursive(styleDir);
      scannedFiles += markdownFiles.length;

      for (const file of markdownFiles) {
        const relativePath = file.path.replace(root.path, '').replace(/^\/+/, '');
        const normalized = normalizeRelativePath(toCanonicalCollectionRelativePath(relativePath));
        if (!normalized || visited.has(normalized)) {
          continue;
        }
        visited.add(normalized);

        const existing = recordByPath.get(normalized);
        const title = path.basename(file.path).replace(/\.md$/i, '');
        const topic = deriveTopicFromArticleTitle(title);
        if (existing && existing.styleId === activeProfile.id && existing.styleName === activeProfile.name) {
          continue;
        }

        await articleRepository.upsert({
          id: existing?.id,
          topic,
          title,
          styleId: activeProfile.id,
          styleName: activeProfile.name,
          relativePath,
          aiCommand: existing?.aiCommand
        });
        indexedRecords += 1;
      }
    }
  }

  return { indexedRecords, scannedFiles };
}

function scheduleSuggestionTrigger(
  event: vscode.TextDocumentChangeEvent,
  timers: Map<string, NodeJS.Timeout>
): void {
  const editor = vscode.window.activeTextEditor;
  if (!editor || editor.document.uri.toString() !== event.document.uri.toString()) {
    return;
  }
  if (!isSuggestableDocument(event.document)) {
    return;
  }
  if (event.contentChanges.length === 0) {
    return;
  }

  const latestChange = event.contentChanges[event.contentChanges.length - 1];
  const changedText = latestChange.text || '';
  if (!/[\p{L}\p{N}\u4E00-\u9FFF]/u.test(changedText) && !/[，。！？,.!?；;:：、]/u.test(changedText)) {
    return;
  }

  const key = event.document.uri.toString();
  const existing = timers.get(key);
  if (existing) {
    clearTimeout(existing);
  }

  const timer = setTimeout(() => {
    timers.delete(key);
    const activeEditor = vscode.window.activeTextEditor;
    if (!activeEditor || activeEditor.document.uri.toString() !== key) {
      return;
    }
    void vscode.commands.executeCommand('editor.action.triggerSuggest');
  }, getSuggestionDelayMs());
  timers.set(key, timer);
}

async function collectCompletionSuggestions(
  query: string,
  styleRepository: StyleRepository,
  materialRepository: MaterialRepository,
  articleRepository: ArticleRepository,
  cache: Map<string, SuggestionBucketCache>
): Promise<SuggestionCandidate[]> {
  const profiles = styleRepository.listProfiles();
  if (profiles.length === 0) {
    return [];
  }

  const activeStyleId = styleRepository.getActiveProfileId();
  const orderedProfiles = activeStyleId
    ? [
      ...profiles.filter(profile => profile.id === activeStyleId),
      ...profiles.filter(profile => profile.id !== activeStyleId)
    ]
    : profiles;

  const now = Date.now();
  const pool: SuggestionCandidate[] = [];
  for (const profile of orderedProfiles) {
    const cached = cache.get(profile.id);
    if (cached && now - cached.builtAt <= SUGGESTION_CACHE_TTL_MS) {
      pool.push(...cached.items);
      continue;
    }

    const bucket = await buildSuggestionBucketForStyle(
      profile,
      profile.id === activeStyleId,
      materialRepository,
      articleRepository
    );
    cache.set(profile.id, { items: bucket, builtAt: now });
    pool.push(...bucket);
  }

  const dedup = new Set<string>();
  const scored: Array<{ candidate: SuggestionCandidate; score: number }> = [];
  for (const candidate of pool) {
    const key = `${candidate.styleId}|${candidate.sourceType}|${candidate.text}`;
    if (dedup.has(key)) {
      continue;
    }
    dedup.add(key);

    const score = scoreSuggestionCandidate(query, candidate);
    if (score <= 0) {
      continue;
    }
    scored.push({ candidate, score });
  }

  return scored
    .sort((left, right) => {
      if (right.score !== left.score) {
        return right.score - left.score;
      }
      if (right.candidate.quality !== left.candidate.quality) {
        return right.candidate.quality - left.candidate.quality;
      }
      return left.candidate.text.length - right.candidate.text.length;
    })
    .slice(0, COMPLETION_MAX_ITEMS)
    .map(entry => entry.candidate);
}

function resolveSuggestionReplaceRange(
  document: vscode.TextDocument,
  position: vscode.Position,
  query: string
): vscode.Range | undefined {
  if (!query) {
    return undefined;
  }
  const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
  const tokenStart = findSuggestionTokenStart(linePrefix);
  if (tokenStart >= position.character) {
    return undefined;
  }
  return new vscode.Range(new vscode.Position(position.line, tokenStart), position);
}

function createCompletionItem(
  candidate: SuggestionCandidate,
  index: number,
  replaceRange?: vscode.Range
): vscode.CompletionItem {
  const item = new vscode.CompletionItem(
    summarizeSuggestionLabel(candidate.text),
    vscode.CompletionItemKind.Snippet
  );
  item.insertText = candidate.text;
  item.filterText = candidate.text;
  item.sortText = `${index}`.padStart(2, '0');
  item.preselect = index === 0;
  item.detail = `${candidate.isCurrentStyle ? '当前风格' : '其他风格'} · ${candidate.styleName} · ${candidate.sourceType === 'material' ? '素材库' : '文集'}`;
  if (replaceRange) {
    item.range = replaceRange;
  }
  return item;
}

function extractSuggestionQuery(linePrefix: string): string {
  const tail = linePrefix.slice(-80).replace(/\t/g, ' ');
  const matched = tail.match(/[\p{L}\p{N}\u4E00-\u9FFF_-]{2,}$/u);
  if (!matched) {
    return '';
  }
  return matched[0].trim().slice(-24);
}

function findSuggestionTokenStart(linePrefix: string): number {
  let index = linePrefix.length;
  while (index > 0) {
    const prevChar = linePrefix[index - 1];
    if (/[\s,，。！？!?；;:：、()[\]{}<>《》【】“”"‘’'`~|\\/]/u.test(prevChar)) {
      break;
    }
    index -= 1;
  }
  return index;
}

function getSuggestionDelayMs(): number {
  const rawDelay = vscode.workspace.getConfiguration('writingAgent').get<number>('suggestionDelay', 500);
  if (!Number.isFinite(rawDelay)) {
    return 500;
  }
  return Math.min(2500, Math.max(80, Math.round(rawDelay)));
}

function getSuggestionMinChars(): number {
  const rawMinChars = vscode.workspace.getConfiguration('writingAgent').get<number>('suggestionMinChars', 6);
  if (!Number.isFinite(rawMinChars)) {
    return 6;
  }
  return Math.min(20, Math.max(2, Math.round(rawMinChars)));
}

function isSuggestableDocument(document: vscode.TextDocument): boolean {
  if (document.isUntitled) {
    return document.languageId === 'markdown' || document.languageId === 'plaintext';
  }
  return document.uri.scheme === 'file' && (document.languageId === 'markdown' || document.languageId === 'plaintext');
}

async function buildSuggestionBucketForStyle(
  profile: StyleProfile,
  isCurrentStyle: boolean,
  materialRepository: MaterialRepository,
  articleRepository: ArticleRepository
): Promise<SuggestionCandidate[]> {
  const candidates: SuggestionCandidate[] = [];
  const seen = new Set<string>();
  const addCandidate = (text: string, sourceType: SuggestionSourceType, quality: number): void => {
    const normalized = normalizeSuggestionText(text);
    if (!isMeaningfulSuggestionText(normalized, 8)) {
      return;
    }
    const dedupKey = `${sourceType}|${normalized}`;
    if (seen.has(dedupKey)) {
      return;
    }
    seen.add(dedupKey);
    candidates.push({
      text: normalized,
      styleId: profile.id,
      styleName: profile.name,
      sourceType,
      isCurrentStyle,
      quality
    });
  };

  const styleMaterials = materialRepository.search(profile.id, { limit: 120 });
  for (const material of styleMaterials) {
    addCandidate(material.content, 'material', material.metadata.quality || material.features.quality || 60);
  }

  const styleArticles = articleRepository.listByStyle(profile.id).slice(0, 30);
  for (const article of styleArticles) {
    addCandidate(article.topic, 'article', 62);
    addCandidate(article.title, 'article', 62);
  }

  const styleFavoriteWords = filterMeaningfulAnalysisWords(profile.style.vocabulary?.favoriteWords || []).slice(0, 24);
  for (const word of styleFavoriteWords) {
    addCandidate(word, 'article', 58);
  }

  return candidates.slice(0, 260);
}

function normalizeSuggestionText(text: string): string {
  return (text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\t/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/[ ]{2,}/g, ' ')
    .trim();
}

function isMeaningfulSuggestionText(text: string, minCoreChars: number): boolean {
  if (!text) {
    return false;
  }
  const compact = text.replace(/\s+/g, '');
  if (!compact) {
    return false;
  }
  if (/^[-_*`#>~|]+$/.test(compact)) {
    return false;
  }
  const core = compact.replace(/[^\p{L}\p{N}\u4E00-\u9FFF]/gu, '');
  return core.length >= minCoreChars;
}

function scoreSuggestionCandidate(query: string, candidate: SuggestionCandidate): number {
  const normalizedQuery = normalizeQueryToken(query);
  if (!normalizedQuery) {
    return 0;
  }
  const normalizedText = normalizeQueryToken(candidate.text);
  if (!normalizedText) {
    return 0;
  }

  let score = 0;
  if (normalizedText.startsWith(normalizedQuery)) {
    score += 120;
  } else if (normalizedText.includes(normalizedQuery)) {
    score += 80;
  }

  const tokens = normalizedQuery.split(/\s+/).filter(Boolean);
  if (tokens.length > 1) {
    let hit = 0;
    for (const token of tokens) {
      if (normalizedText.includes(token)) {
        hit += 1;
      }
    }
    if (hit === tokens.length) {
      score += 28;
    } else if (hit > 0) {
      score += 12;
    }
  }

  score += candidate.isCurrentStyle ? 1000 : 300;
  score += candidate.sourceType === 'material' ? 20 : 12;
  score += Math.min(candidate.quality, 100) / 10;
  return score;
}

function normalizeQueryToken(value: string): string {
  return (value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function summarizeSuggestionLabel(text: string): string {
  if (text.length <= 36) {
    return text;
  }
  return `${text.slice(0, 36)}…`;
}

function normalizeArticleForMaterialExtraction(text: string): string {
  return (text || '')
    .replace(/\r\n/g, '\n')
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/^[-*_]{3,}\s*$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function normalizeStyleAnalysisText(text: string): string {
  return (text || '')
    .replace(/\r\n/g, '\n')
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/^[-*_]{3,}\s*$/gm, '')
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/^>\s+/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

async function readTextFromUriSafely(uri: vscode.Uri): Promise<string> {
  try {
    const buffer = await vscode.workspace.fs.readFile(uri);
    return Buffer.from(buffer).toString('utf8');
  } catch {
    return '';
  }
}

async function collectImportableDocuments(selections: vscode.Uri[]): Promise<vscode.Uri[]> {
  const results: vscode.Uri[] = [];
  const seen = new Set<string>();

  for (const selection of selections) {
    let stat: vscode.FileStat;
    try {
      stat = await vscode.workspace.fs.stat(selection);
    } catch {
      continue;
    }

    if (stat.type === vscode.FileType.File) {
      if (!isSupportedImportDocument(selection)) {
        continue;
      }
      const key = (selection.fsPath || selection.path).toLowerCase();
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      results.push(selection);
      continue;
    }

    if (stat.type === vscode.FileType.Directory) {
      const docs = await listSupportedDocsRecursive(selection);
      for (const doc of docs) {
        const key = (doc.fsPath || doc.path).toLowerCase();
        if (seen.has(key)) {
          continue;
        }
        seen.add(key);
        results.push(doc);
      }
    }
  }

  return results;
}

async function listSupportedDocsRecursive(dir: vscode.Uri): Promise<vscode.Uri[]> {
  const files: vscode.Uri[] = [];
  let entries: [string, vscode.FileType][];
  try {
    entries = await vscode.workspace.fs.readDirectory(dir);
  } catch {
    return files;
  }

  for (const [name, fileType] of entries) {
    const target = vscode.Uri.joinPath(dir, name);
    if (fileType === vscode.FileType.File) {
      if (isSupportedImportDocument(target)) {
        files.push(target);
      }
      continue;
    }
    if (fileType === vscode.FileType.Directory) {
      files.push(...await listSupportedDocsRecursive(target));
    }
  }
  return files;
}

async function listFilesRecursive(dir: vscode.Uri): Promise<vscode.Uri[]> {
  const files: vscode.Uri[] = [];
  let entries: [string, vscode.FileType][];
  try {
    entries = await vscode.workspace.fs.readDirectory(dir);
  } catch {
    return files;
  }

  for (const [name, fileType] of entries) {
    const target = vscode.Uri.joinPath(dir, name);
    if (fileType === vscode.FileType.File) {
      files.push(target);
      continue;
    }
    if (fileType === vscode.FileType.Directory) {
      files.push(...await listFilesRecursive(target));
    }
  }
  return files;
}

async function listMarkdownFilesRecursive(dir: vscode.Uri): Promise<vscode.Uri[]> {
  const files: vscode.Uri[] = [];
  let entries: [string, vscode.FileType][];
  try {
    entries = await vscode.workspace.fs.readDirectory(dir);
  } catch {
    return files;
  }

  for (const [name, fileType] of entries) {
    const target = vscode.Uri.joinPath(dir, name);
    if (fileType === vscode.FileType.File) {
      if (/\.md$/i.test(name)) {
        files.push(target);
      }
      continue;
    }
    if (fileType === vscode.FileType.Directory) {
      files.push(...await listMarkdownFilesRecursive(target));
    }
  }

  return files;
}

function isSupportedImportDocument(uri: vscode.Uri): boolean {
  const lowerPath = uri.path.toLowerCase();
  return (
    lowerPath.endsWith('.txt')
    || lowerPath.endsWith('.md')
    || lowerPath.endsWith('.doc')
    || lowerPath.endsWith('.docx')
  );
}

async function readImportDocumentText(uri: vscode.Uri): Promise<string> {
  const lowerPath = uri.path.toLowerCase();
  const content = await vscode.workspace.fs.readFile(uri);
  const buffer = Buffer.from(content);

  if (lowerPath.endsWith('.docx')) {
    return normalizeImportedText(await extractTextFromDocx(buffer));
  }
  if (lowerPath.endsWith('.doc')) {
    return normalizeImportedText(extractTextFromLegacyDoc(buffer));
  }
  return normalizeImportedText(decodeTextBuffer(buffer));
}

function decodeTextBuffer(buffer: Buffer): string {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return buffer.slice(3).toString('utf8');
  }
  if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
    return buffer.slice(2).toString('utf16le');
  }
  return buffer.toString('utf8');
}

async function extractTextFromDocx(buffer: Buffer): Promise<string> {
  const mammothModule = loadMammothModule();
  if (!mammothModule) {
    throw new Error('DOCX 解析组件不可用，请重新安装插件后重试');
  }

  try {
    const result = await mammothModule.extractRawText({ buffer });
    return result.value || '';
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`DOCX 解析失败：${message}`);
  }
}

function loadMammothModule(): MammothModule | null {
  if (cachedMammothModule !== undefined) {
    return cachedMammothModule;
  }

  try {
    // 延迟加载可选依赖，避免宿主环境缺包时阻断插件激活。
    // eslint-disable-next-line @typescript-eslint/no-var-requires, global-require
    cachedMammothModule = require('mammoth') as MammothModule;
    return cachedMammothModule;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[writing-agent] mammoth 加载失败：${message}`);
    cachedMammothModule = null;
    return null;
  }
}

function extractTextFromLegacyDoc(buffer: Buffer): string {
  const unicodeText = buffer
    .toString('utf16le')
    .replace(/\u0000/g, '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ' ');
  const latinText = buffer
    .toString('latin1')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ' ')
    .replace(/\u0000/g, ' ');

  const unicodeReadable = extractReadableTextChunks(unicodeText);
  const latinReadable = extractReadableTextChunks(latinText);
  const merged = unicodeReadable.length >= latinReadable.length ? unicodeReadable : latinReadable;
  return merged;
}

function extractReadableTextChunks(input: string): string {
  const normalized = input.replace(/\r\n/g, '\n');
  const chunks = normalized.match(/[\u4e00-\u9fa5A-Za-z0-9，。！？；：、“”‘’（）()《》【】…,.!?;:\-\n ]{4,}/g) || [];
  return chunks
    .map(chunk => chunk.replace(/[ \t]{2,}/g, ' ').trim())
    .filter(Boolean)
    .join('\n');
}

function normalizeImportedText(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\u0000/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

async function pickStyleProfile(
  styleRepository: StyleRepository,
  placeHolder: string
): Promise<StyleProfile | undefined> {
  const profiles = styleRepository.listProfiles();
  if (profiles.length === 0) {
    notifyWarn('暂无可用风格。');
    return undefined;
  }

  const picked = await vscode.window.showQuickPick(
    profiles.map(profile => ({
      label: profile.name,
      description: `${profile.articleCount} 篇样本`,
      detail: `最近更新：${profile.updatedAt.toLocaleString()}`,
      profileId: profile.id
    })),
    { placeHolder }
  );
  if (!picked) {
    return undefined;
  }
  return styleRepository.getProfile(picked.profileId);
}

async function selectImportTarget(
  styleRepository: StyleRepository
): Promise<ImportTargetDecision | undefined> {
  const profiles = styleRepository.listProfiles();
  const options = profiles.length > 0
    ? [
      { label: '新建风格', value: 'create' as const },
      { label: '更新现有风格', value: 'update' as const }
    ]
    : [{ label: '新建风格', value: 'create' as const }];

  const strategy = await vscode.window.showQuickPick(options, {
    placeHolder: '导入策略：新建风格或更新现有风格'
  });
  if (!strategy) {
    return undefined;
  }

  if (strategy.value === 'create') {
    const name = await vscode.window.showInputBox({
      prompt: '请输入新风格名称',
      value: `风格 ${new Date().toLocaleDateString()}`
    });
    if (!name) {
      return undefined;
    }
    return { mode: 'create', createName: name.trim() || '新风格' };
  }

  const picked = await vscode.window.showQuickPick(
    profiles.map(profile => ({
      label: profile.name,
      description: `${profile.articleCount} 篇样本`,
      detail: `最近更新：${profile.updatedAt.toLocaleString()}`,
      profileId: profile.id
    })),
    { placeHolder: '选择要更新的风格' }
  );
  if (!picked) {
    return undefined;
  }

  return { mode: 'update', profileId: picked.profileId };
}

async function restoreActiveStyle(
  styleRepository: StyleRepository,
  styleEngine: StyleEngine,
  profileId: string | null
): Promise<void> {
  if (profileId) {
    const profile = styleRepository.getProfile(profileId);
    if (profile) {
      await styleRepository.setActiveProfile(profile.id);
      styleEngine.setCurrentStyle(profile.style);
      return;
    }
  }
  styleEngine.resetStyle();
}

function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function getPrimaryWorkspaceRootUri(): vscode.Uri | null {
  const folders = vscode.workspace.workspaceFolders;
  if (!folders || folders.length === 0) {
    return null;
  }
  return folders[0].uri;
}

function normalizeStorageRootPath(rawPath: string): string {
  const homeExpanded = rawPath.startsWith('~')
    ? path.join(os.homedir(), rawPath.slice(1))
    : rawPath;
  return path.resolve(homeExpanded);
}

function canWriteDirectory(targetPath: string): boolean {
  try {
    fs.mkdirSync(targetPath, { recursive: true });
    fs.accessSync(targetPath, fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

function canWriteExistingFile(filePath: string): boolean {
  if (!fs.existsSync(filePath)) {
    return true;
  }
  try {
    const handle = fs.openSync(filePath, 'r+');
    fs.closeSync(handle);
    return true;
  } catch {
    return false;
  }
}

function canWriteStorageRoot(rootPath: string): boolean {
  if (!canWriteDirectory(rootPath)) {
    return false;
  }
  const dataDir = path.join(rootPath, STORAGE_DATA_RELATIVE_PREFIX);
  if (!canWriteDirectory(dataDir)) {
    return false;
  }

  // 用真实写入探测，规避“目录可访问但文件不可写”的 EPERM 场景（如 Downloads 权限受限）。
  const probeFile = path.join(dataDir, `.write-probe-${Date.now()}-${Math.random().toString(36).slice(2)}.tmp`);
  try {
    fs.writeFileSync(probeFile, 'ok', 'utf8');
    try {
      fs.unlinkSync(probeFile);
    } catch {
      // 删除失败不影响可写判断
    }
  } catch {
    return false;
  }

  if (!canWriteExistingFile(path.join(dataDir, 'styles.json'))) {
    return false;
  }
  if (!canWriteExistingFile(path.join(dataDir, 'materials.by-style.json'))) {
    return false;
  }
  if (!canWriteExistingFile(path.join(dataDir, 'articles.index.json'))) {
    return false;
  }
  return true;
}

function loadCachedResolvedStorageRoot(preferredPath: string): string | null {
  const cache = activeExtensionContext?.globalState.get<StorageRootResolutionCache | null>(
    STORAGE_ROOT_RESOLUTION_CACHE_KEY,
    null
  );
  if (!cache || cache.preferredPath !== preferredPath || !cache.resolvedPath) {
    return null;
  }
  const resolvedPath = cache.resolvedPath;
  if (!canWriteDirectory(resolvedPath)) {
    return null;
  }
  const dataDir = path.join(resolvedPath, STORAGE_DATA_RELATIVE_PREFIX);
  if (!canWriteDirectory(dataDir)) {
    return null;
  }
  if (!canWriteExistingFile(path.join(dataDir, 'styles.json'))) {
    return null;
  }
  if (!canWriteExistingFile(path.join(dataDir, 'materials.by-style.json'))) {
    return null;
  }
  if (!canWriteExistingFile(path.join(dataDir, 'articles.index.json'))) {
    return null;
  }
  return resolvedPath;
}

function persistResolvedStorageRoot(preferredPath: string, resolvedPath: string): void {
  if (!activeExtensionContext) {
    return;
  }
  const cache: StorageRootResolutionCache = { preferredPath, resolvedPath };
  void activeExtensionContext.globalState.update(STORAGE_ROOT_RESOLUTION_CACHE_KEY, cache);
}

function computeStorageRootPath(): string {
  const config = vscode.workspace.getConfiguration('writingAgent');
  const configured = (config.get<string>('storage.path', '') || '').trim();
  const preferred = configured
    ? normalizeStorageRootPath(configured)
    : path.join(os.homedir(), 'Downloads', STORAGE_DEFAULT_DIR_NAME);
  const allowCache = !configured;
  if (allowCache) {
    const cached = loadCachedResolvedStorageRoot(preferred);
    if (cached) {
      return cached;
    }
  }
  if (canWriteStorageRoot(preferred)) {
    if (allowCache) {
      persistResolvedStorageRoot(preferred, preferred);
    }
    return preferred;
  }

  const globalStoragePath = activeExtensionContext?.globalStorageUri.fsPath;
  if (globalStoragePath && canWriteStorageRoot(globalStoragePath)) {
    if (allowCache) {
      persistResolvedStorageRoot(preferred, globalStoragePath);
    }
    console.warn(`[writing-agent] 存储目录不可写（${preferred}），已回退到扩展目录：${globalStoragePath}`);
    return globalStoragePath;
  }

  return preferred;
}

function resolveStorageRootPath(): string {
  if (!storageRootPathCache) {
    storageRootPathCache = computeStorageRootPath();
  }
  return storageRootPathCache;
}

function resolveStorageRootUri(): vscode.Uri {
  return vscode.Uri.file(resolveStorageRootPath());
}

function resolveStorageDataUri(): vscode.Uri {
  return vscode.Uri.joinPath(resolveStorageRootUri(), STORAGE_DATA_RELATIVE_PREFIX);
}

async function ensureStorageDirectories(): Promise<void> {
  const root = resolveStorageRootUri();
  await vscode.workspace.fs.createDirectory(root);
  await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(root, COLLECTION_RELATIVE_PREFIX));
  await vscode.workspace.fs.createDirectory(resolveStorageDataUri());
}

function collectArticleSearchRoots(): vscode.Uri[] {
  const roots: vscode.Uri[] = [resolveStorageRootUri()];
  const workspaceRoot = getPrimaryWorkspaceRootUri();
  if (workspaceRoot && workspaceRoot.path !== roots[0].path) {
    roots.push(workspaceRoot);
  }
  return roots;
}

function getCollectionRelativePrefixes(): string[] {
  return [COLLECTION_RELATIVE_PREFIX, LEGACY_COLLECTION_RELATIVE_PREFIX];
}

function splitRelativePathSegments(relativePath: string): string[] {
  return (relativePath || '')
    .replace(/^\/+/, '')
    .replace(/\\/g, '/')
    .split('/')
    .filter(Boolean);
}

function joinUriByRelativePath(root: vscode.Uri, relativePath: string): vscode.Uri {
  const segments = splitRelativePathSegments(relativePath);
  return segments.length > 0 ? vscode.Uri.joinPath(root, ...segments) : root;
}

function toCanonicalCollectionRelativePath(relativePath: string): string {
  const normalized = splitRelativePathSegments(relativePath).join('/');
  if (!normalized) {
    return '';
  }
  const legacyPrefix = `${LEGACY_COLLECTION_RELATIVE_PREFIX}/`;
  if (normalized.startsWith(legacyPrefix)) {
    return `${COLLECTION_RELATIVE_PREFIX}/${normalized.slice(legacyPrefix.length)}`;
  }
  return normalized;
}

function getCollectionRelativePathCandidates(relativePath: string): string[] {
  const normalized = splitRelativePathSegments(relativePath).join('/');
  if (!normalized) {
    return [];
  }
  const currentPrefix = `${COLLECTION_RELATIVE_PREFIX}/`;
  const legacyPrefix = `${LEGACY_COLLECTION_RELATIVE_PREFIX}/`;
  if (normalized.startsWith(currentPrefix)) {
    const suffix = normalized.slice(currentPrefix.length);
    return [normalized, `${LEGACY_COLLECTION_RELATIVE_PREFIX}/${suffix}`];
  }
  if (normalized.startsWith(legacyPrefix)) {
    const suffix = normalized.slice(legacyPrefix.length);
    return [normalized, `${COLLECTION_RELATIVE_PREFIX}/${suffix}`];
  }
  return [normalized];
}

function resolveArticleUriByRelativePath(relativePath: string): vscode.Uri | null {
  const relativePathCandidates = getCollectionRelativePathCandidates(relativePath);
  if (relativePathCandidates.length === 0) {
    return null;
  }

  const roots = collectArticleSearchRoots();
  for (const root of roots) {
    for (const candidateRelativePath of relativePathCandidates) {
      const candidate = joinUriByRelativePath(root, candidateRelativePath);
      if (candidate.scheme === 'file' && fs.existsSync(candidate.fsPath)) {
        return candidate;
      }
    }
  }
  return joinUriByRelativePath(resolveStorageRootUri(), relativePathCandidates[0]);
}

async function resolveArticleUriForDelete(relativePath: string, preferredRoot: vscode.Uri): Promise<vscode.Uri> {
  const relativePathCandidates = getCollectionRelativePathCandidates(relativePath);
  if (relativePathCandidates.length === 0) {
    return preferredRoot;
  }

  for (const candidateRelativePath of relativePathCandidates) {
    const preferred = joinUriByRelativePath(preferredRoot, candidateRelativePath);
    if (await uriExists(preferred)) {
      return preferred;
    }
  }

  const workspaceRoot = getPrimaryWorkspaceRootUri();
  if (workspaceRoot && workspaceRoot.path !== preferredRoot.path) {
    for (const candidateRelativePath of relativePathCandidates) {
      const fallback = joinUriByRelativePath(workspaceRoot, candidateRelativePath);
      if (await uriExists(fallback)) {
        return fallback;
      }
    }
  }

  return joinUriByRelativePath(preferredRoot, relativePathCandidates[0]);
}

function getConfigurationTarget(): vscode.ConfigurationTarget {
  return vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0
    ? vscode.ConfigurationTarget.Workspace
    : vscode.ConfigurationTarget.Global;
}

function isApiProvider(value: string): value is ApiProvider {
  return ['iflow', 'openrouter', 'kimi', 'deepseek', 'minimax', 'qwen', 'custom'].includes(value);
}

function getProviderPreset(provider: ApiProvider): ProviderPreset {
  return PROVIDER_PRESETS[provider];
}

function getAllApiProviders(): ApiProvider[] {
  return Object.keys(PROVIDER_PRESETS) as ApiProvider[];
}

function normalizeProviderOrder(
  rawOrder: unknown,
  fallbackProvider?: ApiProvider
): ApiProvider[] {
  const ordered: ApiProvider[] = [];
  const pushIfValid = (value: unknown): void => {
    if (typeof value !== 'string' || !isApiProvider(value)) {
      return;
    }
    if (!ordered.includes(value)) {
      ordered.push(value);
    }
  };
  if (Array.isArray(rawOrder)) {
    rawOrder.forEach(pushIfValid);
  }
  if (fallbackProvider) {
    pushIfValid(fallbackProvider);
  }
  getAllApiProviders().forEach(pushIfValid);
  return ordered;
}

function normalizeModelCandidates(candidates: string[]): string[] {
  const normalized: string[] = [];
  for (const model of candidates) {
    const trimmed = model.trim();
    if (!trimmed || normalized.includes(trimmed)) {
      continue;
    }
    normalized.push(trimmed);
  }
  return normalized;
}

function parseModelCandidatesInput(rawInput: string, fallbackModels: string[]): string[] {
  const rawCandidates = rawInput
    .split(/[\n,，]/g)
    .map(item => item.trim())
    .filter(Boolean);
  const parsed = normalizeModelCandidates(rawCandidates);
  if (parsed.length > 0) {
    return parsed;
  }
  const fallback = normalizeModelCandidates(fallbackModels);
  return fallback;
}

function formatModelChainLabel(models: string[]): string {
  const normalized = normalizeModelCandidates(models);
  if (normalized.length === 0) {
    return '未配置模型';
  }
  if (normalized.length <= 2) {
    return normalized.join('→');
  }
  return `${normalized.slice(0, 2).join('→')}→+${normalized.length - 2}`;
}

async function promptManualModelCandidates(
  displayName: string,
  defaultModels: string[]
): Promise<string[] | undefined> {
  const defaultValue = normalizeModelCandidates(defaultModels).join(', ');
  const modelInput = await vscode.window.showInputBox({
    prompt: `模型列表（${displayName}，按优先级；支持逗号或换行分隔）`,
    value: defaultValue
  });
  if (modelInput === undefined) {
    return undefined;
  }
  const fallbackModels = normalizeModelCandidates(defaultModels);
  const parsed = parseModelCandidatesInput(modelInput, fallbackModels);
  return parsed.length > 0 ? parsed : fallbackModels;
}

function getModelOrderEditorWebview(models: string[]): string {
  const initialModels = JSON.stringify(models);
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <style>
    body {
      font-family: var(--vscode-font-family);
      color: var(--vscode-foreground);
      background: var(--vscode-editor-background);
      margin: 0;
      padding: 16px;
    }
    .hint {
      color: var(--vscode-descriptionForeground);
      margin-bottom: 12px;
      font-size: 12px;
    }
    ul {
      list-style: none;
      padding: 0;
      margin: 0;
      border: 1px solid var(--vscode-panel-border);
      border-radius: 6px;
      overflow: hidden;
    }
    li {
      padding: 10px 12px;
      border-bottom: 1px solid var(--vscode-panel-border);
      background: var(--vscode-editorWidget-background);
      cursor: grab;
      user-select: none;
    }
    li:last-child {
      border-bottom: none;
    }
    li.dragging {
      opacity: 0.55;
    }
    li.over {
      outline: 2px solid var(--vscode-focusBorder);
      outline-offset: -2px;
    }
    .actions {
      margin-top: 14px;
      display: flex;
      gap: 8px;
    }
    button {
      border: 0;
      border-radius: 4px;
      padding: 7px 12px;
      cursor: pointer;
      font-size: 12px;
    }
    #save {
      background: var(--vscode-button-background);
      color: var(--vscode-button-foreground);
    }
    #cancel {
      background: var(--vscode-button-secondaryBackground);
      color: var(--vscode-button-secondaryForeground);
    }
  </style>
</head>
<body>
  <div class="hint">拖拽排序后，调用会按从上到下顺序依次尝试。</div>
  <ul id="models"></ul>
  <div class="actions">
    <button id="save">保存顺序</button>
    <button id="cancel">取消</button>
  </div>
  <script>
    const vscode = acquireVsCodeApi();
    let models = ${initialModels};
    let dragIndex = -1;

    const list = document.getElementById('models');
    const saveBtn = document.getElementById('save');
    const cancelBtn = document.getElementById('cancel');

    function render() {
      list.innerHTML = '';
      models.forEach((model, index) => {
        const li = document.createElement('li');
        li.draggable = true;
        li.dataset.index = String(index);
        li.textContent = model;
        li.addEventListener('dragstart', event => {
          dragIndex = Number(li.dataset.index);
          li.classList.add('dragging');
          if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', li.dataset.index || '0');
          }
        });
        li.addEventListener('dragend', () => {
          dragIndex = -1;
          li.classList.remove('dragging');
          document.querySelectorAll('li.over').forEach(item => item.classList.remove('over'));
        });
        li.addEventListener('dragover', event => {
          event.preventDefault();
          li.classList.add('over');
        });
        li.addEventListener('dragleave', () => {
          li.classList.remove('over');
        });
        li.addEventListener('drop', event => {
          event.preventDefault();
          li.classList.remove('over');
          const dropIndex = Number(li.dataset.index);
          if (!Number.isInteger(dragIndex) || dragIndex < 0 || dragIndex === dropIndex) {
            return;
          }
          const next = [...models];
          const moved = next.splice(dragIndex, 1)[0];
          next.splice(dropIndex, 0, moved);
          models = next;
          render();
        });
        list.appendChild(li);
      });
    }

    saveBtn.addEventListener('click', () => {
      vscode.postMessage({ type: 'save', models });
    });
    cancelBtn.addEventListener('click', () => {
      vscode.postMessage({ type: 'cancel' });
    });

    render();
  </script>
</body>
</html>`;
}

async function editModelOrderByDrag(
  title: string,
  models: string[]
): Promise<string[] | undefined> {
  const normalized = normalizeModelCandidates(models);
  if (normalized.length <= 1) {
    return normalized;
  }

  const panel = vscode.window.createWebviewPanel(
    'writingAgent.modelOrderEditor',
    title,
    vscode.ViewColumn.Active,
    { enableScripts: true }
  );
  panel.webview.html = getModelOrderEditorWebview(normalized);

  return new Promise<string[] | undefined>(resolve => {
    let settled = false;
    const finish = (value: string[] | undefined): void => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(value);
    };

    const receiveDisposable = panel.webview.onDidReceiveMessage(message => {
      const payload = message as { type?: string; models?: unknown };
      if (payload.type === 'save') {
        const next = Array.isArray(payload.models)
          ? payload.models.filter((item): item is string => typeof item === 'string')
          : [];
        const parsed = normalizeModelCandidates(next);
        finish(parsed.length > 0 ? parsed : normalized);
        panel.dispose();
        return;
      }
      if (payload.type === 'cancel') {
        finish(undefined);
        panel.dispose();
      }
    });
    const disposeDisposable = panel.onDidDispose(() => {
      receiveDisposable.dispose();
      finish(undefined);
    });
    void receiveDisposable;
    void disposeDisposable;
  });
}

function getProviderOrderEditorWebview(providers: ApiProvider[]): string {
  const initialProviders = JSON.stringify(providers);
  const providerMeta = JSON.stringify(
    getAllApiProviders().reduce<Record<string, { label: string; detail: string }>>((acc, provider) => {
      const preset = getProviderPreset(provider);
      acc[provider] = {
        label: preset.displayName,
        detail: `${provider} · ${preset.defaultBaseUrl}`
      };
      return acc;
    }, {})
  );
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <style>
    body { font-family: var(--vscode-font-family); color: var(--vscode-foreground); padding: 16px; }
    .hint { color: var(--vscode-descriptionForeground); margin-bottom: 12px; font-size: 12px; }
    ul { list-style: none; padding: 0; margin: 0; border: 1px solid var(--vscode-panel-border); border-radius: 6px; overflow: hidden; }
    li { padding: 10px 12px; border-bottom: 1px solid var(--vscode-panel-border); background: var(--vscode-editorWidget-background); cursor: grab; user-select: none; }
    li:last-child { border-bottom: none; }
    li.dragging { opacity: 0.55; }
    li.over { outline: 2px solid var(--vscode-focusBorder); outline-offset: -2px; }
    .provider-name { font-weight: 600; }
    .provider-detail { color: var(--vscode-descriptionForeground); font-size: 12px; margin-top: 2px; }
    .actions { margin-top: 14px; display: flex; gap: 8px; }
    button { border: 0; border-radius: 4px; padding: 7px 12px; cursor: pointer; font-size: 12px; }
    #save { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
    #cancel { background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); }
  </style>
</head>
<body>
  <div class="hint">拖拽排序后，调用会按从上到下顺序依次尝试各 API 提供商。</div>
  <ul id="providers"></ul>
  <div class="actions">
    <button id="save">保存顺序</button>
    <button id="cancel">取消</button>
  </div>
  <script>
    const vscode = acquireVsCodeApi();
    const providerMeta = ${providerMeta};
    let providers = ${initialProviders};
    let dragIndex = -1;
    const list = document.getElementById('providers');
    const saveBtn = document.getElementById('save');
    const cancelBtn = document.getElementById('cancel');

    function render() {
      list.innerHTML = '';
      providers.forEach((provider, index) => {
        const li = document.createElement('li');
        li.draggable = true;
        li.dataset.index = String(index);

        const name = document.createElement('div');
        name.className = 'provider-name';
        name.textContent = providerMeta[provider]?.label || provider;

        const detail = document.createElement('div');
        detail.className = 'provider-detail';
        detail.textContent = providerMeta[provider]?.detail || provider;

        li.appendChild(name);
        li.appendChild(detail);
        li.addEventListener('dragstart', event => {
          dragIndex = Number(li.dataset.index);
          li.classList.add('dragging');
          if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', li.dataset.index || '0');
          }
        });
        li.addEventListener('dragend', () => {
          dragIndex = -1;
          li.classList.remove('dragging');
          document.querySelectorAll('li.over').forEach(item => item.classList.remove('over'));
        });
        li.addEventListener('dragover', event => {
          event.preventDefault();
          li.classList.add('over');
        });
        li.addEventListener('dragleave', () => {
          li.classList.remove('over');
        });
        li.addEventListener('drop', event => {
          event.preventDefault();
          li.classList.remove('over');
          const dropIndex = Number(li.dataset.index);
          if (!Number.isInteger(dragIndex) || dragIndex < 0 || dragIndex === dropIndex) {
            return;
          }
          const next = [...providers];
          const moved = next.splice(dragIndex, 1)[0];
          next.splice(dropIndex, 0, moved);
          providers = next;
          render();
        });
        list.appendChild(li);
      });
    }

    saveBtn.addEventListener('click', () => vscode.postMessage({ type: 'save', providers }));
    cancelBtn.addEventListener('click', () => vscode.postMessage({ type: 'cancel' }));
    render();
  </script>
</body>
</html>`;
}

async function editProviderOrderByDrag(
  title: string,
  providers: ApiProvider[]
): Promise<ApiProvider[] | undefined> {
  const normalized = normalizeProviderOrder(providers);
  if (normalized.length <= 1) {
    return normalized;
  }
  const panel = vscode.window.createWebviewPanel(
    'writingAgent.providerOrderEditor',
    title,
    vscode.ViewColumn.Active,
    { enableScripts: true }
  );
  panel.webview.html = getProviderOrderEditorWebview(normalized);

  return new Promise<ApiProvider[] | undefined>(resolve => {
    let settled = false;
    const finish = (value: ApiProvider[] | undefined): void => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(value);
    };

    const receiveDisposable = panel.webview.onDidReceiveMessage(message => {
      const payload = message as { type?: string; providers?: unknown };
      if (payload.type === 'save') {
        const next = normalizeProviderOrder(payload.providers);
        finish(next.length > 0 ? next : normalized);
        panel.dispose();
        return;
      }
      if (payload.type === 'cancel') {
        finish(undefined);
        panel.dispose();
      }
    });
    const disposeDisposable = panel.onDidDispose(() => {
      receiveDisposable.dispose();
      finish(undefined);
    });
    void receiveDisposable;
    void disposeDisposable;
  });
}

async function pickProviderModelCandidates(
  providerDisplayName: string,
  apiKey: string,
  baseUrl: string,
  defaultModels: string[],
  timeoutMs: number
): Promise<string[] | undefined> {
  let fetchedModels: string[] = [];
  if (apiKey.trim()) {
    try {
      await vscode.window.withProgress(
        {
          location: vscode.ProgressLocation.Notification,
          title: `正在加载 ${providerDisplayName} 模型列表...`,
          cancellable: false
        },
        async () => {
          fetchedModels = await listOpenAiCompatibleModels({
            apiKey: apiKey.trim(),
            baseUrl,
            timeoutMs
          });
        }
      );
    } catch (error) {
      const message = error instanceof Error ? summarizeError(error.message, 60) : '请求失败';
      notifyWarn(`获取 ${providerDisplayName} 模型列表失败，将回退本地候选：${message}`);
    }
  }

  const candidates = normalizeModelCandidates([...fetchedModels, ...defaultModels]);
  if (candidates.length === 0) {
    return undefined;
  }

  const picked = await vscode.window.showQuickPick(
    candidates.map(model => ({
      label: model,
      description: defaultModels.includes(model) ? '当前/默认' : ''
    })),
    {
      canPickMany: true,
      placeHolder: `选择一个或多个 ${providerDisplayName} 模型（按列表顺序作为优先级）`,
      ignoreFocusOut: true
    }
  );
  if (!picked) {
    return undefined;
  }
  const selectedModels = normalizeModelCandidates(picked.map(item => item.label));
  return selectedModels.length > 0 ? selectedModels : undefined;
}

function getProviderSecretKey(provider: ApiProvider): string {
  return `writingAgent.apiKey.${provider}`;
}

function getProviderEnvKey(provider: ApiProvider): string {
  switch (provider) {
    case 'iflow':
      return 'IFLOW_API_KEY';
    case 'openrouter':
      return 'OPENROUTER_API_KEY';
    case 'kimi':
      return 'KIMI_API_KEY';
    case 'deepseek':
      return 'DEEPSEEK_API_KEY';
    case 'minimax':
      return 'MINIMAX_API_KEY';
    case 'qwen':
      return 'QWEN_API_KEY';
    case 'custom':
      return 'WRITING_AGENT_API_KEY';
    default:
      return 'WRITING_AGENT_API_KEY';
  }
}

async function resolveProviderApiKey(
  context: vscode.ExtensionContext,
  provider: ApiProvider
): Promise<string> {
  const config = vscode.workspace.getConfiguration('writingAgent');
  const providerSecret = ((await context.secrets.get(getProviderSecretKey(provider))) || '').trim();
  const legacyIflowSecret = provider === 'iflow'
    ? ((await context.secrets.get(LEGACY_IFLOW_SECRET_KEY)) || '').trim()
    : '';
  const legacyConfigKey = provider === 'iflow'
    ? (config.get<string>('iflow.apiKey', '') || '').trim()
    : '';
  const envProviderKey = (process.env[getProviderEnvKey(provider)] || '').trim();
  const envGenericKey = (process.env.WRITING_AGENT_API_KEY || '').trim();
  return providerSecret || legacyIflowSecret || legacyConfigKey || envProviderKey || envGenericKey;
}

function getConfiguredProviderOrder(config: vscode.WorkspaceConfiguration, currentProvider: ApiProvider): ApiProvider[] {
  const raw = config.get<unknown>('api.providerOrder', []);
  return normalizeProviderOrder(raw, currentProvider);
}

async function updateProviderOrderConfig(
  providerOrder: ApiProvider[],
  target: vscode.ConfigurationTarget
): Promise<void> {
  const config = vscode.workspace.getConfiguration('writingAgent');
  const normalized = normalizeProviderOrder(providerOrder);
  await config.update('api.providerOrder', normalized, target);
}

async function pickApiProvider(currentProvider: ApiProvider): Promise<ApiProvider | undefined> {
  const config = vscode.workspace.getConfiguration('writingAgent');
  const providerOrder = getConfiguredProviderOrder(config, currentProvider);
  const options = providerOrder.map((provider, index) => {
    const preset = getProviderPreset(provider);
    return {
      label: preset.displayName,
      description: provider === currentProvider ? '当前' : `顺序 ${index + 1}`,
      detail: `${provider} · ${preset.defaultBaseUrl}`,
      provider
    };
  });
  const picked = await vscode.window.showQuickPick(options, {
    placeHolder: '选择 API 提供商'
  });
  return picked?.provider;
}

async function resolveApiSettings(
  context: vscode.ExtensionContext
): Promise<RuntimeApiSettings | null> {
  const config = vscode.workspace.getConfiguration('writingAgent');
  const enabled = config.get<boolean>('api.enabled', config.get<boolean>('iflow.enabled', true));
  const rawProvider = config.get<string>('api.provider', 'iflow');
  const provider: ApiProvider = isApiProvider(rawProvider) ? rawProvider : 'iflow';
  const providerOrder = getConfiguredProviderOrder(config, provider);

  const rawProviderConfigs = config.get<unknown>('api.providerConfigs', {});
  const providerConfigs = rawProviderConfigs && typeof rawProviderConfigs === 'object'
    ? rawProviderConfigs as Record<string, unknown>
    : {};

  const resolvedByProvider = new Map<ApiProvider, RuntimeProviderCandidate>();
  for (const item of providerOrder) {
    const preset = getProviderPreset(item);
    const rawConfig = providerConfigs[item];
    const providerConfig = rawConfig && typeof rawConfig === 'object'
      ? rawConfig as Record<string, unknown>
      : {};

    const configuredBaseUrl = typeof providerConfig.baseUrl === 'string'
      ? providerConfig.baseUrl
      : '';
    const baseUrl = item === provider
      ? config.get<string>(
        'api.baseUrl',
        item === 'iflow'
          ? config.get<string>('iflow.baseUrl', configuredBaseUrl || preset.defaultBaseUrl)
          : configuredBaseUrl || preset.defaultBaseUrl
      )
      : (configuredBaseUrl || preset.defaultBaseUrl);

    const configuredModel = typeof providerConfig.model === 'string' ? providerConfig.model : '';
    const model = item === provider
      ? config.get<string>(
        'api.model',
        item === 'iflow'
          ? config.get<string>('iflow.model', configuredModel || preset.defaultModel)
          : configuredModel || preset.defaultModel
      )
      : (configuredModel || preset.defaultModel);

    const configuredModelsRaw = providerConfig.models;
    const configuredModels = Array.isArray(configuredModelsRaw)
      ? configuredModelsRaw.filter((value): value is string => typeof value === 'string')
      : [];
    const activeModels = item === provider
      ? (Array.isArray(config.get<unknown>('api.models', []))
        ? (config.get<unknown>('api.models', []) as unknown[]).filter((value): value is string => typeof value === 'string')
        : [])
      : configuredModels;
    const parsedModels = normalizeModelCandidates([model, ...activeModels]);
    const models = parsedModels.length > 0 ? parsedModels : [preset.defaultModel];
    const primaryModel = models[0];

    const apiKey = await resolveProviderApiKey(context, item);
    resolvedByProvider.set(item, {
      provider: item,
      apiKey,
      baseUrl,
      model: primaryModel,
      models
    });
  }

  const activeCandidate = resolvedByProvider.get(provider) || {
    provider,
    apiKey: await resolveProviderApiKey(context, provider),
    baseUrl: getProviderPreset(provider).defaultBaseUrl,
    model: getProviderPreset(provider).defaultModel,
    models: [getProviderPreset(provider).defaultModel]
  };
  const providerCandidates = providerOrder
    .map(item => resolvedByProvider.get(item))
    .filter((item): item is RuntimeProviderCandidate => Boolean(item));

  const temperature = config.get<number>('api.temperature', config.get<number>('iflow.temperature', 0.7));
  const maxTokens = config.get<number>('api.maxTokens', config.get<number>('iflow.maxTokens', 10000));
  const timeoutMs = config.get<number>('api.timeoutMs', config.get<number>('iflow.timeoutMs', 60000));

  if (!enabled) {
    return {
      enabled: false,
      provider,
      providerOrder,
      providerCandidates,
      apiKey: '',
      baseUrl: activeCandidate.baseUrl,
      model: activeCandidate.model,
      models: activeCandidate.models,
      temperature,
      maxTokens,
      timeoutMs
    };
  }

  return {
    enabled,
    provider,
    providerOrder,
    providerCandidates,
    apiKey: activeCandidate.apiKey,
    baseUrl: activeCandidate.baseUrl,
    model: activeCandidate.model,
    models: activeCandidate.models,
    temperature,
    maxTokens,
    timeoutMs
  };
}

function buildApiValidationFingerprint(settings: RuntimeApiSettings): string {
  return [
    settings.providerOrder.join('>'),
    ...settings.providerCandidates.map(candidate => `${candidate.provider}|${candidate.baseUrl.trim()}|${candidate.models.join(',')}|${candidate.apiKey ? '1' : '0'}`)
  ].join('|');
}

function loadApiValidationState(context: vscode.ExtensionContext): ApiValidationState | null {
  return context.workspaceState.get<ApiValidationState | null>(API_VALIDATION_STATE_KEY, null);
}

async function saveApiValidationState(
  context: vscode.ExtensionContext,
  state: ApiValidationState
): Promise<void> {
  await context.workspaceState.update(API_VALIDATION_STATE_KEY, state);
}

function clearApiValidationState(context: vscode.ExtensionContext): void {
  void context.workspaceState.update(API_VALIDATION_STATE_KEY, null);
}

function summarizeError(message: string, maxLength: number): string {
  const singleLine = message.replace(/\s+/g, ' ').trim();
  if (singleLine.length <= maxLength) {
    return singleLine;
  }
  return `${singleLine.slice(0, Math.max(0, maxLength - 1))}…`;
}

function buildApiFailureHint(message: string): string | null {
  const normalized = message.toLowerCase();
  if (
    normalized.includes('401')
    || normalized.includes('unauthorized')
    || normalized.includes('invalid api key')
    || normalized.includes('no auth')
    || normalized.includes('api key')
    || normalized.includes('认证')
    || normalized.includes('密钥')
  ) {
    return '请检查 API Key 是否正确，且与当前提供商匹配。';
  }
  if (
    normalized.includes('429')
    || normalized.includes('quota')
    || normalized.includes('insufficient')
    || normalized.includes('rate limit')
    || normalized.includes('额度')
    || normalized.includes('余额')
    || normalized.includes('限流')
  ) {
    return '可能触发额度/限流，请稍后重试或更换备用模型。';
  }
  if (normalized.includes('timeout') || normalized.includes('超时')) {
    return '请求可能超时，可提高 `writingAgent.api.timeoutMs` 或更换响应更快的模型。';
  }
  return null;
}

function hasAnyApiProviderKey(settings: RuntimeApiSettings): boolean {
  if (settings.apiKey.trim()) {
    return true;
  }
  return settings.providerCandidates.some(candidate => candidate.apiKey.trim().length > 0);
}

function buildApiFallbackNotice(result: ApiGenerationResult): string | null {
  if (!result.fallbackUsed || !result.fallbackFromModel) {
    return null;
  }
  const fromProviderLabel = result.fallbackFromProvider
    ? getProviderPreset(result.fallbackFromProvider).displayName
    : '';
  const toProviderLabel = getProviderPreset(result.provider).displayName;
  if (result.fallbackFromProvider && result.fallbackFromProvider !== result.provider) {
    return `主调用“${fromProviderLabel}/${result.fallbackFromModel}”不可用，已自动切换到“${toProviderLabel}/${result.model}”。`;
  }
  return `主模型“${result.fallbackFromModel}”不可用，已自动切换到“${result.model}”。`;
}

function shouldTryModelDiscoveryForValidation(settings: RuntimeApiSettings, failureMessage: string): boolean {
  if (!settings.apiKey.trim()) {
    return false;
  }
  if (settings.provider !== 'iflow' && settings.provider !== 'openrouter') {
    return false;
  }
  const normalized = failureMessage.toLowerCase();
  return (
    normalized.includes('所有模型均不可用')
    || normalized.includes('model')
    || normalized.includes('不支持')
    || normalized.includes('not found')
    || normalized.includes('invalid model')
    || normalized.includes('does not exist')
  );
}

function pickDiscoveryCandidatesForValidation(
  provider: ApiProvider,
  discoveredModels: string[],
  existingModels: string[]
): string[] {
  const normalizedDiscovered = normalizeModelCandidates(discoveredModels);
  const normalizedExisting = normalizeModelCandidates(existingModels);
  if (normalizedDiscovered.length === 0) {
    return [];
  }
  if (provider === 'openrouter') {
    const freeModels = normalizedDiscovered.filter(model => model.toLowerCase().includes(':free'));
    const regularModels = normalizedDiscovered.filter(model => !model.toLowerCase().includes(':free'));
    const ranked = normalizeModelCandidates([...freeModels, ...regularModels]).slice(0, 6);
    return normalizeModelCandidates([...ranked, ...normalizedExisting]);
  }
  return normalizeModelCandidates([...normalizedDiscovered.slice(0, 6), ...normalizedExisting]);
}

async function validateApiConnection(
  context: vscode.ExtensionContext,
  settings: RuntimeApiSettings,
  persist: boolean
): Promise<{ valid: boolean; message: string }> {
  const prompt = '请仅回复 OK';
  try {
    const result = await generateWithFallbackModel(settings, prompt, {
      temperature: 0,
      maxTokens: 16,
      timeoutMs: Math.min(
        Math.max(settings.timeoutMs, API_VALIDATION_TIMEOUT_MIN_MS),
        API_VALIDATION_TIMEOUT_MAX_MS
      )
    });
    const providerLabel = getProviderPreset(result.provider).displayName;
    const fallbackNotice = buildApiFallbackNotice(result);
    const message = fallbackNotice
      ? `连接成功（${providerLabel}/${result.model}；${fallbackNotice}）`
      : `连接成功（${providerLabel}/${result.model}）`;
    if (persist) {
      await saveApiValidationState(context, {
        fingerprint: buildApiValidationFingerprint(settings),
        valid: true,
        message,
        checkedAt: new Date().toISOString()
      });
    }
    return { valid: true, message };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (shouldTryModelDiscoveryForValidation(settings, message)) {
      try {
        const discovered = await listOpenAiCompatibleModels({
          apiKey: settings.apiKey,
          baseUrl: settings.baseUrl,
          timeoutMs: Math.min(
            Math.max(settings.timeoutMs, API_VALIDATION_TIMEOUT_MIN_MS),
            API_VALIDATION_TIMEOUT_MAX_MS
          )
        });
        const discoveredModels = pickDiscoveryCandidatesForValidation(
          settings.provider,
          discovered,
          settings.models
        );
        if (discoveredModels.length > 0) {
          const activeCandidate: RuntimeProviderCandidate = {
            provider: settings.provider,
            apiKey: settings.apiKey,
            baseUrl: settings.baseUrl,
            model: discoveredModels[0],
            models: discoveredModels
          };
          const retryResult = await generateWithFallbackModel(
            {
              ...settings,
              providerOrder: [settings.provider],
              providerCandidates: [activeCandidate],
              model: activeCandidate.model,
              models: activeCandidate.models
            },
            prompt,
            {
              temperature: 0,
              maxTokens: 16,
              timeoutMs: Math.min(
                Math.max(settings.timeoutMs, API_VALIDATION_TIMEOUT_MIN_MS),
                API_VALIDATION_TIMEOUT_MAX_MS
              )
            }
          );
          const validatedModels = normalizeModelCandidates([retryResult.model, ...discoveredModels]);
          const config = vscode.workspace.getConfiguration('writingAgent');
          const target = getConfigurationTarget();
          await config.update('api.model', validatedModels[0], getConfigurationTarget());
          await config.update('api.models', validatedModels, getConfigurationTarget());
          const rawProviderConfigs = config.get<unknown>('api.providerConfigs', {});
          const providerConfigs = rawProviderConfigs && typeof rawProviderConfigs === 'object'
            ? { ...(rawProviderConfigs as Record<string, unknown>) }
            : {};
          providerConfigs[settings.provider] = {
            baseUrl: settings.baseUrl,
            model: validatedModels[0],
            models: validatedModels
          };
          await config.update('api.providerConfigs', providerConfigs, target);

          const providerLabel = getProviderPreset(settings.provider).displayName;
          const autoMessage = `连接成功（${providerLabel}/${retryResult.model}，已自动刷新模型列表）`;
          if (persist) {
            await saveApiValidationState(context, {
              fingerprint: buildApiValidationFingerprint({
                ...settings,
                providerCandidates: settings.providerCandidates.map(candidate => candidate.provider === settings.provider
                  ? { ...candidate, model: validatedModels[0], models: validatedModels }
                  : candidate),
                model: validatedModels[0],
                models: validatedModels
              }),
              valid: true,
              message: autoMessage,
              checkedAt: new Date().toISOString()
            });
          }
          return { valid: true, message: autoMessage };
        }
      } catch {
        // 自动发现失败时返回原始错误
      }
    }
    if (persist) {
      await saveApiValidationState(context, {
        fingerprint: buildApiValidationFingerprint(settings),
        valid: false,
        message,
        checkedAt: new Date().toISOString()
      });
    }
    return { valid: false, message };
  }
}

async function removeArticlesByStyle(
  styleId: string,
  articleRepository: ArticleRepository,
  storageRoot: vscode.Uri
): Promise<DeleteStyleArticlesResult> {
  const records = articleRepository.listByStyle(styleId);
  const result: DeleteStyleArticlesResult = {
    removedRecords: 0,
    deletedFiles: 0,
    missingFiles: 0
  };

  for (const record of records) {
    const articleUri = await resolveArticleUriForDelete(record.relativePath, storageRoot);
    try {
      if (await uriExists(articleUri)) {
        await vscode.workspace.fs.delete(articleUri, { recursive: false, useTrash: false });
        result.deletedFiles += 1;
      } else {
        result.missingFiles += 1;
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error(`[writing-agent] 删除风格文章文件失败: ${record.relativePath}`, message);
    }

    const removed = await articleRepository.deleteById(record.id);
    if (removed) {
      result.removedRecords += 1;
    }
  }
  return result;
}

async function moveArticlesToStyle(
  sourceStyle: StyleProfile,
  targetStyle: StyleProfile,
  articleRepository: ArticleRepository,
  storageRoot: vscode.Uri
): Promise<MergeStyleArticlesResult> {
  const records = articleRepository.listByStyle(sourceStyle.id);
  const result: MergeStyleArticlesResult = {
    updatedRecords: 0,
    movedFiles: 0,
    renamedFiles: 0,
    missingFiles: 0
  };

  const targetDir = vscode.Uri.joinPath(storageRoot, COLLECTION_RELATIVE_PREFIX, sanitizePathSegment(targetStyle.name));
  await vscode.workspace.fs.createDirectory(targetDir);

  for (const record of records) {
    let nextRelativePath = record.relativePath;
    let nextTitle = record.title;

    const sourceUri = await resolveArticleUriForDelete(record.relativePath, storageRoot);
    try {
      if (await uriExists(sourceUri)) {
        const sourceName = sourceUri.path.slice(sourceUri.path.lastIndexOf('/') + 1);
        const targetUri = vscode.Uri.joinPath(targetDir, sourceName);
        const finalTarget = await ensureUniqueTargetUri(targetUri);
        if (finalTarget.path !== targetUri.path) {
          result.renamedFiles += 1;
        }
        await vscode.workspace.fs.rename(sourceUri, finalTarget, { overwrite: false });
        nextRelativePath = finalTarget.path.replace(storageRoot.path, '').replace(/^\//, '');
        nextTitle = finalTarget.path.slice(finalTarget.path.lastIndexOf('/') + 1).replace(/\.md$/i, '');
        result.movedFiles += 1;
      } else {
        result.missingFiles += 1;
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error(`[writing-agent] 迁移风格文章文件失败: ${record.relativePath}`, message);
    }

    await articleRepository.upsert({
      id: record.id,
      topic: record.topic,
      title: nextTitle,
      styleId: targetStyle.id,
      styleName: targetStyle.name,
      relativePath: nextRelativePath,
      aiCommand: record.aiCommand,
      createdAt: record.createdAt,
      updatedAt: new Date()
    });
    result.updatedRecords += 1;
  }

  return result;
}

async function resolveArticleRecordTarget(
  arg: unknown,
  articleRepository: ArticleRepository
): Promise<ArticleRecord | undefined> {
  const fromNode = arg as { record?: ArticleRecord } | undefined;
  if (fromNode?.record?.id) {
    return fromNode.record;
  }

  const fromRecord = arg as ArticleRecord | undefined;
  if (fromRecord?.id && fromRecord.relativePath) {
    return fromRecord;
  }

  const fromId = typeof arg === 'string' ? articleRepository.getById(arg) : undefined;
  if (fromId) {
    return fromId;
  }

  const all = articleRepository.listAll();
  if (all.length === 0) {
    notifyWarn('文集为空，暂无可删除文章。');
    return undefined;
  }
  const picked = await vscode.window.showQuickPick(
    all.map(record => ({
      label: record.title,
      description: record.styleName,
      detail: record.relativePath,
      recordId: record.id
    })),
    { placeHolder: '选择要删除的文集文章' }
  );
  if (!picked) {
    return undefined;
  }
  return articleRepository.getById(picked.recordId);
}

async function ensureUniqueTargetUri(uri: vscode.Uri): Promise<vscode.Uri> {
  if (!(await uriExists(uri))) {
    return uri;
  }
  const originalPath = uri.path;
  const extensionIndex = originalPath.lastIndexOf('.');
  const hasExtension = extensionIndex > originalPath.lastIndexOf('/');
  const base = hasExtension ? originalPath.slice(0, extensionIndex) : originalPath;
  const extension = hasExtension ? originalPath.slice(extensionIndex) : '';
  let index = 2;
  while (true) {
    const candidate = uri.with({ path: `${base} (${index})${extension}` });
    if (!(await uriExists(candidate))) {
      return candidate;
    }
    index += 1;
  }
}

async function generateWithFallbackModel(
  settings: RuntimeApiSettings,
  prompt: string,
  overrides?: {
    temperature?: number;
    maxTokens?: number;
    timeoutMs?: number;
  }
): Promise<ApiGenerationResult> {
  const providerCandidates = settings.providerCandidates.length > 0
    ? settings.providerCandidates
    : [
      {
        provider: settings.provider,
        apiKey: settings.apiKey,
        baseUrl: settings.baseUrl,
        model: settings.model,
        models: settings.models
      }
    ];

  const firstProvider = providerCandidates[0];
  const firstProviderModels = normalizeModelCandidates(firstProvider.models.length > 0 ? firstProvider.models : [firstProvider.model]);
  const firstModel = firstProviderModels[0];

  const failedReasons: string[] = [];
  let hasUsableKey = false;
  for (let providerIndex = 0; providerIndex < providerCandidates.length; providerIndex += 1) {
    const candidateProvider = providerCandidates[providerIndex];
    const providerName = getProviderPreset(candidateProvider.provider).displayName;
    if (!candidateProvider.apiKey.trim()) {
      failedReasons.push(`【${providerName}】未配置 API Key`);
      continue;
    }
    hasUsableKey = true;

    const models = normalizeModelCandidates(
      candidateProvider.models.length > 0 ? candidateProvider.models : [candidateProvider.model]
    );
    if (models.length === 0) {
      failedReasons.push(`【${providerName}】未配置可用模型`);
      continue;
    }

    for (let modelIndex = 0; modelIndex < models.length; modelIndex += 1) {
      const candidateModel = models[modelIndex];
      try {
        const result = await generateWithOpenAiCompatibleApi({
          apiKey: candidateProvider.apiKey,
          baseUrl: candidateProvider.baseUrl,
          model: candidateModel,
          prompt,
          temperature: overrides?.temperature ?? settings.temperature,
          maxTokens: overrides?.maxTokens ?? settings.maxTokens,
          timeoutMs: overrides?.timeoutMs ?? settings.timeoutMs
        });
        const resolvedModel = result.model || candidateModel;
        const providerChanged = providerIndex > 0;
        const modelChanged = modelIndex > 0;
        return {
          content: result.content,
          provider: candidateProvider.provider,
          model: resolvedModel,
          fallbackUsed: providerChanged || modelChanged,
          fallbackFromProvider: providerChanged ? firstProvider.provider : undefined,
          fallbackFromModel: providerChanged
            ? firstModel
            : (modelChanged ? models[0] : undefined)
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        failedReasons.push(`【${providerName}/${candidateModel}】${summarizeError(message, 80)}`);
      }
    }
  }

  if (!hasUsableKey) {
    throw new Error('未配置可用 API Key。请先配置至少一个提供商的 API Key。');
  }

  const detail = failedReasons.join('；');
  const hint = buildApiFailureHint(detail);
  if (hint) {
    throw new Error(`所有模型均不可用：${detail}。${hint}`);
  }
  throw new Error(`所有模型均不可用：${detail}`);
}

async function revealWritingAgentView(options: {
  primaryFocusCommand: string;
  explorerFocusCommand: string;
  failMessage: string;
}): Promise<void> {
  try {
    await vscode.commands.executeCommand('workbench.view.extension.writing-agent');
    try {
      await vscode.commands.executeCommand(options.primaryFocusCommand);
    } catch {
      // 某些宿主不支持 treeView.focus，保持容器已打开即可，避免触发 view quick-open。
    }
    return;
  } catch (primaryError) {
    const primaryMessage = primaryError instanceof Error ? primaryError.message : String(primaryError);
    try {
      await vscode.commands.executeCommand('workbench.view.explorer');
      try {
        await vscode.commands.executeCommand(options.explorerFocusCommand);
      } catch {
        // Explorer 已打开时不再调用 openView，避免弹出 view 选择器。
      }
      return;
    } catch (fallbackError) {
      const fallbackMessage = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
      throw new Error(`${options.failMessage}。ActivityBar: ${primaryMessage}；Explorer: ${fallbackMessage}`);
    }
  }
}

async function createArticleFile(
  workspaceRoot: vscode.Uri,
  styleName: string,
  topic: string,
  content: string
): Promise<{ uri: vscode.Uri; relativePath: string; title: string }> {
  const safeStyle = sanitizePathSegment(styleName || '默认风格');
  const safeTopic = sanitizePathSegment(topic || '未命名主题');
  const timestamp = formatTimestamp(new Date());
  const topicMaxLength = Math.max(12, 120 - timestamp.length - 1);
  const topicPart = safeTopic.slice(0, topicMaxLength);
  const fileBase = `${topicPart}-${timestamp}`;
  const collectionDir = vscode.Uri.joinPath(workspaceRoot, COLLECTION_RELATIVE_PREFIX, safeStyle);
  await vscode.workspace.fs.createDirectory(collectionDir);

  let fileName = `${fileBase}.md`;
  let fileUri = vscode.Uri.joinPath(collectionDir, fileName);
  let index = 2;
  while (await uriExists(fileUri)) {
    fileName = `${fileBase} (${index}).md`;
    fileUri = vscode.Uri.joinPath(collectionDir, fileName);
    index += 1;
  }

  await vscode.workspace.fs.writeFile(fileUri, Buffer.from(content, 'utf-8'));
  const relativePath = fileUri.path.replace(workspaceRoot.path, '').replace(/^\//, '');
  const title = fileName.replace(/\.md$/i, '');
  return { uri: fileUri, relativePath, title };
}

function formatTimestamp(date: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function sanitizePathSegment(value: string): string {
  return value
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 60) || 'untitled';
}

function normalizeRelativePath(relativePath: string): string {
  return (relativePath || '')
    .replace(/^\/+/, '')
    .replace(/\\/g, '/')
    .replace(/\/+/g, '/')
    .trim()
    .toLowerCase();
}

function deriveTopicFromArticleTitle(title: string): string {
  const normalized = (title || '').trim();
  if (!normalized) {
    return '未命名主题';
  }
  const withoutTimestamp = normalized.replace(/-\d{8}-\d{6}( \(\d+\))?$/u, '').trim();
  return withoutTimestamp || normalized;
}

function buildMaterialExportMarkdown(styleName: string, materials: WritingMaterial[]): string {
  const sorted = [...materials].sort((a, b) => b.metadata.updatedAt.getTime() - a.metadata.updatedAt.getTime());
  const typeCounts = new Map<MaterialType, number>();
  sorted.forEach(material => {
    typeCounts.set(material.type, (typeCounts.get(material.type) || 0) + 1);
  });

  const summary = [...typeCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([type, count]) => `- ${getMaterialTypeName(type)}：${count} 条`)
    .join('\n');
  const blocks = sorted.map((material, index) => {
    return [
      `## ${index + 1}. ${material.name}`,
      '',
      `- 类型：${getMaterialTypeName(material.type)}`,
      `- 质量：${material.metadata.quality}`,
      `- 更新时间：${material.metadata.updatedAt.toLocaleString()}`,
      `- 来源：${material.metadata.source}`,
      `- 标签：${(material.tags || []).join('、') || '无'}`,
      '',
      '```text',
      material.content,
      '```',
      ''
    ].join('\n');
  }).join('\n');

  return [
    `# 素材导出 · ${styleName || '未命名风格'}`,
    '',
    `- 导出时间：${new Date().toLocaleString()}`,
    `- 导出数量：${sorted.length}`,
    '',
    '## 类型统计',
    summary || '- 无',
    '',
    '## 素材列表',
    '',
    blocks
  ].join('\n');
}

function getMaterialTypeName(type: MaterialType): string {
  switch (type) {
    case MaterialType.SENTENCE:
      return '精彩句子';
    case MaterialType.PARAGRAPH:
      return '优秀段落';
    case MaterialType.QUOTE:
      return '引用名言';
    case MaterialType.METAPHOR:
      return '比喻修辞';
    case MaterialType.OPENING:
      return '文章开头';
    case MaterialType.ENDING:
      return '文章结尾';
    case MaterialType.TRANSITION:
      return '过渡表达';
    case MaterialType.IDEA:
      return '创意观点';
    case MaterialType.STYLE_SAMPLE:
      return '风格样本';
    default:
      return '其他';
  }
}

function getParentUri(uri: vscode.Uri): vscode.Uri {
  const normalized = uri.path.replace(/\/+$/, '');
  const index = normalized.lastIndexOf('/');
  if (index <= 0) {
    return uri.with({ path: '/' });
  }
  return uri.with({ path: normalized.slice(0, index) });
}

async function uriExists(uri: vscode.Uri): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch {
    return false;
  }
}

function buildTopicWritingPrompt(
  topic: string,
  style: WritingStyle,
  styleName: string,
  materials: Array<{ name: string; content: string; type: MaterialType }>,
  outputMode: TopicOutputMode
): string {
  const tone = style.tone?.type || 'neutral';
  const sentenceLength = style.sentenceStructure?.avgLength?.toFixed(1) || '0.0';
  const richness = ((style.vocabulary?.uniqueWordRatio || 0) * 100).toFixed(1);
  const structure = style.contentStructure?.structurePatterns?.slice(0, 3).join('、') || '总分总';
  const transitions = style.contentStructure?.transitionWords?.slice(0, 8).join('、') || '按语义自然衔接';
  const favoriteWords = style.vocabulary?.favoriteWords?.slice(0, 8).join('、') || '无';
  const avoidWords = style.vocabulary?.colloquialisms?.slice(0, 6).join('、') || '无';
  const rhetoricHints = style.rhetoric?.metaphor?.slice(0, 5).join('、') || '按需使用';

  const materialBlock = materials.length > 0
    ? materials
      .slice(0, 5)
      .map((item, index) => `${index + 1}. [${item.type}] ${item.name}\n   ${item.content.slice(0, 120)}`)
      .join('\n')
    : '当前没有可参考素材，可自行补充事实、案例或细节。';

  const styleRequirements = [
    `- 风格名称：${styleName}`,
    `- 语气：${tone}`,
    `- 句长：平均约 ${sentenceLength} 字，长短句自然交替`,
    `- 结构偏好：${structure}`,
    `- 常用过渡词：${transitions}`,
    `- 偏好词汇：${favoriteWords}`,
    `- 口语词控制：${avoidWords}`,
    `- 修辞倾向：${rhetoricHints}`,
    `- 词汇丰富度参考：${richness}%`
  ].join('\n');

  if (outputMode === 'outline') {
    return [
      `请围绕主题「${topic}」输出写作提纲。`,
      '仅遵循下列写作风格要求：',
      styleRequirements,
      '',
      '输出要求：',
      '1) 至少给出 2 套不同思路的提纲（建议 2-3 套）。',
      '2) 每套提纲包含：标题、核心论点、3-5 个小节、每节可写细节提示。',
      '3) 直接输出提纲，不要解释提示词。',
      '',
      '可参考素材：',
      materialBlock
    ].join('\n');
  }

  return [
    `请围绕主题「${topic}」写一篇中文文章。`,
    '仅遵循下列写作风格要求：',
    styleRequirements,
    '',
    '输出要求：',
    '1) 内容自然连贯，避免模板化表达。',
    '2) 直接输出正文，不要解释提示词。',
    '',
    '可参考素材：',
    materialBlock
  ].join('\n');
}

function buildSelectionRewritePrompt(
  selectedText: string,
  instruction: string,
  style?: WritingStyle,
  styleName?: string
): string {
  const tone = style?.tone?.type || 'neutral';
  const sentenceLength = style?.sentenceStructure?.avgLength?.toFixed(1) || '0.0';
  const transitions = style?.contentStructure?.transitionWords?.slice(0, 6).join('、') || '（按原文需要）';
  const favoriteWords = style?.vocabulary?.favoriteWords?.slice(0, 8).join('、') || '（按原文需要）';

  return [
    '你是中文写作编辑助手。',
    '任务：根据用户指令改写给定文本，并直接输出改写后的最终文本。',
    '',
    `用户指令：${instruction}`,
    `目标风格：${styleName || '当前文稿风格'}`,
    `语气参考：${tone}；平均句长参考：${sentenceLength} 字`,
    `过渡词参考：${transitions}`,
    `高频词参考：${favoriteWords}`,
    '',
    '输出要求：',
    '1) 只输出改写后的正文，不要解释，不要加“说明/分析/注释”。',
    '2) 不要使用代码块标记（```）。',
    '3) 改写后可直接替换原文，保持语义连贯。',
    '',
    '待改写文本：',
    '<<<TEXT',
    selectedText,
    'TEXT>>>'
  ].join('\n');
}

function normalizeRewriteOutput(text: string): string {
  let output = text.trim();
  output = output.replace(/^```(?:markdown|md|text)?\s*/i, '').replace(/```$/i, '').trim();
  if (!output) {
    return '';
  }
  return output;
}

function normalizeTextForRewriteCompare(text: string): string {
  return (text || '')
    .trim()
    .toLowerCase()
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/\s+/g, '')
    .replace(/[，。！？、；：,.!?;:()（）【】\[\]“”"'<>\-—]/g, '');
}

function textDiceSimilarity(left: string, right: string): number {
  if (!left && !right) {
    return 1;
  }
  if (!left || !right) {
    return 0;
  }
  if (left === right) {
    return 1;
  }
  const buildBigrams = (value: string): Map<string, number> => {
    const map = new Map<string, number>();
    if (value.length < 2) {
      map.set(value, 1);
      return map;
    }
    for (let i = 0; i < value.length - 1; i += 1) {
      const key = value.slice(i, i + 2);
      map.set(key, (map.get(key) || 0) + 1);
    }
    return map;
  };
  const leftBigrams = buildBigrams(left);
  const rightBigrams = buildBigrams(right);
  let intersection = 0;
  leftBigrams.forEach((leftCount, key) => {
    const rightCount = rightBigrams.get(key) || 0;
    intersection += Math.min(leftCount, rightCount);
  });
  const leftSize = Array.from(leftBigrams.values()).reduce((sum, count) => sum + count, 0);
  const rightSize = Array.from(rightBigrams.values()).reduce((sum, count) => sum + count, 0);
  if (leftSize + rightSize === 0) {
    return 0;
  }
  return (2 * intersection) / (leftSize + rightSize);
}

function isRewriteEffectivelyUnchanged(original: string, rewritten: string): boolean {
  const left = normalizeTextForRewriteCompare(original);
  const right = normalizeTextForRewriteCompare(rewritten);
  if (!left || !right) {
    return false;
  }
  if (left === right) {
    return true;
  }
  if (left.length >= 24 && right.includes(left)) {
    return true;
  }
  if (right.length >= 24 && left.includes(right)) {
    return true;
  }
  return textDiceSimilarity(left, right) >= 0.985;
}

function buildStrictRewritePrompt(
  basePrompt: string,
  originalText: string
): string {
  return [
    basePrompt,
    '',
    '强制改写要求：',
    '1) 严禁直接复述原文；不得连续复用原文 12 个以上字符。',
    '2) 必须重组句式与表达方式，保证和原文有明显差异。',
    '3) 保持原意，但请输出可直接替换的改写文本。',
    '',
    '原文（再次提供，仅用于对照）：',
    '<<<ORIGINAL',
    originalText,
    'ORIGINAL>>>'
  ].join('\n');
}

function buildArticleDraft(
  topic: string,
  styleName: string,
  aiPrompt: string
): string {
  const createdAt = new Date().toLocaleString();
  return [
    `# ${topic}`,
    '',
    `> 风格：${styleName}`,
    `> 创建时间：${createdAt}`,
    '> 状态：草稿（支持使用 Cursor/CodeArts 的 AI 能力续写）',
    '',
    '## 写作说明',
    '在“正文”区域写作或改写。可执行“写作助手: 使用 IDE AI 续写当前文稿”，或直接在编辑器使用 Agent/Codex/补全。',
    '',
    '## AI 提示词（可粘贴到宿主 AI）',
    '```text',
    aiPrompt,
    '```',
    '',
    '## 正文（追加到文末）',
    '',
    '（生成内容会追加到本文件末尾，便于持续迭代。）'
  ].join('\n');
}

function focusArticleBody(editor: vscode.TextEditor): void {
  const lastLine = editor.document.lineCount - 1;
  const lastChar = editor.document.lineAt(lastLine).text.length;
  const position = new vscode.Position(lastLine, lastChar);
  editor.selection = new vscode.Selection(position, position);
  editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenter);
}

async function ensureEditableEditor(editor: vscode.TextEditor): Promise<vscode.TextEditor> {
  if (!editor.document.isClosed) {
    return editor;
  }
  const reopened = await vscode.workspace.openTextDocument(editor.document.uri);
  return vscode.window.showTextDocument(reopened, vscode.ViewColumn.One);
}

async function appendGeneratedDraftToEnd(
  editor: vscode.TextEditor,
  bodyText: string,
  conclusionText: string,
  style: WritingStyle,
  topic: string,
  sourceLabel: string
): Promise<void> {
  const body = dedupeParagraphBlocks(
    enforceStyleConsistency(
      humanizeZh(bodyText, { style, topic, strength: 0.92 }),
      style
    )
  );
  const rawConclusion = dedupeParagraphBlocks(
    enforceStyleConsistency(
      humanizeZh(conclusionText, { style, topic, strength: 0.82 }),
      style
    )
  );
  const conclusion = chooseConclusion(body, rawConclusion || conclusionText, topic);
  const stamp = new Date().toLocaleString();
  const block = [
    '',
    '---',
    '',
    `## 正文追加（${sourceLabel} · ${stamp}）`,
    '',
    body.trim(),
    '',
    '### 结语',
    '',
    conclusion.trim(),
    ''
  ].join('\n');

  const targetEditor = await ensureEditableEditor(editor);
  const lastLine = targetEditor.document.lineCount - 1;
  const lastChar = targetEditor.document.lineAt(lastLine).text.length;
  const position = new vscode.Position(lastLine, lastChar);
  const applied = await targetEditor.edit(editBuilder => {
    editBuilder.insert(position, block);
  });
  if (!applied) {
    throw new Error('写入正文失败：编辑器不可写，请重试。');
  }
  const endLine = targetEditor.document.lineCount - 1;
  const endChar = targetEditor.document.lineAt(endLine).text.length;
  const end = new vscode.Position(endLine, endChar);
  targetEditor.selection = new vscode.Selection(end, end);
  targetEditor.revealRange(new vscode.Range(end, end), vscode.TextEditorRevealType.InCenter);
}

async function appendOutlineToEnd(
  editor: vscode.TextEditor,
  outlineText: string,
  sourceLabel: string
): Promise<void> {
  const normalized = normalizeOutlineOutput(outlineText);
  const stamp = new Date().toLocaleString();
  const block = [
    '',
    '---',
    '',
    `## 写作提纲（${sourceLabel} · ${stamp}）`,
    '',
    normalized || '（提纲为空，请重试或调整主题）',
    ''
  ].join('\n');

  const targetEditor = await ensureEditableEditor(editor);
  const lastLine = targetEditor.document.lineCount - 1;
  const lastChar = targetEditor.document.lineAt(lastLine).text.length;
  const position = new vscode.Position(lastLine, lastChar);
  const applied = await targetEditor.edit(editBuilder => {
    editBuilder.insert(position, block);
  });
  if (!applied) {
    throw new Error('写入提纲失败：编辑器不可写，请重试。');
  }
}

function generateLocalTopicOutlines(
  topic: string,
  style: WritingStyle,
  materials: Array<{ name: string; content: string; type: MaterialType }>
): string {
  const transitions = style.contentStructure?.transitionWords?.slice(0, 3).join('、') || '自然推进';
  const favoriteWords = style.vocabulary?.favoriteWords?.slice(0, 3).join('、') || '核心概念';
  const reference = materials[0]?.content
    ? materials[0].content.replace(/\s+/g, ' ').slice(0, 80)
    : '补充一个具体案例（含时间、动作、结果）';

  return [
    '### 提纲方案 A｜问题拆解型',
    `- 标题建议：把「${topic}」讲透：从问题边界到执行动作`,
    `- 核心论点：围绕「${topic}」先明确边界，再给路径，最后定义评估标准。`,
    '- 小节 1：主题背景与现实约束（谁在什么场景遇到什么问题）。',
    '- 小节 2：关键矛盾拆解（目标、资源、风险三条线）。',
    `- 小节 3：执行方案（按“步骤-证据-结果”推进，过渡词可用：${transitions}）。`,
    '- 小节 4：复盘与迭代（下一步动作 + 指标）。',
    `- 可写细节提示：优先嵌入词汇「${favoriteWords}」。`,
    '',
    '### 提纲方案 B｜案例推进型',
    `- 标题建议：用一个真实案例理解「${topic}」`,
    `- 核心论点：先讲案例，再抽象方法，最后回到可执行清单。`,
    '- 小节 1：案例起点（时间、角色、目标）。',
    '- 小节 2：关键动作与取舍（为何这样做，放弃了什么）。',
    '- 小节 3：结果与偏差（成效、问题、反思）。',
    '- 小节 4：迁移方法（在其他场景如何复用）。',
    `- 可写细节提示：可参考素材片段「${reference}」。`
  ].join('\n');
}

function normalizeOutlineOutput(text: string): string {
  let output = text.trim();
  output = output.replace(/^```(?:markdown|md|text)?\s*/i, '').replace(/```$/i, '').trim();
  return output;
}

function countOutlineVariants(text: string): number {
  const headingMatches = text.match(/(?:^|\n)#{1,4}\s*(?:提纲|方案|思路)[^\n]*/g);
  if (headingMatches && headingMatches.length > 0) {
    return headingMatches.length;
  }
  const fallbackMatches = text.match(/(?:^|\n)(?:方案|思路)\s*[A-Za-z一二三四五六七八九十]/g);
  return fallbackMatches ? fallbackMatches.length : 0;
}

function buildSupplementOutlineVariant(
  topic: string,
  style: WritingStyle,
  materials: Array<{ name: string; content: string; type: MaterialType }>
): string {
  const structure = style.contentStructure?.structurePatterns?.slice(0, 2).join('、') || '总分总';
  const materialName = materials[0]?.name || '一个真实样本';
  return [
    '### 提纲方案 C｜观点对照型',
    `- 标题建议：「${topic}」的三种常见路径与适用边界`,
    `- 核心论点：同一主题在不同约束下应采取不同策略，不必追求单一答案。`,
    '- 小节 1：路径一（适用条件 + 优势 + 风险）。',
    '- 小节 2：路径二（适用条件 + 优势 + 风险）。',
    '- 小节 3：路径三（适用条件 + 优势 + 风险）。',
    `- 小节 4：决策建议（结合风格结构偏好：${structure}）。`,
    `- 可写细节提示：对照素材「${materialName}」补齐证据。`
  ].join('\n');
}

function ensureOutlineVariants(
  text: string,
  topic: string,
  style: WritingStyle,
  materials: Array<{ name: string; content: string; type: MaterialType }>
): string {
  const normalized = normalizeOutlineOutput(text);
  if (!normalized) {
    return generateLocalTopicOutlines(topic, style, materials);
  }
  if (countOutlineVariants(normalized) >= 2) {
    return normalized;
  }
  const supplement = buildSupplementOutlineVariant(topic, style, materials);
  return `${normalized}\n\n${supplement}`.trim();
}

function looksLikeArticleContent(text: string): boolean {
  if (isLikelyPromptTemplate(text)) {
    return false;
  }
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (normalized.length < 160) {
    return false;
  }
  const paragraphCount = text.split(/\n{2,}/).filter(item => item.trim().length > 0).length;
  const sentenceLike = (text.match(/[。！？!?]/g) || []).length;
  return paragraphCount >= 2 || sentenceLike >= 4;
}

function isLikelyPromptTemplate(text: string): boolean {
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (!normalized) {
    return false;
  }

  const startsLikePrompt =
    /^你是\s*IDE\s*内的写作助手/.test(normalized)
    || /^请围绕主题「.+」(?:写一篇中文文章|输出写作提纲)/.test(normalized);
  const containsPromptSections =
    normalized.includes('可参考素材') &&
    (
      (normalized.includes('仅遵循下列写作风格要求') && normalized.includes('输出要求'))
      || (normalized.includes('硬性要求') && normalized.includes('风格名：'))
    );

  const assignmentLikeLines = [
    '主题：',
    '风格名：',
    '语气：',
    '结构偏好：',
    '建议过渡词：',
    '高频偏好词：',
    '风格名称：',
    '句长：',
    '常用过渡词：',
    '偏好词汇：'
  ].filter(item => normalized.includes(item)).length;

  return startsLikePrompt || containsPromptSections || assignmentLikeLines >= 4;
}

async function applyAiResultToDraft(
  editor: vscode.TextEditor,
  aiText: string,
  style: WritingStyle,
  topic: string
): Promise<void> {
  const parsed = parseAiArticle(aiText);
  await appendGeneratedDraftToEnd(editor, parsed.body, parsed.conclusion, style, topic, 'AI 输出');
}

function enforceStyleConsistency(text: string, style: WritingStyle): string {
  let output = text.trim();
  if (!output) {
    return output;
  }

  output = enforceToneVocabulary(output, style.tone.type);
  output = enforceTransitionWords(output, style.contentStructure.transitionWords || []);
  output = alignSentenceLengthByStyle(output, style.sentenceStructure.avgLength || 0);

  return output.replace(/\n{3,}/g, '\n\n').trim();
}

function enforceToneVocabulary(text: string, tone: WritingStyle['tone']['type']): string {
  if (tone === 'serious' || tone === 'formal') {
    return text
      .replace(/(?:哈哈|呵呵|笑死了?)/g, '')
      .replace(/(挺|很|非常)重要/g, '更关键')
      .replace(/有点/g, '略有')
      .replace(/[!！]{2,}/g, '。');
  }
  if (tone === 'casual') {
    return text
      .replace(/需要注意的是/g, '要注意的是')
      .replace(/在此基础上/g, '顺着这个思路')
      .replace(/由此可见/g, '所以能看出来');
  }
  return text;
}

function enforceTransitionWords(text: string, transitions: string[]): string {
  const words = transitions.filter(item => item && item.trim().length >= 1).slice(0, 4);
  if (words.length === 0) {
    return text;
  }

  const paragraphs = text
    .split(/\n{2,}/)
    .map(item => item.trim())
    .filter(Boolean);
  if (paragraphs.length < 2) {
    return text;
  }

  const hasTransitionOpening = paragraphs.some(paragraph =>
    words.some(word => paragraph.startsWith(`${word}，`) || paragraph.startsWith(`${word},`) || paragraph.startsWith(word))
  );
  if (hasTransitionOpening) {
    return text;
  }

  const result = [...paragraphs];
  for (let i = 1; i < result.length; i += 1) {
    const paragraph = result[i];
    if (paragraph.startsWith('#')) {
      continue;
    }
    result[i] = `${words[0]}，${paragraph}`;
    break;
  }
  return result.join('\n\n');
}

function alignSentenceLengthByStyle(text: string, targetAvgLength: number): string {
  if (targetAvgLength <= 0) {
    return text;
  }

  if (targetAvgLength < 18) {
    return text.replace(/([^。！？!?]{30,})，/g, '$1。');
  }

  if (targetAvgLength > 30) {
    return text
      .replace(/。(\n{2,}[^#\n]{0,14}[。！？!?])/g, '，并且$1')
      .replace(/。\n([^\n#]{0,14}[。！？!?])/g, '，$1');
  }

  return text;
}

function parseAiArticle(aiText: string): { body: string; conclusion: string } {
  let text = aiText.trim();
  text = text.replace(/^```(?:markdown|md|text)?\s*/i, '').replace(/```$/i, '').trim();
  text = text.replace(/^#\s+.+\n+/i, '').trim();

  const conclusionHeading = text.match(/(?:^|\n)#{1,3}\s*(结语|总结|结论)\s*\n/i);
  if (conclusionHeading && typeof conclusionHeading.index === 'number') {
    const index = conclusionHeading.index;
    const body = text.slice(0, index).trim();
    const conclusion = text.slice(index).replace(/(?:^|\n)#{1,3}\s*(结语|总结|结论)\s*\n/i, '').trim();
    return {
      body: body || text,
      conclusion: conclusion || ''
    };
  }

  const paragraphs = splitParagraphs(text);
  if (paragraphs.length >= 2) {
    const conclusion = paragraphs.pop() || '';
    const body = paragraphs.join('\n\n').trim();
    return {
      body: body || text,
      conclusion: conclusion || ''
    };
  }

  return {
    body: text,
    conclusion: ''
  };
}

function splitParagraphs(text: string): string[] {
  return text
    .split(/\n{2,}/)
    .map(item => item.trim())
    .filter(Boolean);
}

function normalizeParagraphFingerprint(text: string): string {
  return text
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/\s+/g, '')
    .replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '')
    .toLowerCase();
}

function isNearDuplicateParagraph(a: string, b: string): boolean {
  const left = normalizeParagraphFingerprint(a);
  const right = normalizeParagraphFingerprint(b);
  if (!left || !right) {
    return false;
  }
  if (left === right) {
    return true;
  }
  const [shorter, longer] = left.length <= right.length ? [left, right] : [right, left];
  return shorter.length >= 24 && longer.includes(shorter);
}

function dedupeParagraphBlocks(text: string): string {
  const paragraphs = splitParagraphs(text);
  if (paragraphs.length <= 1) {
    return text.trim();
  }
  const unique: string[] = [];
  for (const paragraph of paragraphs) {
    if (unique.some(existing => isNearDuplicateParagraph(existing, paragraph))) {
      continue;
    }
    unique.push(paragraph);
  }
  return unique.join('\n\n').trim();
}

function buildDefaultConclusion(topic: string): string {
  return `下一步建议：围绕${topic}先列出 3 个可执行动作，并在一周内完成一次复盘。`;
}

function chooseConclusion(body: string, conclusion: string, topic: string): string {
  const bodyParagraphs = splitParagraphs(body);
  const conclusionCandidate = splitParagraphs(conclusion)[0] || conclusion.trim();
  if (!conclusionCandidate) {
    return buildDefaultConclusion(topic);
  }
  if (bodyParagraphs.length === 0) {
    return conclusionCandidate;
  }
  if (bodyParagraphs.some(item => isNearDuplicateParagraph(item, conclusionCandidate))) {
    return buildDefaultConclusion(topic);
  }
  return conclusionCandidate;
}

function resolveRewriteOverlayAnchor(
  editor: vscode.TextEditor,
  selectionRange: vscode.Range
): vscode.Position {
  const visible = editor.visibleRanges[0];
  const hasVisible = Boolean(visible);
  const linesAbove = hasVisible ? Math.max(0, selectionRange.start.line - visible.start.line) : selectionRange.start.line;
  const linesBelow = hasVisible
    ? Math.max(0, visible.end.line - selectionRange.end.line)
    : Math.max(0, editor.document.lineCount - selectionRange.end.line - 1);
  const preferAbove = linesAbove >= linesBelow;
  return preferAbove ? selectionRange.start : selectionRange.end;
}

function clearRewriteApplyDecorations(): void {
  if (!rewriteApplyDecorationType) {
    return;
  }
  for (const editor of vscode.window.visibleTextEditors) {
    editor.setDecorations(rewriteApplyDecorationType, []);
  }
}

function sanitizeForMarkdownCodeBlock(text: string): string {
  return (text || '').replace(/```/g, '``\\`');
}

function resolvePendingRewriteApplyDecision(mode: RewriteApplyMode | undefined): void {
  const pending = pendingRewriteApplyDecision;
  if (!pending) {
    return;
  }
  pendingRewriteApplyDecision = null;
  clearTimeout(pending.timeoutHandle);
  clearRewriteApplyDecorations();
  pending.resolve(mode);
}

async function pickRewriteApplyModeWithPreview(
  editor: vscode.TextEditor,
  selectionRange: vscode.Range,
  rewritten: string
): Promise<'replace' | 'insertAfter' | undefined> {
  resolvePendingRewriteApplyDecision(undefined);
  if (!rewriteApplyDecorationType) {
    return 'replace';
  }

  const previewLimit = 1800;
  const previewText = rewritten.length > previewLimit
    ? `${rewritten.slice(0, previewLimit)}\n\n（预览已截断，全文 ${rewritten.length} 字）`
    : rewritten;
  const previewBlock = sanitizeForMarkdownCodeBlock(previewText);
  const markdown = new vscode.MarkdownString([
    `已生成改写文本（${rewritten.length} 字）。`,
    '',
    '预览：',
    '```text',
    previewBlock,
    '```',
    '',
    '[替换选区](command:writingAgent.rewriteApplyReplace) · [插入到选区后](command:writingAgent.rewriteApplyInsertAfter) · [取消](command:writingAgent.rewriteApplyCancel)'
  ].join('\n'));
  markdown.isTrusted = true;
  markdown.supportHtml = false;

  editor.setDecorations(rewriteApplyDecorationType, [
    {
      range: selectionRange,
      hoverMessage: markdown
    }
  ]);

  const anchor = resolveRewriteOverlayAnchor(editor, selectionRange);
  void vscode.commands.executeCommand('editor.action.showHover', {
    lineNumber: anchor.line + 1,
    pos: anchor.character + 1
  });

  return new Promise<'replace' | 'insertAfter' | undefined>(resolve => {
    const timeoutHandle = setTimeout(() => {
      resolvePendingRewriteApplyDecision(undefined);
      notifyWarn('改写应用选择已超时，未写入文稿。可重新执行“AI编辑选中文本”。');
    }, REWRITE_APPLY_OVERLAY_TIMEOUT_MS);
    pendingRewriteApplyDecision = {
      editorUri: editor.document.uri.toString(),
      selectionRange,
      resolve,
      timeoutHandle
    };
  });
}

function escapeHtml(value: string): string {
  return (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatPercent(value: number, digits = 1): string {
  return `${(Math.max(0, Math.min(1, Number.isFinite(value) ? value : 0)) * 100).toFixed(digits)}%`;
}

function renderTags(words: string[]): string {
  if (!words || words.length === 0) {
    return '<span class="muted">暂无</span>';
  }
  return words.map(word => `<span class="tag">${escapeHtml(word)}</span>`).join('');
}

function filterMeaningfulAnalysisWords(words: string[]): string[] {
  const stopwords = new Set([
    '标题', '正文', '核心论点', '适用条件', '优势', '风险', '创建时间', '状态', '草稿', '新风格',
    '的', '了', '和', '是', '在', '与', '并', '而', '也', '及', 'the', 'and', 'for'
  ]);
  const seen = new Set<string>();
  const result: string[] = [];
  for (const word of words || []) {
    const normalized = (word || '').trim();
    if (!normalized) {
      continue;
    }
    if (stopwords.has(normalized) || stopwords.has(normalized.toLowerCase())) {
      continue;
    }
    if (/^\d+([.)、:：-]\d+)*$/.test(normalized)) {
      continue;
    }
    if (/^[#*_`~\-+=|\\/]+$/.test(normalized)) {
      continue;
    }
    const core = normalized.replace(/[^\p{L}\p{N}\u4E00-\u9FFF]/gu, '');
    if (core.length < 2) {
      continue;
    }
    const dedup = normalized.toLowerCase();
    if (seen.has(dedup)) {
      continue;
    }
    seen.add(dedup);
    result.push(normalized);
  }
  return result.slice(0, 40);
}

function buildStyleFingerprints(style: WritingStyle): string[] {
  const fingerprints: string[] = [];
  const transitions = (style.contentStructure?.transitionWords || []).slice(0, 4);
  const sentencePatterns = (style.sentenceStructure?.sentencePatterns || []).slice(0, 4);
  const punctuation = (style.punctuation?.specialPunctuation || []).slice(0, 3);
  const favoriteWords = filterMeaningfulAnalysisWords(style.vocabulary?.favoriteWords || []).slice(0, 6);

  transitions.forEach(item => fingerprints.push(`标志性过渡词：${item}`));
  sentencePatterns.forEach(item => fingerprints.push(`句式模板：${item}`));
  punctuation.forEach(item => fingerprints.push(`个性化标点：${item}`));
  for (let i = 0; i + 1 < favoriteWords.length && i < 3; i += 1) {
    fingerprints.push(`固定搭配候选：${favoriteWords[i]} + ${favoriteWords[i + 1]}`);
  }
  return Array.from(new Set(fingerprints)).slice(0, 10);
}

function getStyleAnalysisWebview(style: WritingStyle): string {
  const favoriteWords = filterMeaningfulAnalysisWords(style.vocabulary?.favoriteWords || []);
  const terminology = filterMeaningfulAnalysisWords(style.vocabulary?.terminology || []);
  const colloquialisms = filterMeaningfulAnalysisWords(style.vocabulary?.colloquialisms || []);
  const transitionWords = filterMeaningfulAnalysisWords(style.contentStructure?.transitionWords || []);
  const structurePatterns = filterMeaningfulAnalysisWords(style.contentStructure?.structurePatterns || []);
  const sentencePatterns = filterMeaningfulAnalysisWords(style.sentenceStructure?.sentencePatterns || []);
  const connectives = filterMeaningfulAnalysisWords(style.logicType?.connectives || []);
  const toneMarkers = filterMeaningfulAnalysisWords(style.tone?.markers || []);
  const temporalMarkers = filterMeaningfulAnalysisWords(style.timePerspective?.temporalMarkers || []);
  const emotionalWords = filterMeaningfulAnalysisWords(style.emotionalArc?.emotionalWords || []);
  const specialPunctuation = filterMeaningfulAnalysisWords(style.punctuation?.specialPunctuation || []);
  const fingerprints = buildStyleFingerprints(style);

  const sentenceLength = style.sentenceStructure?.avgLength || 0;
  const uniqueWordRatio = style.vocabulary?.uniqueWordRatio || 0;
  const complexRatio = style.sentenceStructure?.complexSentenceRatio || 0;
  const interactionLevel = style.audienceAwareness?.engagementLevel || 0;
  const certainty = 1 - Math.min(1, (style.tone?.type === 'neutral' ? 0.4 : 0.2) + (style.tone?.intensity || 0) * 0.3);

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <style>
        body {
          padding: 20px;
          font-family: var(--vscode-font-family);
          color: var(--vscode-editor-foreground);
          line-height: 1.6;
        }
        h2, h3 { margin: 0 0 10px; }
        .section {
          margin: 18px 0;
          padding: 14px;
          border: 1px solid var(--vscode-panel-border);
          border-radius: 8px;
          background: var(--vscode-editor-background);
        }
        .grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
          gap: 10px 16px;
        }
        .metric {
          margin: 6px 0;
        }
        .metric-name {
          color: var(--vscode-descriptionForeground);
          margin-right: 4px;
        }
        .metric-bar {
          height: 8px;
          background: var(--vscode-progressBar-background);
          border-radius: 4px;
          margin-top: 4px;
        }
        .tag {
          display: inline-block;
          background: var(--vscode-button-secondaryBackground);
          color: var(--vscode-button-secondaryForeground);
          padding: 2px 8px;
          margin: 2px 4px 2px 0;
          border-radius: 999px;
          font-size: 12px;
        }
        .muted {
          color: var(--vscode-descriptionForeground);
        }
        ul {
          margin: 6px 0 0 18px;
          padding: 0;
        }
      </style>
    </head>
    <body>
      <h2>写作风格分析</h2>

      <div class="section">
        <h3>1. 词汇层面</h3>
        <div class="grid">
          <div class="metric">
            <span class="metric-name">用词偏好：</span>
            <span>平均词长 ${style.vocabulary?.avgWordLength?.toFixed(1) || '0.0'}，词汇丰富度 ${formatPercent(uniqueWordRatio)}</span>
            <div class="metric-bar" style="width:${Math.min(100, uniqueWordRatio * 100)}%"></div>
          </div>
          <div class="metric">
            <span class="metric-name">情感色彩：</span>
            <span>${escapeHtml(style.emotionalArc?.overallTrend || 'neutral')}（情感词 ${emotionalWords.length}）</span>
          </div>
          <div class="metric">
            <span class="metric-name">专业术语密度：</span>
            <span>${terminology.length} 项</span>
          </div>
          <div class="metric">
            <span class="metric-name">缩略与简称：</span>
            <span>${terminology.filter(item => /[A-Z]{2,}|[0-9]/.test(item)).length} 项</span>
          </div>
        </div>
        <div><strong>常用词汇（已过滤无意义词）</strong><br/>${renderTags(favoriteWords)}</div>
        <div><strong>术语/黑话</strong><br/>${renderTags(terminology)}</div>
        <div><strong>口语表达</strong><br/>${renderTags(colloquialisms)}</div>
      </div>

      <div class="section">
        <h3>2. 句法层面</h3>
        <div class="grid">
          <div class="metric">
            <span class="metric-name">句子长度分布：</span>
            <span>平均 ${sentenceLength.toFixed(1)} 字</span>
            <div class="metric-bar" style="width:${Math.min(sentenceLength * 2, 100)}%"></div>
          </div>
          <div class="metric">
            <span class="metric-name">句式结构：</span>
            <span>复杂句占比 ${formatPercent(complexRatio)}，疑问句 ${formatPercent(style.sentenceStructure?.questionRatio || 0)}，感叹句 ${formatPercent(style.sentenceStructure?.exclamationRatio || 0)}</span>
          </div>
          <div class="metric">
            <span class="metric-name">修辞手法：</span>
            <span>比喻 ${style.rhetoric?.metaphor?.length || 0}，排比 ${style.rhetoric?.parallelism?.length || 0}，引用 ${style.rhetoric?.quotes?.length || 0}</span>
          </div>
          <div class="metric">
            <span class="metric-name">标点运用：</span>
            <span>逗号频率 ${(style.punctuation?.commaFrequency || 0).toFixed(3)}，句号频率 ${(style.punctuation?.periodFrequency || 0).toFixed(3)}</span>
          </div>
        </div>
        <div><strong>常见句式</strong><br/>${renderTags(sentencePatterns)}</div>
        <div><strong>特殊标点</strong><br/>${renderTags(specialPunctuation)}</div>
      </div>

      <div class="section">
        <h3>3. 语篇层面</h3>
        <div class="grid">
          <div class="metric"><span class="metric-name">论证结构：</span><span>${escapeHtml(style.logicType?.type || 'narrative')}（论据风格：${escapeHtml(style.logicType?.evidenceStyle || 'general')}）</span></div>
          <div class="metric"><span class="metric-name">段落组织：</span><span>段落均长 ${(style.contentStructure?.paragraphLength || 0).toFixed(1)}，开头 ${escapeHtml(style.contentStructure?.openingStyle || '标准开头')}，结尾 ${escapeHtml(style.contentStructure?.endingStyle || '标准结尾')}</span></div>
          <div class="metric"><span class="metric-name">信息密度：</span><span>密度 ${formatPercent(style.informationDensity?.density || 0)}，复杂度 ${formatPercent(style.informationDensity?.complexity || 0)}</span></div>
          <div class="metric"><span class="metric-name">叙事视角：</span><span>${escapeHtml(style.spacePerspective?.viewpoint || 'third-person')}（时间主导：${escapeHtml(style.timePerspective?.dominant || 'present')}）</span></div>
        </div>
        <div><strong>结构模式</strong><br/>${renderTags(structurePatterns)}</div>
        <div><strong>逻辑连接词</strong><br/>${renderTags(connectives)}</div>
      </div>

      <div class="section">
        <h3>4. 语气与视角</h3>
        <div class="grid">
          <div class="metric"><span class="metric-name">人称与视角：</span><span>${escapeHtml(style.spacePerspective?.viewpoint || 'third-person')}，叙事声音 ${escapeHtml(style.narrativeVoice?.voice || 'neutral')}</span></div>
          <div class="metric"><span class="metric-name">确定性程度：</span><span>${formatPercent(Math.max(0, Math.min(1, certainty)))}（估计）</span></div>
          <div class="metric">
            <span class="metric-name">互动性：</span><span>${formatPercent(interactionLevel)}（受众意识）</span>
            <div class="metric-bar" style="width:${Math.min(100, interactionLevel * 100)}%"></div>
          </div>
          <div class="metric"><span class="metric-name">权力距离：</span><span>${escapeHtml(style.culturalContext?.register || 'informal')} / ${escapeHtml(style.narrativeVoice?.stance || '平等探讨')}</span></div>
        </div>
        <div><strong>语气标记词</strong><br/>${renderTags(toneMarkers)}</div>
        <div><strong>时间标记词</strong><br/>${renderTags(temporalMarkers)}</div>
      </div>

      <div class="section">
        <h3>5. 独特标识（指纹特征）</h3>
        ${fingerprints.length > 0
    ? `<ul>${fingerprints.map(item => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
    : '<span class="muted">暂无稳定指纹，可在更多样本后重试分析。</span>'}
        <div style="margin-top:8px;"><strong>标志性过渡词</strong><br/>${renderTags(transitionWords)}</div>
      </div>
    </body>
    </html>
  `;
}

export function deactivate(): void {
  console.log('写作 Agent 插件已停用');
}
