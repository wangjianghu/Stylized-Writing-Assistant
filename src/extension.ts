import * as vscode from 'vscode';
import * as mammoth from 'mammoth';
import { MaterialType, StyleProfile, WritingStyle } from './core/types';
import { StyleEngine } from './core/styleEngine';
import { MaterialExtractor } from './core/materialManager/extractor';
import { MaterialRepository } from './core/materialManager/repository';
import { StyleRepository } from './core/styleManager/repository';
import { ArticleRecord, ArticleRepository } from './core/articleManager/repository';
import { generateStyleAlignedDraft } from './core/articleManager/generator';
import { humanizeZh } from './core/articleManager/humanizerZh';
import {
  generateWithOpenAiCompatibleApi,
  isModelNotSupportedApiError
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

type ApiProvider = 'iflow' | 'kimi' | 'deepseek' | 'minimax' | 'qwen' | 'custom';
type TopicOutputMode = 'article' | 'outline';
type AiExecutionChannel = 'api' | 'ide-chat';

interface ProviderPreset {
  provider: ApiProvider;
  displayName: string;
  defaultBaseUrl: string;
  defaultModel: string;
}

interface RuntimeApiSettings {
  enabled: boolean;
  provider: ApiProvider;
  apiKey: string;
  baseUrl: string;
  model: string;
  temperature: number;
  maxTokens: number;
  timeoutMs: number;
}

interface ApiGenerationResult {
  content: string;
  model: string;
  fallbackUsed: boolean;
  fallbackFromModel?: string;
}

interface ApiValidationState {
  fingerprint: string;
  valid: boolean;
  message: string;
  checkedAt: string;
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

const LEGACY_IFLOW_SECRET_KEY = 'writingAgent.iflow.apiKey';
const API_VALIDATION_STATE_KEY = 'writingAgent.api.validation.v1';
const LEGACY_COLLECTION_RELATIVE_PREFIX = '.writing-agent/文集';
const COLLECTION_RELATIVE_PREFIX = '文集';
const STATUS_MESSAGE_TIMEOUT_MS = 5000;
const SHOW_ASSISTANT_DEBOUNCE_MS = 1200;
let showAssistantViewLockUntil = 0;
let showAssistantViewsInFlight: Promise<void> | null = null;
const PROVIDER_PRESETS: Record<ApiProvider, ProviderPreset> = {
  iflow: {
    provider: 'iflow',
    displayName: '心流',
    defaultBaseUrl: 'https://apis.iflow.cn/v1',
    defaultModel: 'Qwen/Qwen3-8B'
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

  const styleEngine = new StyleEngine();
  const materialExtractor = new MaterialExtractor();
  const materialRepository = new MaterialRepository(context);
  const styleRepository = new StyleRepository(context);
  const articleRepository = new ArticleRepository(context);
  const quickEntryProvider = new QuickEntryProvider(async () => {
    const settings = await resolveApiSettings(context);
    if (!settings || !settings.enabled) {
      return { configured: false, text: '已关闭', iconId: 'circle-slash' };
    }
    if (settings.apiKey) {
      const providerName = PROVIDER_PRESETS[settings.provider].displayName;
      const fingerprint = buildApiValidationFingerprint(settings);
      const validation = loadApiValidationState(context);
      if (!validation || validation.fingerprint !== fingerprint) {
        return { configured: true, text: `已配置·${providerName}/${settings.model}·未测试`, iconId: 'warning' };
      }
      if (validation.valid) {
        return { configured: true, text: `已配置·${providerName}/${settings.model}·有效`, iconId: 'check' };
      }
      const brief = summarizeError(validation.message, 36);
      return { configured: true, text: `已配置·${providerName}/${settings.model}·无效（${brief}）`, iconId: 'error' };
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
      const root = getPrimaryWorkspaceRootUri();
      if (!root) {
        return null;
      }
      return vscode.Uri.joinPath(root, ...relativePath.split('/'));
    }
  );

  const activeProfile = styleRepository.getActiveProfile();
  if (activeProfile) {
    styleEngine.setCurrentStyle(activeProfile.style);
  } else {
    styleEngine.resetStyle();
  }

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
  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration(event => {
      if (event.affectsConfiguration('writingAgent.api') || event.affectsConfiguration('writingAgent.iflow')) {
        quickEntryProvider.refresh();
      }
    })
  );
  void (async () => {
    await migrateArticleCollectionToVisibleRoot(articleRepository, articleCollectionProvider);
    await syncArticleCollectionWithDisk(articleRepository, articleCollectionProvider);
    articleCollectionProvider.refresh();
  })();

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
        const root = getPrimaryWorkspaceRootUri();
        if (!root) {
          throw new Error('请先在开发宿主中打开工作区，再生成文章。');
        }

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
        } else if (apiSettings?.enabled && apiSettings.apiKey) {
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
            const providerName = PROVIDER_PRESETS[apiSettings.provider].displayName;
            sourceLabel = `${providerName} API · ${apiResult.model}`;
            aiCommand = `${apiSettings.provider}:${apiResult.model}`;
            if (apiResult.fallbackUsed && apiResult.fallbackFromModel) {
              notifyWarn(
                `当前模型“${apiResult.fallbackFromModel}”不受支持，已自动改用“${apiResult.model}”。`
              );
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

        progress.report({ increment: 100, message: '完成' });
        const message = outputMode === 'outline'
          ? `已在文集创建《${articleFile.title}》，并在文件末尾写入写作提纲。`
          : `已在文集创建《${articleFile.title}》，并在文件末尾写入正文。`;
        const actions = ['打开文集'];
        if (shouldPromptApiConfig) {
          actions.push('配置 API');
        }
        if (routedToIdeChat) {
          actions.push('写入 AI 结果');
        }
        const action = await vscode.window.showInformationMessage(message, ...actions);
        if (action === '打开文集') {
          await vscode.commands.executeCommand('writingAgent.showArticleCollection');
        } else if (action === '配置 API') {
          await vscode.commands.executeCommand('writingAgent.configureApi');
        } else if (action === '写入 AI 结果') {
          await vscode.commands.executeCommand('writingAgent.applyClipboardToDraft');
        }
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
    const workspaceRoot = getPrimaryWorkspaceRootUri();

    let materialResult: { moved: number; merged: number } | undefined;
    let articleDeleteResult: DeleteStyleArticlesResult | undefined;
    let articleMergeResult: MergeStyleArticlesResult | undefined;
    if (strategy.value === 'merge-all' && mergeTarget) {
      materialResult = await materialRepository.moveStyleMaterials(target.id, mergeTarget.id);
      articleMergeResult = await moveArticlesToStyle(target, mergeTarget, articleRepository, workspaceRoot);
    } else {
      await materialRepository.deleteStyleMaterials(target.id);
      articleDeleteResult = await removeArticlesByStyle(target.id, articleRepository, workspaceRoot);
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
    materialLibraryProvider.refresh();
  });

  registerSafeCommand(context, 'writingAgent.refreshArticleCollection', async () => {
    await syncArticleCollectionWithDisk(articleRepository, articleCollectionProvider, { notifyOnCleanup: true });
    articleCollectionProvider.refresh();
  });

  registerSafeCommand(context, 'writingAgent.deleteArticle', async (arg?: unknown) => {
    const root = getPrimaryWorkspaceRootUri();
    if (!root) {
      notifyWarn('请先打开工作区，再删除文集文章。');
      return;
    }

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

    const articleUri = vscode.Uri.joinPath(root, ...target.relativePath.split('/'));
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
    const executionChannel = resolveAiExecutionChannel();
    if (executionChannel === 'ide-chat') {
      const dispatched = await dispatchPromptToIdeChat(prompt);
      if (dispatched) {
        notifyInfo('已将续写提示词填充到 IDE Chat。');
      } else {
        notifyWarn('当前宿主不支持自动填充 Chat，提示词已复制到剪贴板。');
      }
      return;
    }

    const settings = await resolveApiSettings(context);
    if (settings?.enabled && settings.apiKey) {
      try {
        const result = await generateWithFallbackModel(settings, prompt);
        const parsed = parseAiArticle(result.content);
        await appendGeneratedDraftToEnd(
          editor,
          parsed.body,
          parsed.conclusion,
          style,
          topic,
          `${PROVIDER_PRESETS[settings.provider].displayName} API · ${result.model}`
        );
        if (result.fallbackUsed && result.fallbackFromModel) {
          notifyWarn(
            `当前模型“${result.fallbackFromModel}”不受支持，已自动改用“${result.model}”。`
          );
        }
        notifyInfo('已通过 API 写入当前文稿末尾。');
        return;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.error('[writing-agent] API 续写失败，改用本地续写', message);
        notifyWarn(`API 续写失败，已回退本地续写：${message}`);
        const action = await vscode.window.showWarningMessage(
          'API 续写失败，是否改用 IDE AI 对话并自动填充提示词？',
          '改用 IDE AI 对话',
          '继续本地续写'
        );
        if (action === '改用 IDE AI 对话') {
          const dispatched = await dispatchPromptToIdeChat(prompt);
          if (dispatched) {
            notifyInfo('已改用 IDE Chat，请在聊天中继续生成并回填文稿。');
          } else {
            notifyWarn('无法自动填充 IDE Chat，提示词已复制到剪贴板。');
          }
          return;
        }
      }
    } else if (settings?.enabled) {
      const action = await vscode.window.showWarningMessage(
        '未配置 API Key。是否改用 IDE AI 对话，或现在配置 API？',
        '改用 IDE AI 对话',
        '配置 API'
      );
      if (action === '改用 IDE AI 对话') {
        const dispatched = await dispatchPromptToIdeChat(prompt);
        if (dispatched) {
          notifyInfo('已将续写提示词填充到 IDE Chat。');
        } else {
          notifyWarn('无法自动填充 IDE Chat，提示词已复制到剪贴板。');
        }
        return;
      }
      if (action === '配置 API') {
        await vscode.commands.executeCommand('writingAgent.configureApi');
      }
    }

    const fallback = generateStyleAlignedDraft(
      topic,
      style,
      active ? materialRepository.search(active.id, { semantic: topic, limit: 8 }) : []
    );
    await appendGeneratedDraftToEnd(editor, fallback.body, fallback.conclusion, style, topic, '本地续写');
    await vscode.env.clipboard.writeText(prompt);
    notifyInfo('已回退本地续写，并复制提示词供手动使用。');
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

    const selectedText = editor.document.getText(editor.selection).trim();
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
    const selectionRange = new vscode.Range(editor.selection.start, editor.selection.end);
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
    if (!settings.apiKey) {
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
      let fallbackInfo: { from: string; to: string } | undefined;
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
          if (!rewritten) {
            throw new Error('API 返回为空，未生成可替换内容');
          }
          if (result.fallbackUsed && result.fallbackFromModel) {
            fallbackInfo = {
              from: result.fallbackFromModel,
              to: result.model
            };
          }
          progress.report({ increment: 100, message: '完成' });
        }
      );

      const outputMode = await pickRewriteApplyModeWithPreview(rewritten);
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
        notifyWarn(`当前模型“${fallbackInfo.from}”不受支持，已自动改用“${fallbackInfo.to}”。`);
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
        { label: '仅修改模型与地址', value: 'configOnly' }
      ],
      { placeHolder: `配置 API（${preset.displayName}）` }
    );
    if (!action) {
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

    const modelDefault = provider === currentProvider && current?.model
      ? current.model
      : preset.defaultModel;
    const modelInput = await vscode.window.showInputBox({
      prompt: `模型名称（${preset.displayName}）`,
      value: modelDefault
    });
    if (modelInput === undefined) {
      return;
    }

    await config.update('api.enabled', true, target);
    await config.update('api.provider', provider, target);
    await config.update('api.baseUrl', baseUrlInput.trim() || preset.defaultBaseUrl, target);
    await config.update('api.model', modelInput.trim() || preset.defaultModel, target);

    const latest = await resolveApiSettings(context);
    if (latest?.enabled && latest.apiKey) {
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

  registerSafeCommand(context, 'writingAgent.testApiConnection', async () => {
    const settings = await resolveApiSettings(context);
    if (!settings?.enabled) {
      notifyWarn('API 已关闭，请先开启后再测试。');
      return;
    }
    if (!settings.apiKey) {
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
    notifyInfo('已将剪贴板内容写入当前文稿。');
  });

  registerSafeCommand(context, 'writingAgent.showAssistantViews', async () => {
    const now = Date.now();
    if (now < showAssistantViewLockUntil) {
      return;
    }
    if (showAssistantViewsInFlight) {
      await showAssistantViewsInFlight;
      return;
    }
    showAssistantViewLockUntil = now + SHOW_ASSISTANT_DEBOUNCE_MS;

    showAssistantViewsInFlight = (async () => {
      try {
        await vscode.commands.executeCommand('workbench.view.extension.writing-agent');
        try {
          await vscode.commands.executeCommand('workbench.actions.treeView.quickActions.focus');
        } catch {
          // 某些兼容层不支持该 focus 命令，容器已打开即可。
        }
        return;
      } catch (primaryError) {
        const primaryMessage = primaryError instanceof Error ? primaryError.message : String(primaryError);
        try {
          await vscode.commands.executeCommand('workbench.view.explorer');
          try {
            await vscode.commands.executeCommand('workbench.actions.treeView.quickActionsExplorer.focus');
          } catch {
            await vscode.commands.executeCommand('workbench.action.openView', 'quickActionsExplorer');
          }
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

  registerSafeCommand(context, 'writingAgent.refreshStyleProfile', async () => {
    styleProfileProvider.refresh();
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

async function dispatchPromptToIdeChat(prompt: string): Promise<boolean> {
  const normalized = prompt.trim();
  if (!normalized) {
    return false;
  }

  const commandSet = new Set(await vscode.commands.getCommands(true));
  const promptAttempts: Array<{ command: string; args: unknown }> = [
    { command: 'workbench.action.chat.open', args: { query: normalized } },
    { command: 'workbench.action.chat.open', args: { inputValue: normalized } },
    { command: 'workbench.action.chat.open', args: { initialQuery: normalized } },
    { command: 'workbench.action.chat.open', args: normalized },
    { command: 'vscode.chat.open', args: { query: normalized } },
    { command: 'cursor.chat.new', args: normalized },
    { command: 'cursor.chat.open', args: normalized },
    { command: 'codearts.agent.chat.open', args: { query: normalized } },
    { command: 'codearts.chat.open', args: { query: normalized } }
  ];

  for (const attempt of promptAttempts) {
    if (!commandSet.has(attempt.command)) {
      continue;
    }
    try {
      await vscode.commands.executeCommand(attempt.command, attempt.args);
      return true;
    } catch {
      // 尝试下一种宿主命令。
    }
  }

  const focusAttempts = [
    'workbench.panel.chat.view.focus',
    'workbench.action.chat.open',
    'cursor.chat.open',
    'codearts.agent.chat.open'
  ];
  for (const command of focusAttempts) {
    if (!commandSet.has(command)) {
      continue;
    }
    try {
      await vscode.commands.executeCommand(command);
      break;
    } catch {
      // 忽略聚焦失败，最终至少回退剪贴板。
    }
  }

  await vscode.env.clipboard.writeText(normalized);
  return false;
}

async function syncArticleCollectionWithDisk(
  articleRepository: ArticleRepository,
  articleCollectionProvider: ArticleCollectionProvider,
  options?: { notifyOnCleanup?: boolean }
): Promise<void> {
  const root = getPrimaryWorkspaceRootUri();
  if (!root) {
    return;
  }
  const result = await articleRepository.pruneMissingFiles(root);
  if (result.removed > 0) {
    articleCollectionProvider.refresh();
    if (options?.notifyOnCleanup) {
      notifyInfo(`文集已清理 ${result.removed} 条失效索引。`);
    }
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
  try {
    const result = await mammoth.extractRawText({ buffer });
    return result.value || '';
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`DOCX 解析失败：${message}`);
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

function getConfigurationTarget(): vscode.ConfigurationTarget {
  return vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0
    ? vscode.ConfigurationTarget.Workspace
    : vscode.ConfigurationTarget.Global;
}

function isApiProvider(value: string): value is ApiProvider {
  return ['iflow', 'kimi', 'deepseek', 'minimax', 'qwen', 'custom'].includes(value);
}

function getProviderPreset(provider: ApiProvider): ProviderPreset {
  return PROVIDER_PRESETS[provider];
}

function getProviderSecretKey(provider: ApiProvider): string {
  return `writingAgent.apiKey.${provider}`;
}

function getProviderEnvKey(provider: ApiProvider): string {
  switch (provider) {
    case 'iflow':
      return 'IFLOW_API_KEY';
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

async function pickApiProvider(currentProvider: ApiProvider): Promise<ApiProvider | undefined> {
  const options = (Object.keys(PROVIDER_PRESETS) as ApiProvider[]).map(provider => {
    const preset = getProviderPreset(provider);
    return {
      label: preset.displayName,
      description: provider === currentProvider ? '当前' : '',
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
  const preset = getProviderPreset(provider);

  const baseUrl = config.get<string>(
    'api.baseUrl',
    provider === 'iflow'
      ? config.get<string>('iflow.baseUrl', preset.defaultBaseUrl)
      : preset.defaultBaseUrl
  );
  const model = config.get<string>(
    'api.model',
    provider === 'iflow'
      ? config.get<string>('iflow.model', preset.defaultModel)
      : preset.defaultModel
  );
  const temperature = config.get<number>('api.temperature', config.get<number>('iflow.temperature', 0.7));
  const maxTokens = config.get<number>('api.maxTokens', config.get<number>('iflow.maxTokens', 2200));
  const timeoutMs = config.get<number>('api.timeoutMs', config.get<number>('iflow.timeoutMs', 60000));

  const providerSecret = ((await context.secrets.get(getProviderSecretKey(provider))) || '').trim();
  const legacyIflowSecret = provider === 'iflow'
    ? ((await context.secrets.get(LEGACY_IFLOW_SECRET_KEY)) || '').trim()
    : '';
  const legacyConfigKey = provider === 'iflow'
    ? (config.get<string>('iflow.apiKey', '') || '').trim()
    : '';
  const envProviderKey = (process.env[getProviderEnvKey(provider)] || '').trim();
  const envGenericKey = (process.env.WRITING_AGENT_API_KEY || '').trim();
  const apiKey = providerSecret || legacyIflowSecret || legacyConfigKey || envProviderKey || envGenericKey;

  if (!enabled) {
    return {
      enabled: false,
      provider,
      apiKey: '',
      baseUrl,
      model,
      temperature,
      maxTokens,
      timeoutMs
    };
  }

  return {
    enabled,
    provider,
    apiKey,
    baseUrl,
    model,
    temperature,
    maxTokens,
    timeoutMs
  };
}

function buildApiValidationFingerprint(settings: RuntimeApiSettings): string {
  return [
    settings.provider,
    settings.baseUrl.trim(),
    settings.model.trim()
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

async function validateApiConnection(
  context: vscode.ExtensionContext,
  settings: RuntimeApiSettings,
  persist: boolean
): Promise<{ valid: boolean; message: string }> {
  const prompt = '请仅回复 OK';
  try {
    const result = await generateWithOpenAiCompatibleApi({
      apiKey: settings.apiKey,
      baseUrl: settings.baseUrl,
      model: settings.model,
      prompt,
      temperature: 0,
      maxTokens: 16,
      timeoutMs: Math.min(Math.max(settings.timeoutMs, 8000), 10000)
    });
    const message = `连接成功（${result.model || settings.model}）`;
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
  workspaceRoot: vscode.Uri | null
): Promise<DeleteStyleArticlesResult> {
  const records = articleRepository.listByStyle(styleId);
  const result: DeleteStyleArticlesResult = {
    removedRecords: 0,
    deletedFiles: 0,
    missingFiles: 0
  };

  for (const record of records) {
    if (workspaceRoot) {
      const articleUri = vscode.Uri.joinPath(workspaceRoot, ...record.relativePath.split('/'));
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
  workspaceRoot: vscode.Uri | null
): Promise<MergeStyleArticlesResult> {
  const records = articleRepository.listByStyle(sourceStyle.id);
  const result: MergeStyleArticlesResult = {
    updatedRecords: 0,
    movedFiles: 0,
    renamedFiles: 0,
    missingFiles: 0
  };

  const targetDir = workspaceRoot
    ? vscode.Uri.joinPath(workspaceRoot, COLLECTION_RELATIVE_PREFIX, sanitizePathSegment(targetStyle.name))
    : null;
  if (targetDir) {
    await vscode.workspace.fs.createDirectory(targetDir);
  }

  for (const record of records) {
    let nextRelativePath = record.relativePath;
    let nextTitle = record.title;

    if (workspaceRoot && targetDir) {
      const sourceUri = vscode.Uri.joinPath(workspaceRoot, ...record.relativePath.split('/'));
      try {
        if (await uriExists(sourceUri)) {
          const sourceName = sourceUri.path.slice(sourceUri.path.lastIndexOf('/') + 1);
          const targetUri = vscode.Uri.joinPath(targetDir, sourceName);
          const finalTarget = await ensureUniqueTargetUri(targetUri);
          if (finalTarget.path !== targetUri.path) {
            result.renamedFiles += 1;
          }
          await vscode.workspace.fs.rename(sourceUri, finalTarget, { overwrite: false });
          nextRelativePath = finalTarget.path.replace(workspaceRoot.path, '').replace(/^\//, '');
          nextTitle = finalTarget.path.slice(finalTarget.path.lastIndexOf('/') + 1).replace(/\.md$/i, '');
          result.movedFiles += 1;
        } else {
          result.missingFiles += 1;
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.error(`[writing-agent] 迁移风格文章文件失败: ${record.relativePath}`, message);
      }
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

async function migrateArticleCollectionToVisibleRoot(
  articleRepository: ArticleRepository,
  articleCollectionProvider: ArticleCollectionProvider
): Promise<void> {
  const root = getPrimaryWorkspaceRootUri();
  if (!root) {
    return;
  }

  const legacyDir = vscode.Uri.joinPath(root, ...LEGACY_COLLECTION_RELATIVE_PREFIX.split('/'));
  const currentDir = vscode.Uri.joinPath(root, COLLECTION_RELATIVE_PREFIX);
  const legacyExists = await uriExists(legacyDir);
  if (!legacyExists) {
    return;
  }

  try {
    if (!(await uriExists(currentDir))) {
      await vscode.workspace.fs.rename(legacyDir, currentDir, { overwrite: false });
    } else {
      const files = await listFilesRecursive(legacyDir);
      for (const file of files) {
        const relative = file.path.slice(legacyDir.path.length).replace(/^\/+/, '');
        const target = vscode.Uri.joinPath(currentDir, ...relative.split('/'));
        const finalTarget = await ensureUniqueTargetUri(target);
        await vscode.workspace.fs.createDirectory(getParentUri(finalTarget));
        await vscode.workspace.fs.rename(file, finalTarget, { overwrite: false });
      }
      await vscode.workspace.fs.delete(legacyDir, { recursive: true, useTrash: false });
    }
    const changed = await articleRepository.replacePathPrefix(
      `${LEGACY_COLLECTION_RELATIVE_PREFIX}/`,
      `${COLLECTION_RELATIVE_PREFIX}/`
    );
    if (changed > 0) {
      articleCollectionProvider.refresh();
      notifyInfo(`已迁移 ${changed} 篇文集记录到“${COLLECTION_RELATIVE_PREFIX}”目录。`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('[writing-agent] 文集迁移失败', message);
    notifyWarn(`文集迁移失败：${message}`);
  }
}

async function listFilesRecursive(dir: vscode.Uri): Promise<vscode.Uri[]> {
  const files: vscode.Uri[] = [];
  const entries = await vscode.workspace.fs.readDirectory(dir);
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

async function generateWithFallbackModel(settings: RuntimeApiSettings, prompt: string): Promise<ApiGenerationResult> {
  const firstResult = await generateWithOpenAiCompatibleApi({
    apiKey: settings.apiKey,
    baseUrl: settings.baseUrl,
    model: settings.model,
    prompt,
    temperature: settings.temperature,
    maxTokens: settings.maxTokens,
    timeoutMs: settings.timeoutMs
  }).catch(error => ({ error } as const));

  if (!('error' in firstResult)) {
    return {
      content: firstResult.content,
      model: firstResult.model || settings.model,
      fallbackUsed: false
    };
  }

  if (!isModelNotSupportedApiError(firstResult.error)) {
    throw firstResult.error;
  }

  const presetModel = getProviderPreset(settings.provider).defaultModel.trim();
  const currentModel = settings.model.trim();
  if (!presetModel || presetModel === currentModel) {
    throw new Error(`当前模型不受支持：${currentModel}。请执行“写作助手: 配置 API”更换模型后重试。`);
  }

  const retryResult = await generateWithOpenAiCompatibleApi({
    apiKey: settings.apiKey,
    baseUrl: settings.baseUrl,
    model: presetModel,
    prompt,
    temperature: settings.temperature,
    maxTokens: settings.maxTokens,
    timeoutMs: settings.timeoutMs
  });

  const config = vscode.workspace.getConfiguration('writingAgent');
  await config.update('api.model', presetModel, getConfigurationTarget());
  return {
    content: retryResult.content,
    model: retryResult.model || presetModel,
    fallbackUsed: true,
    fallbackFromModel: currentModel
  };
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

  const lastLine = editor.document.lineCount - 1;
  const lastChar = editor.document.lineAt(lastLine).text.length;
  const position = new vscode.Position(lastLine, lastChar);
  await editor.edit(editBuilder => {
    editBuilder.insert(position, block);
  });
  const endLine = editor.document.lineCount - 1;
  const endChar = editor.document.lineAt(endLine).text.length;
  const end = new vscode.Position(endLine, endChar);
  editor.selection = new vscode.Selection(end, end);
  editor.revealRange(new vscode.Range(end, end), vscode.TextEditorRevealType.InCenter);
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

  const lastLine = editor.document.lineCount - 1;
  const lastChar = editor.document.lineAt(lastLine).text.length;
  const position = new vscode.Position(lastLine, lastChar);
  await editor.edit(editBuilder => {
    editBuilder.insert(position, block);
  });
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

async function pickRewriteApplyModeWithPreview(
  rewritten: string
): Promise<'replace' | 'insertAfter' | undefined> {
  const previewLimit = 2600;
  const preview =
    rewritten.length > previewLimit
      ? `${rewritten.slice(0, previewLimit)}\n\n（预览已截断，全文 ${rewritten.length} 字）`
      : rewritten;
  const action = await vscode.window.showInformationMessage(
    '已生成改写文本，请确认应用方式。',
    { modal: true, detail: preview },
    '替换选区',
    '插入到选区后'
  );
  if (action === '替换选区') {
    return 'replace';
  }
  if (action === '插入到选区后') {
    return 'insertAfter';
  }
  return undefined;
}

function getStyleAnalysisWebview(style: WritingStyle): string {
  return `
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {
                    padding: 20px;
                    font-family: var(--vscode-font-family);
                    line-height: 1.6;
                }
                h2 { color: var(--vscode-editor-foreground); }
                .metric {
                    background: var(--vscode-editor-background);
                    padding: 15px;
                    margin: 10px 0;
                    border-radius: 5px;
                }
                .metric-bar {
                    height: 8px;
                    background: var(--vscode-progressBar-background);
                    border-radius: 4px;
                    margin-top: 5px;
                }
                .tag {
                    display: inline-block;
                    background: var(--vscode-button-background);
                    color: var(--vscode-button-foreground);
                    padding: 3px 8px;
                    margin: 3px;
                    border-radius: 3px;
                    font-size: 12px;
                }
            </style>
        </head>
        <body>
            <h2>写作风格分析</h2>
            
            <div class="metric">
                <label>平均句长: ${style.sentenceStructure?.avgLength?.toFixed(1) || 0} 字</label>
                <div class="metric-bar" style="width: ${Math.min((style.sentenceStructure?.avgLength || 0) / 2, 100)}%"></div>
            </div>
            
            <div class="metric">
                <label>词汇丰富度: ${((style.vocabulary?.uniqueWordRatio || 0) * 100).toFixed(1)}%</label>
                <div class="metric-bar" style="width: ${(style.vocabulary?.uniqueWordRatio || 0) * 100}%"></div>
            </div>
            
            <div class="metric">
                <label>语气类型: ${style.tone?.type || 'neutral'}</label>
            </div>
            
            <div class="metric">
                <label>情感倾向: ${style.emotionalArc?.overallTrend || 'neutral'}</label>
            </div>
            
            <h3>常用词汇</h3>
            <div>
                ${(style.vocabulary?.favoriteWords || []).map((w: string) =>
    `<span class="tag">${w}</span>`
  ).join('')}
            </div>
            
            <h3>结构特征</h3>
            <div>
                ${(style.contentStructure?.structurePatterns || []).map((p: string) =>
    `<span class="tag">${p}</span>`
  ).join('')}
            </div>
            
            <h3>过渡词</h3>
            <div>
                ${(style.contentStructure?.transitionWords || []).map((w: string) =>
    `<span class="tag">${w}</span>`
  ).join('')}
            </div>
        </body>
        </html>
    `;
}

export function deactivate(): void {
  console.log('写作 Agent 插件已停用');
}
