import * as vscode from 'vscode';
import * as path from 'path';
import { MaterialType, SearchQuery, WritingMaterial } from '../core/types';
import { MaterialRepository } from '../core/materialManager/repository';

interface SearchPanelOptions {
  presetTypes?: MaterialType[];
  materialId?: string;
  initialQuery?: string;
  styleId?: string;
  styleName?: string;
}

interface SearchRequestMessage {
  command: 'search';
  payload?: {
    query?: string;
    types?: MaterialType[];
    focusMaterialId?: string;
  };
}

interface CopyMessage {
  command: 'copyMaterial';
  payload?: { id?: string };
}

interface InsertMessage {
  command: 'insertMaterial';
  payload?: { id?: string };
}

interface DetailMessage {
  command: 'openMaterialDetail';
  payload?: { id?: string };
}

interface ExportMarkdownMessage {
  command: 'exportMarkdown';
  payload?: {
    query?: string;
    types?: MaterialType[];
  };
}

type InboundMessage = SearchRequestMessage | CopyMessage | InsertMessage | DetailMessage | ExportMarkdownMessage;

interface WebMaterial {
  id: string;
  name: string;
  content: string;
  type: MaterialType;
  typeLabel: string;
  tags: string[];
  quality: number;
  source: string;
  updatedAt: string;
}

export class MaterialSearchPanel {
  private static readonly VIEW_TYPE = 'writingAgent.materialSearch';
  private static currentPanel: MaterialSearchPanel | undefined;

  static show(
    context: vscode.ExtensionContext,
    repository: MaterialRepository,
    options?: SearchPanelOptions
  ): void {
    if (MaterialSearchPanel.currentPanel) {
      MaterialSearchPanel.currentPanel.panel.reveal(vscode.ViewColumn.One);
      MaterialSearchPanel.currentPanel.updateOptions(options);
      return;
    }

    const panel = vscode.window.createWebviewPanel(
      MaterialSearchPanel.VIEW_TYPE,
      '素材工作台',
      vscode.ViewColumn.One,
      { enableScripts: true }
    );

    MaterialSearchPanel.currentPanel = new MaterialSearchPanel(panel, context, repository, options);
  }

  private readonly panel: vscode.WebviewPanel;
  private readonly repository: MaterialRepository;
  private options: SearchPanelOptions;
  private readonly disposables: vscode.Disposable[] = [];

  private constructor(
    panel: vscode.WebviewPanel,
    context: vscode.ExtensionContext,
    repository: MaterialRepository,
    options?: SearchPanelOptions
  ) {
    this.panel = panel;
    this.repository = repository;
    this.options = options || {};
    this.updatePanelTitle();

    this.panel.webview.html = this.getHtml();
    this.bindEvents();
    this.postBootstrap(context.extensionUri);
    this.searchAndPost({
      keywords: this.extractKeywords(this.options.initialQuery || ''),
      types: this.options.presetTypes,
      limit: 100
    }, this.options.materialId);
  }

  private bindEvents(): void {
    this.panel.onDidDispose(() => this.dispose(), null, this.disposables);

    this.panel.webview.onDidReceiveMessage(
      async (message: InboundMessage) => {
        try {
          switch (message.command) {
            case 'search':
              await this.handleSearch(message.payload);
              break;
            case 'copyMaterial':
              await this.handleCopy(message.payload?.id);
              break;
            case 'insertMaterial':
              await this.handleInsert(message.payload?.id);
              break;
            case 'openMaterialDetail':
              this.handleDetail(message.payload?.id);
              break;
            case 'exportMarkdown':
              await this.handleExportMarkdown(message.payload);
              break;
            default:
              break;
          }
        } catch (error) {
          const messageText = error instanceof Error ? error.message : String(error);
          vscode.window.showErrorMessage(`素材搜索失败: ${messageText}`);
        }
      },
      null,
      this.disposables
    );
  }

  private async handleSearch(payload?: SearchRequestMessage['payload']): Promise<void> {
    const query = payload?.query || '';
    const request: SearchQuery = {
      keywords: this.extractKeywords(query),
      types: payload?.types && payload.types.length > 0 ? payload.types : undefined,
      limit: 100
    };
    this.searchAndPost(request, payload?.focusMaterialId);
  }

  private async handleCopy(id?: string): Promise<void> {
    if (!id) {
      return;
    }
    const styleId = this.getStyleIdOrWarn();
    if (!styleId) {
      return;
    }
    const material = this.repository.getById(styleId, id);
    if (!material) {
      this.notifyWarn('未找到对应素材');
      return;
    }
    await vscode.env.clipboard.writeText(material.content);
    await this.repository.markUsed(styleId, material.id);
    this.notifyInfo(`已复制素材：${material.name}`);
  }

  private async handleInsert(id?: string): Promise<void> {
    if (!id) {
      return;
    }
    const styleId = this.getStyleIdOrWarn();
    if (!styleId) {
      return;
    }
    const material = this.repository.getById(styleId, id);
    if (!material) {
      this.notifyWarn('未找到对应素材');
      return;
    }

    const editor = await this.resolveEditableEditor();
    if (!editor) {
      this.notifyWarn('请先打开一个可编辑文档');
      return;
    }

    await editor.edit(editBuilder => {
      editBuilder.insert(editor.selection.active, material.content);
    });
    await vscode.window.showTextDocument(editor.document, {
      preserveFocus: false,
      preview: false,
      viewColumn: editor.viewColumn
    });
    await this.repository.markUsed(styleId, material.id);
    this.notifyInfo(`已插入素材：${material.name}`);
  }

  private async resolveEditableEditor(): Promise<vscode.TextEditor | undefined> {
    const active = vscode.window.activeTextEditor;
    if (active && this.isEditableTextEditor(active)) {
      return active;
    }

    const visibleEditors = vscode.window.visibleTextEditors;
    for (const editor of visibleEditors) {
      if (this.isEditableTextEditor(editor)) {
        return editor;
      }
    }

    const groups = [vscode.window.tabGroups.activeTabGroup, ...vscode.window.tabGroups.all]
      .filter((group, index, array) => array.findIndex(item => item.viewColumn === group.viewColumn) === index);
    for (const group of groups) {
      const tabs = [...group.tabs].reverse();
      for (const tab of tabs) {
        const uri = this.extractTabUri(tab);
        if (!uri || !this.isPotentiallyEditableUri(uri)) {
          continue;
        }
        try {
          const doc = await vscode.workspace.openTextDocument(uri);
          if (!this.isEditableTextDocument(doc)) {
            continue;
          }
          const editor = await vscode.window.showTextDocument(doc, {
            preserveFocus: true,
            preview: false,
            viewColumn: group.viewColumn
          });
          if (this.isEditableTextEditor(editor)) {
            return editor;
          }
        } catch {
          // ignore and try next tab candidate
        }
      }
    }

    return undefined;
  }

  private isEditableTextEditor(editor: vscode.TextEditor): boolean {
    return this.isEditableTextDocument(editor.document);
  }

  private isEditableTextDocument(document: vscode.TextDocument): boolean {
    if (document.isClosed) {
      return false;
    }
    const readonlyFlag = (document as unknown as { isReadonly?: boolean }).isReadonly;
    if (readonlyFlag === true) {
      return false;
    }
    const blockedSchemes = new Set(['output', 'debug', 'vscode', 'vscode-notebook-cell']);
    if (blockedSchemes.has(document.uri.scheme)) {
      return false;
    }
    return true;
  }

  private isPotentiallyEditableUri(uri: vscode.Uri): boolean {
    return uri.scheme === 'file' || uri.scheme === 'untitled';
  }

  private extractTabUri(tab: vscode.Tab): vscode.Uri | undefined {
    const input = tab.input as unknown as { uri?: unknown; modified?: unknown };
    if (input?.uri instanceof vscode.Uri) {
      return input.uri;
    }
    if (input?.modified instanceof vscode.Uri) {
      return input.modified;
    }
    return undefined;
  }

  private handleDetail(id?: string): void {
    if (!id) {
      return;
    }
    const styleId = this.options.styleId;
    if (!styleId) {
      return;
    }
    const material = this.repository.getById(styleId, id);
    if (!material) {
      return;
    }
    this.panel.webview.postMessage({
      type: 'materialDetail',
      material: this.toWebMaterial(material)
    });
  }

  private async handleExportMarkdown(payload?: ExportMarkdownMessage['payload']): Promise<void> {
    const styleId = this.getStyleIdOrWarn();
    if (!styleId) {
      return;
    }

    const queryText = payload?.query || '';
    const types = payload?.types && payload.types.length > 0 ? payload.types : undefined;
    const request: SearchQuery = {
      keywords: this.extractKeywords(queryText),
      types,
      limit: 1000
    };
    const materials = this.repository.search(styleId, request);
    if (materials.length === 0) {
      this.notifyWarn('当前筛选结果为空，暂无可导出素材。');
      return;
    }

    const styleName = this.options.styleName || '未命名风格';
    const timestamp = this.formatFileTimestamp(new Date());
    const defaultName = `${this.sanitizeFileName(styleName)}-素材导出-${timestamp}.md`;
    const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri;
    const defaultUri = workspaceRoot
      ? vscode.Uri.joinPath(workspaceRoot, '.writing-agent', 'exports', defaultName)
      : undefined;

    const targetUri = await vscode.window.showSaveDialog({
      saveLabel: '导出素材',
      defaultUri,
      filters: { Markdown: ['md'] }
    });
    if (!targetUri) {
      return;
    }

    const parentDir = targetUri.with({ path: path.posix.dirname(targetUri.path) });
    await vscode.workspace.fs.createDirectory(parentDir);

    const content = this.buildMarkdownExport(styleName, queryText, types, materials);
    await vscode.workspace.fs.writeFile(targetUri, Buffer.from(content, 'utf-8'));
    const action = await vscode.window.showInformationMessage(
      `已导出 ${materials.length} 条素材到 ${targetUri.fsPath}`,
      '打开文件'
    );
    if (action === '打开文件') {
      await vscode.commands.executeCommand('vscode.open', targetUri);
    }
  }

  private searchAndPost(query: SearchQuery, focusMaterialId?: string): void {
    const styleId = this.options.styleId;
    if (!styleId) {
      this.panel.webview.postMessage({
        type: 'searchResults',
        results: [],
        focusMaterialId
      });
      this.panel.webview.postMessage({
        type: 'typeStats',
        totalAll: 0,
        totalFiltered: 0,
        stats: []
      });
      return;
    }

    const matchedMaterials = this.repository.search(styleId, query);
    const results = matchedMaterials.map(material => this.toWebMaterial(material));
    this.panel.webview.postMessage({
      type: 'searchResults',
      results,
      focusMaterialId
    });
    this.postTypeStats(styleId, matchedMaterials);
  }

  private postTypeStats(styleId: string, matchedMaterials: WritingMaterial[]): void {
    const overallMap = this.repository.getTypeCounts(styleId);
    const filteredMap = new Map<MaterialType, number>();

    for (const material of matchedMaterials) {
      filteredMap.set(material.type, (filteredMap.get(material.type) || 0) + 1);
    }

    const allTypes = new Set<MaterialType>([
      ...overallMap.keys(),
      ...filteredMap.keys()
    ]);
    const stats = [...allTypes]
      .map(type => ({
        type,
        label: this.getTypeName(type),
        overall: overallMap.get(type) || 0,
        filtered: filteredMap.get(type) || 0
      }))
      .sort((a, b) => b.filtered - a.filtered || b.overall - a.overall);

    this.panel.webview.postMessage({
      type: 'typeStats',
      totalAll: this.repository.getAll(styleId).length,
      totalFiltered: matchedMaterials.length,
      stats
    });
  }

  private buildMarkdownExport(
    styleName: string,
    queryText: string,
    types: MaterialType[] | undefined,
    materials: WritingMaterial[]
  ): string {
    const typeCounts = new Map<MaterialType, number>();
    for (const material of materials) {
      typeCounts.set(material.type, (typeCounts.get(material.type) || 0) + 1);
    }

    const typeSummary = [...typeCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([type, count]) => `- ${this.getTypeName(type)}：${count} 条`)
      .join('\n');

    const materialBlocks = materials.map((material, index) => [
      `## ${index + 1}. ${material.name}`,
      '',
      `- 类型：${this.getTypeName(material.type)}`,
      `- 质量：${material.metadata.quality}`,
      `- 更新时间：${material.metadata.updatedAt.toLocaleString()}`,
      `- 来源：${material.metadata.source}`,
      `- 标签：${(material.tags || []).join('、') || '无'}`,
      '',
      '```text',
      material.content,
      '```',
      ''
    ].join('\n')).join('\n');

    return [
      `# 素材导出 · ${styleName}`,
      '',
      `- 导出时间：${new Date().toLocaleString()}`,
      `- 搜索关键词：${queryText || '（空）'}`,
      `- 类型筛选：${types && types.length > 0 ? types.map(item => this.getTypeName(item)).join('、') : '全部'}`,
      `- 导出数量：${materials.length}`,
      '',
      '## 类型统计',
      typeSummary || '- 无',
      '',
      '## 素材列表',
      '',
      materialBlocks
    ].join('\n');
  }

  private sanitizeFileName(name: string): string {
    return name
      .trim()
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
      .replace(/\s+/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '')
      .slice(0, 80) || 'materials';
  }

  private formatFileTimestamp(date: Date): string {
    const pad = (value: number): string => String(value).padStart(2, '0');
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
  }

  private updateOptions(options?: SearchPanelOptions): void {
    if (!options) {
      return;
    }
    this.options = options;
    this.updatePanelTitle();
    this.panel.webview.postMessage({
      type: 'applyPreset',
      preset: {
        query: options.initialQuery || '',
        types: options.presetTypes || [],
        focusMaterialId: options.materialId,
        styleName: options.styleName || ''
      }
    });
    this.postStyleContext();
    this.searchAndPost(
      {
        keywords: this.extractKeywords(options.initialQuery || ''),
        types: options.presetTypes,
        limit: 100
      },
      options.materialId
    );
  }

  private postBootstrap(extensionUri: vscode.Uri): void {
    this.panel.webview.postMessage({
      type: 'bootstrap',
      preset: {
        query: this.options.initialQuery || '',
        types: this.options.presetTypes || [],
        focusMaterialId: this.options.materialId,
        styleName: this.options.styleName || ''
      },
      availableTypes: Object.values(MaterialType),
      extensionRoot: extensionUri.toString()
    });
    this.postStyleContext();
  }

  private postStyleContext(): void {
    this.panel.webview.postMessage({
      type: 'styleContext',
      styleName: this.options.styleName || '',
      hasActiveStyle: Boolean(this.options.styleId)
    });
  }

  private extractKeywords(query: string): string[] {
    if (!query.trim()) {
      return [];
    }
    return query.split(/[\s,，;；]+/).map(item => item.trim()).filter(Boolean);
  }

  private updatePanelTitle(): void {
    this.panel.title = this.options.styleName
      ? `素材工作台 · ${this.options.styleName}`
      : '素材工作台';
  }

  private toWebMaterial(material: WritingMaterial): WebMaterial {
    return {
      id: material.id,
      name: material.name,
      content: material.content,
      type: material.type,
      typeLabel: this.getTypeName(material.type),
      tags: material.tags || [],
      quality: material.metadata.quality,
      source: material.metadata.source,
      updatedAt: material.metadata.updatedAt.toLocaleString()
    };
  }

  private getTypeName(type: MaterialType): string {
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

  private getHtml(): string {
    const initialPreset = JSON.stringify({
      query: this.options.initialQuery || '',
      types: this.options.presetTypes || [],
      focusMaterialId: this.options.materialId
    }).replace(/</g, '\\u003c');

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>素材工作台</title>
  <style>
    :root {
      color-scheme: light dark;
    }
    body {
      margin: 0;
      padding: 16px;
      font-family: var(--vscode-font-family);
      color: var(--vscode-foreground);
      background: var(--vscode-editor-background);
    }
    .toolbar {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 8px;
      margin-bottom: 12px;
    }
    .toolbar-actions {
      display: flex;
      gap: 8px;
    }
    input[type="text"] {
      border: 1px solid var(--vscode-input-border);
      background: var(--vscode-input-background);
      color: var(--vscode-input-foreground);
      padding: 8px 10px;
      border-radius: 6px;
    }
    button {
      border: 0;
      background: var(--vscode-button-background);
      color: var(--vscode-button-foreground);
      padding: 8px 12px;
      border-radius: 6px;
      cursor: pointer;
    }
    button:hover {
      background: var(--vscode-button-hoverBackground);
    }
    button.secondary {
      background: var(--vscode-button-secondaryBackground);
      color: var(--vscode-button-secondaryForeground);
    }
    button.secondary:hover {
      background: var(--vscode-button-secondaryHoverBackground);
    }
    .filters {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 12px;
      padding: 8px;
      background: color-mix(in srgb, var(--vscode-editorWidget-background) 90%, transparent);
      border-radius: 6px;
    }
    .stats-panel {
      border: 1px solid var(--vscode-panel-border);
      border-radius: 8px;
      margin-bottom: 12px;
      overflow: hidden;
    }
    .stats-header {
      padding: 8px 12px;
      font-weight: 600;
      background: color-mix(in srgb, var(--vscode-sideBar-background) 85%, transparent);
      border-bottom: 1px solid var(--vscode-panel-border);
    }
    .stats-summary {
      padding: 8px 12px;
      font-size: 12px;
      opacity: 0.85;
      border-bottom: 1px solid var(--vscode-panel-border);
    }
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 8px;
      padding: 10px;
    }
    .stats-item {
      border: 1px solid var(--vscode-panel-border);
      border-radius: 6px;
      padding: 8px;
      background: color-mix(in srgb, var(--vscode-editor-background) 92%, transparent);
      cursor: pointer;
    }
    .stats-item:hover {
      background: color-mix(in srgb, var(--vscode-list-hoverBackground) 92%, transparent);
    }
    .stats-item strong {
      display: block;
      margin-bottom: 4px;
    }
    .layout {
      display: grid;
      grid-template-columns: 1.2fr 1fr;
      gap: 12px;
    }
    .results, .detail {
      border: 1px solid var(--vscode-panel-border);
      border-radius: 8px;
      overflow: hidden;
      min-height: 420px;
    }
    .results-header, .detail-header {
      padding: 8px 12px;
      font-weight: 600;
      background: color-mix(in srgb, var(--vscode-sideBar-background) 85%, transparent);
      border-bottom: 1px solid var(--vscode-panel-border);
    }
    .result-list {
      max-height: 560px;
      overflow: auto;
    }
    .result-item {
      padding: 10px 12px;
      border-bottom: 1px solid var(--vscode-panel-border);
      cursor: pointer;
    }
    .result-item:hover {
      background: color-mix(in srgb, var(--vscode-list-hoverBackground) 95%, transparent);
    }
    .result-item.active {
      background: color-mix(in srgb, var(--vscode-list-activeSelectionBackground) 85%, transparent);
    }
    .item-title {
      display: flex;
      justify-content: space-between;
      gap: 8px;
      font-weight: 600;
      margin-bottom: 6px;
    }
    .item-meta {
      opacity: 0.8;
      font-size: 12px;
      margin-bottom: 6px;
    }
    .item-actions {
      display: flex;
      gap: 8px;
      margin-top: 8px;
    }
    .item-actions button {
      padding: 4px 8px;
      font-size: 12px;
    }
    .tags {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-top: 6px;
    }
    .tag {
      font-size: 12px;
      padding: 2px 8px;
      border-radius: 999px;
      background: color-mix(in srgb, var(--vscode-button-background) 70%, transparent);
    }
    .detail-content {
      padding: 12px;
      white-space: pre-wrap;
      line-height: 1.65;
    }
    .empty {
      padding: 16px;
      opacity: 0.8;
    }
    @media (max-width: 900px) {
      .layout {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <input id="query" type="text" placeholder="输入关键词，如：TypeScript 类型系统" />
    <div class="toolbar-actions">
      <button id="searchBtn">搜索</button>
      <button id="exportBtn" class="secondary">导出 MD</button>
    </div>
  </div>
  <div id="styleHint" class="empty" style="padding:0 0 10px 0;"></div>
  <div id="filters" class="filters"></div>
  <section class="stats-panel">
    <div class="stats-header">类型统计（同页可筛选）</div>
    <div id="statsSummary" class="stats-summary">总计 0 条，当前筛选 0 条</div>
    <div id="statsGrid" class="stats-grid"></div>
  </section>
  <div class="layout">
    <section class="results">
      <div class="results-header">搜索结果</div>
      <div id="resultList" class="result-list"></div>
    </section>
    <section class="detail">
      <div class="detail-header">素材详情</div>
      <div id="detailContent" class="detail-content">请选择左侧素材查看详情</div>
    </section>
  </div>

  <script>
    const vscode = acquireVsCodeApi();
    const state = {
      preset: ${initialPreset},
      availableTypes: [],
      selectedTypes: new Set(),
      results: [],
      activeId: undefined,
      totalAll: 0,
      totalFiltered: 0,
      typeStats: []
    };

    const queryInput = document.getElementById('query');
    const searchBtn = document.getElementById('searchBtn');
    const exportBtn = document.getElementById('exportBtn');
    const resultList = document.getElementById('resultList');
    const detailContent = document.getElementById('detailContent');
    const filtersEl = document.getElementById('filters');
    const styleHint = document.getElementById('styleHint');
    const statsSummary = document.getElementById('statsSummary');
    const statsGrid = document.getElementById('statsGrid');

    function toLabel(type) {
      const map = {
        sentence: '精彩句子',
        paragraph: '优秀段落',
        quote: '引用名言',
        metaphor: '比喻修辞',
        opening: '文章开头',
        ending: '文章结尾',
        transition: '过渡表达',
        idea: '创意观点',
        'style-sample': '风格样本'
      };
      return map[type] || type;
    }

    function escapeHtml(value) {
      return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    function renderFilters() {
      filtersEl.innerHTML = '';
      for (const type of state.availableTypes) {
        const label = document.createElement('label');
        label.style.display = 'inline-flex';
        label.style.alignItems = 'center';
        label.style.gap = '4px';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = state.selectedTypes.has(type);
        checkbox.addEventListener('change', () => {
          if (checkbox.checked) {
            state.selectedTypes.add(type);
          } else {
            state.selectedTypes.delete(type);
          }
          runSearch();
        });

        const text = document.createElement('span');
        text.textContent = toLabel(type);

        label.appendChild(checkbox);
        label.appendChild(text);
        filtersEl.appendChild(label);
      }
    }

    function renderStats() {
      statsSummary.textContent = '总计 ' + state.totalAll + ' 条，当前筛选 ' + state.totalFiltered + ' 条';
      if (!state.typeStats.length) {
        statsGrid.innerHTML = '<div class="empty">暂无类型统计。</div>';
        return;
      }

      statsGrid.innerHTML = state.typeStats.map(item => \`
        <div class="stats-item" data-type="\${item.type}">
          <strong>\${escapeHtml(item.label)}</strong>
          <div>当前筛选：\${item.filtered} 条</div>
          <div style="opacity:0.8;">总库存：\${item.overall} 条</div>
        </div>
      \`).join('');

      statsGrid.querySelectorAll('.stats-item').forEach(item => {
        item.addEventListener('click', () => {
          const type = item.getAttribute('data-type');
          if (!type) return;
          if (state.selectedTypes.size === 1 && state.selectedTypes.has(type)) {
            state.selectedTypes.clear();
          } else {
            state.selectedTypes = new Set([type]);
          }
          renderFilters();
          runSearch();
        });
      });
    }

    function renderResults() {
      if (!state.results.length) {
        resultList.innerHTML = '<div class="empty">暂无匹配素材，换个关键词试试。</div>';
        detailContent.textContent = '请选择左侧素材查看详情';
        return;
      }

      resultList.innerHTML = state.results.map(material => {
        const tags = (material.tags || []).slice(0, 4).map(tag => '<span class="tag">' + escapeHtml(tag) + '</span>').join('');
        const preview = material.content.length > 90 ? material.content.slice(0, 90) + '...' : material.content;
        const activeClass = state.activeId === material.id ? 'active' : '';
        return \`
          <div class="result-item \${activeClass}" data-id="\${material.id}">
            <div class="item-title">
              <span>\${escapeHtml(material.name)}</span>
              <span>质量 \${material.quality}</span>
            </div>
            <div class="item-meta">\${escapeHtml(material.typeLabel)} · \${escapeHtml(material.updatedAt)}</div>
            <div>\${escapeHtml(preview)}</div>
            <div class="tags">\${tags}</div>
            <div class="item-actions">
              <button data-action="copy" data-id="\${material.id}">复制</button>
              <button data-action="insert" data-id="\${material.id}">插入</button>
            </div>
          </div>
        \`;
      }).join('');

      resultList.querySelectorAll('.result-item').forEach(item => {
        item.addEventListener('click', () => {
          const id = item.getAttribute('data-id');
          if (!id) return;
          state.activeId = id;
          const material = state.results.find(entry => entry.id === id);
          if (material) {
            showDetail(material);
          }
          vscode.postMessage({ command: 'openMaterialDetail', payload: { id } });
          renderResults();
        });
      });

      resultList.querySelectorAll('button[data-action="copy"]').forEach(button => {
        button.addEventListener('click', event => {
          event.stopPropagation();
          const id = button.getAttribute('data-id');
          if (!id) return;
          vscode.postMessage({ command: 'copyMaterial', payload: { id } });
        });
      });

      resultList.querySelectorAll('button[data-action="insert"]').forEach(button => {
        button.addEventListener('click', event => {
          event.stopPropagation();
          const id = button.getAttribute('data-id');
          if (!id) return;
          vscode.postMessage({ command: 'insertMaterial', payload: { id } });
        });
      });
    }

    function showDetail(material) {
      detailContent.innerHTML = \`
        <div><strong>\${escapeHtml(material.name)}</strong></div>
        <div style="opacity:0.8;margin-top:6px;">\${escapeHtml(material.typeLabel)} · 来源：\${escapeHtml(material.source)}</div>
        <div style="margin-top:12px;white-space:pre-wrap;">\${escapeHtml(material.content)}</div>
      \`;
    }

    function runSearch(focusMaterialId) {
      vscode.postMessage({
        command: 'search',
        payload: {
          query: queryInput.value || '',
          types: Array.from(state.selectedTypes),
          focusMaterialId
        }
      });
    }

    searchBtn.addEventListener('click', () => runSearch());
    exportBtn.addEventListener('click', () => {
      vscode.postMessage({
        command: 'exportMarkdown',
        payload: {
          query: queryInput.value || '',
          types: Array.from(state.selectedTypes)
        }
      });
    });
    queryInput.addEventListener('keydown', event => {
      if (event.key === 'Enter') {
        runSearch();
      }
    });

    window.addEventListener('message', event => {
      const message = event.data;
      if (!message || !message.type) {
        return;
      }

      if (message.type === 'bootstrap') {
        state.availableTypes = message.availableTypes || [];
        const preset = message.preset || {};
        queryInput.value = preset.query || '';
        state.selectedTypes = new Set(preset.types || []);
        renderFilters();
      }

      if (message.type === 'applyPreset') {
        const preset = message.preset || {};
        queryInput.value = preset.query || '';
        state.selectedTypes = new Set(preset.types || []);
        renderFilters();
        runSearch(preset.focusMaterialId);
      }

      if (message.type === 'searchResults') {
        state.results = message.results || [];
        state.activeId = message.focusMaterialId || state.activeId;
        if (!state.activeId && state.results.length > 0) {
          state.activeId = state.results[0].id;
        }
        const active = state.results.find(item => item.id === state.activeId);
        if (active) {
          showDetail(active);
        }
        renderResults();
      }

      if (message.type === 'typeStats') {
        state.totalAll = message.totalAll || 0;
        state.totalFiltered = message.totalFiltered || 0;
        state.typeStats = message.stats || [];
        renderStats();
      }

      if (message.type === 'styleContext') {
        const styleName = message.styleName || '未设置';
        if (message.hasActiveStyle) {
          styleHint.textContent = '当前风格：' + styleName;
        } else {
          styleHint.textContent = '未激活风格，无法搜索素材。请先导入样本数据或切换风格。';
        }
      }

      if (message.type === 'materialDetail' && message.material) {
        state.activeId = message.material.id;
        showDetail(message.material);
        renderResults();
      }
    });
  </script>
</body>
</html>`;
  }

  private getStyleIdOrWarn(): string | null {
    if (this.options.styleId) {
      return this.options.styleId;
    }
    this.notifyWarn('当前没有激活风格，请先导入样本数据或切换风格。');
    return null;
  }

  private notifyInfo(message: string): void {
    vscode.window.setStatusBarMessage(`$(info) 写作助手 ${message}`, 5000);
  }

  private notifyWarn(message: string): void {
    vscode.window.setStatusBarMessage(`$(warning) 写作助手 ${message}`, 5000);
  }

  private dispose(): void {
    MaterialSearchPanel.currentPanel = undefined;
    while (this.disposables.length > 0) {
      const disposable = this.disposables.pop();
      disposable?.dispose();
    }
  }
}
