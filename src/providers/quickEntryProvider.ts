import * as vscode from 'vscode';

interface QuickEntryItem {
  id: string;
  label: string;
  description?: string;
  command: vscode.Command;
  iconId: string;
}

interface ApiStatusInfo {
  configured: boolean;
  text: string;
  iconId?: string;
}

export class QuickEntryProvider implements vscode.TreeDataProvider<QuickEntryItem> {
  private readonly changeEmitter = new vscode.EventEmitter<QuickEntryItem | undefined | null | void>();
  readonly onDidChangeTreeData = this.changeEmitter.event;

  constructor(private readonly getApiStatus: () => Promise<ApiStatusInfo>) {}

  getTreeItem(element: QuickEntryItem): vscode.TreeItem {
    const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
    item.id = element.id;
    item.description = element.description;
    item.tooltip = element.description || element.label;
    item.iconPath = new vscode.ThemeIcon(element.iconId);
    item.command = element.command;
    return item;
  }

  async getChildren(): Promise<QuickEntryItem[]> {
    const status = await this.getApiStatus();
    return [
      {
        id: 'iflow-config',
        label: '配置 API',
        description: status.text,
        iconId: status.iconId || (status.configured ? 'check' : 'warning'),
        command: {
          command: 'writingAgent.configureApi',
          title: '配置 API'
        }
      },
      {
        id: 'import',
        label: '导入样本数据',
        description: '学习写作风格',
        iconId: 'cloud-upload',
        command: {
          command: 'writingAgent.importArticle',
          title: '导入样本数据'
        }
      },
      {
        id: 'generate',
        label: '根据主题写作',
        description: '支持文章或多方案提纲',
        iconId: 'edit',
        command: {
          command: 'writingAgent.generateArticle',
          title: '生成风格化文章'
        }
      },
      {
        id: 'selection-rewrite',
        label: '改写选中文本',
        description: '默认 API，可切 IDE Chat',
        iconId: 'wand',
        command: {
          command: 'writingAgent.rewriteSelectionWithIFlow',
          title: 'AI编辑选中文本'
        }
      },
      {
        id: 'search',
        label: '搜索素材',
        description: '关键词 + 类型筛选',
        iconId: 'search',
        command: {
          command: 'writingAgent.searchMaterial',
          title: '搜索素材'
        }
      },
      {
        id: 'suggestion-min-chars',
        label: '补全触发字数',
        description: '设置自动补全最小触发字符数',
        iconId: 'symbol-number',
        command: {
          command: 'writingAgent.configureSuggestionMinChars',
          title: '设置自动补全最小触发字符数'
        }
      },
      {
        id: 'apply-ai-result',
        label: '写入 AI 结果',
        description: '把聊天复制内容落盘到当前文稿',
        iconId: 'clippy',
        command: {
          command: 'writingAgent.applyClipboardToDraft',
          title: '从剪贴板写入当前文稿'
        }
      },
      {
        id: 'material',
        label: '打开素材库',
        description: '类型统计 + 搜索 + 导出 MD',
        iconId: 'library',
        command: {
          command: 'writingAgent.searchMaterial',
          title: '打开素材库'
        }
      },
      {
        id: 'style',
        label: '打开风格配置',
        description: '切换与管理风格',
        iconId: 'symbol-class',
        command: {
          command: 'writingAgent.openStyleProfile',
          title: '打开风格配置'
        }
      }
    ];
  }

  refresh(): void {
    this.changeEmitter.fire();
  }
}
