import * as vscode from 'vscode';
import { MaterialType, WritingMaterial } from '../core/types';
import { MaterialRepository } from '../core/materialManager/repository';

type RootNodeKind = 'summary' | 'types-section' | 'recent-section';
type NodeKind = RootNodeKind | 'type-item' | 'material-item';
type ActiveStyle = { id: string; name: string };

interface MaterialNode {
  kind: NodeKind;
  label: string;
  description?: string;
  material?: WritingMaterial;
  type?: MaterialType;
  count?: number;
}

export class MaterialLibraryProvider implements vscode.TreeDataProvider<MaterialNode> {
  private readonly onDidChangeTreeDataEmitter = new vscode.EventEmitter<MaterialNode | undefined | null | void>();
  readonly onDidChangeTreeData = this.onDidChangeTreeDataEmitter.event;

  constructor(
    private readonly repository: MaterialRepository,
    private readonly resolveActiveStyle: () => ActiveStyle | null
  ) {}

  refresh(): void {
    this.onDidChangeTreeDataEmitter.fire();
  }

  getTreeItem(element: MaterialNode): vscode.TreeItem {
    switch (element.kind) {
      case 'summary': {
        const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
        item.description = element.description;
        item.iconPath = new vscode.ThemeIcon('library');
        item.tooltip = element.description || element.label;
        return item;
      }
      case 'types-section':
      case 'recent-section': {
        const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.Expanded);
        item.description = element.description;
        item.iconPath = new vscode.ThemeIcon(element.kind === 'types-section' ? 'symbol-enum' : 'history');
        item.tooltip = element.description || element.label;
        return item;
      }
      case 'type-item': {
        const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
        item.description = element.description;
        item.iconPath = new vscode.ThemeIcon('tag');
        item.tooltip = element.description || element.label;
        item.command = {
          command: 'writingAgent.searchMaterial',
          title: '搜索素材',
          arguments: [{ presetTypes: element.type ? [element.type] : [] }]
        };
        return item;
      }
      case 'material-item': {
        const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
        item.description = element.description;
        const baseTooltip = element.description || element.label;
        const preview = element.material?.content?.trim();
        item.tooltip = preview ? `${baseTooltip}\n${preview}` : baseTooltip;
        item.iconPath = new vscode.ThemeIcon('note');
        item.command = {
          command: 'writingAgent.searchMaterial',
          title: '查看素材详情',
          arguments: [{ materialId: element.material?.id }]
        };
        return item;
      }
      default:
        return new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
    }
  }

  getChildren(element?: MaterialNode): MaterialNode[] {
    const activeStyle = this.resolveActiveStyle();
    if (!activeStyle) {
      return [];
    }

    if (!element) {
      const total = this.repository.getAll(activeStyle.id).length;
      if (total === 0) {
        return [];
      }
      const recentCount = this.repository.listRecent(activeStyle.id, 8).length;
      return [
        {
          kind: 'summary',
          label: `${activeStyle.name} · 素材总数 ${total}`,
          description: '当前风格'
        },
        {
          kind: 'types-section',
          label: '类型统计',
          description: `${this.repository.getTypeCounts(activeStyle.id).size} 类`
        },
        {
          kind: 'recent-section',
          label: '最近素材',
          description: `${recentCount} 条`
        }
      ];
    }

    if (element.kind === 'types-section') {
      return [...this.repository.getTypeCounts(activeStyle.id).entries()]
        .sort((a, b) => b[1] - a[1])
        .map(([type, count]) => ({
          kind: 'type-item',
          type,
          count,
          label: this.getTypeName(type),
          description: `${count} 条`
        }));
    }

    if (element.kind === 'recent-section') {
      return this.repository.listRecent(activeStyle.id, 8).map(material => ({
        kind: 'material-item',
        material,
        label: material.name,
        description: `${this.getTypeName(material.type)} · 质量 ${material.metadata.quality}`
      }));
    }

    return [];
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
}
