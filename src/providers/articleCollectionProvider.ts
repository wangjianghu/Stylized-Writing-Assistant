import * as vscode from 'vscode';
import { ArticleRecord, ArticleRepository } from '../core/articleManager/repository';

type ActiveStyle = { id: string; name: string };
type CollectionNodeKind = 'summary' | 'recent-section' | 'style-section' | 'article-item';

interface CollectionNode {
  kind: CollectionNodeKind;
  label: string;
  description?: string;
  styleId?: string;
  styleCount?: number;
  isActiveStyle?: boolean;
  record?: ArticleRecord;
}

export class ArticleCollectionProvider implements vscode.TreeDataProvider<CollectionNode> {
  private readonly onDidChangeTreeDataEmitter = new vscode.EventEmitter<CollectionNode | undefined | null | void>();
  readonly onDidChangeTreeData = this.onDidChangeTreeDataEmitter.event;

  constructor(
    private readonly repository: ArticleRepository,
    private readonly resolveActiveStyle: () => ActiveStyle | null,
    private readonly resolveUriByRelativePath: (relativePath: string) => vscode.Uri | null
  ) {}

  refresh(): void {
    this.onDidChangeTreeDataEmitter.fire();
  }

  getTreeItem(element: CollectionNode): vscode.TreeItem {
    if (element.kind === 'summary') {
      const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
      item.description = element.description;
      item.iconPath = new vscode.ThemeIcon('book');
      item.tooltip = element.description || element.label;
      return item;
    }

    if (element.kind === 'recent-section') {
      const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.Expanded);
      item.description = element.description;
      item.iconPath = new vscode.ThemeIcon('history');
      item.tooltip = element.description || element.label;
      return item;
    }

    if (element.kind === 'style-section') {
      const item = new vscode.TreeItem(
        element.label,
        element.isActiveStyle ? vscode.TreeItemCollapsibleState.Expanded : vscode.TreeItemCollapsibleState.Collapsed
      );
      item.description = `${element.styleCount || 0} 篇`;
      item.iconPath = new vscode.ThemeIcon(element.isActiveStyle ? 'star-full' : 'symbol-key');
      item.contextValue = 'writingAgent.articleStyleItem';
      item.tooltip = element.description || item.description || element.label;
      if (element.styleId && !element.isActiveStyle) {
        item.command = {
          command: 'writingAgent.switchStyle',
          title: '切换风格',
          arguments: [{ styleId: element.styleId }]
        };
      }
      return item;
    }

    const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
    item.description = element.description;
    item.iconPath = new vscode.ThemeIcon('markdown');
    item.contextValue = 'writingAgent.articleItem';

    if (element.record) {
      const uri = this.resolveUriByRelativePath(element.record.relativePath);
      if (uri) {
        item.command = {
          command: 'vscode.open',
          title: '打开文稿',
          arguments: [uri]
        };
        item.resourceUri = uri;
        const baseTooltip = element.description || element.record.styleName || element.label;
        item.tooltip = `${baseTooltip}\n${element.record.relativePath}`;
      } else {
        item.tooltip = '未检测到工作区，无法定位文稿文件';
      }
    } else {
      item.tooltip = element.description || element.label;
    }

    return item;
  }

  getChildren(element?: CollectionNode): CollectionNode[] {
    const activeStyle = this.resolveActiveStyle();
    const all = this.repository.listAll();

    if (!element) {
      if (all.length === 0) {
        return [];
      }

      const currentStyleCount = activeStyle ? this.repository.listByStyle(activeStyle.id).length : 0;
      const summaryText = activeStyle ? `当前风格「${activeStyle.name}」${currentStyleCount} 篇` : `共 ${all.length} 篇`;

      const styleMap = new Map<string, { styleName: string; count: number; latestUpdatedAt: number }>();
      for (const record of all) {
        const existing = styleMap.get(record.styleId);
        if (!existing) {
          styleMap.set(record.styleId, {
            styleName: record.styleName,
            count: 1,
            latestUpdatedAt: record.updatedAt.getTime()
          });
          continue;
        }
        existing.count += 1;
        existing.latestUpdatedAt = Math.max(existing.latestUpdatedAt, record.updatedAt.getTime());
      }

      const styleSections: CollectionNode[] = Array.from(styleMap.entries())
        .map(([styleId, info]): CollectionNode => {
          return {
            kind: 'style-section',
            label: activeStyle?.id === styleId ? `当前风格：${info.styleName}` : info.styleName,
            styleId,
            styleCount: info.count,
            isActiveStyle: activeStyle?.id === styleId,
            description: `最近更新：${new Date(info.latestUpdatedAt).toLocaleString()}`
          };
        })
        .sort((a, b) => {
          if (a.isActiveStyle && !b.isActiveStyle) {
            return -1;
          }
          if (!a.isActiveStyle && b.isActiveStyle) {
            return 1;
          }
          return (b.styleCount || 0) - (a.styleCount || 0);
        });

      return [
        {
          kind: 'summary',
          label: `文章总数 ${all.length}`,
          description: summaryText
        },
        {
          kind: 'recent-section',
          label: '最近生成',
          description: `${Math.min(8, all.length)} 篇`
        },
        ...styleSections
      ];
    }

    if (element.kind === 'recent-section') {
      return all.slice(0, 8).map(record => ({
        kind: 'article-item',
        label: record.title,
        description: `${record.styleName} · ${record.updatedAt.toLocaleString()}`,
        record
      }));
    }

    if (element.kind === 'style-section') {
      if (!element.styleId) {
        return [];
      }
      return this.repository.listByStyle(element.styleId).map(record => ({
        kind: 'article-item',
        label: record.title,
        description: record.updatedAt.toLocaleString(),
        record
      }));
    }

    return [];
  }
}
