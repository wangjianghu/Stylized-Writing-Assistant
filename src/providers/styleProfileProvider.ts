import * as vscode from 'vscode';
import { StyleProfile } from '../core/types';
import { StyleRepository } from '../core/styleManager/repository';

type StyleNodeKind = 'current-summary' | 'summary-metric' | 'profiles-section' | 'profile-item';

interface StyleNode {
  kind: StyleNodeKind;
  label: string;
  description?: string;
  iconId?: string;
  profileId?: string;
}

export class StyleProfileProvider implements vscode.TreeDataProvider<StyleNode> {
  private readonly onDidChangeTreeDataEmitter = new vscode.EventEmitter<StyleNode | undefined | null | void>();
  readonly onDidChangeTreeData = this.onDidChangeTreeDataEmitter.event;

  constructor(private readonly styleRepository: StyleRepository) {}

  refresh(): void {
    this.onDidChangeTreeDataEmitter.fire();
  }

  getTreeItem(element: StyleNode): vscode.TreeItem {
    if (element.kind === 'current-summary') {
      const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.Expanded);
      item.description = element.description;
      item.iconPath = new vscode.ThemeIcon('star-full');
      item.tooltip = element.description || element.label;
      return item;
    }

    if (element.kind === 'profiles-section') {
      const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.Expanded);
      item.description = element.description;
      item.iconPath = new vscode.ThemeIcon('list-unordered');
      item.tooltip = element.description || element.label;
      return item;
    }

    const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
    item.description = element.description;
    if (element.iconId) {
      item.iconPath = new vscode.ThemeIcon(element.iconId);
    }
    item.tooltip = element.description || element.label;

    if (element.kind === 'profile-item' && element.profileId) {
      item.command = {
        command: 'writingAgent.switchStyle',
        title: '切换风格',
        arguments: [{ styleId: element.profileId }]
      };
    }
    return item;
  }

  getChildren(element?: StyleNode): StyleNode[] {
    const profiles = this.styleRepository.listProfiles();
    const activeProfile = this.styleRepository.getActiveProfile();

    if (!element) {
      if (profiles.length === 0) {
        return [];
      }
      return [
        {
          kind: 'current-summary',
          label: activeProfile ? `当前风格：${activeProfile.name}` : '当前风格：未设置',
          description: activeProfile ? `${activeProfile.articleCount} 篇样本` : '请选择或创建风格'
        },
        {
          kind: 'profiles-section',
          label: '风格列表',
          description: `${profiles.length} 个`
        }
      ];
    }

    if (element.kind === 'current-summary') {
      return this.getActiveSummaryNodes(activeProfile);
    }

    if (element.kind === 'profiles-section') {
      const activeId = this.styleRepository.getActiveProfileId();
      return profiles.map(profile => ({
        kind: 'profile-item' as const,
        profileId: profile.id,
        label: profile.id === activeId ? `${profile.name} (当前)` : profile.name,
        description: `${profile.articleCount} 篇 · ${profile.updatedAt.toLocaleDateString()}`,
        iconId: profile.id === activeId ? 'check' : 'circle-large-outline'
      }));
    }

    return [];
  }

  private getActiveSummaryNodes(activeProfile: StyleProfile | null): StyleNode[] {
    if (!activeProfile) {
      return [
        {
          kind: 'summary-metric',
          label: '学习状态',
          description: '未学习',
          iconId: 'warning'
        }
      ];
    }

    return [
      {
        kind: 'summary-metric',
        label: '学习状态',
        description: '已学习',
        iconId: 'check'
      },
      {
        kind: 'summary-metric',
        label: '语气类型',
        description: activeProfile.style.tone.type,
        iconId: 'comment-discussion'
      },
      {
        kind: 'summary-metric',
        label: '平均句长',
        description: `${activeProfile.style.sentenceStructure.avgLength.toFixed(1)} 字`,
        iconId: 'list-unordered'
      },
      {
        kind: 'summary-metric',
        label: '词汇丰富度',
        description: `${(activeProfile.style.vocabulary.uniqueWordRatio * 100).toFixed(1)}%`,
        iconId: 'symbol-key'
      },
      {
        kind: 'summary-metric',
        label: '样本数量',
        description: `${activeProfile.articleCount} 篇`,
        iconId: 'book'
      },
      {
        kind: 'summary-metric',
        label: '更新时间',
        description: activeProfile.updatedAt.toLocaleString(),
        iconId: 'history'
      }
    ];
  }
}
