import * as vscode from 'vscode';
import { v4 as uuidv4 } from 'uuid';

export interface ArticleRecord {
  id: string;
  topic: string;
  title: string;
  styleId: string;
  styleName: string;
  relativePath: string;
  aiCommand?: string;
  createdAt: Date;
  updatedAt: Date;
}

type SerializedArticleRecord = Omit<ArticleRecord, 'createdAt' | 'updatedAt'> & {
  createdAt: string;
  updatedAt: string;
};

export class ArticleRepository {
  private static readonly STORAGE_KEY = 'writingAgent.articleCollection.v1';
  private records: ArticleRecord[];

  constructor(private readonly context: vscode.ExtensionContext) {
    this.records = this.load();
  }

  listAll(): ArticleRecord[] {
    return [...this.records].sort((a, b) => b.updatedAt.getTime() - a.updatedAt.getTime());
  }

  listByStyle(styleId: string): ArticleRecord[] {
    return this.listAll().filter(record => record.styleId === styleId);
  }

  getById(id: string): ArticleRecord | undefined {
    return this.records.find(record => record.id === id);
  }

  async upsert(
    input: Omit<ArticleRecord, 'id' | 'createdAt' | 'updatedAt'> & {
      id?: string;
      createdAt?: Date;
      updatedAt?: Date;
    }
  ): Promise<ArticleRecord> {
    const now = new Date();
    const existing = this.records.find(record => record.relativePath === input.relativePath || record.id === input.id);

    if (existing) {
      existing.topic = input.topic;
      existing.title = input.title;
      existing.styleId = input.styleId;
      existing.styleName = input.styleName;
      existing.relativePath = input.relativePath;
      existing.aiCommand = input.aiCommand;
      existing.updatedAt = input.updatedAt || now;
      await this.persist();
      return existing;
    }

    const created: ArticleRecord = {
      id: input.id || uuidv4(),
      topic: input.topic,
      title: input.title,
      styleId: input.styleId,
      styleName: input.styleName,
      relativePath: input.relativePath,
      aiCommand: input.aiCommand,
      createdAt: input.createdAt || now,
      updatedAt: input.updatedAt || now
    };

    this.records.push(created);
    await this.persist();
    return created;
  }

  async touch(id: string): Promise<void> {
    const target = this.getById(id);
    if (!target) {
      return;
    }
    target.updatedAt = new Date();
    await this.persist();
  }

  async deleteById(id: string): Promise<ArticleRecord | undefined> {
    const index = this.records.findIndex(record => record.id === id);
    if (index < 0) {
      return undefined;
    }
    const [removed] = this.records.splice(index, 1);
    await this.persist();
    return removed;
  }

  async replacePathPrefix(fromPrefix: string, toPrefix: string): Promise<number> {
    const normalizedFrom = fromPrefix.replace(/^\/+/, '').replace(/\/+$/, '');
    const normalizedTo = toPrefix.replace(/^\/+/, '').replace(/\/+$/, '');
    let changed = 0;
    for (const record of this.records) {
      const relative = record.relativePath.replace(/^\/+/, '');
      if (!relative.startsWith(`${normalizedFrom}/`)) {
        continue;
      }
      record.relativePath = `${normalizedTo}/${relative.slice(normalizedFrom.length + 1)}`;
      record.updatedAt = new Date();
      changed += 1;
    }
    if (changed > 0) {
      await this.persist();
    }
    return changed;
  }

  async pruneMissingFiles(workspaceRoot: vscode.Uri): Promise<{ removed: number; checked: number }> {
    const kept: ArticleRecord[] = [];
    let removed = 0;

    for (const record of this.records) {
      const targetUri = vscode.Uri.joinPath(workspaceRoot, ...record.relativePath.split('/'));
      try {
        await vscode.workspace.fs.stat(targetUri);
        kept.push(record);
      } catch {
        removed += 1;
      }
    }

    if (removed > 0) {
      this.records = kept;
      await this.persist();
    }

    return { removed, checked: this.records.length + removed };
  }

  private load(): ArticleRecord[] {
    const raw = this.context.workspaceState.get<SerializedArticleRecord[]>(ArticleRepository.STORAGE_KEY, []);
    return raw.map(item => ({
      ...item,
      createdAt: new Date(item.createdAt),
      updatedAt: new Date(item.updatedAt)
    }));
  }

  private async persist(): Promise<void> {
    const payload: SerializedArticleRecord[] = this.records.map(record => ({
      ...record,
      createdAt: record.createdAt.toISOString(),
      updatedAt: record.updatedAt.toISOString()
    }));
    await this.context.workspaceState.update(ArticleRepository.STORAGE_KEY, payload);
  }
}
