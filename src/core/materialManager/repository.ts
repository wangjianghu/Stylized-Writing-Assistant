import * as vscode from 'vscode';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import { MaterialType, SearchQuery, WritingMaterial } from '../types';
import { TextPreprocessor } from '../../utils/textPreprocessor';
import { fileExists, readJsonFile, writeJsonFile } from '../storage/jsonFileStore';

type SerializedWritingMaterial = Omit<WritingMaterial, 'metadata'> & {
  metadata: Omit<WritingMaterial['metadata'], 'createdAt' | 'updatedAt' | 'lastUsedAt'> & {
    createdAt: string;
    updatedAt: string;
    lastUsedAt?: string;
  };
};

type SerializedMaterialBuckets = Record<string, SerializedWritingMaterial[]>;

/**
 * 素材仓库：按风格隔离素材持久化和检索
 */
export class MaterialRepository {
  private static readonly STORAGE_KEY = 'writingAgent.materials.byStyle.v1';
  private static readonly LEGACY_STORAGE_KEY = 'writingAgent.materials.v1';
  private static readonly LEGACY_BUCKET_ID = '__legacy__';

  private readonly preprocessor: TextPreprocessor;
  private readonly storageFilePath?: string;
  private materialsByStyle: Map<string, WritingMaterial[]>;

  constructor(private readonly context: vscode.ExtensionContext, storageRootPath?: string) {
    if (storageRootPath) {
      this.storageFilePath = path.join(storageRootPath, 'data', 'materials.by-style.json');
    }
    this.preprocessor = new TextPreprocessor();
    this.materialsByStyle = this.loadMaterialsByStyle();
    if (this.storageFilePath && !fileExists(this.storageFilePath) && this.materialsByStyle.size > 0) {
      void this.persist();
    }
  }

  getAll(styleId: string): WritingMaterial[] {
    return [...this.getBucket(styleId)];
  }

  getById(styleId: string, id: string): WritingMaterial | undefined {
    return this.getBucket(styleId).find(material => material.id === id);
  }

  listRecent(styleId: string, limit: number = 10): WritingMaterial[] {
    return [...this.getBucket(styleId)]
      .sort((a, b) => this.getTime(b.metadata.updatedAt) - this.getTime(a.metadata.updatedAt))
      .slice(0, Math.max(1, limit));
  }

  getTypeCounts(styleId: string): Map<MaterialType, number> {
    const counts = new Map<MaterialType, number>();
    for (const material of this.getBucket(styleId)) {
      counts.set(material.type, (counts.get(material.type) || 0) + 1);
    }
    return counts;
  }

  async addMany(styleId: string, materials: WritingMaterial[]): Promise<number> {
    const bucket = this.getBucket(styleId);
    let addedCount = 0;

    for (const material of materials) {
      const prepared = this.ensureMaterial(material);
      if (!this.isMeaningfulContent(prepared.content, 8)) {
        continue;
      }

      const existingIndex = this.findExistingIndex(bucket, prepared);
      if (existingIndex >= 0) {
        const existing = bucket[existingIndex];
        existing.metadata.updatedAt = new Date();
        existing.metadata.quality = Math.max(existing.metadata.quality, prepared.metadata.quality);
        existing.tags = this.mergeTags(existing.tags, prepared.tags);
        existing.features.keywords = this.mergeTags(existing.features.keywords, prepared.features.keywords);
      } else {
        bucket.push(prepared);
        addedCount += 1;
      }
    }

    if (addedCount > 0 || materials.length > 0) {
      await this.persist();
    }
    return addedCount;
  }

  async addOneFromSelection(styleId: string, content: string, source: string): Promise<WritingMaterial> {
    const cleanContent = content.trim();
    const type = cleanContent.length > 80 || cleanContent.includes('\n')
      ? MaterialType.PARAGRAPH
      : MaterialType.SENTENCE;
    const keywords = this.preprocessor.extractKeywords(cleanContent, 6);
    const quality = this.scoreSelectedMaterial(cleanContent);
    const sentiment = this.detectSentiment(cleanContent);

    const material: WritingMaterial = {
      id: uuidv4(),
      name: this.buildName(cleanContent),
      content: cleanContent,
      type,
      tags: this.mergeTags(keywords.slice(0, 4), ['手动保存']),
      category: type === MaterialType.PARAGRAPH ? 'paragraph' : 'sentence',
      metadata: {
        source,
        createdAt: new Date(),
        updatedAt: new Date(),
        usedCount: 0,
        quality
      },
      features: {
        keywords,
        sentiment,
        style: 'manual',
        quality
      }
    };

    const addedCount = await this.addMany(styleId, [material]);
    if (addedCount > 0) {
      return material;
    }

    const existing = this.findByFingerprint(styleId, this.materialFingerprint(material));
    return existing ?? material;
  }

  async markUsed(styleId: string, id: string): Promise<void> {
    const material = this.getById(styleId, id);
    if (!material) {
      return;
    }

    material.metadata.usedCount += 1;
    material.metadata.lastUsedAt = new Date();
    material.metadata.updatedAt = new Date();
    await this.persist();
  }

  search(styleId: string, query: SearchQuery): WritingMaterial[] {
    const materials = this.getBucket(styleId);
    const limit = Math.max(1, query.limit || 50);
    const normalizedKeywords = this.normalizeKeywords(query.keywords, query.semantic);
    const tagFilters = (query.tags || []).map(tag => tag.toLowerCase().trim()).filter(Boolean);
    const typeSet = query.types && query.types.length > 0 ? new Set(query.types) : undefined;

    const scored: Array<{ material: WritingMaterial; score: number }> = [];

    for (const material of materials) {
      if (typeSet && !typeSet.has(material.type)) {
        continue;
      }

      if (tagFilters.length > 0) {
        const tags = material.tags.map(tag => tag.toLowerCase());
        const hasTag = tagFilters.some(tag => tags.includes(tag));
        if (!hasTag) {
          continue;
        }
      }

      const baseScore = material.metadata.quality / 100;
      let score = baseScore;

      if (normalizedKeywords.length > 0) {
        const keywordScore = this.scoreByKeywords(material, normalizedKeywords);
        if (keywordScore <= 0) {
          continue;
        }
        score += keywordScore;
      }

      scored.push({ material, score });
    }

    return scored
      .sort((a, b) => {
        if (b.score !== a.score) {
          return b.score - a.score;
        }
        if (b.material.metadata.quality !== a.material.metadata.quality) {
          return b.material.metadata.quality - a.material.metadata.quality;
        }
        return this.getTime(b.material.metadata.updatedAt) - this.getTime(a.material.metadata.updatedAt);
      })
      .slice(0, limit)
      .map(entry => entry.material);
  }

  async deleteStyleMaterials(styleId: string): Promise<void> {
    this.materialsByStyle.delete(styleId);
    await this.persist();
  }

  async moveStyleMaterials(
    sourceStyleId: string,
    targetStyleId: string
  ): Promise<{ moved: number; merged: number }> {
    if (!sourceStyleId || !targetStyleId || sourceStyleId === targetStyleId) {
      return { moved: 0, merged: 0 };
    }

    const sourceBucket = this.materialsByStyle.get(sourceStyleId);
    if (!sourceBucket || sourceBucket.length === 0) {
      this.materialsByStyle.delete(sourceStyleId);
      await this.persist();
      return { moved: 0, merged: 0 };
    }

    const targetBucket = this.getBucket(targetStyleId);
    let moved = 0;
    let merged = 0;

    for (const material of sourceBucket) {
      const prepared = this.ensureMaterial(material);
      const existingIndex = this.findExistingIndex(targetBucket, prepared);
      if (existingIndex >= 0) {
        const existing = targetBucket[existingIndex];
        existing.metadata.updatedAt = new Date();
        existing.metadata.quality = Math.max(existing.metadata.quality, prepared.metadata.quality);
        existing.tags = this.mergeTags(existing.tags, prepared.tags);
        existing.features.keywords = this.mergeTags(existing.features.keywords, prepared.features.keywords);
        merged += 1;
        continue;
      }
      targetBucket.push(prepared);
      moved += 1;
    }

    this.materialsByStyle.delete(sourceStyleId);
    await this.persist();
    return { moved, merged };
  }

  private getBucket(styleId: string): WritingMaterial[] {
    const existing = this.materialsByStyle.get(styleId);
    if (existing) {
      const cleaned = existing.filter(material => this.isMeaningfulContent(material.content, 8));
      if (cleaned.length !== existing.length) {
        this.materialsByStyle.set(styleId, cleaned);
        void this.persist();
        return cleaned;
      }
      return existing;
    }

    const legacyBucket = this.materialsByStyle.get(MaterialRepository.LEGACY_BUCKET_ID);
    if (legacyBucket && legacyBucket.length > 0) {
      this.materialsByStyle.set(styleId, legacyBucket);
      this.materialsByStyle.delete(MaterialRepository.LEGACY_BUCKET_ID);
      void this.persist();
      return legacyBucket;
    }

    if (!this.materialsByStyle.has(styleId)) {
      this.materialsByStyle.set(styleId, []);
    }
    return this.materialsByStyle.get(styleId)!;
  }

  private scoreByKeywords(material: WritingMaterial, keywords: string[]): number {
    const name = material.name.toLowerCase();
    const content = material.content.toLowerCase();
    const tags = material.tags.map(tag => tag.toLowerCase());
    const featureKeywords = material.features.keywords.map(keyword => keyword.toLowerCase());

    let score = 0;
    for (const keyword of keywords) {
      let hitScore = 0;
      if (name.includes(keyword)) {
        hitScore += 4;
      }
      if (content.includes(keyword)) {
        hitScore += 3;
      }
      if (tags.some(tag => tag.includes(keyword))) {
        hitScore += 2;
      }
      if (featureKeywords.some(item => item.includes(keyword))) {
        hitScore += 2;
      }
      score += hitScore;
    }
    return score;
  }

  private normalizeKeywords(keywords?: string[], semantic?: string): string[] {
    const list: string[] = [];
    if (Array.isArray(keywords)) {
      list.push(...keywords);
    }
    if (semantic) {
      list.push(...semantic.split(/[\s,，;；]+/));
    }
    return list
      .map(keyword => keyword.toLowerCase().trim())
      .filter(Boolean);
  }

  private buildName(content: string): string {
    const firstLine = content.split('\n')[0].trim();
    const clipped = firstLine.length > 24 ? `${firstLine.slice(0, 24)}...` : firstLine;
    return clipped || '未命名素材';
  }

  private scoreSelectedMaterial(content: string): number {
    let score = 55;
    if (content.length > 30) {
      score += 10;
    }
    if (content.length > 100) {
      score += 10;
    }
    if (/[。！？!?]/.test(content)) {
      score += 10;
    }
    if (/[，,；;]/.test(content)) {
      score += 5;
    }
    return Math.min(score, 95);
  }

  private detectSentiment(text: string): string {
    const positiveWords = ['好', '棒', '优秀', '喜欢', '开心', '快乐', '成功', '美好'];
    const negativeWords = ['坏', '差', '糟糕', '讨厌', '难过', '悲伤', '失败', '痛苦'];

    let positiveCount = 0;
    let negativeCount = 0;

    for (const word of positiveWords) {
      if (text.includes(word)) {
        positiveCount += 1;
      }
    }
    for (const word of negativeWords) {
      if (text.includes(word)) {
        negativeCount += 1;
      }
    }

    if (positiveCount > negativeCount) {
      return 'positive';
    }
    if (negativeCount > positiveCount) {
      return 'negative';
    }
    return 'neutral';
  }

  private mergeTags(left: string[], right: string[]): string[] {
    const merged = new Set<string>();
    for (const item of left) {
      merged.add(item);
    }
    for (const item of right) {
      merged.add(item);
    }
    return [...merged].slice(0, 12);
  }

  private ensureMaterial(material: WritingMaterial): WritingMaterial {
    return {
      ...material,
      id: material.id || uuidv4(),
      name: material.name || this.buildName(material.content),
      tags: material.tags || [],
      metadata: {
        ...material.metadata,
        source: material.metadata?.source || 'unknown',
        createdAt: material.metadata?.createdAt ? new Date(material.metadata.createdAt) : new Date(),
        updatedAt: material.metadata?.updatedAt ? new Date(material.metadata.updatedAt) : new Date(),
        usedCount: material.metadata?.usedCount || 0,
        lastUsedAt: material.metadata?.lastUsedAt ? new Date(material.metadata.lastUsedAt) : undefined,
        quality: material.metadata?.quality || 60
      },
      features: {
        ...material.features,
        keywords: material.features?.keywords || [],
        sentiment: material.features?.sentiment || 'neutral',
        style: material.features?.style || 'unknown',
        quality: material.features?.quality || material.metadata?.quality || 60
      }
    };
  }

  private findExistingIndex(bucket: WritingMaterial[], material: WritingMaterial): number {
    const fingerprint = this.materialFingerprint(material);
    return bucket.findIndex(existing =>
      existing.id === material.id || this.materialFingerprint(existing) === fingerprint
    );
  }

  private findByFingerprint(styleId: string, fingerprint: string): WritingMaterial | undefined {
    return this.getBucket(styleId).find(material => this.materialFingerprint(material) === fingerprint);
  }

  private materialFingerprint(material: WritingMaterial): string {
    const source = material.metadata?.source || 'unknown';
    return `${material.type}|${source}|${material.content.trim()}`;
  }

  private getTime(date: Date): number {
    return date instanceof Date ? date.getTime() : new Date(date).getTime();
  }

  private loadMaterialsByStyle(): Map<string, WritingMaterial[]> {
    const fromBuckets = this.storageFilePath
      ? readJsonFile<SerializedMaterialBuckets>(
        this.storageFilePath,
        this.context.workspaceState.get<SerializedMaterialBuckets>(
          MaterialRepository.STORAGE_KEY,
          {}
        )
      )
      : this.context.workspaceState.get<SerializedMaterialBuckets>(
        MaterialRepository.STORAGE_KEY,
        {}
      );

    const hasBuckets = Object.keys(fromBuckets).length > 0;
    if (hasBuckets) {
      return new Map<string, WritingMaterial[]>(
        Object.entries(fromBuckets).map(([styleId, items]) => [
          styleId,
          (items || []).map(item => this.deserialize(item))
        ])
      );
    }

    // 兼容旧版单桶数据结构，迁移到 legacy 分桶。
    const legacy = this.context.workspaceState.get<SerializedWritingMaterial[]>(
      MaterialRepository.LEGACY_STORAGE_KEY,
      []
    );
    if (legacy.length === 0) {
      return new Map<string, WritingMaterial[]>();
    }

    return new Map<string, WritingMaterial[]>([
      [
        MaterialRepository.LEGACY_BUCKET_ID,
        legacy.map(item => this.deserialize(item))
      ]
    ]);
  }

  private async persist(): Promise<void> {
    const payload: SerializedMaterialBuckets = {};
    this.materialsByStyle.forEach((materials, styleId) => {
      payload[styleId] = materials.map(material => this.serialize(material));
    });
    if (this.storageFilePath) {
      writeJsonFile(this.storageFilePath, payload);
      return;
    }
    await this.context.workspaceState.update(MaterialRepository.STORAGE_KEY, payload);
  }

  private serialize(material: WritingMaterial): SerializedWritingMaterial {
    return {
      ...material,
      metadata: {
        ...material.metadata,
        createdAt: material.metadata.createdAt.toISOString(),
        updatedAt: material.metadata.updatedAt.toISOString(),
        lastUsedAt: material.metadata.lastUsedAt?.toISOString()
      }
    };
  }

  private deserialize(material: SerializedWritingMaterial): WritingMaterial {
    return {
      ...material,
      metadata: {
        ...material.metadata,
        createdAt: new Date(material.metadata.createdAt),
        updatedAt: new Date(material.metadata.updatedAt),
        lastUsedAt: material.metadata.lastUsedAt ? new Date(material.metadata.lastUsedAt) : undefined
      }
    };
  }

  private isMeaningfulContent(content: string, minCoreChars: number): boolean {
    const compact = (content || '').replace(/\s+/g, '');
    if (!compact) {
      return false;
    }

    if (/^[-_*`#>~|]+$/.test(compact)) {
      return false;
    }

    const core = compact.replace(/[^\p{L}\p{N}\u4E00-\u9FFF]/gu, '');
    return core.length >= minCoreChars;
  }
}
