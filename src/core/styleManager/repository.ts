import * as vscode from 'vscode';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import { StyleProfile, WritingStyle } from '../types';
import { fileExists, readJsonFile, writeJsonFile } from '../storage/jsonFileStore';

type SerializedStyleProfile = Omit<StyleProfile, 'createdAt' | 'updatedAt' | 'style'> & {
  createdAt: string;
  updatedAt: string;
  style: SerializedWritingStyle;
};

type SerializedWritingStyle = Omit<WritingStyle, 'createdAt' | 'updatedAt' | 'sentenceStructure' | 'timePerspective'> & {
  sentenceStructure: Omit<WritingStyle['sentenceStructure'], 'lengthDistribution'> & {
    lengthDistribution: Record<string, number>;
  };
  timePerspective: Omit<WritingStyle['timePerspective'], 'tenseUsage'> & {
    tenseUsage: Record<string, number>;
  };
  createdAt?: string;
  updatedAt?: string;
};

export class StyleRepository {
  private static readonly STYLES_KEY = 'writingAgent.styles.v1';
  private static readonly ACTIVE_STYLE_ID_KEY = 'writingAgent.styles.activeId';
  private readonly stylesFilePath?: string;
  private readonly stateFilePath?: string;
  private loaded = false;
  private profiles: StyleProfile[];
  private activeProfileId: string | null;

  constructor(private readonly context: vscode.ExtensionContext, storageRootPath?: string) {
    if (storageRootPath) {
      const dataDir = path.join(storageRootPath, 'data');
      this.stylesFilePath = path.join(dataDir, 'styles.json');
      this.stateFilePath = path.join(dataDir, 'styles.state.json');
    }
    this.profiles = [];
    this.activeProfileId = null;
  }

  listProfiles(): StyleProfile[] {
    this.ensureLoaded();
    return [...this.profiles].sort((a, b) => b.updatedAt.getTime() - a.updatedAt.getTime());
  }

  getProfile(id: string): StyleProfile | undefined {
    this.ensureLoaded();
    return this.profiles.find(profile => profile.id === id);
  }

  getActiveProfileId(): string | null {
    this.ensureLoaded();
    return this.activeProfileId;
  }

  getActiveProfile(): StyleProfile | null {
    this.ensureLoaded();
    if (!this.activeProfileId) {
      return null;
    }
    return this.getProfile(this.activeProfileId) || null;
  }

  async setActiveProfile(id: string): Promise<void> {
    this.ensureLoaded();
    const profile = this.getProfile(id);
    if (!profile) {
      throw new Error('目标风格不存在');
    }
    this.activeProfileId = id;
    await this.persistActiveProfileState();
  }

  async createProfile(name: string, style?: WritingStyle): Promise<StyleProfile> {
    this.ensureLoaded();
    const now = new Date();
    const uniqueName = this.toUniqueName(name || '新风格');
    const profile: StyleProfile = {
      id: uuidv4(),
      name: uniqueName,
      style: style ? this.normalizeStyle(style) : this.createEmptyStyle(),
      createdAt: now,
      updatedAt: now,
      articleCount: 0
    };
    this.profiles.push(profile);
    await this.persistProfiles();
    return profile;
  }

  async updateProfile(id: string, style: WritingStyle, source?: string): Promise<StyleProfile> {
    this.ensureLoaded();
    const profile = this.getProfile(id);
    if (!profile) {
      throw new Error('目标风格不存在');
    }

    profile.style = this.normalizeStyle(style);
    profile.updatedAt = new Date();
    profile.articleCount += 1;
    profile.lastSource = source || profile.lastSource;
    await this.persistProfiles();
    return profile;
  }

  async refreshProfileStyle(id: string, style: WritingStyle, source?: string): Promise<StyleProfile> {
    this.ensureLoaded();
    const profile = this.getProfile(id);
    if (!profile) {
      throw new Error('目标风格不存在');
    }

    profile.style = this.normalizeStyle(style);
    profile.updatedAt = new Date();
    profile.lastSource = source || profile.lastSource;
    await this.persistProfiles();
    return profile;
  }

  async renameProfile(id: string, name: string): Promise<void> {
    this.ensureLoaded();
    const profile = this.getProfile(id);
    if (!profile) {
      throw new Error('目标风格不存在');
    }

    const trimmedName = name.trim();
    if (!trimmedName) {
      throw new Error('风格名称不能为空');
    }
    profile.name = this.toUniqueName(trimmedName, id);
    profile.updatedAt = new Date();
    await this.persistProfiles();
  }

  async deleteProfile(id: string): Promise<void> {
    this.ensureLoaded();
    const existingIndex = this.profiles.findIndex(profile => profile.id === id);
    if (existingIndex < 0) {
      throw new Error('目标风格不存在');
    }

    this.profiles.splice(existingIndex, 1);

    if (this.activeProfileId === id) {
      const fallback = this.listProfiles()[0];
      this.activeProfileId = fallback?.id || null;
      await this.persistActiveProfileState();
    }

    await this.persistProfiles();
  }

  private ensureLoaded(): void {
    if (this.loaded) {
      return;
    }
    this.loaded = true;
    this.profiles = this.loadProfiles();
    this.activeProfileId = this.loadActiveProfileId();
    this.ensureValidActiveProfile();
    if (this.stylesFilePath && !fileExists(this.stylesFilePath) && this.profiles.length > 0) {
      void this.persistProfiles();
    }
    if (this.stateFilePath && !fileExists(this.stateFilePath) && this.activeProfileId) {
      void this.persistActiveProfileState();
    }
  }

  private ensureValidActiveProfile(): void {
    if (!this.activeProfileId) {
      return;
    }
    const exists = this.profiles.some(profile => profile.id === this.activeProfileId);
    if (!exists) {
      this.activeProfileId = null;
      void this.persistActiveProfileState();
    }
  }

  private toUniqueName(name: string, excludedId?: string): string {
    const base = name.trim() || '新风格';
    const normalized = base.toLowerCase();
    const hasName = (targetName: string): boolean => {
      const normalizedTarget = targetName.toLowerCase();
      return this.profiles.some(profile =>
        profile.id !== excludedId && profile.name.toLowerCase() === normalizedTarget
      );
    };

    if (!hasName(normalized)) {
      return base;
    }

    let index = 2;
    while (hasName(`${base} (${index})`)) {
      index += 1;
    }
    return `${base} (${index})`;
  }

  private loadProfiles(): StyleProfile[] {
    const rawProfiles = this.stylesFilePath
      ? readJsonFile<SerializedStyleProfile[]>(
        this.stylesFilePath,
        this.context.workspaceState.get<SerializedStyleProfile[]>(StyleRepository.STYLES_KEY, [])
      )
      : this.context.workspaceState.get<SerializedStyleProfile[]>(StyleRepository.STYLES_KEY, []);
    return rawProfiles.map(raw => this.deserializeProfile(raw));
  }

  private async persistProfiles(): Promise<void> {
    const rawProfiles = this.profiles.map(profile => this.serializeProfile(profile));
    if (this.stylesFilePath) {
      writeJsonFile(this.stylesFilePath, rawProfiles);
      return;
    }
    await this.context.workspaceState.update(StyleRepository.STYLES_KEY, rawProfiles);
  }

  private loadActiveProfileId(): string | null {
    if (this.stateFilePath) {
      const state = readJsonFile<{ activeProfileId?: string | null }>(
        this.stateFilePath,
        {
          activeProfileId: this.context.workspaceState.get<string | null>(
            StyleRepository.ACTIVE_STYLE_ID_KEY,
            null
          )
        }
      );
      return state.activeProfileId || null;
    }
    return this.context.workspaceState.get<string | null>(StyleRepository.ACTIVE_STYLE_ID_KEY, null);
  }

  private async persistActiveProfileState(): Promise<void> {
    if (this.stateFilePath) {
      writeJsonFile(this.stateFilePath, { activeProfileId: this.activeProfileId });
      return;
    }
    await this.context.workspaceState.update(StyleRepository.ACTIVE_STYLE_ID_KEY, this.activeProfileId);
  }

  private serializeProfile(profile: StyleProfile): SerializedStyleProfile {
    return {
      ...profile,
      createdAt: profile.createdAt.toISOString(),
      updatedAt: profile.updatedAt.toISOString(),
      style: this.serializeStyle(profile.style)
    };
  }

  private deserializeProfile(raw: SerializedStyleProfile): StyleProfile {
    return {
      ...raw,
      createdAt: new Date(raw.createdAt),
      updatedAt: new Date(raw.updatedAt),
      style: this.deserializeStyle(raw.style)
    };
  }

  private serializeStyle(style: WritingStyle): SerializedWritingStyle {
    return {
      ...style,
      createdAt: style.createdAt?.toISOString(),
      updatedAt: style.updatedAt?.toISOString(),
      sentenceStructure: {
        ...style.sentenceStructure,
        lengthDistribution: this.mapNumberToRecord(style.sentenceStructure.lengthDistribution)
      },
      timePerspective: {
        ...style.timePerspective,
        tenseUsage: this.mapStringToRecord(style.timePerspective.tenseUsage)
      }
    };
  }

  private deserializeStyle(raw: SerializedWritingStyle): WritingStyle {
    return {
      ...raw,
      createdAt: raw.createdAt ? new Date(raw.createdAt) : undefined,
      updatedAt: raw.updatedAt ? new Date(raw.updatedAt) : undefined,
      sentenceStructure: {
        ...raw.sentenceStructure,
        lengthDistribution: this.recordToNumberMap(raw.sentenceStructure.lengthDistribution)
      },
      timePerspective: {
        ...raw.timePerspective,
        tenseUsage: this.recordToStringMap(raw.timePerspective.tenseUsage)
      }
    };
  }

  private mapNumberToRecord(map: Map<number, number>): Record<string, number> {
    const result: Record<string, number> = {};
    map.forEach((value, key) => {
      result[String(key)] = value;
    });
    return result;
  }

  private mapStringToRecord(map: Map<string, number>): Record<string, number> {
    const result: Record<string, number> = {};
    map.forEach((value, key) => {
      result[key] = value;
    });
    return result;
  }

  private recordToNumberMap(record: Record<string, number>): Map<number, number> {
    return new Map<number, number>(
      Object.entries(record || {}).map(([key, value]) => [Number(key), value])
    );
  }

  private recordToStringMap(record: Record<string, number>): Map<string, number> {
    return new Map<string, number>(Object.entries(record || {}));
  }

  private normalizeStyle(style: WritingStyle): WritingStyle {
    const serialized = this.serializeStyle(style);
    return this.deserializeStyle(serialized);
  }

  private createEmptyStyle(): WritingStyle {
    const now = new Date();
    return {
      vocabulary: {
        avgWordLength: 0,
        uniqueWordRatio: 0,
        favoriteWords: [],
        rareWords: [],
        terminology: [],
        colloquialisms: []
      },
      sentenceStructure: {
        avgLength: 0,
        lengthDistribution: new Map<number, number>(),
        complexSentenceRatio: 0,
        questionRatio: 0,
        exclamationRatio: 0,
        sentencePatterns: []
      },
      tone: {
        type: 'neutral',
        intensity: 0,
        consistency: 0,
        markers: []
      },
      contentStructure: {
        paragraphLength: 0,
        structurePatterns: [],
        transitionWords: [],
        openingStyle: '',
        endingStyle: ''
      },
      rhetoric: {
        metaphor: [],
        parallelism: [],
        personification: [],
        exaggeration: [],
        quotes: [],
        rhetoricalDevices: []
      },
      timePerspective: {
        dominant: 'present',
        tenseUsage: new Map<string, number>(),
        temporalMarkers: [],
        flashbackUsage: false
      },
      spacePerspective: {
        viewpoint: 'third-person',
        spatialMarkers: [],
        sceneDescription: []
      },
      logicType: {
        type: 'narrative',
        reasoningPatterns: [],
        evidenceStyle: '',
        connectives: []
      },
      informationDensity: {
        density: 0,
        complexity: 0,
        abstraction: 0,
        concreteness: 0
      },
      emotionalArc: {
        overallTrend: 'neutral',
        peaks: [],
        valleys: [],
        emotionalWords: [],
        sentimentScore: 0
      },
      audienceAwareness: {
        targetAudience: '',
        engagementLevel: 0,
        accessibility: 0,
        jargonUsage: []
      },
      culturalContext: {
        genre: '',
        culturalReferences: [],
        register: 'informal',
        conventions: []
      },
      narrativeVoice: {
        voice: '',
        identity: '',
        stance: '',
        personality: []
      },
      rhythmControl: {
        pace: 'medium',
        variation: 0,
        pauseMarkers: [],
        flowPatterns: []
      },
      punctuation: {
        commaFrequency: 0,
        periodFrequency: 0,
        specialPunctuation: []
      },
      createdAt: now,
      updatedAt: now
    };
  }
}
