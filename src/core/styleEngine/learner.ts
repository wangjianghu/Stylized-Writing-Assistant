import { WritingStyle } from '../types';
import { StyleAnalyzer } from './analyzer';

/**
 * 风格学习器
 */
export class StyleLearner {
  private analyzer: StyleAnalyzer;
  private currentStyle: WritingStyle | null = null;

  constructor() {
    this.analyzer = new StyleAnalyzer();
  }

  /**
   * 从文本学习风格（增量学习）
   */
  async learnFromText(text: string, alpha: number = 0.3): Promise<void> {
    const newFeatures = await this.analyzer.analyze(text);

    if (!this.currentStyle) {
      this.currentStyle = newFeatures;
      return;
    }

    // 使用指数移动平均进行增量更新
    this.currentStyle = this.mergeStyles(this.currentStyle, newFeatures, alpha);
    this.currentStyle.updatedAt = new Date();
  }

  /**
   * 获取当前风格
   */
  getCurrentStyle(): WritingStyle | null {
    return this.currentStyle;
  }

  /**
   * 设置当前风格
   */
  setCurrentStyle(style: WritingStyle): void {
    this.currentStyle = style;
  }

  /**
   * 重置风格
   */
  reset(): void {
    this.currentStyle = null;
  }

  /**
   * 合并两个风格（使用指数移动平均）
   */
  private mergeStyles(
    oldStyle: WritingStyle,
    newStyle: WritingStyle,
    alpha: number
  ): WritingStyle {
    return {
      // 核心维度
      vocabulary: {
        avgWordLength: this.ema(oldStyle.vocabulary.avgWordLength, newStyle.vocabulary.avgWordLength, alpha),
        uniqueWordRatio: this.ema(oldStyle.vocabulary.uniqueWordRatio, newStyle.vocabulary.uniqueWordRatio, alpha),
        favoriteWords: this.mergeArrays(oldStyle.vocabulary.favoriteWords, newStyle.vocabulary.favoriteWords),
        rareWords: this.mergeArrays(oldStyle.vocabulary.rareWords, newStyle.vocabulary.rareWords),
        terminology: this.mergeArrays(oldStyle.vocabulary.terminology, newStyle.vocabulary.terminology),
        colloquialisms: this.mergeArrays(oldStyle.vocabulary.colloquialisms, newStyle.vocabulary.colloquialisms)
      },

      sentenceStructure: {
        avgLength: this.ema(oldStyle.sentenceStructure.avgLength, newStyle.sentenceStructure.avgLength, alpha),
        lengthDistribution: this.mergeMapsNumber(oldStyle.sentenceStructure.lengthDistribution, newStyle.sentenceStructure.lengthDistribution),
        complexSentenceRatio: this.ema(oldStyle.sentenceStructure.complexSentenceRatio, newStyle.sentenceStructure.complexSentenceRatio, alpha),
        questionRatio: this.ema(oldStyle.sentenceStructure.questionRatio, newStyle.sentenceStructure.questionRatio, alpha),
        exclamationRatio: this.ema(oldStyle.sentenceStructure.exclamationRatio, newStyle.sentenceStructure.exclamationRatio, alpha),
        sentencePatterns: this.mergeArrays(oldStyle.sentenceStructure.sentencePatterns, newStyle.sentenceStructure.sentencePatterns)
      },

      tone: {
        type: newStyle.tone.type, // 使用最新的类型
        intensity: this.ema(oldStyle.tone.intensity, newStyle.tone.intensity, alpha),
        consistency: this.ema(oldStyle.tone.consistency, newStyle.tone.consistency, alpha),
        markers: this.mergeArrays(oldStyle.tone.markers, newStyle.tone.markers)
      },

      contentStructure: {
        paragraphLength: this.ema(oldStyle.contentStructure.paragraphLength, newStyle.contentStructure.paragraphLength, alpha),
        structurePatterns: this.mergeArrays(oldStyle.contentStructure.structurePatterns, newStyle.contentStructure.structurePatterns),
        transitionWords: this.mergeArrays(oldStyle.contentStructure.transitionWords, newStyle.contentStructure.transitionWords),
        openingStyle: newStyle.contentStructure.openingStyle,
        endingStyle: newStyle.contentStructure.endingStyle
      },

      rhetoric: {
        metaphor: this.mergeArrays(oldStyle.rhetoric.metaphor, newStyle.rhetoric.metaphor),
        parallelism: this.mergeArrays(oldStyle.rhetoric.parallelism, newStyle.rhetoric.parallelism),
        personification: this.mergeArrays(oldStyle.rhetoric.personification, newStyle.rhetoric.personification),
        exaggeration: this.mergeArrays(oldStyle.rhetoric.exaggeration, newStyle.rhetoric.exaggeration),
        quotes: this.mergeArrays(oldStyle.rhetoric.quotes, newStyle.rhetoric.quotes),
        rhetoricalDevices: this.mergeArrays(oldStyle.rhetoric.rhetoricalDevices, newStyle.rhetoric.rhetoricalDevices)
      },

      // 深度维度
      timePerspective: {
        dominant: newStyle.timePerspective.dominant,
        tenseUsage: this.mergeMaps(oldStyle.timePerspective.tenseUsage, newStyle.timePerspective.tenseUsage),
        temporalMarkers: this.mergeArrays(oldStyle.timePerspective.temporalMarkers, newStyle.timePerspective.temporalMarkers),
        flashbackUsage: newStyle.timePerspective.flashbackUsage
      },

      spacePerspective: {
        viewpoint: newStyle.spacePerspective.viewpoint,
        spatialMarkers: this.mergeArrays(oldStyle.spacePerspective.spatialMarkers, newStyle.spacePerspective.spatialMarkers),
        sceneDescription: this.mergeArrays(oldStyle.spacePerspective.sceneDescription, newStyle.spacePerspective.sceneDescription)
      },

      logicType: {
        type: newStyle.logicType.type,
        reasoningPatterns: this.mergeArrays(oldStyle.logicType.reasoningPatterns, newStyle.logicType.reasoningPatterns),
        evidenceStyle: newStyle.logicType.evidenceStyle,
        connectives: this.mergeArrays(oldStyle.logicType.connectives, newStyle.logicType.connectives)
      },

      informationDensity: {
        density: this.ema(oldStyle.informationDensity.density, newStyle.informationDensity.density, alpha),
        complexity: this.ema(oldStyle.informationDensity.complexity, newStyle.informationDensity.complexity, alpha),
        abstraction: this.ema(oldStyle.informationDensity.abstraction, newStyle.informationDensity.abstraction, alpha),
        concreteness: this.ema(oldStyle.informationDensity.concreteness, newStyle.informationDensity.concreteness, alpha)
      },

      emotionalArc: {
        overallTrend: newStyle.emotionalArc.overallTrend,
        peaks: newStyle.emotionalArc.peaks,
        valleys: newStyle.emotionalArc.valleys,
        emotionalWords: this.mergeArrays(oldStyle.emotionalArc.emotionalWords, newStyle.emotionalArc.emotionalWords),
        sentimentScore: this.ema(oldStyle.emotionalArc.sentimentScore, newStyle.emotionalArc.sentimentScore, alpha)
      },

      // 语境维度
      audienceAwareness: {
        targetAudience: newStyle.audienceAwareness.targetAudience,
        engagementLevel: this.ema(oldStyle.audienceAwareness.engagementLevel, newStyle.audienceAwareness.engagementLevel, alpha),
        accessibility: this.ema(oldStyle.audienceAwareness.accessibility, newStyle.audienceAwareness.accessibility, alpha),
        jargonUsage: this.mergeArrays(oldStyle.audienceAwareness.jargonUsage, newStyle.audienceAwareness.jargonUsage)
      },

      culturalContext: {
        genre: newStyle.culturalContext.genre,
        culturalReferences: this.mergeArrays(oldStyle.culturalContext.culturalReferences, newStyle.culturalContext.culturalReferences),
        register: newStyle.culturalContext.register,
        conventions: this.mergeArrays(oldStyle.culturalContext.conventions, newStyle.culturalContext.conventions)
      },

      narrativeVoice: {
        voice: newStyle.narrativeVoice.voice,
        identity: newStyle.narrativeVoice.identity,
        stance: newStyle.narrativeVoice.stance,
        personality: this.mergeArrays(oldStyle.narrativeVoice.personality, newStyle.narrativeVoice.personality)
      },

      rhythmControl: {
        pace: newStyle.rhythmControl.pace,
        variation: this.ema(oldStyle.rhythmControl.variation, newStyle.rhythmControl.variation, alpha),
        pauseMarkers: this.mergeArrays(oldStyle.rhythmControl.pauseMarkers, newStyle.rhythmControl.pauseMarkers),
        flowPatterns: this.mergeArrays(oldStyle.rhythmControl.flowPatterns, newStyle.rhythmControl.flowPatterns)
      },

      // 辅助维度
      punctuation: {
        commaFrequency: this.ema(oldStyle.punctuation.commaFrequency, newStyle.punctuation.commaFrequency, alpha),
        periodFrequency: this.ema(oldStyle.punctuation.periodFrequency, newStyle.punctuation.periodFrequency, alpha),
        specialPunctuation: this.mergeArrays(oldStyle.punctuation.specialPunctuation, newStyle.punctuation.specialPunctuation)
      },

      // 元数据
      id: oldStyle.id,
      name: oldStyle.name,
      createdAt: oldStyle.createdAt,
      updatedAt: new Date()
    };
  }

  /**
   * 指数移动平均
   */
  private ema(oldValue: number, newValue: number, alpha: number): number {
    return oldValue * (1 - alpha) + newValue * alpha;
  }

  /**
   * 合并数组（去重）
   */
  private mergeArrays(arr1: string[], arr2: string[]): string[] {
    const merged = [...new Set([...arr1, ...arr2])];
    return merged.slice(0, 20); // 限制长度
  }

  /**
   * 合并 Map（字符串键）
   */
  private mergeMaps(map1: Map<string, number>, map2: Map<string, number>): Map<string, number> {
    const merged = new Map(map1);
    map2.forEach((value, key) => {
      merged.set(key, (merged.get(key) || 0) + value);
    });
    return merged;
  }

  /**
   * 合并 Map（数字键）
   */
  private mergeMapsNumber(map1: Map<number, number>, map2: Map<number, number>): Map<number, number> {
    const merged = new Map(map1);
    map2.forEach((value, key) => {
      merged.set(key, (merged.get(key) || 0) + value);
    });
    return merged;
  }
}
