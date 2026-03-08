import { WritingStyle } from '../types';
import { StyleAnalyzer } from './analyzer';
import { StyleLearner } from './learner';

/**
 * 风格引擎
 */
export class StyleEngine {
  private analyzer: StyleAnalyzer;
  private learner: StyleLearner;

  constructor() {
    this.analyzer = new StyleAnalyzer();
    this.learner = new StyleLearner();
  }

  /**
   * 分析文本风格
   */
  async analyzeText(text: string): Promise<WritingStyle> {
    return await this.analyzer.analyze(text);
  }

  /**
   * 从文本学习风格
   */
  async learnFromText(text: string): Promise<void> {
    await this.learner.learnFromText(text);
  }

  /**
   * 获取当前风格
   */
  getCurrentStyle(): WritingStyle | null {
    return this.learner.getCurrentStyle();
  }

  /**
   * 设置当前风格
   */
  setCurrentStyle(style: WritingStyle): void {
    this.learner.setCurrentStyle(style);
  }

  /**
   * 重置风格
   */
  resetStyle(): void {
    this.learner.reset();
  }

  /**
   * 计算风格相似度
   */
  calculateSimilarity(style1: WritingStyle, style2: WritingStyle): number {
    // 核心维度权重 60%
    const coreScore = this.calculateCoreSimilarity(style1, style2) * 0.6;

    // 深度维度权重 40%
    const deepScore = this.calculateDeepSimilarity(style1, style2) * 0.4;

    return coreScore + deepScore;
  }

  /**
   * 核心维度相似度
   */
  private calculateCoreSimilarity(style1: WritingStyle, style2: WritingStyle): number {
    let score = 0;
    let count = 0;

    // 词汇特征
    score += this.numericSimilarity(style1.vocabulary.avgWordLength, style2.vocabulary.avgWordLength);
    score += this.numericSimilarity(style1.vocabulary.uniqueWordRatio, style2.vocabulary.uniqueWordRatio);
    count += 2;

    // 句式结构
    score += this.numericSimilarity(style1.sentenceStructure.avgLength, style2.sentenceStructure.avgLength);
    score += this.numericSimilarity(style1.sentenceStructure.complexSentenceRatio, style2.sentenceStructure.complexSentenceRatio);
    count += 2;

    // 语气语调
    score += style1.tone.type === style2.tone.type ? 1 : 0;
    score += this.numericSimilarity(style1.tone.intensity, style2.tone.intensity);
    count += 2;

    // 内容结构
    score += this.numericSimilarity(style1.contentStructure.paragraphLength, style2.contentStructure.paragraphLength);
    count += 1;

    // 信息密度
    score += this.numericSimilarity(style1.informationDensity.density, style2.informationDensity.density);
    score += this.numericSimilarity(style1.informationDensity.complexity, style2.informationDensity.complexity);
    count += 2;

    return score / count;
  }

  /**
   * 深度维度相似度
   */
  private calculateDeepSimilarity(style1: WritingStyle, style2: WritingStyle): number {
    let score = 0;
    let count = 0;

    // 时间视角
    score += style1.timePerspective.dominant === style2.timePerspective.dominant ? 1 : 0;
    count += 1;

    // 空间视角
    score += style1.spacePerspective.viewpoint === style2.spacePerspective.viewpoint ? 1 : 0;
    count += 1;

    // 逻辑类型
    score += style1.logicType.type === style2.logicType.type ? 1 : 0;
    count += 1;

    // 情感曲线
    score += this.numericSimilarity(style1.emotionalArc.sentimentScore, style2.emotionalArc.sentimentScore);
    count += 1;

    // 节奏控制
    score += style1.rhythmControl.pace === style2.rhythmControl.pace ? 1 : 0;
    score += this.numericSimilarity(style1.rhythmControl.variation, style2.rhythmControl.variation);
    count += 2;

    return score / count;
  }

  /**
   * 数值相似度（归一化）
   */
  private numericSimilarity(value1: number, value2: number): number {
    if (value1 === 0 && value2 === 0) return 1;
    const diff = Math.abs(value1 - value2);
    const max = Math.max(Math.abs(value1), Math.abs(value2));
    return 1 - (diff / max);
  }

  /**
   * 导出风格配置
   */
  exportStyle(): string {
    const style = this.getCurrentStyle();
    if (!style) {
      throw new Error('No style to export');
    }
    return JSON.stringify(style, null, 2);
  }

  /**
   * 导入风格配置
   */
  importStyle(styleJson: string): void {
    try {
      const style = JSON.parse(styleJson) as WritingStyle;
      // 转换 Map 对象
      if (style.sentenceStructure.lengthDistribution) {
        const mapObj = style.sentenceStructure.lengthDistribution as any;
        style.sentenceStructure.lengthDistribution = new Map<number, number>(
          Object.entries(mapObj).map(([k, v]) => [parseInt(k), v as number])
        );
      }
      if (style.timePerspective.tenseUsage) {
        const mapObj = style.timePerspective.tenseUsage as any;
        style.timePerspective.tenseUsage = new Map<string, number>(
          Object.entries(mapObj)
        );
      }
      this.setCurrentStyle(style);
    } catch (error) {
      throw new Error('Invalid style format');
    }
  }
}

export { StyleAnalyzer } from './analyzer';
export { StyleLearner } from './learner';
