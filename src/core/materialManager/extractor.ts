import { WritingMaterial, MaterialType } from '../types';
import { TextPreprocessor } from '../../utils/textPreprocessor';
import { v4 as uuidv4 } from 'uuid';

/**
 * 素材提取器
 */
export class MaterialExtractor {
  private preprocessor: TextPreprocessor;

  constructor() {
    this.preprocessor = new TextPreprocessor();
  }

  /**
   * 智能提取素材
   */
  async extractMaterials(text: string, source: string = 'unknown'): Promise<WritingMaterial[]> {
    const materials: WritingMaterial[] = [];

    // 1. 提取精彩句子
    materials.push(...await this.extractSentences(text, source));

    // 2. 提取优秀段落
    materials.push(...await this.extractParagraphs(text, source));

    // 3. 提取引用和名言
    materials.push(...await this.extractQuotes(text, source));

    // 4. 提取比喻和修辞
    materials.push(...await this.extractRhetoric(text, source));

    // 5. 提取开头和结尾
    materials.push(...await this.extractStructure(text, source));

    return materials;
  }

  /**
   * 提取精彩句子
   */
  private async extractSentences(text: string, source: string): Promise<WritingMaterial[]> {
    const sentences = this.preprocessor.splitSentences(text);
    const materials: WritingMaterial[] = [];

    for (const sentence of sentences) {
      const cleaned = this.normalizeCandidateText(sentence);
      if (!this.isMeaningfulText(cleaned, 10)) {
        continue;
      }

      const score = this.scoreSentence(cleaned);
      if (score >= 50) {
        materials.push({
          id: uuidv4(),
          name: this.buildSnippetName(cleaned, 30),
          content: cleaned,
          type: MaterialType.SENTENCE,
          tags: this.extractTags(cleaned),
          category: 'sentence',
          metadata: {
            source,
            createdAt: new Date(),
            updatedAt: new Date(),
            usedCount: 0,
            quality: score
          },
          features: {
            keywords: this.preprocessor.extractKeywords(cleaned, 5),
            sentiment: this.detectSentiment(cleaned),
            style: 'unknown',
            quality: score
          }
        });
      }
    }

    return materials;
  }

  /**
   * 提取优秀段落
   */
  private async extractParagraphs(text: string, source: string): Promise<WritingMaterial[]> {
    const paragraphs = this.preprocessor.splitParagraphs(text);
    const materials: WritingMaterial[] = [];

    for (const paragraph of paragraphs) {
      const cleaned = this.normalizeCandidateText(paragraph);
      if (!this.isMeaningfulText(cleaned, 20)) {
        continue;
      }

      const score = this.scoreParagraph(cleaned);
      if (score >= 60) {
        materials.push({
          id: uuidv4(),
          name: this.buildSnippetName(cleaned, 50),
          content: cleaned,
          type: MaterialType.PARAGRAPH,
          tags: this.extractTags(cleaned),
          category: 'paragraph',
          metadata: {
            source,
            createdAt: new Date(),
            updatedAt: new Date(),
            usedCount: 0,
            quality: score
          },
          features: {
            keywords: this.preprocessor.extractKeywords(cleaned, 5),
            sentiment: this.detectSentiment(cleaned),
            style: 'unknown',
            quality: score
          }
        });
      }
    }

    return materials;
  }

  /**
   * 提取引用和名言
   */
  private async extractQuotes(text: string, source: string): Promise<WritingMaterial[]> {
    const materials: WritingMaterial[] = [];
    
    // 提取引号内容
    const quoteMatches = text.match(/[""「」『』]([^""「」『』]+)[""「」『』]/g);
    if (quoteMatches) {
      for (const match of quoteMatches) {
        const content = this.normalizeCandidateText(match.replace(/[""「」『』]/g, ''));
        if (this.isMeaningfulText(content, 10)) {
          materials.push({
            id: uuidv4(),
            name: this.buildSnippetName(content, 30),
            content,
            type: MaterialType.QUOTE,
            tags: ['引用'],
            category: 'quote',
            metadata: {
              source,
              createdAt: new Date(),
              updatedAt: new Date(),
              usedCount: 0,
              quality: 70
            },
            features: {
              keywords: this.preprocessor.extractKeywords(content, 5),
              sentiment: this.detectSentiment(content),
              style: 'unknown',
              quality: 70
            }
          });
        }
      }
    }

    return materials;
  }

  /**
   * 提取比喻和修辞
   */
  private async extractRhetoric(text: string, source: string): Promise<WritingMaterial[]> {
    const materials: WritingMaterial[] = [];
    
    // 提取比喻句
    const metaphorMarkers = ['像', '如同', '仿佛', '好比', '似'];
    const sentences = this.preprocessor.splitSentences(text);

    for (const sentence of sentences) {
      const cleaned = this.normalizeCandidateText(sentence);
      if (!this.isMeaningfulText(cleaned, 10)) {
        continue;
      }

      for (const marker of metaphorMarkers) {
        if (cleaned.includes(marker)) {
          materials.push({
            id: uuidv4(),
            name: this.buildSnippetName(cleaned, 30),
            content: cleaned,
            type: MaterialType.METAPHOR,
            tags: ['比喻', marker],
            category: 'rhetoric',
            metadata: {
              source,
              createdAt: new Date(),
              updatedAt: new Date(),
              usedCount: 0,
              quality: 75
            },
            features: {
              keywords: this.preprocessor.extractKeywords(cleaned, 5),
              sentiment: this.detectSentiment(cleaned),
              style: 'unknown',
              quality: 75
            }
          });
          break;
        }
      }
    }

    return materials;
  }

  /**
   * 提取开头和结尾
   */
  private async extractStructure(text: string, source: string): Promise<WritingMaterial[]> {
    const materials: WritingMaterial[] = [];
    const paragraphs = this.preprocessor.splitParagraphs(text);

    if (paragraphs.length > 0) {
      const opening = this.findMeaningfulParagraph(paragraphs, true);
      const ending = this.findMeaningfulParagraph(paragraphs, false);

      // 提取开头
      if (opening) {
        materials.push({
          id: uuidv4(),
          name: '文章开头',
          content: opening,
          type: MaterialType.OPENING,
          tags: ['开头'],
          category: 'structure',
          metadata: {
            source,
            createdAt: new Date(),
            updatedAt: new Date(),
            usedCount: 0,
            quality: 80
          },
          features: {
            keywords: this.preprocessor.extractKeywords(opening, 5),
            sentiment: this.detectSentiment(opening),
            style: 'unknown',
            quality: 80
          }
        });
      }

      // 提取结尾
      if (ending && ending !== opening) {
        materials.push({
          id: uuidv4(),
          name: '文章结尾',
          content: ending,
          type: MaterialType.ENDING,
          tags: ['结尾'],
          category: 'structure',
          metadata: {
            source,
            createdAt: new Date(),
            updatedAt: new Date(),
            usedCount: 0,
            quality: 80
          },
          features: {
            keywords: this.preprocessor.extractKeywords(ending, 5),
            sentiment: this.detectSentiment(ending),
            style: 'unknown',
            quality: 80
          }
        });
      }
    }

    return materials;
  }

  /**
   * 句子质量评分
   */
  private scoreSentence(sentence: string): number {
    let score = 0;

    // 长度适中 (20分)
    if (sentence.length >= 20 && sentence.length <= 100) {
      score += 20;
    }

    // 包含修辞 (30分)
    if (this.hasMetaphor(sentence)) score += 30;
    if (this.hasParallelism(sentence)) score += 25;

    // 情感强度 (25分)
    score += this.getEmotionalIntensity(sentence) * 25;

    // 信息密度 (25分)
    const keywords = this.preprocessor.extractKeywords(sentence, 5);
    score += Math.min(keywords.length * 5, 25);

    return score;
  }

  /**
   * 段落质量评分
   */
  private scoreParagraph(paragraph: string): number {
    if (!this.isMeaningfulText(paragraph, 20)) {
      return 0;
    }

    let score = 0;

    // 长度适中 (20分)
    if (paragraph.length >= 100 && paragraph.length <= 500) {
      score += 20;
    }

    // 结构完整 (30分)
    const sentences = this.preprocessor.splitSentences(paragraph);
    if (sentences.length >= 3) {
      score += 30;
    }

    // 主题明确 (25分)
    const keywords = this.preprocessor.extractKeywords(paragraph, 5);
    if (keywords.length >= 3) {
      score += 25;
    }

    // 情感丰富 (25分)
    score += this.getEmotionalIntensity(paragraph) * 25;

    return score;
  }

  /**
   * 检测是否包含比喻
   */
  private hasMetaphor(text: string): boolean {
    const markers = ['像', '如同', '仿佛', '好比', '似'];
    return markers.some(m => text.includes(m));
  }

  /**
   * 检测是否包含排比
   */
  private hasParallelism(text: string): boolean {
    // 简化检测：重复词语
    const words = text.split(/[，。！？、]/);
    const wordSet = new Set(words);
    return words.length > wordSet.size + 2;
  }

  /**
   * 获取情感强度
   */
  private getEmotionalIntensity(text: string): number {
    const emotionalWords = ['好', '棒', '优秀', '喜欢', '开心', '快乐', '成功', '美好', '坏', '差', '糟糕', '讨厌', '难过', '悲伤', '失败', '痛苦'];
    let count = 0;
    emotionalWords.forEach(w => {
      if (text.includes(w)) count++;
    });
    return Math.min(count / 3, 1);
  }

  /**
   * 检测情感倾向
   */
  private detectSentiment(text: string): string {
    const positiveWords = ['好', '棒', '优秀', '喜欢', '开心', '快乐', '成功', '美好'];
    const negativeWords = ['坏', '差', '糟糕', '讨厌', '难过', '悲伤', '失败', '痛苦'];

    let positiveCount = 0;
    let negativeCount = 0;

    positiveWords.forEach(w => {
      if (text.includes(w)) positiveCount++;
    });

    negativeWords.forEach(w => {
      if (text.includes(w)) negativeCount++;
    });

    if (positiveCount > negativeCount) return 'positive';
    if (negativeCount > positiveCount) return 'negative';
    return 'neutral';
  }

  /**
   * 提取标签
   */
  private extractTags(text: string): string[] {
    const tags: string[] = [];

    // 基于关键词提取标签
    const keywords = this.preprocessor.extractKeywords(text, 3);
    tags.push(...keywords);

    // 基于特征提取标签
    if (this.hasMetaphor(text)) tags.push('比喻');
    if (text.includes('！') || text.includes('!')) tags.push('感叹');
    if (text.includes('？') || text.includes('?')) tags.push('疑问');

    return [...new Set(tags)];
  }

  private normalizeCandidateText(text: string): string {
    return text.replace(/\r\n/g, '\n').trim();
  }

  private buildSnippetName(text: string, maxLength: number): string {
    if (text.length <= maxLength) {
      return text;
    }
    return `${text.slice(0, maxLength)}...`;
  }

  private findMeaningfulParagraph(paragraphs: string[], fromStart: boolean): string | null {
    const list = fromStart ? paragraphs : [...paragraphs].reverse();
    for (const paragraph of list) {
      const cleaned = this.normalizeCandidateText(paragraph);
      if (this.isMeaningfulText(cleaned, 20)) {
        return cleaned;
      }
    }
    return null;
  }

  private isMeaningfulText(text: string, minCoreChars: number): boolean {
    if (!text) {
      return false;
    }

    const compact = text.replace(/\s+/g, '');
    if (!compact) {
      return false;
    }

    if (/^[-_*`#>~|]+$/.test(compact)) {
      return false;
    }

    const markdownRuleOnly = compact
      .replace(/^#+/g, '')
      .replace(/^>+/g, '')
      .replace(/[-_*`~|]/g, '');
    if (!markdownRuleOnly) {
      return false;
    }

    const core = compact.replace(/[^\p{L}\p{N}\u4E00-\u9FFF]/gu, '');
    if (core.length < minCoreChars) {
      return false;
    }

    return true;
  }
}
