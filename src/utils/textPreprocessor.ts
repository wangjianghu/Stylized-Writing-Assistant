type NodeJiebaLike = {
  cut: (text: string) => string[];
};

let cachedNodeJieba: NodeJiebaLike | null | undefined;

function loadNodeJieba(): NodeJiebaLike | null {
  if (cachedNodeJieba !== undefined) {
    return cachedNodeJieba;
  }

  try {
    // 延迟加载可选原生依赖，避免宿主不支持时导致插件激活失败。
    // eslint-disable-next-line @typescript-eslint/no-var-requires, global-require
    cachedNodeJieba = require('nodejieba') as NodeJiebaLike;
    return cachedNodeJieba;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(`[writing-agent] nodejieba 加载失败，降级使用基础分词：${message}`);
    cachedNodeJieba = null;
    return null;
  }
}

/**
 * 文本预处理器
 */
export class TextPreprocessor {
  /**
   * 中文分词
   */
  tokenize(text: string): string[] {
    const nodeJieba = loadNodeJieba();
    if (!nodeJieba) {
      return this.simpleTokenize(text);
    }

    try {
      return nodeJieba.cut(text);
    } catch (error) {
      // 如果 nodejieba 不可用，使用简单的分词
      return this.simpleTokenize(text);
    }
  }

  /**
   * 简单分词（备用方案）
   */
  private simpleTokenize(text: string): string[] {
    // 移除标点符号
    const cleanText = text.replace(/[，。！？、；：""''（）【】《》\s]/g, ' ');
    // 按空格分割
    return cleanText.split(/\s+/).filter(word => word.length > 0);
  }

  /**
   * 句子分割
   */
  splitSentences(text: string): string[] {
    // 处理中英文标点
    const sentenceEnders = /[。！？!?.]+/g;
    const sentences = text.split(sentenceEnders).filter(s => s.trim());
    return sentences.map(s => s.trim());
  }

  /**
   * 段落分割
   */
  splitParagraphs(text: string): string[] {
    return text.split(/\n\s*\n/).filter(p => p.trim());
  }

  /**
   * 清理文本
   */
  clean(text: string): string {
    return text
      .replace(/\s+/g, ' ')
      .replace(/[\u0000-\u001F\u007F-\u009F]/g, '')
      .trim();
  }

  /**
   * 统计文本信息
   */
  getStatistics(text: string) {
    const tokens = this.tokenize(text);
    const sentences = this.splitSentences(text);
    const paragraphs = this.splitParagraphs(text);

    return {
      totalChars: text.length,
      totalWords: tokens.length,
      totalSentences: sentences.length,
      totalParagraphs: paragraphs.length,
      avgWordLength: tokens.length > 0 
        ? tokens.reduce((sum, t) => sum + t.length, 0) / tokens.length 
        : 0,
      avgSentenceLength: sentences.length > 0 
        ? sentences.reduce((sum, s) => sum + s.length, 0) / sentences.length 
        : 0,
      avgParagraphLength: paragraphs.length > 0 
        ? paragraphs.reduce((sum, p) => sum + p.length, 0) / paragraphs.length 
        : 0
    };
  }

  /**
   * 提取关键词
   */
  extractKeywords(text: string, topN: number = 10): string[] {
    const tokens = this.tokenize(text);
    const wordFreq = new Map<string, number>();

    // 统计词频
    tokens.forEach(token => {
      if (token.length > 1) { // 过滤单字
        wordFreq.set(token, (wordFreq.get(token) || 0) + 1);
      }
    });

    // 排序并返回前N个
    return Array.from(wordFreq.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, topN)
      .map(([word]) => word);
  }

  /**
   * 移除停用词
   */
  removeStopwords(tokens: string[]): string[] {
    const stopwords = new Set([
      '的', '了', '和', '是', '在', '有', '我', '他', '她', '它',
      '这', '那', '就', '也', '都', '而', '及', '与', '或', '但',
      '如', '若', '因', '为', '所', '以', '对', '把', '被', '让',
      '给', '向', '从', '到', '由', '于', '等', '很', '更', '最',
      'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to',
      'for', 'of', 'with', 'by', 'from', 'as', 'is', 'was', 'are', 'were'
    ]);

    return tokens.filter(token => !stopwords.has(token.toLowerCase()));
  }
}
