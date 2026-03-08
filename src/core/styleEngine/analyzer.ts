import { WritingStyle } from '../types';
import { TextPreprocessor } from '../../utils/textPreprocessor';

/**
 * 风格分析器（14维度分析）
 */
export class StyleAnalyzer {
  private preprocessor: TextPreprocessor;

  constructor() {
    this.preprocessor = new TextPreprocessor();
  }

  /**
   * 分析文本风格（14维度完整分析）
   */
  async analyze(text: string): Promise<WritingStyle> {
    const tokens = this.preprocessor.tokenize(text);
    const sentences = this.preprocessor.splitSentences(text);
    const paragraphs = this.preprocessor.splitParagraphs(text);

    return {
      // 核心维度
      vocabulary: this.analyzeVocabulary(tokens),
      sentenceStructure: this.analyzeSentenceStructure(sentences),
      tone: this.analyzeTone(text, tokens),
      contentStructure: this.analyzeContentStructure(paragraphs, text),
      rhetoric: this.analyzeRhetoric(text),

      // 深度维度
      timePerspective: this.analyzeTimePerspective(text, tokens),
      spacePerspective: this.analyzeSpacePerspective(text, tokens),
      logicType: this.analyzeLogicType(text, tokens),
      informationDensity: this.analyzeInformationDensity(text, tokens),
      emotionalArc: this.analyzeEmotionalArc(text, tokens),

      // 语境维度
      audienceAwareness: this.analyzeAudienceAwareness(text, tokens),
      culturalContext: this.analyzeCulturalContext(text),
      narrativeVoice: this.analyzeNarrativeVoice(text, tokens),
      rhythmControl: this.analyzeRhythmControl(sentences, text),

      // 辅助维度
      punctuation: this.analyzePunctuation(text),

      // 元数据
      createdAt: new Date(),
      updatedAt: new Date()
    };
  }

  /**
   * 1. 分析词汇特征
   */
  private analyzeVocabulary(tokens: string[]): WritingStyle['vocabulary'] {
    const uniqueTokens = new Set(tokens);
    const wordLengths = tokens.map(t => t.length);

    // 计算词频
    const wordFreq = new Map<string, number>();
    tokens.forEach(t => {
      wordFreq.set(t, (wordFreq.get(t) || 0) + 1);
    });

    // 高频词（排除停用词）
    const stopwords = new Set(['的', '是', '在', '了', '和', '等', 'the', 'a', 'an', 'is', 'are']);
    const topWords = Array.from(wordFreq.entries())
      .filter(([word]) => !stopwords.has(word.toLowerCase()) && word.length > 1)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 50)
      .map(([word]) => word);

    // 识别专业术语（简单规则：包含大写字母、数字或特殊字符）
    const terminology = tokens.filter(t => 
      /[A-Z]{2,}/.test(t) || /\d/.test(t) || /[\-_]/.test(t)
    );

    // 识别口语化表达
    const colloquialisms = tokens.filter(t => 
      ['嗯', '啊', '吧', '呢', '嘛', '哦'].includes(t)
    );

    return {
      avgWordLength: this.average(wordLengths),
      uniqueWordRatio: uniqueTokens.size / Math.max(tokens.length, 1),
      favoriteWords: topWords.slice(0, 20),
      rareWords: topWords.slice(-10),
      terminology: [...new Set(terminology)].slice(0, 20),
      colloquialisms: [...new Set(colloquialisms)]
    };
  }

  /**
   * 2. 分析句式结构
   */
  private analyzeSentenceStructure(sentences: string[]): WritingStyle['sentenceStructure'] {
    const lengths = sentences.map(s => s.length);
    const lengthDistribution = new Map<number, number>();
    
    lengths.forEach(len => {
      const bucket = Math.floor(len / 10) * 10;
      lengthDistribution.set(bucket, (lengthDistribution.get(bucket) || 0) + 1);
    });

    const complexCount = sentences.filter(s => 
      s.includes('，') || s.includes(',') || s.includes('；') || s.includes(';')
    ).length;

    const questionCount = sentences.filter(s => 
      s.includes('?') || s.includes('？')
    ).length;

    const exclamationCount = sentences.filter(s => 
      s.includes('!') || s.includes('！')
    ).length;

    return {
      avgLength: this.average(lengths),
      lengthDistribution,
      complexSentenceRatio: complexCount / Math.max(sentences.length, 1),
      questionRatio: questionCount / Math.max(sentences.length, 1),
      exclamationRatio: exclamationCount / Math.max(sentences.length, 1),
      sentencePatterns: this.extractSentencePatterns(sentences)
    };
  }

  /**
   * 3. 分析语气语调
   */
  private analyzeTone(text: string, _tokens: string[]): WritingStyle['tone'] {
    // 简单的语气判断
    const formalMarkers = ['因此', '然而', '此外', '综上所述', '由此可见'];
    const casualMarkers = ['哈哈', '嗯嗯', '好吧', '其实', '反正'];
    const humorousMarkers = ['哈哈', '呵呵', '笑死', '逗'];
    const seriousMarkers = ['必须', '务必', '严格', '认真', '重要'];

    let formalScore = 0;
    let casualScore = 0;
    let humorousScore = 0;
    let seriousScore = 0;

    formalMarkers.forEach(m => {
      if (text.includes(m)) formalScore++;
    });
    casualMarkers.forEach(m => {
      if (text.includes(m)) casualScore++;
    });
    humorousMarkers.forEach(m => {
      if (text.includes(m)) humorousScore++;
    });
    seriousMarkers.forEach(m => {
      if (text.includes(m)) seriousScore++;
    });

    const maxScore = Math.max(formalScore, casualScore, humorousScore, seriousScore);
    let type: 'formal' | 'casual' | 'humorous' | 'serious' | 'neutral' = 'neutral';
    
    if (maxScore > 0) {
      if (formalScore === maxScore) type = 'formal';
      else if (casualScore === maxScore) type = 'casual';
      else if (humorousScore === maxScore) type = 'humorous';
      else if (seriousScore === maxScore) type = 'serious';
    }

    return {
      type,
      intensity: maxScore / 5,
      consistency: 0.8, // 简化计算
      markers: [...formalMarkers, ...casualMarkers].filter(m => text.includes(m))
    };
  }

  /**
   * 4. 分析内容结构
   */
  private analyzeContentStructure(paragraphs: string[], text: string): WritingStyle['contentStructure'] {
    const lengths = paragraphs.map(p => p.length);
    
    // 检测结构模式
    const structurePatterns: string[] = [];
    if (text.includes('首先') && text.includes('其次') && text.includes('最后')) {
      structurePatterns.push('递进式');
    }
    if (text.includes('一方面') && text.includes('另一方面')) {
      structurePatterns.push('对比式');
    }
    if (text.includes('总之') || text.includes('综上所述')) {
      structurePatterns.push('总分总');
    }

    // 提取过渡词（按出现频次排序，避免固定词表过于机械）
    const transitionCandidates = [
      '因此', '所以', '然而', '但是', '另外', '此外', '同时', '接着', '随后',
      '不过', '反过来', '相较之下', '换句话说', '具体来说', '进一步看', '再看',
      '最后', '与此同时', '总体来看', '由此可见', '从这个角度'
    ];
    const transitionWords = transitionCandidates
      .map(word => ({ word, count: this.countOccurrences(text, word) }))
      .filter(item => item.count > 0)
      .sort((a, b) => b.count - a.count)
      .slice(0, 10)
      .map(item => item.word);

    // 分析开头和结尾
    const firstParagraph = paragraphs[0] || '';
    const lastParagraph = paragraphs[paragraphs.length - 1] || '';

    return {
      paragraphLength: this.average(lengths),
      structurePatterns,
      transitionWords,
      openingStyle: this.detectOpeningStyle(firstParagraph),
      endingStyle: this.detectEndingStyle(lastParagraph)
    };
  }

  /**
   * 5. 分析修辞手法
   */
  private analyzeRhetoric(text: string): WritingStyle['rhetoric'] {
    // 检测比喻
    const metaphorMarkers = ['像', '如同', '仿佛', '好比', '似'];
    const metaphors = metaphorMarkers.filter(m => text.includes(m));

    // 检测排比（简化：检测重复句式）
    const parallelism = this.detectParallelism(text);

    // 检测引用
    const quotes = this.extractQuotes(text);

    return {
      metaphor: metaphors,
      parallelism,
      personification: [], // 需要更复杂的NLP分析
      exaggeration: [], // 需要更复杂的NLP分析
      quotes,
      rhetoricalDevices: []
    };
  }

  /**
   * 6. 分析时间视角
   */
  private analyzeTimePerspective(text: string, _tokens: string[]): WritingStyle['timePerspective'] {
    const pastMarkers = ['曾', '已经', '过去', '以前', '昨天', '去年'];
    const presentMarkers = ['现在', '目前', '正在', '今天', '今年'];
    const futureMarkers = ['将', '将要', '未来', '明天', '明年'];

    let pastCount = 0;
    let presentCount = 0;
    let futureCount = 0;

    pastMarkers.forEach(m => { if (text.includes(m)) pastCount++; });
    presentMarkers.forEach(m => { if (text.includes(m)) presentCount++; });
    futureMarkers.forEach(m => { if (text.includes(m)) futureCount++; });

    const tenseUsage = new Map<string, number>();
    tenseUsage.set('past', pastCount);
    tenseUsage.set('present', presentCount);
    tenseUsage.set('future', futureCount);

    const max = Math.max(pastCount, presentCount, futureCount);
    let dominant: 'past' | 'present' | 'future' = 'present';
    if (max === pastCount) dominant = 'past';
    else if (max === futureCount) dominant = 'future';

    return {
      dominant,
      tenseUsage,
      temporalMarkers: [...pastMarkers, ...presentMarkers, ...futureMarkers].filter(m => text.includes(m)),
      flashbackUsage: text.includes('回忆') || text.includes('回想') || text.includes('那时')
    };
  }

  /**
   * 7. 分析空间视角
   */
  private analyzeSpacePerspective(text: string, _tokens: string[]): WritingStyle['spacePerspective'] {
    // 检测叙事视角
    let viewpoint: 'first-person' | 'third-person' | 'omniscient' = 'third-person';
    
    const firstPersonPronouns = ['我', '我们', '咱', '咱们'];
    const thirdPersonPronouns = ['他', '她', '它', '他们', '她们', '它们'];

    let firstPersonCount = 0;
    let thirdPersonCount = 0;

    firstPersonPronouns.forEach(p => {
      const regex = new RegExp(p, 'g');
      const matches = text.match(regex);
      if (matches) firstPersonCount += matches.length;
    });

    thirdPersonPronouns.forEach(p => {
      const regex = new RegExp(p, 'g');
      const matches = text.match(regex);
      if (matches) thirdPersonCount += matches.length;
    });

    if (firstPersonCount > thirdPersonCount * 2) {
      viewpoint = 'first-person';
    } else if (firstPersonCount === 0 && thirdPersonCount > 0) {
      viewpoint = 'omniscient';
    }

    const spatialMarkers = ['这里', '那里', '上面', '下面', '左边', '右边', '前方', '后方']
      .filter(m => text.includes(m));

    return {
      viewpoint,
      spatialMarkers,
      sceneDescription: [] // 需要更复杂的分析
    };
  }

  /**
   * 8. 分析逻辑类型
   */
  private analyzeLogicType(text: string, _tokens: string[]): WritingStyle['logicType'] {
    // 检测逻辑类型
    let type: 'deductive' | 'inductive' | 'abductive' | 'narrative' = 'narrative';

    if (text.includes('因此') || text.includes('所以') || text.includes('由此可见')) {
      type = 'deductive';
    } else if (text.includes('例如') || text.includes('比如') || text.includes('举例')) {
      type = 'inductive';
    } else if (text.includes('可能') || text.includes('也许') || text.includes('推测')) {
      type = 'abductive';
    }

    const connectives = ['因为', '所以', '但是', '然而', '因此', '由此', '于是', '那么']
      .filter(c => text.includes(c));

    return {
      type,
      reasoningPatterns: [],
      evidenceStyle: text.includes('数据') || text.includes('研究') ? 'data-driven' : 'logic-based',
      connectives
    };
  }

  /**
   * 9. 分析信息密度与复杂度
   */
  private analyzeInformationDensity(text: string, tokens: string[]): WritingStyle['informationDensity'] {
    // 信息密度：单位长度内的信息量（用独特词汇比例近似）
    const uniqueTokens = new Set(tokens);
    const density = uniqueTokens.size / Math.max(tokens.length, 1);

    // 复杂度：平均句长和词汇复杂度
    const sentences = this.preprocessor.splitSentences(text);
    const avgSentenceLength = sentences.length > 0 
      ? sentences.reduce((sum, s) => sum + s.length, 0) / sentences.length 
      : 0;
    const complexity = Math.min(avgSentenceLength / 50, 1);

    // 抽象程度：抽象词汇比例
    const abstractWords = tokens.filter(t => t.length > 4);
    const abstraction = abstractWords.length / Math.max(tokens.length, 1);

    return {
      density,
      complexity,
      abstraction,
      concreteness: 1 - abstraction
    };
  }

  /**
   * 10. 分析情感曲线
   */
  private analyzeEmotionalArc(text: string, _tokens: string[]): WritingStyle['emotionalArc'] {
    // 简单的情感分析
    const positiveWords = ['好', '棒', '优秀', '喜欢', '开心', '快乐', '成功', '美好'];
    const negativeWords = ['坏', '差', '糟糕', '讨厌', '难过', '悲伤', '失败', '痛苦'];

    let positiveCount = 0;
    let negativeCount = 0;

    positiveWords.forEach(w => {
      const regex = new RegExp(w, 'g');
      const matches = text.match(regex);
      if (matches) positiveCount += matches.length;
    });

    negativeWords.forEach(w => {
      const regex = new RegExp(w, 'g');
      const matches = text.match(regex);
      if (matches) negativeCount += matches.length;
    });

    const sentimentScore = (positiveCount - negativeCount) / Math.max(positiveCount + negativeCount, 1);

    return {
      overallTrend: sentimentScore > 0.2 ? 'positive' : sentimentScore < -0.2 ? 'negative' : 'neutral',
      peaks: [], // 需要更细粒度的分析
      valleys: [],
      emotionalWords: [...positiveWords, ...negativeWords].filter(w => text.includes(w)),
      sentimentScore
    };
  }

  /**
   * 11. 分析受众意识
   */
  private analyzeAudienceAwareness(text: string, tokens: string[]): WritingStyle['audienceAwareness'] {
    // 检测互动程度
    const engagementMarkers = ['你', '您', '大家', '各位', '读者', '听众'];
    let engagementLevel = 0;
    engagementMarkers.forEach(m => {
      if (text.includes(m)) engagementLevel += 0.2;
    });
    engagementLevel = Math.min(engagementLevel, 1);

    // 检测行话使用
    const jargonUsage = tokens.filter(t => 
      /[A-Z]{2,}/.test(t) || /\d/.test(t) || t.length > 6
    );

    return {
      targetAudience: engagementLevel > 0.5 ? 'general' : 'professional',
      engagementLevel,
      accessibility: 1 - (jargonUsage.length / Math.max(tokens.length, 1)),
      jargonUsage: [...new Set(jargonUsage)].slice(0, 10)
    };
  }

  /**
   * 12. 分析文化语境与文体类型
   */
  private analyzeCulturalContext(text: string): WritingStyle['culturalContext'] {
    // 检测文体类型
    let genre = 'general';
    if (text.includes('代码') || text.includes('函数') || text.includes('API')) {
      genre = 'technical';
    } else if (text.includes('研究') || text.includes('论文') || text.includes('实验')) {
      genre = 'academic';
    } else if (text.includes('故事') || text.includes('小说') || text.includes('人物')) {
      genre = 'narrative';
    }

    // 检测语域
    let register: 'formal' | 'informal' | 'colloquial' = 'informal';
    if (text.includes('尊敬') || text.includes('阁下') || text.includes('谨')) {
      register = 'formal';
    } else if (text.includes('哈哈') || text.includes('嗯嗯') || text.includes('好吧')) {
      register = 'colloquial';
    }

    return {
      genre,
      culturalReferences: [],
      register,
      conventions: []
    };
  }

  /**
   * 13. 分析叙事声音与身份
   */
  private analyzeNarrativeVoice(text: string, _tokens: string[]): WritingStyle['narrativeVoice'] {
    // 检测立场态度
    const stanceMarkers = ['认为', '觉得', '相信', '主张', '反对', '支持'];
    const stance = stanceMarkers.filter(m => text.includes(m)).join(', ');

    return {
      voice: 'neutral',
      identity: '',
      stance,
      personality: []
    };
  }

  /**
   * 14. 分析节奏控制
   */
  private analyzeRhythmControl(sentences: string[], text: string): WritingStyle['rhythmControl'] {
    // 计算节奏快慢（基于平均句长）
    const avgLength = sentences.length > 0 
      ? sentences.reduce((sum, s) => sum + s.length, 0) / sentences.length 
      : 0;

    let pace: 'fast' | 'medium' | 'slow' = 'medium';
    if (avgLength < 20) pace = 'fast';
    else if (avgLength > 50) pace = 'slow';

    // 计算节奏变化
    const lengths = sentences.map(s => s.length);
    const variation = this.standardDeviation(lengths);

    return {
      pace,
      variation,
      pauseMarkers: ['，', '。', '；', '：'].filter(m => text.includes(m)),
      flowPatterns: []
    };
  }

  /**
   * 分析标点符号
   */
  private analyzePunctuation(text: string): WritingStyle['punctuation'] {
    const totalLength = text.length;
    
    return {
      commaFrequency: (text.match(/，/g) || []).length / Math.max(totalLength, 1),
      periodFrequency: (text.match(/。/g) || []).length / Math.max(totalLength, 1),
      specialPunctuation: ['！', '？', '……', '——'].filter(p => text.includes(p))
    };
  }

  // ===== 辅助方法 =====

  private countOccurrences(text: string, keyword: string): number {
    if (!keyword) {
      return 0;
    }
    const escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, 'g');
    return (text.match(regex) || []).length;
  }

  private average(numbers: number[]): number {
    if (numbers.length === 0) return 0;
    return numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
  }

  private standardDeviation(numbers: number[]): number {
    if (numbers.length === 0) return 0;
    const avg = this.average(numbers);
    const squareDiffs = numbers.map(n => Math.pow(n - avg, 2));
    return Math.sqrt(this.average(squareDiffs));
  }

  private extractSentencePatterns(sentences: string[]): string[] {
    // 简化：提取常见句式模式
    const patterns: string[] = [];
    sentences.forEach(s => {
      if (s.includes('因为') && s.includes('所以')) patterns.push('因果句');
      if (s.includes('虽然') && s.includes('但是')) patterns.push('转折句');
      if (s.includes('不仅') && s.includes('而且')) patterns.push('递进句');
    });
    return [...new Set(patterns)];
  }

  private detectOpeningStyle(paragraph: string): string {
    if (paragraph.length < 50) return '简洁开头';
    if (paragraph.includes('介绍') || paragraph.includes('概述')) return '概述式开头';
    if (paragraph.includes('问题') || paragraph.includes('为什么')) return '问题式开头';
    return '标准开头';
  }

  private detectEndingStyle(paragraph: string): string {
    if (paragraph.includes('总之') || paragraph.includes('综上所述')) return '总结式结尾';
    if (paragraph.includes('希望') || paragraph.includes('期待')) return '展望式结尾';
    return '标准结尾';
  }

  private detectParallelism(text: string): string[] {
    // 简化的排比检测
    const patterns: string[] = [];
    if ((text.match(/首先/g) || []).length >= 2) patterns.push('首先...其次...最后');
    if ((text.match(/一方面/g) || []).length >= 2) patterns.push('一方面...另一方面');
    return patterns;
  }

  private extractQuotes(text: string): string[] {
    const quotes: string[] = [];
    // 提取引号内容
    const quoteMatches = text.match(/[""「」『』]([^""「」『』]+)[""「」『』]/g);
    if (quoteMatches) {
      quotes.push(...quoteMatches.map(q => q.replace(/[""「」『』]/g, '')));
    }
    return quotes.slice(0, 10);
  }
}
