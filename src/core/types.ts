/**
 * 写作风格类型定义（14维度分析框架）
 */
export interface WritingStyle {
  // ===== 核心维度 =====

  // 1. 词汇特征
  vocabulary: {
    avgWordLength: number;        // 平均词长
    uniqueWordRatio: number;      // 独特词汇比例
    favoriteWords: string[];      // 高频偏好词汇
    rareWords: string[];          // 稀有词汇使用
    terminology: string[];        // 专业术语使用
    colloquialisms: string[];     // 口语化表达
  };

  // 2. 句式结构
  sentenceStructure: {
    avgLength: number;            // 平均句长
    lengthDistribution: Map<number, number>; // 句长分布
    complexSentenceRatio: number; // 复杂句比例
    questionRatio: number;        // 疑问句比例
    exclamationRatio: number;     // 感叹句比例
    sentencePatterns: string[];   // 常用句式模式
  };

  // 3. 语气语调
  tone: {
    type: 'formal' | 'casual' | 'humorous' | 'serious' | 'neutral';
    intensity: number;            // 语气强度
    consistency: number;          // 语气一致性
    markers: string[];            // 语气标记词
  };

  // 4. 内容结构
  contentStructure: {
    paragraphLength: number;      // 平均段落长度
    structurePatterns: string[];  // 结构模式（总分总、递进等）
    transitionWords: string[];    // 过渡词使用
    openingStyle: string;         // 开头方式
    endingStyle: string;          // 结尾方式
  };

  // 5. 修辞手法
  rhetoric: {
    metaphor: string[];           // 比喻用法
    parallelism: string[];        // 排比句式
    personification: string[];    // 拟人手法
    exaggeration: string[];       // 夸张表达
    quotes: string[];             // 引用习惯
    rhetoricalDevices: string[];  // 其他修辞手法
  };

  // ===== 深度维度 =====

  // 6. 时间视角
  timePerspective: {
    dominant: 'past' | 'present' | 'future'; // 主导时态
    tenseUsage: Map<string, number>;         // 时态使用分布
    temporalMarkers: string[];               // 时间标记词
    flashbackUsage: boolean;                 // 倒叙使用
  };

  // 7. 空间视角
  spacePerspective: {
    viewpoint: 'first-person' | 'third-person' | 'omniscient'; // 叙事视角
    spatialMarkers: string[];     // 空间标记词
    sceneDescription: string[];   // 场景描写方式
  };

  // 8. 逻辑类型
  logicType: {
    type: 'deductive' | 'inductive' | 'abductive' | 'narrative'; // 逻辑类型
    reasoningPatterns: string[];  // 推理模式
    evidenceStyle: string;        // 论据风格
    connectives: string[];        // 逻辑连接词
  };

  // 9. 信息密度与复杂度
  informationDensity: {
    density: number;              // 信息密度
    complexity: number;           // 复杂度
    abstraction: number;          // 抽象程度
    concreteness: number;         // 具体程度
  };

  // 10. 情感曲线
  emotionalArc: {
    overallTrend: string;         // 整体趋势
    peaks: number[];              // 情感高峰位置
    valleys: number[];            // 情感低谷位置
    emotionalWords: string[];     // 情感词汇
    sentimentScore: number;       // 情感倾向分数
  };

  // ===== 语境维度 =====

  // 11. 受众意识
  audienceAwareness: {
    targetAudience: string;       // 目标受众
    engagementLevel: number;      // 互动程度
    accessibility: number;        // 可理解性
    jargonUsage: string[];        // 行话使用
  };

  // 12. 文化语境与文体类型
  culturalContext: {
    genre: string;                // 文体类型
    culturalReferences: string[]; // 文化引用
    register: 'formal' | 'informal' | 'colloquial'; // 语域
    conventions: string[];        // 文体惯例
  };

  // 13. 叙事声音与身份
  narrativeVoice: {
    voice: string;                // 叙事声音
    identity: string;             // 作者身份表达
    stance: string;               // 立场态度
    personality: string[];        // 个性特征
  };

  // 14. 节奏控制
  rhythmControl: {
    pace: 'fast' | 'medium' | 'slow'; // 节奏快慢
    variation: number;            // 节奏变化
    pauseMarkers: string[];       // 停顿标记
    flowPatterns: string[];       // 流畅度模式
  };

  // 标点特征（辅助维度）
  punctuation: {
    commaFrequency: number;       // 逗号使用频率
    periodFrequency: number;      // 句号使用频率
    specialPunctuation: string[]; // 特殊标点使用
  };

  // 元数据
  id?: string;
  name?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

/**
 * 风格档案（多风格管理）
 */
export interface StyleProfile {
  id: string;
  name: string;
  style: WritingStyle;
  createdAt: Date;
  updatedAt: Date;
  articleCount: number;
  lastSource?: string;
}

/**
 * 素材类型
 */
export enum MaterialType {
  SENTENCE = 'sentence',          // 精彩句子
  PARAGRAPH = 'paragraph',        // 优秀段落
  QUOTE = 'quote',                // 引用
  METAPHOR = 'metaphor',          // 比喻
  OPENING = 'opening',            // 开头
  ENDING = 'ending',              // 结尾
  TRANSITION = 'transition',      // 过渡
  IDEA = 'idea',                  // 创意点
  STYLE_SAMPLE = 'style-sample'   // 风格样本
}

/**
 * 写作素材
 */
export interface WritingMaterial {
  // 基本信息
  id: string;
  name: string;
  content: string;
  type: MaterialType;
  tags: string[];
  category: string;

  // 元数据
  metadata: {
    source: string;
    createdAt: Date;
    updatedAt: Date;
    usedCount: number;
    lastUsedAt?: Date;
    quality: number;
  };

  // 自动提取的特征
  features: {
    keywords: string[];
    sentiment: string;
    style: string;
    quality: number;
  };

  // 风格特征（14维度分析结果）
  styleFeatures?: Partial<WritingStyle>;

  // 风格规则
  rules?: {
    mustFollow: string[];
    recommended: string[];
    avoid: string[];
  };

  // 关联信息
  relations?: {
    similarMaterials: string[];
    relatedStyles: string[];
    sceneAssociations: string[];
    tagAssociations: string[];
  };

  // 向量表示
  embedding?: number[];

  // 聚类信息
  clusterInfo?: {
    clusterId: string;
    isRepresentative: boolean;
    distance: number;
  };
}

/**
 * 文章生成选项
 */
export interface GenerationOptions {
  length?: number;              // 目标字数
  structure?: string;           // 指定结构类型
  tone?: string;                // 指定语气
  includeMaterials?: boolean;   // 是否引用素材库
  aiEnhanced?: boolean;         // 是否使用 AI 增强
}

/**
 * 生成的文章
 */
export interface GeneratedArticle {
  content: string;
  styleScore: number;           // 风格一致性评分 (0-1)
  usedMaterials: string[];      // 使用的素材ID
  metadata: {
    topic: string;
    generatedAt: Date;
    styleProfile: string;
  };
}

/**
 * 文章结构
 */
export interface ArticleStructure {
  type: string;
  sections: SectionDefinition[];
}

/**
 * 章节定义
 */
export interface SectionDefinition {
  type: string;
  [key: string]: any;
}

/**
 * 处理后的文本
 */
export interface ProcessedText {
  content: string;
  source: string;
  format: string;
  metadata: {
    originalFormat?: string;
    length: number;
    processedAt: Date;
  };
}

/**
 * 搜索查询
 */
export interface SearchQuery {
  keywords?: string[];
  semantic?: string;
  tags?: string[];
  types?: MaterialType[];
  limit?: number;
}

/**
 * 推荐上下文
 */
export interface RecommendationContext {
  targetStyle?: Partial<WritingStyle>;
  scene?: string;
  tags?: string[];
  keywords?: string[];
}
