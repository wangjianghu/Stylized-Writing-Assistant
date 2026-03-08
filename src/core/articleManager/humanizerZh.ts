import { WritingStyle } from '../types';

export interface HumanizerZhOptions {
  style?: WritingStyle;
  topic?: string;
  strength?: number;
}

// Rules aligned with op7418/Humanizer-zh project (MIT):
// https://github.com/op7418/Humanizer-zh
const AI_WORD_REPLACEMENTS: Array<{ pattern: RegExp; replacement: string }> = [
  { pattern: /赋能/g, replacement: '支持' },
  { pattern: /抓手/g, replacement: '关键点' },
  { pattern: /闭环/g, replacement: '完整流程' },
  { pattern: /打通/g, replacement: '连接' },
  { pattern: /链路/g, replacement: '流程' },
  { pattern: /颗粒度/g, replacement: '细节程度' },
  { pattern: /抽象度/g, replacement: '抽象程度' },
  { pattern: /体系化/g, replacement: '系统化' },
  { pattern: /认知(?![障症])/g, replacement: '理解' },
  { pattern: /范式/g, replacement: '方式' },
  { pattern: /场景化/g, replacement: '结合场景' },
  { pattern: /高频地|高频/g, replacement: '经常' }
];

const SLOGAN_REPLACEMENTS: Array<{ pattern: RegExp; replacement: string }> = [
  { pattern: /在当今时代[，,]?/g, replacement: '' },
  { pattern: /毋庸置疑[，,]?/g, replacement: '' },
  { pattern: /值得一提的是[，,]?/g, replacement: '' },
  { pattern: /总的来说[，,]?/g, replacement: '' },
  { pattern: /综上所述[，,]?/g, replacement: '' },
  { pattern: /让我们(?:一起来)?(?:看一下|看看|来看看)/g, replacement: '' },
  { pattern: /我们可以看到/g, replacement: '可以看到' },
  { pattern: /大家都知道/g, replacement: '' },
  { pattern: /不难发现/g, replacement: '' },
  { pattern: /显而易见/g, replacement: '' },
  { pattern: /极其|非常|特别|相当|十分/g, replacement: '较' },
  { pattern: /革命性的?/g, replacement: '明显的' },
  { pattern: /颠覆性的?/g, replacement: '重要的' },
  { pattern: /前所未有的?/g, replacement: '新的' }
];

function normalizePunctuation(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/[—–]{2,}/g, '，')
    .replace(/…{2,}/g, '…')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/，{2,}/g, '，')
    .replace(/。{2,}/g, '。')
    .trim();
}

function dedupeNeighborLines(text: string): string {
  const lines = text.split('\n');
  const result: string[] = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed && result.length > 0 && result[result.length - 1].trim() === trimmed) {
      continue;
    }
    result.push(line);
  }
  return result.join('\n');
}

function reduceMechanicalSequence(text: string): string {
  return text
    .replace(/首先[，,、]?\s*/g, '先看关键问题：')
    .replace(/其次[，,、]?\s*/g, '再看推进条件：')
    .replace(/再次[，,、]?\s*/g, '继续看执行细节：')
    .replace(/最后[，,、]?\s*/g, '收束到结论：');
}

function cleanupMarkdownNoise(text: string): string {
  return text
    .replace(/^\s*[\-\*]\s+\*\*([^*]+)\*\*[:：]\s*/gm, '$1：')
    .replace(/^\s*[\-\*]\s+/gm, '')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/[`]{3,}[\s\S]*?[`]{3,}/g, match => match.replace(/\*\*/g, ''));
}

function removeEmojiAndVisualNoise(text: string): string {
  return text
    .replace(/[\u{1F300}-\u{1FAFF}]/gu, '')
    .replace(/[▶►◆◇■□⭐✨🔥💡✅❌⚠️📌📍🎯]/gu, '')
    .replace(/[|｜]{2,}/g, '|');
}

function cleanupAiMetaExpression(text: string): string {
  return text
    .replace(/作为(?:一个)?(?:AI|人工智能)(?:模型|助手)?[，,]?\s*/gi, '')
    .replace(/我(?:们)?(?:将|会)?(从以下几个方面|分三点|展开说明)[：:]/g, '')
    .replace(/以下是(?:我的|本次)?(?:回答|分析|建议)[：:]/g, '')
    .replace(/总之[，,]?/g, '')
    .replace(/总而言之[，,]?/g, '')
    .replace(/此外[，,]此外[，,]?/g, '此外，')
    .replace(/同时[，,]同时[，,]?/g, '同时，');
}

function reduceParallelAndVerbose(text: string): string {
  return text
    .replace(/不仅[^，。！？!?]{0,30}，还[^，。！？!?]{0,30}/g, match =>
      match.replace(/^不仅/, '').replace('，还', '，并且'))
    .replace(/通过([^，。！？!?]{1,24})，实现([^，。！？!?]{1,24})/g, '通过$1，完成$2')
    .replace(/为了([^，。！？!?]{1,24})，我们需要([^，。！？!?]{1,24})/g, '要$1，需要$2')
    .replace(/(在[^，。！？!?]{2,30}方面)[，,](在[^，。！？!?]{2,30}方面)/g, '$1和$2');
}

function softenAbsolutes(text: string): string {
  return text
    .replace(/必须/g, '需要')
    .replace(/一定要/g, '建议')
    .replace(/完全/g, '尽量')
    .replace(/绝对/g, '通常')
    .replace(/唯一/g, '主要');
}

function stripGenericParagraphs(text: string): string {
  const genericPatterns = [
    /^本文将(?:介绍|讨论|分析)/,
    /^关于.+我们(?:可以|能够)/,
    /^在.+方面，.+具有(?:重要|可观)价值/,
    /^总之，?.+是一个(?:重要|值得)主题/,
    /^这是一个(?:重要|复杂)的话题/
  ];

  const paragraphs = text
    .split(/\n{2,}/)
    .map(item => item.trim())
    .filter(Boolean);

  const filtered = paragraphs.filter(item => {
    if (item.length > 90) {
      return true;
    }
    return !genericPatterns.some(pattern => pattern.test(item));
  });

  return filtered.join('\n\n');
}

function diversifySentenceOpeners(text: string): string {
  const replacements: Array<[RegExp, string]> = [
    [/^因此，/gm, '换个角度看，'],
    [/^另外，/gm, '再往下看，'],
    [/^同时，/gm, '并且，'],
    [/^此外，/gm, '再补一层，']
  ];

  let output = text;
  for (const [pattern, replacement] of replacements) {
    output = output.replace(pattern, replacement);
  }
  return output;
}

function alignTone(text: string, tone: WritingStyle['tone']['type']): string {
  if (tone === 'serious' || tone === 'formal') {
    return text
      .replace(/！/g, '。')
      .replace(/太/g, '较')
      .replace(/真的/g, '确实')
      .replace(/有点/g, '略有')
      .replace(/哈哈|呵呵|嘛|啦/g, '');
  }
  if (tone === 'casual') {
    return text
      .replace(/可以观察到/g, '能看到')
      .replace(/需要注意的是/g, '要注意的是')
      .replace(/在实践中/g, '落到实际场景里');
  }
  return text;
}

function alignSentenceByLength(text: string, targetLength: number): string {
  if (targetLength <= 0) {
    return text;
  }
  if (targetLength < 18) {
    return text.replace(/([^。！？!?]{28,})，/g, '$1。');
  }
  if (targetLength > 30) {
    return text
      .replace(/。(\n{1,2}[^#\n]{0,14}[。！？!?])/g, '，并且$1')
      .replace(/。\n([^\n#]{0,12}[。！？!?])/g, '，$1');
  }
  return text;
}

function applyReplaceRules(text: string, rules: Array<{ pattern: RegExp; replacement: string }>): string {
  let output = text;
  for (const rule of rules) {
    output = output.replace(rule.pattern, rule.replacement);
  }
  return output;
}

export function humanizeZh(text: string, options: HumanizerZhOptions = {}): string {
  let output = normalizePunctuation(text);

  output = removeEmojiAndVisualNoise(output);
  output = cleanupMarkdownNoise(output);
  output = cleanupAiMetaExpression(output);
  output = applyReplaceRules(output, AI_WORD_REPLACEMENTS);
  output = applyReplaceRules(output, SLOGAN_REPLACEMENTS);

  output = dedupeNeighborLines(output);
  output = reduceMechanicalSequence(output);
  output = reduceParallelAndVerbose(output);
  output = softenAbsolutes(output);
  output = stripGenericParagraphs(output);
  output = diversifySentenceOpeners(output);
  output = alignTone(output, options.style?.tone?.type || 'neutral');
  output = alignSentenceByLength(output, options.style?.sentenceStructure?.avgLength || 0);

  return normalizePunctuation(output);
}
