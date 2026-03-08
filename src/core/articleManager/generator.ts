import { MaterialType, WritingMaterial, WritingStyle } from '../types';
import { humanizeZh } from './humanizerZh';

export interface StyleDraft {
  body: string;
  conclusion: string;
}

function pickMaterial(materials: WritingMaterial[], preferredTypes: MaterialType[]): WritingMaterial | undefined {
  for (const type of preferredTypes) {
    const target = materials.find(item => item.type === type && item.content.trim().length >= 20);
    if (target) {
      return target;
    }
  }
  return materials.find(item => item.content.trim().length >= 20);
}

function snippet(text: string, maxLength: number): string {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (clean.length <= maxLength) {
    return clean;
  }
  return `${clean.slice(0, maxLength)}...`;
}

function toneLine(tone: WritingStyle['tone']['type']): string {
  switch (tone) {
    case 'serious':
      return '先把边界和前提说清楚，再谈结论。';
    case 'formal':
      return '建议以可验证的事实为主线，避免情绪化判断。';
    case 'casual':
      return '直接说人话：这个主题的价值，得落在真实场景里。';
    case 'humorous':
      return '可以稍微轻一点，但核心观点必须落地。';
    default:
      return '先把问题讲清楚，再推进到方案层面。';
  }
}

export function generateStyleAlignedDraft(
  topic: string,
  style: WritingStyle,
  materials: WritingMaterial[]
): StyleDraft {
  const transitions = style.contentStructure.transitionWords.slice(0, 6);
  const transitionA = transitions[0] || '因此';
  const transitionB = transitions[1] || '另外';
  const transitionC = transitions[2] || '最后';
  const favoriteWords = style.vocabulary.favoriteWords.slice(0, 3);
  const keywordLine = favoriteWords.length > 0 ? `优先围绕 ${favoriteWords.join('、')} 这些词组织表达。` : '';

  const openingMaterial = pickMaterial(materials, [MaterialType.OPENING, MaterialType.PARAGRAPH, MaterialType.SENTENCE]);
  const evidenceMaterial = pickMaterial(materials, [MaterialType.PARAGRAPH, MaterialType.SENTENCE, MaterialType.QUOTE]);
  const endingMaterial = pickMaterial(materials, [MaterialType.ENDING, MaterialType.PARAGRAPH, MaterialType.SENTENCE]);

  const openingEvidence = openingMaterial
    ? `可借鉴的开场信息：${snippet(openingMaterial.content, 90)}`
    : '可先补一段背景：这个主题为什么现在值得写。';
  const mainEvidence = evidenceMaterial
    ? `素材参考：${snippet(evidenceMaterial.content, 120)}`
    : '建议补充一条真实案例或数据，避免泛泛描述。';
  const endingEvidence = endingMaterial
    ? `收束时可参考：${snippet(endingMaterial.content, 90)}`
    : '结尾建议落到“下一步行动”和“评估指标”。';

  const body = [
    '### 问题边界与判断标准',
    `${toneLine(style.tone.type)} 围绕「${topic}」，建议先回答三件事：对象是谁、目标是什么、约束是什么。${transitionA}，写作时不要只写“是什么”，要写“如何判断是否做对”。`,
    `${openingEvidence} ${keywordLine}`.trim(),
    '',
    '### 典型场景与可执行方案',
    `把「${topic}」拆成可执行的步骤：先定义输入，再说明过程，最后给出产出。${transitionB}，每个步骤都要附一条可验证证据，比如指标、案例或反例。`,
    `${mainEvidence} 如果当前素材不足，至少补一段“失败成本”与“替代方案”的比较。`,
    '',
    '### 风险、取舍与迭代',
    `写到这个层面时，重点不再是堆观点，而是说明取舍逻辑：什么情况下该做，什么情况下不该做。${transitionC}，把迭代节奏写清楚，避免“一次性完成”的假设。`,
    `${endingEvidence}`
  ].join('\n\n');

  const conclusion = [
    `围绕「${topic}」的写作，如果能同时覆盖“边界、证据、取舍、迭代”，文章就会更像真实经验，而不是模板总结。`,
    '下一步建议：补齐一个真实场景样本，再把结论压缩成可执行清单。'
  ].join('\n');

  return {
    body: humanizeZh(body, { style, topic, strength: 0.8 }),
    conclusion: humanizeZh(conclusion, { style, topic, strength: 0.7 })
  };
}

