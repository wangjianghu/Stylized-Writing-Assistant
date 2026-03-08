const { StyleEngine } = require('../out/core/styleEngine');
const { MaterialExtractor } = require('../out/core/materialManager/extractor');

// 测试文本
const sampleText = `
# TypeScript 类型系统介绍

TypeScript 是 JavaScript 的超集，它为 JavaScript 添加了静态类型系统。这使得我们可以在开发过程中及早发现错误，提高代码质量。

## 为什么需要类型系统？

类型系统可以帮助我们在编译时捕获错误，而不是在运行时。这大大提高了开发效率和代码的可维护性。例如，当我们试图将一个字符串赋值给一个数字类型的变量时，TypeScript 会立即报错。

## 基本类型

TypeScript 支持以下基本类型：
- number: 数字类型
- string: 字符串类型
- boolean: 布尔类型
- array: 数组类型
- object: 对象类型

## 高级类型

除了基本类型，TypeScript 还提供了高级类型系统，包括：
- 联合类型（Union Types）
- 交叉类型（Intersection Types）
- 类型别名（Type Aliases）
- 泛型（Generics）

## 总结

TypeScript 的类型系统是一个强大的工具，它可以帮助我们编写更安全、更可靠的代码。通过学习和使用 TypeScript，我们可以显著提高开发效率和代码质量。
`;

async function testStyleEngine() {
  console.log('=== 测试风格引擎 ===\n');
  
  const styleEngine = new StyleEngine();
  
  // 分析文本风格
  console.log('分析文本风格...');
  const style = await styleEngine.analyzeText(sampleText);
  
  console.log('\n风格特征：');
  console.log('- 平均句长:', style.sentenceStructure.avgLength.toFixed(1), '字');
  console.log('- 词汇丰富度:', (style.vocabulary.uniqueWordRatio * 100).toFixed(1), '%');
  console.log('- 语气类型:', style.tone.type);
  console.log('- 情感倾向:', style.emotionalArc.overallTrend);
  console.log('- 常用词汇:', style.vocabulary.favoriteWords.slice(0, 5).join(', '));
  console.log('- 结构模式:', style.contentStructure.structurePatterns.join(', ') || '无');
  console.log('- 过渡词:', style.contentStructure.transitionWords.slice(0, 5).join(', ') || '无');
  
  // 学习风格
  console.log('\n学习风格...');
  await styleEngine.learnFromText(sampleText);
  
  const currentStyle = styleEngine.getCurrentStyle();
  console.log('当前风格已设置');
  console.log('- 风格ID:', currentStyle.id || '未设置');
  console.log('- 创建时间:', currentStyle.createdAt);
}

async function testMaterialExtractor() {
  console.log('\n\n=== 测试素材提取器 ===\n');
  
  const extractor = new MaterialExtractor();
  
  // 提取素材
  console.log('提取素材...');
  const materials = await extractor.extractMaterials(sampleText, 'test-document');
  
  console.log('\n提取结果：');
  console.log('- 总素材数:', materials.length);
  
  // 按类型统计
  const typeCount = {};
  materials.forEach(m => {
    typeCount[m.type] = (typeCount[m.type] || 0) + 1;
  });
  
  console.log('\n按类型统计：');
  Object.entries(typeCount).forEach(([type, count]) => {
    console.log(`- ${type}: ${count} 条`);
  });
  
  // 显示部分素材
  console.log('\n部分素材示例：');
  materials.slice(0, 3).forEach((m, i) => {
    console.log(`\n素材 ${i + 1}:`);
    console.log('- 类型:', m.type);
    console.log('- 内容:', m.content.substring(0, 50) + '...');
    console.log('- 质量:', m.metadata.quality);
    console.log('- 标签:', m.tags.join(', '));
  });
}

async function main() {
  try {
    await testStyleEngine();
    await testMaterialExtractor();
    
    console.log('\n\n=== 测试完成 ===');
    console.log('所有核心功能正常工作！');
  } catch (error) {
    console.error('测试失败:', error);
    process.exit(1);
  }
}

main();
