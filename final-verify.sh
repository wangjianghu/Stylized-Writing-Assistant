#!/bin/bash

# 最终完整验证脚本

echo "========================================="
echo "VS Code 个人写作 Agent - 最终验证"
echo "========================================="
echo ""

cd "/Users/caomoumou/Documents/CodeArts 华为云码道/写作Agent"

# 1. 验证项目结构
echo "1️⃣ 验证项目结构..."
if [ -f "package.json" ] && [ -f "tsconfig.json" ] && [ -d "src" ]; then
    echo "   ✅ 项目结构完整"
else
    echo "   ❌ 项目结构不完整"
    exit 1
fi

# 2. 验证 npm 路径
echo ""
echo "2️⃣ 验证 npm 路径..."
if [ -f "/Users/caomoumou/.deskclaw/node/bin/npm" ]; then
    NPM_VERSION=$(/Users/caomoumou/.deskclaw/node/bin/npm --version)
    echo "   ✅ npm 路径正确 (版本 $NPM_VERSION)"
else
    echo "   ❌ npm 路径不存在"
    exit 1
fi

# 3. 验证调试配置
echo ""
echo "3️⃣ 验证调试配置..."
if grep -q '"启动扩展"' .vscode/launch.json && grep -q '"compile"' .vscode/tasks.json; then
    echo "   ✅ 调试配置正确"
else
    echo "   ❌ 调试配置错误"
    exit 1
fi

# 4. 验证任务配置
echo ""
echo "4️⃣ 验证任务配置..."
if grep -q "/Users/caomoumou/.deskclaw/node/bin/npm" .vscode/tasks.json; then
    echo "   ✅ 任务配置正确（使用 npm 绝对路径）"
else
    echo "   ❌ 任务配置错误"
    exit 1
fi

# 5. 测试编译
echo ""
echo "5️⃣ 测试编译..."
/Users/caomoumou/.deskclaw/node/bin/npm run compile > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 编译成功"
else
    echo "   ❌ 编译失败"
    exit 1
fi

# 6. 验证编译输出
echo ""
echo "6️⃣ 验证编译输出..."
if [ -f "out/extension.js" ]; then
    SIZE=$(ls -lh out/extension.js | awk '{print $5}')
    echo "   ✅ out/extension.js 存在 ($SIZE)"
else
    echo "   ❌ out/extension.js 不存在"
    exit 1
fi

# 7. 测试核心功能
echo ""
echo "7️⃣ 测试核心功能..."
node test/test-basic.js > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 功能测试通过"
else
    echo "   ❌ 功能测试失败"
    exit 1
fi

# 8. 验证文档
echo ""
echo "8️⃣ 验证文档..."
DOCS=(
    "README.md"
    "design.md"
    "最终完成报告.md"
    "问题解决总结.md"
    "npm路径已修复.md"
)
ALL_DOCS_EXIST=true
for doc in "${DOCS[@]}"; do
    if [ ! -f "$doc" ]; then
        ALL_DOCS_EXIST=false
        break
    fi
done
if [ "$ALL_DOCS_EXIST" = true ]; then
    echo "   ✅ 文档齐全"
else
    echo "   ⚠️  部分文档缺失"
fi

echo ""
echo "========================================="
echo "✅ 所有验证通过！"
echo "========================================="
echo ""
echo "项目已完全就绪，可以启动调试！"
echo ""
echo "启动步骤："
echo ""
echo "【步骤 1】重新加载窗口（必须！）"
echo "  按 Cmd+Shift+P (Mac) 或 Ctrl+Shift+P (Windows)"
echo "  输入：Developer: Reload Window"
echo "  回车"
echo ""
echo "【步骤 2】启动调试"
echo "  按 F5"
echo "  或在调试面板选择'启动扩展'，点击绿色按钮"
echo ""
echo "【备选方案】如果调试还有问题："
echo "  1. 手动编译：npm run compile"
echo "  2. 选择'启动扩展（不编译）'"
echo ""
echo "【测试功能】在新窗口中："
echo "  Cmd+Shift+I - 导入文章"
echo "  Cmd+Shift+G - 生成文章"
echo "  Cmd+Shift+A - 分析风格"
echo ""
