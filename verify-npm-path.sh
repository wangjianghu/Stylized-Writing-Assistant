#!/bin/bash

# 最终验证和启动脚本

echo "==================================="
echo "✅ npm 路径已修复"
echo "==================================="
echo ""

cd "/Users/caomoumou/Documents/CodeArts 华为云码道/写作Agent"

# 验证 npm 路径
echo "1️⃣ 验证 npm 路径..."
if [ -f "/Users/caomoumou/.deskclaw/node/bin/npm" ]; then
    echo "   ✅ npm 路径正确"
    echo "   版本: $(/Users/caomoumou/.deskclaw/node/bin/npm --version)"
else
    echo "   ❌ npm 路径不存在"
    exit 1
fi

# 验证配置
echo ""
echo "2️⃣ 验证 tasks.json..."
if grep -q "/Users/caomoumou/.deskclaw/node/bin/npm" .vscode/tasks.json; then
    echo "   ✅ tasks.json 配置正确"
else
    echo "   ❌ tasks.json 配置错误"
    exit 1
fi

# 测试编译
echo ""
echo "3️⃣ 测试编译..."
/Users/caomoumou/.deskclaw/node/bin/npm run compile > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 编译成功"
else
    echo "   ❌ 编译失败"
    exit 1
fi

# 验证输出
echo ""
echo "4️⃣ 验证输出..."
if [ -f "out/extension.js" ]; then
    echo "   ✅ out/extension.js 存在 ($(ls -lh out/extension.js | awk '{print $5}'))"
else
    echo "   ❌ out/extension.js 不存在"
    exit 1
fi

# 测试功能
echo ""
echo "5️⃣ 测试功能..."
node test/test-basic.js > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 功能测试通过"
else
    echo "   ❌ 功能测试失败"
    exit 1
fi

echo ""
echo "==================================="
echo "✅ 所有验证通过！"
echo "==================================="
echo ""
echo "现在请在 CodeArts IDE 中："
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
echo "==================================="
echo "如果还有问题，请使用无编译模式："
echo "1. 先手动编译：npm run compile"
echo "2. 选择'启动扩展（不编译）'"
echo "==================================="
