#!/bin/bash

# 一键启动调试脚本

echo "==================================="
echo "写作 Agent - 一键启动调试"
echo "==================================="
echo ""

cd "/Users/caomoumou/Documents/CodeArts 华为云码道/写作Agent"

# 1. 编译项目
echo "1️⃣ 编译项目..."
npm run compile > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 编译成功"
else
    echo "   ❌ 编译失败"
    exit 1
fi

# 2. 验证输出
echo ""
echo "2️⃣ 验证编译输出..."
if [ -f "out/extension.js" ]; then
    echo "   ✅ out/extension.js 存在"
else
    echo "   ❌ out/extension.js 不存在"
    exit 1
fi

# 3. 运行测试
echo ""
echo "3️⃣ 运行功能测试..."
node test/test-basic.js > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 功能测试通过"
else
    echo "   ❌ 功能测试失败"
    exit 1
fi

echo ""
echo "==================================="
echo "✅ 所有检查通过！"
echo "==================================="
echo ""
echo "现在请在 CodeArts IDE 中："
echo ""
echo "1. 重新加载窗口"
echo "   按 Cmd+Shift+P (Mac) 或 Ctrl+Shift+P (Windows)"
echo "   输入：Developer: Reload Window"
echo "   回车"
echo ""
echo "2. 启动调试"
echo "   按 F5"
echo ""
echo "或者使用命令面板："
echo "   按 Cmd+Shift+P"
echo "   输入：Debug: Select and Start Debugging"
echo "   选择：Run Extension"
echo ""
