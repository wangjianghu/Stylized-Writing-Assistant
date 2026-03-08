#!/bin/bash

# 写作 Agent 快速启动脚本

echo "==================================="
echo "写作 Agent - 快速启动"
echo "==================================="
echo ""

# 进入项目目录
cd "/Users/caomoumou/Documents/CodeArts 华为云码道/写作Agent"

# 检查 node_modules
if [ ! -d "node_modules" ]; then
    echo "📦 安装依赖..."
    npm install
    echo ""
fi

# 编译项目
echo "🔨 编译项目..."
npm run compile

if [ $? -eq 0 ]; then
    echo "✅ 编译成功！"
    echo ""
    echo "==================================="
    echo "现在可以："
    echo "1. 在 CodeArts IDE 中打开项目"
    echo "2. 按 F5 启动调试"
    echo "3. 或运行: node test/test-basic.js"
    echo "==================================="
else
    echo "❌ 编译失败，请检查错误信息"
fi
