#!/bin/bash

# 调试启动验证脚本

echo "==================================="
echo "写作 Agent - 调试启动验证"
echo "==================================="
echo ""

cd "/Users/caomoumou/Documents/CodeArts 华为云码道/写作Agent"

# 1. 检查项目结构
echo "1️⃣ 检查项目结构..."
if [ -f "package.json" ]; then
    echo "   ✅ package.json 存在"
else
    echo "   ❌ package.json 不存在"
    exit 1
fi

if [ -f "tsconfig.json" ]; then
    echo "   ✅ tsconfig.json 存在"
else
    echo "   ❌ tsconfig.json 不存在"
    exit 1
fi

# 2. 检查调试配置
echo ""
echo "2️⃣ 检查调试配置..."
if [ -f ".vscode/launch.json" ]; then
    echo "   ✅ .vscode/launch.json 存在"
else
    echo "   ❌ .vscode/launch.json 不存在"
    exit 1
fi

if [ -f ".vscode/tasks.json" ]; then
    echo "   ✅ .vscode/tasks.json 存在"
else
    echo "   ❌ .vscode/tasks.json 不存在"
    exit 1
fi

# 3. 检查依赖
echo ""
echo "3️⃣ 检查依赖..."
if [ -d "node_modules" ]; then
    echo "   ✅ node_modules 存在"
else
    echo "   ⚠️  node_modules 不存在，正在安装..."
    npm install
fi

# 4. 检查编译输出
echo ""
echo "4️⃣ 检查编译输出..."
if [ -f "out/extension.js" ]; then
    echo "   ✅ out/extension.js 存在"
else
    echo "   ⚠️  out/extension.js 不存在，正在编译..."
    npm run compile
fi

# 5. 运行基础测试
echo ""
echo "5️⃣ 运行基础功能测试..."
node test/test-basic.js > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 基础功能测试通过"
else
    echo "   ❌ 基础功能测试失败"
    exit 1
fi

# 6. 检查主入口
echo ""
echo "6️⃣ 检查主入口配置..."
MAIN_ENTRY=$(cat package.json | grep '"main"' | cut -d'"' -f4)
if [ "$MAIN_ENTRY" = "./out/extension.js" ]; then
    echo "   ✅ 主入口配置正确: $MAIN_ENTRY"
else
    echo "   ❌ 主入口配置错误: $MAIN_ENTRY"
    exit 1
fi

echo ""
echo "==================================="
echo "✅ 所有检查通过！"
echo "==================================="
echo ""
echo "现在可以启动调试："
echo ""
echo "方法 1：在 CodeArts IDE 中"
echo "  1. 打开项目文件夹"
echo "  2. 按 F5 启动调试"
echo ""
echo "方法 2：使用命令面板"
echo "  1. 按 Cmd+Shift+P (Mac) 或 Ctrl+Shift+P (Windows)"
echo "  2. 输入 'Debug: Select and Start Debugging'"
echo "  3. 选择 'Run Extension'"
echo ""
echo "方法 3：命令行测试"
echo "  node test/test-basic.js"
echo ""
