#!/bin/bash

# CodeArts IDE 调试启动最终脚本

echo "==================================="
echo "CodeArts IDE 调试启动"
echo "==================================="
echo ""

cd "/Users/caomoumou/Documents/CodeArts 华为云码道/写作Agent"

# 验证所有配置
echo "1️⃣ 验证编译..."
npm run compile > /dev/null 2>&1
if [ $? -eq 0 ]; then
    echo "   ✅ 编译成功"
else
    echo "   ❌ 编译失败"
    exit 1
fi

echo ""
echo "2️⃣ 验证输出..."
if [ -f "out/extension.js" ]; then
    echo "   ✅ out/extension.js 存在 ($(ls -lh out/extension.js | awk '{print $5}'))"
else
    echo "   ❌ out/extension.js 不存在"
    exit 1
fi

echo ""
echo "3️⃣ 验证配置..."
if grep -q '"启动扩展"' .vscode/launch.json; then
    echo "   ✅ 调试配置正确（中文）"
else
    echo "   ❌ 调试配置错误"
    exit 1
fi

echo ""
echo "4️⃣ 验证功能..."
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
echo "【重要】步骤 1：重新加载窗口"
echo "  按 Cmd+Shift+P (Mac) 或 Ctrl+Shift+P (Windows)"
echo "  输入：Developer: Reload Window"
echo "  回车"
echo ""
echo "【重要】步骤 2：打开调试面板"
echo "  点击左侧'运行和调试'图标"
echo "  或按 Cmd+Shift+D (Mac) / Ctrl+Shift+D (Windows)"
echo ""
echo "【重要】步骤 3：选择配置"
echo "  在顶部下拉菜单中应该能看到："
echo "  ✅ 启动扩展"
echo "  ✅ 启动扩展（不编译）"
echo ""
echo "【重要】步骤 4：开始调试"
echo "  选择'启动扩展'"
echo "  点击绿色'开始调试'按钮（▶️）"
echo "  或直接按 F5"
echo ""
echo "==================================="
echo "如果还是看不到配置，请告诉我："
echo "1. 重新加载窗口后，调试面板显示什么？"
echo "2. 下拉菜单中有哪些选项？"
echo "==================================="
