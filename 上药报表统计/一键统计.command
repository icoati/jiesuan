#!/bin/bash
# ====================================================
#  心血管内科报表一键统计（Mac 双击运行）
# ====================================================

# 切换到脚本所在目录（双击时 cwd 是 $HOME，必须切换）
cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  心血管内科报表一键统计"
echo "=================================================="
echo ""

# 检查 Node.js 是否安装
if ! command -v node &> /dev/null; then
    echo "❌ 未检测到 Node.js，请先安装："
    echo "   打开 https://nodejs.org 下载 LTS 版本安装"
    echo "   或终端运行：brew install node"
    echo ""
    echo "按回车键退出..."
    read
    exit 1
fi

echo "✓ 检测到 Node.js: $(node --version)"

# 检查 node_modules（首次使用需安装依赖）
if [ ! -d "node_modules" ]; then
    echo ""
    echo "📦 首次使用，正在安装依赖（只需一次）..."
    npm install xlsx 2>&1

    if [ ! -d "node_modules" ]; then
        echo ""
        echo "❌ 依赖安装失败，请手动执行："
        echo "   cd \"$(pwd)\""
        echo "   npm install xlsx"
        echo ""
        echo "按回车键退出..."
        read
        exit 1
    fi
    echo "✓ 依赖安装完成"
fi

echo ""
echo "🚀 开始生成报表..."
echo "--------------------------------------------------"
node 统计报表.js

echo ""
echo "=================================================="
echo "✅ 完成！报表已保存在当前文件夹"
echo "=================================================="
echo ""
echo "按回车键关闭窗口..."
read
