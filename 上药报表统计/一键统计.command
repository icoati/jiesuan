#!/bin/bash
# ====================================================
#  心血管内科报表一键统计（Mac 双击运行 - Python 原生版）
# ====================================================

# 切换到脚本所在目录（双击时 cwd 是 $HOME，必须切换）
cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  心血管内科报表一键统计 (Python 原生版)"
echo "=================================================="
echo ""

# 检查 Python 运行环境
if command -v python3 &> /dev/null; then
    PY_BIN="python3"
elif command -v python &> /dev/null; then
    PY_BIN="python"
else
    echo "❌ 未检测到 Python，请先安装 Python 3"
    echo "   推荐访问 https://www.python.org 下载安装"
    echo ""
    echo "按回车键退出..."
    read
    exit 1
fi

echo "✓ 检测到 Python: $($PY_BIN --version)"
echo ""
echo "🚀 开始生成报表..."
echo "--------------------------------------------------"
$PY_BIN 统计报表.py

echo ""
echo "=================================================="
echo "✅ 完成！报表已保存在当前文件夹"
echo "=================================================="
echo ""
echo "按回车键关闭窗口..."
read
