#!/usr/bin/env bash
# ==============================================================================
# 医疗健康语料库 - 每月结算一键生成脚本 (Mac 专用一键双击运行)
# ==============================================================================

# 1. 自动切换到当前脚本所在文件夹（Mac 双击必须的步骤）
cd "$(dirname "$0")"

echo "============================================================"
echo "      🚀 医疗健康语料库 - 每月结算一键生成工具 (macOS)        "
echo "============================================================"
echo "📁 当前工作文件夹: $(pwd)"
echo ""

# 2. 检查 Python 3 环境
if ! command -v python3 &> /dev/null; then
    echo "❌ 未检测到 Python 3 环境！"
    echo "👉 请先安装 Python 3 或 Xcode Command Line Tools（可在终端运行: xcode-select --install）"
    read -n 1 -s -r -p "按任意键退出..."
    exit 1
fi

echo "🔍 检测到 Python 版本: $(python3 --version)"

# 3. 检查并自动安装缺失的依赖库
echo "🔍 正在检查运行依赖库..."
python3 -c "import pandas, openpyxl, xlwt, xlrd" &> /dev/null
if [ $? -ne 0 ]; then
    echo "⚠️ 首次运行或检测到缺失依赖，正在自动为您安装所需的 Python 库 (openpyxl, pandas, xlwt, xlrd)..."
    python3 -m pip install --user openpyxl pandas xlwt xlrd
    if [ $? -ne 0 ]; then
        echo "❌ 依赖安装失败，请尝试在终端手动运行: pip3 install openpyxl pandas xlwt xlrd"
        read -n 1 -s -r -p "按任意键退出..."
        exit 1
    fi
    echo "✅ 依赖库安装完成！"
fi

# 4. 执行数据提取与结算汇总
echo ""
python3 settlement_tool.py "$(pwd)"

# 5. 打开 Finder 窗口并高亮显示生成的文件
echo "📂 正在为您打开结果所在文件夹..."
open .

echo ""
echo "🎉 全部操作已完成！可随时关闭本窗口。"
echo "============================================================"
read -n 1 -s -r -p "按任意键关闭窗口..."
exit 0
