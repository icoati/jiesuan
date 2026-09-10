#!/bin/bash
# 进入脚本所在当前目录
cd "$(dirname "$0")"

echo "=================================================="
echo "    雷允上语料库劳务费用明细表一键生成工具 (Mac版)"
echo "=================================================="
echo ""

# 检查 Python 3
if ! command -v python3 &> /dev/null; then
    echo "【错误】未检测到 python3，请先安装 Python 3 (可访问 https://www.python.org 下载安装)"
    read -p "按回车键退出..."
    exit 1
fi

# 检查并安装依赖
echo "正在检查运行环境 (pandas, openpyxl)..."
python3 -c "import pandas, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "正在安装必要依赖包: pandas, openpyxl..."
    python3 -m pip install pandas openpyxl
fi

echo ""
echo "开始自动识别文件并生成劳务费用明细表..."
python3 generate_settlement.py

echo ""
echo "=================================================="
read -p "处理完成！按回车键关闭窗口..."
