#!/bin/bash
cd "$(dirname "$0")"

echo "=================================================="
echo "  北检&华东科普点评电签表一键生成工具 (Mac版)"
echo "=================================================="
echo ""

if ! command -v python3 &> /dev/null; then
    echo "【错误】未检测到 python3，请先安装 Python 3"
    read -p "按回车键退出..."
    exit 1
fi

python3 generate_dianping_sign_table.py

echo ""
echo "=================================================="
read -p "处理完成！按回车键关闭窗口..."
