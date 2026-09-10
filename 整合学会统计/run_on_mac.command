#!/bin/bash
# ============================================================
# 医生劳务费一键结算系统 - macOS 双击启动器
# ============================================================

# 自动切换到当前脚本所在目录
cd "$(dirname "$0")"

echo "============================================================"
echo "          医生劳务费一键结算系统 (macOS 专用)               "
echo "============================================================"
echo "当前目录: $(pwd)"
echo ""

# 1. 检查 Python 3 环境
if ! command -v python3 &> /dev/null; then
    echo "[!] 错误: 未检测到 python3！"
    echo "    请前往 https://www.python.org/downloads/ 安装 Python 3，"
    echo "    或者在终端执行: xcode-select --install"
    echo ""
    read -n 1 -s -r -p "按任意键退出..."
    exit 1
fi

# 2. 检查并自动安装 openpyxl 依赖
if ! python3 -c "import openpyxl" &> /dev/null; then
    echo "[*] 首次运行检测: 正在自动为您安装 openpyxl 依赖库..."
    pip3 install openpyxl
    if [ $? -ne 0 ]; then
        echo "[!] 安装 openpyxl 失败，请检查网络或在终端手动执行: pip3 install openpyxl"
        echo ""
        read -n 1 -s -r -p "按任意键退出..."
        exit 1
    fi
    echo "[√] openpyxl 安装成功！"
    echo ""
fi

# 3. 运行结算主程序
echo "[*] 正在启动结算程序..."
python3 settlement_mac.py

echo ""
echo "============================================================"
read -n 1 -s -r -p "程序执行完毕，按任意键退出终端窗口..."
