#!/bin/bash
# ====================================================
#  调研案例电签表 - Mac 一键启动脚本
# ====================================================

cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  🏥 调研案例电签表 - 一键生成工具 (macOS)"
echo "=================================================="
echo ""

# 寻找可用的 Python 3
PY_BIN=""
CANDIDATES=(
    "/opt/homebrew/bin/python3"
    "/usr/local/bin/python3"
    "python3.13"
    "python3.12"
    "python3.11"
    "python3.10"
    "python3.9"
    "python3"
    "python"
)
for cmd in "${CANDIDATES[@]}"; do
    if command -v "$cmd" &> /dev/null; then
        PY_VER=$("$cmd" -c 'import sys; print(sys.version_info[0])' 2>/dev/null)
        if [ "$PY_VER" = "3" ]; then
            PY_BIN="$cmd"
            break
        fi
    fi
done

if [ -z "$PY_BIN" ]; then
    echo "❌ 未检测到 Python 3 环境，请先安装 Python！"
    read -p "按回车键退出..."
    exit 1
fi

# 检查依赖
$PY_BIN -c "import pandas, openpyxl, xlwt" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️ 正在自动安装运行依赖 (pandas, openpyxl, xlwt)..."
    $PY_BIN -m pip install pandas openpyxl xlwt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn 2>/dev/null || \
    $PY_BIN -m pip install pandas openpyxl xlwt --break-system-packages 2>/dev/null || \
    $PY_BIN -m pip install pandas openpyxl xlwt
fi

# 运行脚本
$PY_BIN settlement_survey_case.py "$(pwd)"

echo ""
read -p "🎉 全部处理完成！按回车键退出..."
