#!/bin/bash
# ====================================================
#  医疗数据统计与劳务结算平台 - Mac 一键启动脚本
# ====================================================

# 确保无论在何处双击，均切换至项目根目录
cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  🏥 医疗数据统计与劳务结算平台 (macOS 本地服务)"
echo "=================================================="
echo ""

# 1. 清理可能残留的旧 Streamlit 进程，确保最新拉取的代码能立即生效
if command -v lsof >/dev/null 2>&1; then
    PIDS=$(lsof -ti:8501 2>/dev/null)
    if [ -n "$PIDS" ]; then
        echo "🧹 正在清理残留的旧服务进程 (释放端口 8501)..."
        echo "$PIDS" | xargs kill -9 >/dev/null 2>&1
        sleep 0.5
    fi
fi
pkill -f "streamlit run app.py" >/dev/null 2>&1

# 2. 检查网络并自动同步 GitHub 远程最新代码
REPO_URL="https://github.com/icoati/jiesuan.git"
if command -v git >/dev/null 2>&1; then
    if curl -I -s --connect-timeout 2 https://github.com >/dev/null 2>&1; then
        echo "🔄 正在检查 GitHub 远程仓库更新..."
        if [ ! -d ".git" ]; then
            git init >/dev/null 2>&1
            git remote add origin "$REPO_URL" >/dev/null 2>&1
        fi
        git fetch origin main >/dev/null 2>&1
        if [ $? -eq 0 ]; then
            git reset --hard origin/main >/dev/null 2>&1
            chmod +x *.command 2>/dev/null
            chmod +x 调研案例电签表/*.command 2>/dev/null
            echo "✓ 已自动同步至 GitHub 最新代码版本！"
        fi
    fi
fi

# 3. 寻找可用的 Python 3 环境
PY_BIN=""

if [ -f ".venv/bin/python" ]; then
    PY_BIN=".venv/bin/python"
    echo "✓ 检测到本地专属虚拟环境: .venv"
fi

if [ -z "$PY_BIN" ]; then
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
fi

if [ -z "$PY_BIN" ]; then
    echo "❌ 未检测到 Python 3 运行环境！"
    echo "   👉 推荐从 Python 官网下载安装: https://www.python.org/downloads/macos/"
    echo "   👉 或打开终端执行命令安装: brew install python"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

echo "✓ 当前 Python 解释器: $($PY_BIN --version 2>&1) [$PY_BIN]"

# 4. 检查依赖完整性
echo "🔍 正在核验运行依赖..."
$PY_BIN -c "import streamlit, pandas, openpyxl, xlwt, xlrd, requests, bs4" 2>/dev/null
DEP_STATUS=$?

if [ $DEP_STATUS -ne 0 ]; then
    echo "⚠️ 首次启动或部分运行依赖尚未安装，正在为您一键配置环境（预计需要 1-2 分钟）..."
    
    if [ ! -d ".venv" ]; then
        echo "📦 正在创建本地独立运行沙盒 (.venv)..."
        $PY_BIN -m venv .venv 2>/dev/null
        if [ -f ".venv/bin/python" ]; then
            PY_BIN=".venv/bin/python"
            echo "✓ 独立环境沙盒创建成功！"
        fi
    fi

    echo "⬇️ 正在自动下载并安装依赖组件..."
    $PY_BIN -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn 2>/dev/null || \
    $PY_BIN -m pip install -r requirements.txt --break-system-packages 2>/dev/null || \
    $PY_BIN -m pip install -r requirements.txt
    
    if [ $? -ne 0 ]; then
        echo "❌ 依赖安装出现异常，请检查网络连接后重试。"
        read -p "按回车键退出..."
        exit 1
    fi
    echo "✅ 全部依赖组件已安装就绪！"
fi

echo ""
echo "=================================================="
echo "🚀 正在启动本地服务..."
echo "📍 浏览器访问地址: http://localhost:8501"
echo "💡 提示：将自动调起浏览器页面。退出时直接在此窗口按 Ctrl + C 即可停止服务。"
echo "=================================================="
echo ""

# 5. 启动 Streamlit
export LOCAL_MODE=1
$PY_BIN -m streamlit run app.py --server.port 8501 --server.headless false