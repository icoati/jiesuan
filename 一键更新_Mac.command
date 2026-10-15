#!/bin/bash
# ====================================================
#  医疗数据统计平台 - Mac 一键拉取远程最新更新脚本
# ====================================================

cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  🔄 正在检查并同步 GitHub 远程仓库最新更新..."
echo "=================================================="
echo ""

# 1. 清理旧进程
if command -v lsof >/dev/null 2>&1; then
    PIDS=$(lsof -ti:8501 2>/dev/null)
    if [ -n "$PIDS" ]; then
        echo "🧹 正在清理可能驻留的旧服务进程 (端口 8501)..."
        echo "$PIDS" | xargs kill -9 >/dev/null 2>&1
    fi
fi
pkill -f "streamlit run app.py" >/dev/null 2>&1

if ! command -v git &> /dev/null; then
    echo "❌ 未检测到 git 命令！"
    echo "👉 首次使用请在终端运行: xcode-select --install 进行安装"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

REPO_URL="https://github.com/icoati/jiesuan.git"

if [ ! -d ".git" ]; then
    echo "⚠️ 当前文件夹尚未与 GitHub 仓库绑定！"
    echo "👉 正在自动为您绑定远程仓库: $REPO_URL ..."
    git init
    git remote add origin "$REPO_URL" 2>/dev/null || git remote set-url origin "$REPO_URL"
fi

echo "⬇️ 正在从远程 main 分支同步最新代码..."
git fetch origin main
if [ $? -eq 0 ]; then
    git reset --hard origin/main
    git branch --set-upstream-to=origin/main main 2>/dev/null
    
    # 确保脚本具备可执行权限
    chmod +x *.command 2>/dev/null
    chmod +x 调研案例电签表/*.command 2>/dev/null
    
    echo ""
    echo "🎉 恭喜！平台代码已成功同步至最新版本！"
    echo "   (所有新功能包括【调研案例电签表】均已就绪)"
else
    echo ""
    echo "⚠️ 拉取可能遇到网络波动，请检查网络连接后重试。"
fi

echo "=================================================="
echo "👉 您现在可以直接双击【启动.command】或【医疗数据统计平台.app】运行最新功能！"
echo ""
read -p "按回车键关闭窗口..."
