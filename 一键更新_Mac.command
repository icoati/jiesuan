#!/bin/bash
# ====================================================
#  医疗数据统计平台 - Mac 一键拉取远程最新更新脚本
# ====================================================

cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  🔄 正在检查并拉取 GitHub 远程仓库最新更新..."
echo "=================================================="
echo ""

if ! command -v git &> /dev/null; then
    echo "❌ 未检测到 git 命令！"
    echo "👉 首次使用请在终端运行: xcode-select --install 进行安装"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

if [ ! -d ".git" ]; then
    echo "⚠️ 当前文件夹尚未与 GitHub 仓库绑定！"
    echo "👉 正在自动为您绑定远程仓库: https://github.com/icoati/jiesuan.git ..."
    git init
    git remote add origin https://github.com/icoati/jiesuan.git 2>/dev/null || git remote set-url origin https://github.com/icoati/jiesuan.git
    git fetch origin main
    git reset --hard origin/main
    git branch --set-upstream-to=origin/main main 2>/dev/null
    echo "✅ 绑定完成！已自动同步至最新版代码！"
else
    echo "⬇️ 正在执行 git pull..."
    git pull
    if [ $? -eq 0 ]; then
        echo ""
        echo "🎉 恭喜！平台代码已成功更新至最新版本！"
    else
        echo ""
        echo "⚠️ 拉取可能遇到冲突或网络波动，请检查网络连接。"
    fi
fi

# 确保脚本具备可执行权限
chmod +x *.command 2>/dev/null
chmod +x 调研案例电签表/*.command 2>/dev/null

echo "=================================================="
echo "👉 您现在可以直接双击【启动.command】运行最新功能！"
echo ""
read -p "按回车键关闭窗口..."
