#!/bin/bash
# ====================================================
#  Mac 权限一键修复脚本 (解决双击提示无权限或权限被拒)
# ====================================================

cd "$(dirname "$0")" || exit 1

echo "=================================================="
echo "  🔧 正在为您一键修复 Mac 脚本与应用执行权限..."
echo "=================================================="

# 移除 macOS 下载隔离属性 (Quarantine)
xattr -d com.apple.quarantine *.command 2>/dev/null
xattr -d -r com.apple.quarantine . 2>/dev/null
xattr -cr . 2>/dev/null

# 修复桌面 App 权限（若已生成）
if [ -d "$HOME/Desktop/医疗数据统计平台.app" ]; then
    xattr -cr "$HOME/Desktop/医疗数据统计平台.app" 2>/dev/null
    chmod -R 755 "$HOME/Desktop/医疗数据统计平台.app" 2>/dev/null
    chmod +x "$HOME/Desktop/医疗数据统计平台.app/Contents/MacOS/launcher" 2>/dev/null
fi

if [ -d "医疗数据统计平台.app" ]; then
    xattr -cr "医疗数据统计平台.app" 2>/dev/null
    chmod -R 755 "医疗数据统计平台.app" 2>/dev/null
    chmod +x "医疗数据统计平台.app/Contents/MacOS/launcher" 2>/dev/null
fi

# 递归为所有 command 脚本赋予可执行权限
chmod +x *.command 2>/dev/null
find . -name "*.command" -exec chmod +x {} + 2>/dev/null
find . -name "*.sh" -exec chmod +x {} + 2>/dev/null

echo "✅ 权限修复完成！"
echo "👉 现在您可以直接双击【医疗数据统计平台.app】或【启动.command】秒开应用了！"
echo "=================================================="
echo ""
read -p "按回车键退出..."
