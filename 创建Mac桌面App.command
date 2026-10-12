#!/bin/bash
# ====================================================
#  自动生成 Mac 独立桌面客户端 (医疗数据统计平台.app)
# ====================================================

# 切换到项目根目录
DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$HOME/Desktop/医疗数据统计平台.app"

echo "=================================================="
echo "  🖥️ 正在为 Mac 自动构建独立桌面 App..."
echo "=================================================="

# 创建 App Bundle 目录结构
mkdir -p "$APP_DIR/Contents/MacOS"
mkdir -p "$APP_DIR/Contents/Resources"

# 1. 写入 Info.plist
cat <<EOF > "$APP_DIR/Contents/Info.plist"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIdentifier</key>
    <string>com.medical.statplatform</string>
    <key>CFBundleName</key>
    <string>医疗数据统计平台</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>LSUIElement</key>
    <false/>
</dict>
</plist>
EOF

# 2. 写入启动执行器 launcher
cat <<EOF > "$APP_DIR/Contents/MacOS/launcher"
#!/bin/bash
PROJECT_DIR="$DIR"
cd "\$PROJECT_DIR" || exit 1

# 寻找 Python 3
PY_BIN=""
if [ -f ".venv/bin/python" ]; then
    PY_BIN=".venv/bin/python"
fi

if [ -z "\$PY_BIN" ]; then
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
    for cmd in "\${CANDIDATES[@]}"; do
        if command -v "\$cmd" &> /dev/null; then
            PY_VER=\$("\$cmd" -c 'import sys; print(sys.version_info[0])' 2>/dev/null)
            if [ "\$PY_VER" = "3" ]; then
                PY_BIN="\$cmd"
                break
            fi
        fi
    done
fi

if [ -z "\$PY_BIN" ]; then
    osascript -e 'display dialog "未在您的 Mac 上检测到 Python 3 环境，请先安装 Python！\n推荐前往 python.org 下载安装。" buttons {"确定"} default button 1 with icon stop'
    exit 1
fi

# 检查服务是否已在运行
if ! curl -s http://127.0.0.1:8501 > /dev/null 2>&1; then
    export LOCAL_MODE=1
    nohup \$PY_BIN -m streamlit run app.py --server.port 8501 --server.headless true > /dev/null 2>&1 &
    # 等待服务端口就绪
    for i in {1..20}; do
        if curl -s http://127.0.0.1:8501 > /dev/null 2>&1; then
            break
        fi
        sleep 0.5
    done
fi

# 优先以独立客户端无边框窗口模式打开（如果安装了 Chrome 或 Edge）
if [ -d "/Applications/Google Chrome.app" ]; then
    open -na "Google Chrome" --args --app="http://127.0.0.1:8501"
elif [ -d "/Applications/Microsoft Edge.app" ]; then
    open -na "Microsoft Edge" --args --app="http://127.0.0.1:8501"
else
    open "http://127.0.0.1:8501"
fi
EOF

# 授予可执行权限
chmod +x "$APP_DIR/Contents/MacOS/launcher"
chmod +x "$DIR/一键启动服务_Mac.command"

echo ""
echo "=================================================="
echo "✅ 成功！已在您的【桌面】生成独立应用："
echo "   👉 ~/Desktop/医疗数据统计平台.app"
echo "=================================================="
echo "💡 您可以直接双击桌面上的图标秒开应用，或将它拖到程序坞 (Dock) 栏中！"
echo ""
read -p "按回车键退出..."