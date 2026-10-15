#!/bin/bash
# ====================================================
#  自动生成 Mac 独立桌面客户端 (医疗数据统计平台.app)
# ====================================================

DIR="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$HOME/Desktop/医疗数据统计平台.app"

echo "=================================================="
echo "  🖥️ 正在为 Mac 自动构建独立桌面 App..."
echo "=================================================="

# 如果本地已有完整打包的 .app，直接同步至桌面
if [ -d "$DIR/医疗数据统计平台.app" ]; then
    rm -rf "$APP_DIR" 2>/dev/null
    cp -R "$DIR/医疗数据统计平台.app" "$HOME/Desktop/"
else
    # 创建 App Bundle 目录结构
    mkdir -p "$APP_DIR/Contents/MacOS"
    mkdir -p "$APP_DIR/Contents/Resources"

    # 复制应用图标
    if [ -f "$DIR/AppIcon.icns" ]; then
        cp "$DIR/AppIcon.icns" "$APP_DIR/Contents/Resources/AppIcon.icns"
    fi

    # 1. 写入 Info.plist
    cat <<EOF > "$APP_DIR/Contents/Info.plist"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.medical.statplatform</string>
    <key>CFBundleName</key>
    <string>医疗数据统计平台</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>2.0</string>
    <key>LSUIElement</key>
    <false/>
</dict>
</plist>
EOF

    # 2. 写入启动执行器 launcher
    cat <<'EOF' > "$APP_DIR/Contents/MacOS/launcher"
#!/bin/bash
# ====================================================
#  医疗数据统计平台 - Mac 独立原生客户端启动器
# ====================================================
export PATH="/opt/homebrew/bin:/opt/homebrew/sbin:/usr/local/bin:/usr/local/sbin:$HOME/.local/bin:$PATH"

# 1. 强制清理残留旧进程，释放 8501 端口
if command -v lsof >/dev/null 2>&1; then
    PIDS=$(lsof -ti:8501 2>/dev/null)
    if [ -n "$PIDS" ]; then
        echo "$PIDS" | xargs kill -9 >/dev/null 2>&1
    fi
fi
pkill -f "streamlit run app.py" >/dev/null 2>&1
sleep 0.3

# 2. 智能探测项目根目录
REPO_URL="https://github.com/icoati/jiesuan.git"
WORK_DIR=""

DIR_UP="$(cd "$(dirname "$0")/../../.." && pwd)"
if [ -f "$DIR_UP/app.py" ]; then
    WORK_DIR="$DIR_UP"
elif [ -f "$HOME/.medical_stat_platform_path" ] && [ -f "$(cat "$HOME/.medical_stat_platform_path")/app.py" ]; then
    WORK_DIR="$(cat "$HOME/.medical_stat_platform_path")"
elif [ -f "$HOME/Desktop/脚本集成/app.py" ]; then
    WORK_DIR="$HOME/Desktop/脚本集成"
elif [ -f "$HOME/Desktop/jiesuan/app.py" ]; then
    WORK_DIR="$HOME/Desktop/jiesuan"
elif [ -f "$HOME/Desktop/医疗数据统计平台/app.py" ]; then
    WORK_DIR="$HOME/Desktop/医疗数据统计平台"
elif [ -f "$HOME/Documents/医疗数据统计平台/app.py" ]; then
    WORK_DIR="$HOME/Documents/医疗数据统计平台"
else
    WORK_DIR="$HOME/Documents/医疗数据统计平台"
    mkdir -p "$WORK_DIR"
    osascript -e 'display notification "首次启动，正在下载平台代码..." with title "医疗数据统计平台"' 2>/dev/null
    git clone "$REPO_URL" "$WORK_DIR" >/dev/null 2>&1
fi

echo "$WORK_DIR" > "$HOME/.medical_stat_platform_path"
cd "$WORK_DIR" || exit 1

# 3. 靠 GitHub 自动同步最新代码
if command -v git >/dev/null 2>&1; then
    if curl -I -s --connect-timeout 2 https://github.com >/dev/null 2>&1; then
        if [ ! -d ".git" ]; then
            git init >/dev/null 2>&1
            git remote add origin "$REPO_URL" >/dev/null 2>&1
        fi
        git fetch origin main >/dev/null 2>&1
        if [ $? -eq 0 ]; then
            git reset --hard origin/main >/dev/null 2>&1
            chmod +x *.command 2>/dev/null
        fi
    fi
fi

# 4. 寻找 Python 3 环境
PY_BIN=""
if [ -f ".venv/bin/python" ]; then
    PY_BIN=".venv/bin/python"
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
    osascript -e 'display dialog "未在您的 Mac 上检测到 Python 3 环境，请先安装 Python！\n推荐前往官网 python.org 下载安装。" buttons {"确定"} default button 1 with icon stop'
    exit 1
fi

# 5. 校验运行依赖
$PY_BIN -c "import streamlit, pandas, openpyxl, xlwt, xlrd, requests, bs4" 2>/dev/null
if [ $? -ne 0 ]; then
    osascript -e 'display notification "正在为您准备依赖组件，即将完成..." with title "医疗数据统计平台"' 2>/dev/null
    if [ ! -d ".venv" ]; then
        $PY_BIN -m venv .venv 2>/dev/null
        if [ -f ".venv/bin/python" ]; then
            PY_BIN=".venv/bin/python"
        fi
    fi
    $PY_BIN -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn 2>/dev/null || \
    $PY_BIN -m pip install -r requirements.txt --break-system-packages 2>/dev/null || \
    $PY_BIN -m pip install -r requirements.txt 2>/dev/null
fi

# 6. 后台启动 Streamlit 服务
export LOCAL_MODE=1
nohup "$PY_BIN" -m streamlit run app.py --server.port 8501 --server.headless true > /dev/null 2>&1 &

# 7. 轮询等待端口就绪 (最多等待 15 秒)
READY=0
for i in {1..30}; do
    if curl -s http://127.0.0.1:8501 > /dev/null 2>&1; then
        READY=1
        break
    fi
    sleep 0.5
done

if [ $READY -eq 1 ]; then
    if [ -d "/Applications/Google Chrome.app" ]; then
        open -na "Google Chrome" --args --app="http://127.0.0.1:8501"
    elif [ -d "/Applications/Microsoft Edge.app" ]; then
        open -na "Microsoft Edge" --args --app="http://127.0.0.1:8501"
    else
        open "http://127.0.0.1:8501"
    fi
else
    osascript -e 'display alert "启动未就绪" message "本地服务启动超时，请在项目目录下双击 一键启动服务_Mac.command 查看详细信息。"' 2>/dev/null
fi
EOF
fi

# 授予可执行权限并移除隔离属性
chmod +x "$APP_DIR/Contents/MacOS/launcher" 2>/dev/null
xattr -cr "$APP_DIR" 2>/dev/null
chmod +x "$DIR"/*.command 2>/dev/null

echo ""
echo "=================================================="
echo "✅ 成功！已在您的【桌面】生成独立应用："
echo "   👉 ~/Desktop/医疗数据统计平台.app"
echo "=================================================="
echo "💡 特性亮点："
echo "   1. 每次打开自动从 GitHub 静默拉取最新版代码并生效"
echo "   2. 自动清理旧残留进程，杜绝新功能未生效问题"
echo "   3. 您可以直接双击桌面图标秒开，或将它拖入程序坞 (Dock)！"
echo ""
read -p "按回车键退出..."