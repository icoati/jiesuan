#!/bin/bash
# ==============================================================================
#  医疗数据统计与劳务结算云平台 - 华为云/VPS 一键全自动部署脚本 (Ubuntu 24.04/22.04)
# ==============================================================================

set -e

echo "=================================================================="
echo "  🚀 正在开始一键自动化部署【医疗数据统计与劳务结算平台】..."
echo "=================================================================="

# 1. 检查 root 权限
if [ "$EUID" -ne 0 ]; then
  echo "❌ 请使用 root 权限运行此脚本！(可先执行: sudo -i)"
  exit 1
fi

# 2. 配置时区为中国北京时间 (UTC+8)
echo "[1/7] 配置系统时区为 Asia/Shanghai..."
timedatectl set-timezone Asia/Shanghai 2>/dev/null || ln -sf /usr/share/zoneinfo/Asia/Shanghai /etc/localtime

# 3. 检查并创建 2GB 虚拟内存 (Swap)，保护 2GB 内存服务器防止 OOM 内存溢出
if [ $(free -m | awk '/^Swap:/ {print $2}') -eq 0 ]; then
    echo "[2/7] 检测到系统无 Swap，正在自动创建 2GB 虚拟内存以保护系统稳定..."
    fallocate -l 2G /swapfile || dd if=/dev/zero of=/swapfile bs=1M count=2048
    chmod 600 /swapfile
    mkswap /swapfile
    swapon /swapfile
    echo '/swapfile none swap sw 0 0' >> /etc/fstab
    echo "✓ 2GB Swap 虚拟内存创建并激活成功！"
else
    echo "[2/7] 系统已有 Swap 虚拟内存，跳过创建。"
fi

# 4. 更新系统软件包源并安装 Python3、pip、git、curl
echo "[3/7] 安装系统基础组件 (Python3, pip, venv, git, curl)..."
apt-get update -y
apt-get install -y python3 python3-pip python3-venv git curl

# 5. 克隆或更新 GitHub 项目代码
REPO_URL="https://github.com/icoati/jiesuan.git"
INSTALL_DIR="/root/jiesuan"

echo "[4/7] 准备项目代码仓库 ($INSTALL_DIR)..."
if [ -d "$INSTALL_DIR/.git" ]; then
    echo "检测到已有项目目录，正在拉取最新代码..."
    cd "$INSTALL_DIR"
    git fetch origin main
    git reset --hard origin/main
else
    echo "正在从 GitHub 克隆项目代码..."
    rm -rf "$INSTALL_DIR" 2>/dev/null
    git clone "$REPO_URL" "$INSTALL_DIR"
    cd "$INSTALL_DIR"
fi

# 6. 配置 Python 独立虚拟环境并安装核心依赖
echo "[5/7] 配置 Python 虚拟沙盒环境 (.venv)..."
if [ ! -d ".venv" ]; then
    python3 -m venv .venv
fi

echo "正在使用国内清华源加速安装 Python 依赖包..."
.venv/bin/pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
.venv/bin/pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn

# 7. 配置 systemd 系统后台守护服务
echo "[6/7] 配置 systemd 后台自动守护进程 (服务名: jiesuan)..."
cat << 'EOF' > /etc/systemd/system/jiesuan.service
[Unit]
Description=Medical Settlement & Statistics Platform (Streamlit)
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=/root/jiesuan
Environment="PATH=/root/jiesuan/.venv/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin"
Environment="PYTHONUNBUFFERED=1"
Environment="TZ=Asia/Shanghai"
ExecStart=/root/jiesuan/.venv/bin/streamlit run app.py \
    --server.port 8501 \
    --server.address 0.0.0.0 \
    --server.headless true \
    --server.maxUploadSize 500 \
    --browser.gatherUsageStats false
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
EOF

# 生成一键更新脚本
cat << 'EOF' > /root/jiesuan/update.sh
#!/bin/bash
echo "🔄 正在从 GitHub 同步最新版本..."
cd /root/jiesuan || exit 1
git fetch origin main
git reset --hard origin/main
echo "正在重启结算平台服务..."
systemctl restart jiesuan
echo "✅ 更新完毕！当前运行状态："
systemctl status jiesuan --no-pager
EOF
chmod +x /root/jiesuan/update.sh
ln -sf /root/jiesuan/update.sh /usr/local/bin/jiesuan-update 2>/dev/null || true

# 重新加载并启动服务
systemctl daemon-reload
systemctl enable jiesuan
systemctl restart jiesuan

# 8. 启动完成，检查健康状态
echo "[7/7] 正在核验服务启动状态..."
sleep 3

if systemctl is-active --quiet jiesuan; then
    SERVER_IP=$(curl -s --connect-timeout 3 ifconfig.me || echo "1.92.81.242")
    echo ""
    echo "=================================================================="
    echo "🎉 恭喜！【医疗数据统计与劳务结算平台】已成功部署并启动！"
    echo "=================================================================="
    echo "🌐 平台外网访问地址: http://${SERVER_IP}:8501"
    echo ""
    echo "💡 华为云特别提示："
    echo "   请务必在【华为云控制台】->【安全组】入方向规则中开放端口 8501！"
    echo "   (放通后即可在任何电脑/手机浏览器直接打开上述地址使用)"
    echo ""
    echo "🔄 后续 GitHub 推送代码后，在服务器执行以下任一命令即可秒级热更新："
    echo "   👉 jiesuan-update"
    echo "   👉 或: /root/jiesuan/update.sh"
    echo "=================================================================="
else
    echo "❌ 服务启动出现异常，请查看日志:"
    journalctl -u jiesuan -n 20 --no-pager
fi
