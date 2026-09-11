# 医疗数据统计与劳务结算云平台

本平台将原本零散的本地一键统计脚本整合为统一的 **Streamlit Web 门户**，既支持在本地浏览器直接运行，也可一键托管部署至 **Streamlit Community Cloud**。

> **核心原则：**
> 底层所有计算逻辑、表头同义词词典、状态审计过滤规则与排版格式均由原脚本驱动，**逻辑 100% 保持不变**。

---

## 支持的业务功能 (访问密码: `910104`)

| 模块名称 | 核心驱动 | 输入文件要求 | 输出结果 |
| :--- | :--- | :--- | :--- |
| **1. 上药雷允上进度表** | Python (`统计报表.py`) | 2 个 Excel：分别包含「答卷记录」与「项目人员」工作表 | `统计总表.xlsx`（包含统计汇总、人员维度、案例维度） |
| **2. 医生劳务费一键结算（语料库）** | Python (`settlement_mac.py`) | 1 个必选 Excel（明细）+ 1 个可选 Excel（确认单） | `劳务明细总表_*.xlsx`、`任务明细表_*.xlsx`、ZIP 打包 |
| **3. 语料库电签一键结算（语料库）** | Python (`settlement_tool.py`) | 3 个 Excel/XLS（支付清单、语料明文、用户明文） | `最终.xlsx`、`最终.xls`、对账筛选表、待核查告警名单、ZIP 打包 |
| **4. 陈菊梅基金会-雷允上结算包** | Python (`generate_settlement.py`) | 3 个 Excel（项目进度表、语料明文表、用户明文表） | `YYYYMMDD-劳务费用明细表.xlsx`（含汇总明细、分项目明细、语料对账总表） |

---

## 🚀 方式一：本地直接运行

### 1. 安装依赖
全套系统现已实现 **100% 纯 Python 原生架构**，无需安装 Node.js 或 npm！
确保本机已安装 Python 3.8+，在项目根目录下打开终端，运行：

```bash
# 一键安装全部依赖（pandas、openpyxl、streamlit 等）
pip install -r requirements.txt
```

### 2. 启动服务
```bash
streamlit run app.py
```
启动后，浏览器会自动打开 `http://localhost:8501`。

---

## ☁️ 方式二：部署到 Streamlit 云端（推荐，全员免配置使用）

只需简单 3 步，即可生成永久在线的网页链接，团队其他人无需安装任何 Python 或 Node.js 即可使用！

### 第一步：推送到 GitHub 仓库
1. 打开 [GitHub.com](https://github.com/) 并新建一个仓库（例如命名为 `medical-settlement-system`）。
   - **注意**：因涉及结算或医生隐私数据格式，建议将仓库设为 **Private（私有）**。
2. 在本地项目根目录初始化并推送代码：
   ```bash
   git init
   git add .
   git commit -m "feat: 整合 Streamlit 云端统计门户"
   git branch -M main
   git remote add origin https://github.com/<你的用户名>/medical-settlement-system.git
   git push -u origin main
   ```

### 第二步：登录 Streamlit Community Cloud
1. 浏览器访问：[share.streamlit.io](https://share.streamlit.io/)
2. 使用你的 **GitHub 账号**授权登录。

### 第三步：一键创建应用 (Deploy an app)
1. 点击右上角 **New app**。
2. 配置项填写：
   - **Repository**：选择你刚刚创建的 GitHub 仓库
   - **Branch**：`main`
   - **Main file path**：`app.py`
   - **App URL**（可选）：自定义二级域名，例如 `my-settlement-portal.streamlit.app`
3. 点击 **Deploy!**。

> **云端依赖说明：**
> 全量业务模块均已统一为 **100% 纯 Python 原生执行**，只需 `requirements.txt` 安装常用数据分析库即可运行，彻底免除了 Node.js 运行时配置。

---

## 📁 项目目录结构说明

```text
脚本集成/
├── app.py                      # Streamlit 统一门户主程序 (100% 纯 Python 全栈引擎)
├── requirements.txt            # Python 依赖清单 (streamlit, pandas, openpyxl, xlwt, xlrd)
├── .gitignore                  # Git 忽略文件（忽略缓存与临时表）
├── 一键启动服务.bat            # Windows 本机与局域网一键双击启动脚本
├── 上药报表统计/
│   ├── 统计报表.py              # [Python 原生版] 心血管内科进度表统计核心脚本
│   ├── 统计报表.js              # [原版 Node.js 脚本保留备份]
│   └── 使用说明.txt
├── 整合学会统计/
│   ├── settlement_mac.py       # [Python 原生版] 劳务费结算核心脚本
│   └── README_MAC.md
├── 语料库电签统计/
│   ├── settlement_tool.py      # [Python 原生版] 语料库月度结算核心脚本
│   └── README_Mac使用说明.md
└── 陈菊梅基金会-雷允上结算包/
    ├── generate_settlement.py  # [Python 原生版] 基金会劳务费用明细导出核心脚本
    ├── README_使用说明.md
    ├── run_windows.bat
    └── run_mac.command
```
