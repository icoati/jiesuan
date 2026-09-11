# 医疗数据统计与劳务结算云平台

本平台将原本零散的本地一键统计脚本整合为统一的 **Streamlit Web 门户**，既支持在本地浏览器直接运行，也可一键托管部署至 **Streamlit Community Cloud**。

> **核心原则：**
> 底层所有计算逻辑、表头同义词词典、状态审计过滤规则与排版格式均由原脚本驱动，**逻辑 100% 保持不变**。

---

## 支持的业务功能

| 模块名称 | 核心驱动 | 输入文件要求 | 输出结果 |
| :--- | :--- | :--- | :--- |
| **模块 1：上药雷允上进度表** | Python (`统计报表.py`) | 2 个 Excel（答卷记录表、人员列表） | `统计总表.xlsx`（包含统计汇总、人员维度、案例维度） |
| **模块 2：语料库电签信息表** | Python (`settlement_tool.py`) | 3 个 Excel/XLS（支付清单、数据明细、用户明细） | `最终.xlsx`、`最终.xls`、对账筛选表、待核查告警名单、ZIP 打包 |
| **模块 3：北京整合-上药雷允上结算包** | Python (`settlement_mac.py`) | 1 个必选 Excel（明细）+ 1 个可选 Excel（确认单） | `劳务明细总表_*.xlsx`、`任务明细表_*.xlsx`、ZIP 打包 |
| **模块 4：陈菊梅基金会-雷允上结算包** | Python (`generate_settlement.py`) | 3 个 Excel（项目进度表、数据明细表、用户明细表） | `YYYYMMDD-劳务费用明细表.xlsx`（含汇总明细、分项目明细、对账总表） |
| **模块 5：老Saas医院导入模板** | Python (`hospital_grade_tool.py`) | 1 个医院管理 Excel 导入表 | 洗稿剔除知识库已有机构、剔除名单导出、`已补充等级_*.xlsx`、300条/批切分分批包、ZIP 打包、保留下拉验证约束 |

---

## 🚀 方式一：本地直接运行

### 1. 安装依赖
全套系统现已实现 **100% 纯 Python 原生架构**，无需安装 Node.js 或 npm！
确保本机已安装 Python 3.8+，在项目根目录下打开终端，运行：

```bash
# 一键安装全部依赖（pandas、openpyxl、streamlit、requests、beautifulsoup4 等）
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
   git commit -m "feat: 整合 Streamlit 云端统计门户与老Saas医院导入模块"
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

> **云端依赖与超大数据库说明：**
> - 全量业务模块均已统一为 **100% 纯 Python 原生执行**，无需 Node.js。
> - 「老Saas医院导入模板」采用双引擎智能回退：云端或 Git 仓库仅需 6.7MB 的 `hospital_base_cache.json.gz` 即可覆盖 49.4 万全国医疗机构，严格控制在 GitHub 100MB 限制以内。本地若有 `hospital_master.db` 则自动走毫秒级 SQLite 索引。

---

## 📁 项目目录说明

```text
├── app.py                      # Web 统一门户主程序 (100% 纯 Python 全栈引擎)
├── requirements.txt            # Python 依赖清单 (streamlit, pandas, openpyxl, xlwt, xlrd, requests, bs4)
├── .gitignore                  # Git 忽略规则 (过滤本地超大 *.db 库)
├── 一键启动服务.bat            # Windows 本机与局域网一键启动脚本
├── 上药雷允上进度表/           # 模块 1 核心驱动
├── 语料库电签信息表/           # 模块 2 核心驱动
├── 北京整合-上药雷允上结算包/  # 模块 3 核心驱动
├── 陈菊梅基金会-雷允上结算包/  # 模块 4 核心驱动
└── 老Saas医院导入模板/         # 模块 5 核心驱动 (49.4万超级库与300条分批切分)
```
