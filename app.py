# -*- coding: utf-8 -*-
"""
======================================================================
医疗数据统计与劳务结算云平台 (Enterprise Precision B2B Edition)
升级特性：
1. 7天免密持久化：借助浏览器 LocalStorage 与防篡改 Token，实现真正的关闭浏览器 7 天内免输入密码
2. 模块命名精准更新：「心血管内科报表统计（上药）」正式更名为「上药雷允上进度表」（侧边栏与顶部横幅完全同步）
3. 彻底根除 HTML 代码泄漏问题：采用 render_html 安全渲染，杜绝 Markdown 代码块误判
4. 修复密码输入框结构与眼睛图标：透明居中、右侧无畸变色块
5. 专属高端矢量 SVG 徽标系统，全面弃用 Emoji 表情
======================================================================
"""

import os
import sys
import io
import re
import time
import datetime
import hashlib
import shutil
import zipfile
import tempfile
import subprocess
from pathlib import Path
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import history_manager

# 项目根目录
ROOT_DIR = os.path.abspath(os.path.dirname(__file__))

# 页面基础配置
st.set_page_config(
    page_title="医疗数据统计与劳务结算云平台",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# 核心 HTML 渲染助手（彻底清除缩进，杜绝 Markdown 将 HTML 误认为代码块）
# -------------------------------------------------------------
def render_html(html_code: str, container=st):
    """清除每行首尾空白，防止 Markdown 语法将 4 空格缩进当做代码块渲染"""
    cleaned = "\n".join(line.strip() for line in html_code.strip().splitlines())
    container.markdown(cleaned, unsafe_allow_html=True)


# -------------------------------------------------------------
# 注入高端企业级 CSS（精确定位密码框、矢量图标体系、单排导航）
# -------------------------------------------------------------
render_html("""
<style>
    /* 引入苹果与现代化系统无衬线字体 */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Segoe UI", Roboto, "Helvetica Neue", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
        -webkit-font-smoothing: antialiased;
    }

    /* 彻底隐藏页面 Header 与侧边栏 Header，根除遮挡与剃头问题 */
    header[data-testid="stHeader"],
    div[data-testid="stSidebarHeader"] {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        visibility: hidden !important;
        pointer-events: none !important;
        padding: 0 !important;
        margin: 0 !important;
    }

    /* 侧边栏内容区顶部留白精细修正 */
    div[data-testid="stSidebarContent"],
    div[data-testid="stSidebarUserContent"] {
        padding-top: 1.6rem !important;
    }

    /* 主工作区外边距与内边距精细化排版（无 Header 遮挡，视界完整通透） */
    .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 3.5rem !important;
        padding-left: 2.5rem !important;
        padding-right: 2.5rem !important;
        max-width: 1240px !important;
    }

    /* ================= 彻底锁定侧边栏常驻，完全禁用收缩展开功能 ================= */
    [data-testid="stSidebarCollapseButton"],
    [data-testid="collapsedControl"],
    button[data-testid="stSidebarCollapseButton"],
    header [data-testid="stSidebarCollapseButton"],
    header [data-testid="collapsedControl"],
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"],
    div[data-testid="stSidebarNav"] button,
    button[aria-label="Close sidebar"],
    button[aria-label="Open sidebar"],
    [data-testid="stSidebarUserContent"] button[aria-label="Close sidebar"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        width: 0 !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    section[data-testid="stSidebar"] {
        min-width: 300px !important;
        max-width: 300px !important;
        width: 300px !important;
        transform: none !important;
        transition: none !important;
        border-right: 1px solid rgba(226, 232, 240, 0.8) !important;
    }
    @media (prefers-color-scheme: dark) {
        section[data-testid="stSidebar"] {
            border-right: 1px solid rgba(255, 255, 255, 0.08) !important;
        }
    }
    [data-theme="dark"] section[data-testid="stSidebar"] {
        border-right: 1px solid rgba(255, 255, 255, 0.08) !important;
    }

    section[data-testid="stSidebar"][aria-expanded="false"] {
        margin-left: 0 !important;
        transform: none !important;
    }
    
    .stSidebar [data-testid="stVerticalBlock"] button {
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 0.75rem 1.1rem !important;
        font-size: 0.94rem !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
        margin-bottom: 8px !important;
        letter-spacing: -0.2px !important;
    }

    /* ================= 专属登录门禁卡片 (高对比清晰呈现) ================= */
    div[data-testid="stForm"] {
        background: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 32px !important;
        padding: 46px 36px 38px 36px !important;
        box-shadow: 0 20px 50px -10px rgba(0, 0, 0, 0.12), 0 2px 6px rgba(0, 0, 0, 0.04) !important;
        max-width: 410px !important;
        width: 100% !important;
        margin: 50px auto 0 auto !important;
        text-align: center !important;
    }

    @media (prefers-color-scheme: dark) {
        div[data-testid="stForm"] {
            background: #1e222b !important;
            border: 1px solid rgba(255, 255, 255, 0.12) !important;
            box-shadow: 0 25px 60px -10px rgba(0, 0, 0, 0.6) !important;
        }
    }
    [data-theme="dark"] div[data-testid="stForm"] {
        background: #1e222b !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
    }

    /* 顶部品牌圆环徽标 */
    .tower-logo-badge {
        width: 80px;
        height: 80px;
        background: rgba(2, 132, 199, 0.08) !important;
        border: 1px solid rgba(2, 132, 199, 0.2) !important;
        border-radius: 24px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: 0 8px 24px rgba(2, 132, 199, 0.12) !important;
        margin-bottom: 22px !important;
    }
    .tower-logo-badge svg {
        width: 44px;
        height: 44px;
    }

    .tower-title {
        font-size: 1.55rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 28px;
        letter-spacing: -0.3px;
    }
    @media (prefers-color-scheme: dark) {
        .tower-title {
            color: #f8fafc !important;
        }
    }
    [data-theme="dark"] .tower-title {
        color: #f8fafc !important;
    }

    /* ================= 密码输入框与表单提示深度清理 ================= */
    /* 彻底隐藏 Streamlit 的表单回车提示英文 (Press Enter to submit form) */
    [data-testid="InputInstructions"],
    div[data-testid="InputInstructions"],
    div[data-baseweb="input"] [data-testid="InputInstructions"] {
        display: none !important;
        visibility: hidden !important;
        font-size: 0 !important;
        width: 0 !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
    }

    /* 1. 外层输入框容器：统一橙色高光边框 */
    div[data-testid="stForm"] [data-testid="stTextInput"] > div[data-baseweb="input"] {
        border: 2px solid #ea580c !important;
        border-radius: 12px !important;
        background-color: #ffffff !important;
        box-shadow: 0 0 0 4px rgba(234, 88, 12, 0.16) !important;
        height: 52px !important;
        display: flex !important;
        align-items: center !important;
        padding: 0 !important;
        overflow: hidden !important;
    }

    /* 2. 文本输入核心区：占满整行、纯净无杂质 */
    div[data-testid="stForm"] [data-testid="stTextInput"] input {
        background: transparent !important;
        border: none !important;
        outline: none !important;
        color: #0f172a !important;
        font-size: 1.05rem !important;
        font-weight: 500 !important;
        padding: 0 16px !important;
        height: 100% !important;
        width: 100% !important;
        flex: 1 1 100% !important;
    }
    div[data-testid="stForm"] [data-testid="stTextInput"] input::placeholder {
        color: #94a3b8 !important;
    }

    /* 3. 彻底隐藏密码框内溢出的英文按钮文字 (visibility) */
    div[data-testid="stForm"] [data-testid="stTextInput"] button {
        display: none !important;
        width: 0 !important;
        height: 0 !important;
        overflow: hidden !important;
    }

    /* 4. 登录提交主大按钮：仅定向作用于提交按钮，绝不污染眼睛按钮 */
    div[data-testid="stForm"] [data-testid="stFormSubmitButton"] > button {
        background: #ea580c !important;
        background-color: #ea580c !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        height: 50px !important;
        line-height: 50px !important;
        padding: 0 !important;
        width: 100% !important;
        box-shadow: 0 8px 22px rgba(234, 88, 12, 0.38) !important;
        margin-top: 10px !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
        cursor: pointer !important;
    }
    div[data-testid="stForm"] [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #c2410c !important;
        transform: translateY(-1.5px) !important;
        box-shadow: 0 12px 28px rgba(234, 88, 12, 0.5) !important;
    }
    div[data-testid="stForm"] [data-testid="stFormSubmitButton"] > button p {
        color: #ffffff !important;
        font-size: 1.1rem !important;
        font-weight: 700 !important;
    }

    /* 7天免密复选框 */
    div[data-testid="stForm"] [data-testid="stCheckbox"] {
        text-align: left !important;
        margin: 12px 0 16px 2px !important;
    }
    div[data-testid="stForm"] [data-testid="stCheckbox"] label span {
        color: #64748b !important;
        font-size: 0.92rem !important;
        font-weight: 500 !important;
    }
    @media (prefers-color-scheme: dark) {
        div[data-testid="stForm"] [data-testid="stCheckbox"] label span {
            color: #94a3b8 !important;
        }
    }

    /* ================= 现代高质感工作台顶栏 (Apple / B2B Precision) ================= */
    /* ================= 现代高质感工作台顶栏 (Apple / B2B Precision) ================= */
    .ios-hero-banner {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%) !important;
        border: 1px solid #e2e8f0 !important;
        border-top: 3px solid #0284c7 !important;
        border-radius: 18px !important;
        padding: 20px 26px !important;
        margin-bottom: 20px !important;
        color: #0f172a !important;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.05), 0 1px 3px rgba(0, 0, 0, 0.02) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        position: relative !important;
        overflow: hidden !important;
    }
    @media (prefers-color-scheme: dark) {
        .ios-hero-banner {
            background: linear-gradient(135deg, #161b22 0%, #1e2430 100%) !important;
            border-color: rgba(255, 255, 255, 0.1) !important;
            border-top: 3px solid #38bdf8 !important;
            color: #f8fafc !important;
            box-shadow: 0 10px 30px -6px rgba(0, 0, 0, 0.4) !important;
        }
    }
    [data-theme="dark"] .ios-hero-banner {
        background: linear-gradient(135deg, #161b22 0%, #1e2430 100%) !important;
        border-color: rgba(255, 255, 255, 0.1) !important;
        border-top: 3px solid #38bdf8 !important;
        color: #f8fafc !important;
    }

    .ios-hero-left {
        display: flex;
        align-items: center;
        gap: 16px;
        min-width: 0;
        flex: 1;
    }
    .ios-hero-icon-badge {
        width: 50px;
        height: 50px;
        border-radius: 14px;
        background: rgba(2, 132, 199, 0.08) !important;
        border: 1px solid rgba(2, 132, 199, 0.2) !important;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }
    .ios-hero-icon-badge svg path,
    .ios-hero-icon-badge svg rect,
    .ios-hero-icon-badge svg line,
    .ios-hero-icon-badge svg polyline {
        stroke: #0284c7 !important;
    }
    @media (prefers-color-scheme: dark) {
        .ios-hero-icon-badge {
            background: rgba(56, 189, 248, 0.15) !important;
            border-color: rgba(56, 189, 248, 0.3) !important;
        }
        .ios-hero-icon-badge svg path,
        .ios-hero-icon-badge svg rect,
        .ios-hero-icon-badge svg line,
        .ios-hero-icon-badge svg polyline {
            stroke: #38bdf8 !important;
        }
    }
    [data-theme="dark"] .ios-hero-icon-badge {
        background: rgba(56, 189, 248, 0.15) !important;
        border-color: rgba(56, 189, 248, 0.3) !important;
    }
    [data-theme="dark"] .ios-hero-icon-badge svg path,
    [data-theme="dark"] .ios-hero-icon-badge svg rect,
    [data-theme="dark"] .ios-hero-icon-badge svg line,
    [data-theme="dark"] .ios-hero-icon-badge svg polyline {
        stroke: #38bdf8 !important;
    }

    .ios-hero-title {
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        margin: 0 0 3px 0 !important;
        letter-spacing: -0.3px !important;
        color: #0f172a !important;
        white-space: nowrap !important;
    }
    @media (prefers-color-scheme: dark) {
        .ios-hero-title {
            color: #f8fafc !important;
        }
    }
    [data-theme="dark"] .ios-hero-title {
        color: #f8fafc !important;
    }

    .ios-hero-subtitle {
        font-size: 0.86rem !important;
        color: #64748b !important;
        margin: 0 !important;
        line-height: 1.4 !important;
    }
    @media (prefers-color-scheme: dark) {
        .ios-hero-subtitle {
            color: #94a3b8 !important;
        }
    }
    [data-theme="dark"] .ios-hero-subtitle {
        color: #94a3b8 !important;
    }

    .ios-hero-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(2, 132, 199, 0.08) !important;
        border: 1px solid rgba(2, 132, 199, 0.2) !important;
        border-radius: 9999px;
        padding: 5px 14px;
        font-size: 0.78rem;
        font-weight: 600;
        color: #0284c7 !important;
        letter-spacing: 0.2px;
        flex-shrink: 0;
        margin-left: 16px;
    }
    .ios-hero-pill-dot {
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: #0284c7 !important;
        box-shadow: 0 0 8px #0284c7 !important;
    }
    @media (prefers-color-scheme: dark) {
        .ios-hero-pill {
            background: rgba(56, 189, 248, 0.12) !important;
            border-color: rgba(56, 189, 248, 0.25) !important;
            color: #38bdf8 !important;
        }
        .ios-hero-pill-dot {
            background: #38bdf8 !important;
            box-shadow: 0 0 8px #38bdf8 !important;
        }
    }
    [data-theme="dark"] .ios-hero-pill {
        background: rgba(56, 189, 248, 0.12) !important;
        border-color: rgba(56, 189, 248, 0.25) !important;
        color: #38bdf8 !important;
    }
    [data-theme="dark"] .ios-hero-pill-dot {
        background: #38bdf8 !important;
        box-shadow: 0 0 8px #38bdf8 !important;
    }

    /* ================= Bento 文件规范卡片体系 ================= */
    .bento-req-container {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 18px 22px;
        margin-bottom: 20px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02);
    }
    @media (prefers-color-scheme: dark) {
        .bento-req-container {
            background: #1e2430 !important;
            border-color: rgba(255, 255, 255, 0.08) !important;
        }
    }
    [data-theme="dark"] .bento-req-container {
        background: #1e2430 !important;
        border-color: rgba(255, 255, 255, 0.08) !important;
    }

    .bento-req-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 14px;
        padding-bottom: 12px;
        border-bottom: 1px solid #e2e8f0;
    }
    @media (prefers-color-scheme: dark) {
        .bento-req-header {
            border-bottom-color: rgba(255, 255, 255, 0.08) !important;
        }
    }
    [data-theme="dark"] .bento-req-header {
        border-bottom-color: rgba(255, 255, 255, 0.08) !important;
    }

    .bento-req-title {
        font-size: 0.95rem;
        font-weight: 700;
        color: #0f172a;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    @media (prefers-color-scheme: dark) {
        .bento-req-title { color: #f8fafc !important; }
    }
    [data-theme="dark"] .bento-req-title { color: #f8fafc !important; }

    .bento-req-badge {
        font-size: 0.78rem;
        font-weight: 600;
        color: #0284c7;
        background: rgba(2, 132, 199, 0.08);
        border: 1px solid rgba(2, 132, 199, 0.2);
        padding: 3px 10px;
        border-radius: 8px;
    }

    .bento-req-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 12px;
    }

    .bento-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 12px 16px;
        display: flex;
        align-items: flex-start;
        gap: 12px;
    }
    @media (prefers-color-scheme: dark) {
        .bento-card {
            background: #161b22 !important;
            border-color: rgba(255, 255, 255, 0.06) !important;
        }
    }
    [data-theme="dark"] .bento-card {
        background: #161b22 !important;
        border-color: rgba(255, 255, 255, 0.06) !important;
    }

    .bento-card-num {
        font-size: 0.85rem;
        font-weight: 800;
        color: #0284c7;
        background: rgba(2, 132, 199, 0.1);
        width: 28px;
        height: 28px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }

    .bento-card-title {
        font-size: 0.92rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 2px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    @media (prefers-color-scheme: dark) {
        .bento-card-title { color: #f8fafc !important; }
    }
    [data-theme="dark"] .bento-card-title { color: #f8fafc !important; }

    .bento-card-desc {
        font-size: 0.82rem;
        color: #64748b;
        line-height: 1.4;
    }
    @media (prefers-color-scheme: dark) {
        .bento-card-desc { color: #94a3b8 !important; }
    }
    [data-theme="dark"] .bento-card-desc { color: #94a3b8 !important; }

    .code-pill {
        background: rgba(2, 132, 199, 0.08);
        color: #0284c7;
        padding: 2px 6px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.82rem;
    }
    .req-tag-must {
        font-size: 0.72rem;
        font-weight: 600;
        color: #ea580c;
        background: rgba(234, 88, 12, 0.1);
        padding: 1px 6px;
        border-radius: 4px;
    }
    .req-tag-opt {
        font-size: 0.72rem;
        font-weight: 600;
        color: #64748b;
        background: rgba(100, 116, 139, 0.1);
        padding: 1px 6px;
        border-radius: 4px;
    }

    /* ================= 文件上传区美化 ================= */
    div[data-testid="stFileUploader"] {
        background: #f8fafc !important;
        border: 2px dashed #cbd5e1 !important;
        border-radius: 16px !important;
        padding: 16px 20px !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02) !important;
        margin-bottom: 18px !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #0284c7 !important;
        background: rgba(2, 132, 199, 0.02) !important;
    }
    @media (prefers-color-scheme: dark) {
        div[data-testid="stFileUploader"] {
            background: #1e2430 !important;
            border-color: #334155 !important;
        }
        div[data-testid="stFileUploader"]:hover {
            border-color: #38bdf8 !important;
            background: rgba(56, 189, 248, 0.04) !important;
        }
    }
    [data-theme="dark"] div[data-testid="stFileUploader"] {
        background: #1e2430 !important;
        border-color: #334155 !important;
    }

    /* ================= 业务主操作按钮深度定制 (告别警示红，拥抱高端商务蓝) ================= */
    div.stButton > button[kind="primary"],
    div.stButton > button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.65rem 2rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.2px !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.32) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
        cursor: pointer !important;
        margin-top: 10px !important;
    }

    div.stButton > button[kind="primary"]:hover,
    div.stButton > button[data-testid="baseButton-primary"]:hover {
        background: linear-gradient(135deg, #0369a1 0%, #075985 100%) !important;
        transform: translateY(-1.5px) !important;
        box-shadow: 0 8px 22px rgba(2, 132, 199, 0.45) !important;
    }

    div.stButton > button[kind="primary"]:disabled,
    div.stButton > button[data-testid="baseButton-primary"]:disabled {
        background: #e2e8f0 !important;
        background-color: #e2e8f0 !important;
        color: #94a3b8 !important;
        box-shadow: none !important;
        cursor: not-allowed !important;
        transform: none !important;
    }

    @media (prefers-color-scheme: dark) {
        div.stButton > button[kind="primary"]:disabled,
        div.stButton > button[data-testid="baseButton-primary"]:disabled {
            background: #334155 !important;
            background-color: #334155 !important;
            color: #64748b !important;
        }
    }
    [data-theme="dark"] div.stButton > button[kind="primary"]:disabled,
    [data-theme="dark"] div.stButton > button[data-testid="baseButton-primary"]:disabled {
        background: #334155 !important;
        background-color: #334155 !important;
        color: #64748b !important;
    }

    /* KPI 指标卡片美化 */
    div[data-testid="stMetric"] {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 14px !important;
        padding: 16px 20px !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03) !important;
    }
    @media (prefers-color-scheme: dark) {
        div[data-testid="stMetric"] {
            background: #1e2430 !important;
            border-color: rgba(255, 255, 255, 0.08) !important;
        }
    }
    [data-theme="dark"] div[data-testid="stMetric"] {
        background: #1e2430 !important;
        border-color: rgba(255, 255, 255, 0.08) !important;
    }
    div[data-testid="stMetricLabel"] p {
        font-size: 0.86rem !important;
        font-weight: 600 !important;
        color: #64748b !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.65rem !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        letter-spacing: -0.5px !important;
    }
    @media (prefers-color-scheme: dark) {
        div[data-testid="stMetricLabel"] p { color: #94a3b8 !important; }
        div[data-testid="stMetricValue"] { color: #f8fafc !important; }
    }
    [data-theme="dark"] div[data-testid="stMetricLabel"] p { color: #94a3b8 !important; }
    [data-theme="dark"] div[data-testid="stMetricValue"] { color: #f8fafc !important; }

    /* 状态药丸标签 */
    .ios-precheck-box {
        background: rgba(125, 125, 125, 0.06);
        border: 1px dashed rgba(125, 125, 125, 0.28);
        border-radius: 14px;
        padding: 16px 20px;
        margin: 14px 0 20px 0;
    }
    .ios-badge-success {
        display: inline-flex;
        align-items: center;
        background: rgba(16, 185, 129, 0.14);
        color: #059669 !important;
        border: 1px solid rgba(16, 185, 129, 0.35);
        padding: 5px 12px;
        border-radius: 8px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-right: 8px;
        margin-bottom: 6px;
    }
    @media (prefers-color-scheme: dark) {
        .ios-badge-success { color: #10b981 !important; }
    }
    [data-theme="dark"] .ios-badge-success { color: #10b981 !important; }

    .ios-badge-pending {
        display: inline-flex;
        align-items: center;
        background: rgba(245, 158, 11, 0.12);
        color: #d97706 !important;
        border: 1px solid rgba(245, 158, 11, 0.35);
        padding: 5px 12px;
        border-radius: 8px;
        font-size: 0.85rem;
        font-weight: 500;
        margin-right: 8px;
        margin-bottom: 6px;
    }
    @media (prefers-color-scheme: dark) {
        .ios-badge-pending { color: #f59e0b !important; }
    }
    [data-theme="dark"] .ios-badge-pending { color: #f59e0b !important; }

    /* 下载按钮统一风格 */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        padding: 0.65rem 1.5rem !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.28) !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-1.5px) !important;
        box-shadow: 0 8px 22px rgba(16, 185, 129, 0.4) !important;
    }
</style>
""")


# -------------------------------------------------------------
# 高端矢量 SVG 图标定义（无 Emoji，专业质感）
# -------------------------------------------------------------
SVG_CARDIO = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M19 14c1.49-1.46 3-3.21 3-5.5A5.5 5.5 0 0 0 16.5 3c-1.76 0-3 .5-4.5 2-1.5-1.5-2.74-2-4.5-2A5.5 5.5 0 0 0 2 8.5c0 2.3 1.5 4.05 3 5.5l7 7Z"/>
    <path d="M3.22 12H9.5l.5-1 2 4.5 2-7 1.5 3.5h5.27"/>
</svg>
"""

SVG_DOCTOR_SETTLE = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <rect x="2" y="5" width="20" height="14" rx="3"/>
    <line x1="2" y1="10" x2="22" y2="10"/>
    <path d="M6 15h2"/>
    <path d="M12 15h6"/>
</svg>
"""

SVG_CORPUS_SIGN = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
    <polyline points="14 2 14 8 20 8"/>
    <path d="M9 15l2 2 4-4"/>
</svg>
"""

SVG_JUMEI_SETTLE = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
    <path d="M12 8v8"/>
    <path d="M8 12h8"/>
</svg>
"""

SVG_BRAND_SIDEBAR = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M12 2v20M2 12h20"/>
    <rect x="3" y="3" width="18" height="18" rx="4"/>
</svg>
"""

SVG_HOSPITAL_SAAS = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M3 21h18"/>
    <path d="M5 21V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v16"/>
    <path d="M9 21v-4a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v4"/>
    <line x1="10" y1="9" x2="14" y2="9"/>
    <line x1="12" y1="7" x2="12" y2="11"/>
</svg>
"""

SVG_BEIJIAN_SETTLE = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M9 11l3 3L22 4"/>
    <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
    <line x1="9" y1="18" x2="15" y2="18"/>
</svg>
"""

SVG_KOPU_VIDEO = """
<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <polygon points="23 7 16 12 23 17 23 7"/>
    <rect x="1" y="5" width="15" height="14" rx="2" ry="2"/>
    <circle cx="8.5" cy="12" r="2.5"/>
</svg>
"""



# -------------------------------------------------------------
# 辅助函数
# -------------------------------------------------------------
def format_size(num_bytes: int) -> str:
    """人类可读文件大小格式化"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} TB"


def create_zip_archive(files_dict: dict) -> bytes:
    """将多个文件打包为内存中的 zip 字节流"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in files_dict.items():
            zf.writestr(fname, data)
    return zip_buffer.getvalue()


# -------------------------------------------------------------
# 门禁控制：7 天内免密持久化引擎 (LocalStorage + HMAC Token)
# -------------------------------------------------------------
# 支持通过 Streamlit Secrets 或环境变量自定义密码，避免将真实密码暴露在公共仓库中
try:
    SYSTEM_PASSWORD = str(st.secrets.get("SYSTEM_PASSWORD", os.getenv("SYSTEM_PASSWORD", "910104"))).strip()
    AUTH_SALT = str(st.secrets.get("AUTH_SALT", "medical_settlement_cloud_salt_2026_secure")).strip()
except Exception:
    SYSTEM_PASSWORD = os.getenv("SYSTEM_PASSWORD", "910104")
    AUTH_SALT = "medical_settlement_cloud_salt_2026_secure"

def generate_auth_token(expiry_ts: int) -> str:
    """生成具备防篡改特性的认证 Token"""
    seed = f"{SYSTEM_PASSWORD}:{expiry_ts}:{AUTH_SALT}"
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()[:32]

def verify_auth_token(expiry_ts: int, token: str) -> bool:
    """验证 Token 是否合法且仍在有效期内"""
    if time.time() > expiry_ts:
        return False
    expected = generate_auth_token(expiry_ts)
    return expected == token


# -------------------------------------------------------------
# 智能环境判断：本地端默认直接免密直达，云端 VPS 保持安全门禁
# -------------------------------------------------------------
IS_LOCAL = (
    os.getenv("LOCAL_MODE", "").lower() in ["1", "true", "yes"] or
    os.getenv("SKIP_AUTH", "").lower() in ["1", "true", "yes"] or
    os.getenv("SYSTEM_PASSWORD", "").lower() in ["none", "off", "0"]
)

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = IS_LOCAL

# 1. 检查 URL 中是否携带了合法的 7 天免密 Token
url_token = st.query_params.get("auth_token", None)
url_exp = st.query_params.get("auth_exp", None)

if not st.session_state["authenticated"]:
    if url_token and url_exp:
        try:
            exp_int = int(url_exp)
            if verify_auth_token(exp_int, url_token):
                st.session_state["authenticated"] = True
        except Exception:
            pass

# 2. 如果当前未认证，先尝试从浏览器 LocalStorage 中读取 7 天免密凭据
if not st.session_state["authenticated"]:
    # 注入轻量 JavaScript 桥接，读取浏览器本地保存的 7 天免密状态
    components.html("""
    <script>
    try {
        const savedAuth = localStorage.getItem("medical_settlement_auth");
        if (savedAuth) {
            const data = JSON.parse(savedAuth);
            if (Date.now() < data.expiry_ms) {
                const url = new URL(window.parent.location.href);
                if (url.searchParams.get("auth_token") !== data.token) {
                    url.searchParams.set("auth_token", data.token);
                    url.searchParams.set("auth_exp", data.exp_ts);
                    window.parent.location.replace(url.href);
                }
            } else {
                localStorage.removeItem("medical_settlement_auth");
            }
        }
    } catch(e) {}
    </script>
    """, height=0, width=0)

    _, col_center, _ = st.columns([1, 1.35, 1])
    with col_center:
        with st.form("login_form", clear_on_submit=False):
            render_html("""
            <div style="text-align: center;">
                <div class="tower-logo-badge">
                    <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 2v20M2 12h20"/>
                        <rect x="3" y="3" width="18" height="18" rx="4"/>
                    </svg>
                </div>
                <div class="tower-title">医疗统计与结算中心</div>
            </div>
            """)

            entered_pwd = st.text_input(
                "请输入访问密码",
                type="password",
                placeholder="请输入访问密码",
                label_visibility="collapsed"
            )
            
            remember_me = st.checkbox("7天内免输入密码", value=True)
            
            submit_login = st.form_submit_button("验证并进入", use_container_width=True)

            if submit_login:
                if entered_pwd == SYSTEM_PASSWORD:
                    st.session_state["authenticated"] = True
                    if remember_me:
                        # 生成未来 7 天的有效期时间戳并写入浏览器 LocalStorage
                        expiry_ts = int(time.time() + 7 * 86400)
                        token = generate_auth_token(expiry_ts)
                        st.query_params["auth_token"] = token
                        st.query_params["auth_exp"] = str(expiry_ts)
                        components.html(f"""
                        <script>
                        try {{
                            localStorage.setItem("medical_settlement_auth", JSON.stringify({{
                                token: "{token}",
                                exp_ts: {expiry_ts},
                                expiry_ms: {expiry_ts * 1000}
                            }}));
                        }} catch(e) {{}}
                        </script>
                        """, height=0, width=0)
                    st.rerun()
                else:
                    st.error("访问密码错误，请重新输入")

    st.stop()


# -------------------------------------------------------------
# 业务功能定义
# -------------------------------------------------------------
MODULE_SHANGYAO = "上药雷允上进度表"
MODULE_CORPUS = "语料库电签信息表"
MODULE_ZHENGHE = "北京整合-上药雷允上结算包"
MODULE_JUMEI = "陈菊梅基金会-雷允上结算包"
MODULE_SAAS = "老Saas医院导入模板"
MODULE_BEIJIAN = "北检&康恩贝结算表"
MODULE_KOPU = "北检&华东科普视频电签表"

MODULE_OPTIONS = [
    MODULE_SHANGYAO,
    MODULE_CORPUS,
    MODULE_ZHENGHE,
    MODULE_JUMEI,
    MODULE_SAAS,
    MODULE_BEIJIAN,
    MODULE_KOPU
]

@st.cache_resource(show_spinner="正在载入全国 49.4 万医院等级知识库与检索引擎...")
def get_hospital_matcher():
    """单例全局缓存全国 49.4 万医院知识库，秒级常驻内存，免重复初始化"""
    from 老Saas医院导入模板.hospital_grade_tool import HospitalGradeMatcher
    return HospitalGradeMatcher()

if "current_module" not in st.session_state or st.session_state["current_module"] not in MODULE_OPTIONS:
    st.session_state["current_module"] = MODULE_SHANGYAO


# -------------------------------------------------------------
# 左侧侧边栏：大按钮直观导航 (单排显示，100% 同步标题)
# -------------------------------------------------------------
with st.sidebar:
    render_html(f"""
    <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 20px;">
        <div style="width: 42px; height: 42px; border-radius: 12px; background: rgba(2, 132, 199, 0.1); display: flex; align-items: center; justify-content: center;">
            {SVG_BRAND_SIDEBAR}
        </div>
        <div>
            <div style="font-weight: 800; font-size: 1.15rem; letter-spacing: -0.3px; color: var(--text-color);">业务统计结算中心</div>
            <div style="font-size: 0.78rem; color: #86868b;">Cloud Settlement Platform</div>
        </div>
    </div>
    """, container=st.sidebar)
    
    st.sidebar.markdown("##### 业务功能导航")
    st.sidebar.caption("点击下方按钮快速切换：")

    # 左侧大尺寸直观切换按钮组 (单排防折行，无 Emoji)
    for mod in MODULE_OPTIONS:
        is_active = (st.session_state["current_module"] == mod)
        btn_type = "primary" if is_active else "secondary"
        if st.sidebar.button(mod, key=f"side_btn_{mod}", use_container_width=True, type=btn_type):
            st.session_state["current_module"] = mod
            st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.markdown("##### 系统环境状态")
    render_html(f'<span class="ios-badge-success">Python {sys.version.split()[0]} 原生全栈引擎</span>', container=st.sidebar)
    render_html(f'<span class="ios-badge-success">{"本地极速直达 (免密运行)" if IS_LOCAL else "云端安全防护 (7天免密)"}</span>', container=st.sidebar)
    render_html('<span class="ios-badge-success">7大业务模块就绪</span>', container=st.sidebar)


current_module = st.session_state["current_module"]


# =============================================================
# 模块一：上药雷允上进度表
# =============================================================
if current_module == MODULE_SHANGYAO:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-badge">
                {SVG_CARDIO}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_SHANGYAO}</div>
                <div class="ios-hero-subtitle">自动识别「答卷记录」与「项目人员」，高保真生成包含【统计汇总】、【人员维度】、【案例维度】的最终业务总表。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            Python 原生引擎
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_SHANGYAO)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">智能嗅探 · 无需重命名</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">答卷记录表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">Excel 内必须包含名为 <code class="code-pill">答卷记录</code> 的工作表（Sheet）</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">项目人员列表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">Excel 内必须包含名为 <code class="code-pill">项目人员</code> 的工作表（Sheet）</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_files = st.file_uploader(
        "拖拽或选择上传 2 个 Excel 源文件 (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="upload_shangyao"
    )

    # 智能实时预检
    has_src = None
    has_staff = None
    if uploaded_files:
        for uf in uploaded_files:
            try:
                xl = pd.ExcelFile(io.BytesIO(uf.getvalue()))
                if "答卷记录" in xl.sheet_names:
                    has_src = uf.name
                if "项目人员" in xl.sheet_names:
                    has_staff = uf.name
            except Exception:
                pass

        render_html('<div class="ios-precheck-box"><b>文件智能预检状态：</b><br>')
        cols_check = st.columns(2)
        with cols_check[0]:
            if has_src:
                render_html(f'<span class="ios-badge-success">已检测到「答卷记录表」</span><br><small style="opacity:0.8;">匹配文件: {has_src}</small>', container=cols_check[0])
            else:
                render_html('<span class="ios-badge-pending">尚未检测到包含「答卷记录」工作表的文件</span>', container=cols_check[0])
        with cols_check[1]:
            if has_staff:
                render_html(f'<span class="ios-badge-success">已检测到「项目人员列表」</span><br><small style="opacity:0.8;">匹配文件: {has_staff}</small>', container=cols_check[1])
            else:
                render_html('<span class="ios-badge-pending">尚未检测到包含「项目人员」工作表的文件</span>', container=cols_check[1])
        render_html('</div>')

    _, col_btn, _ = st.columns([1, 1.8, 1])
    with col_btn:
        ready_to_run = bool(has_src and has_staff)
        run_btn = st.button(
            f"开始生成：{MODULE_SHANGYAO}",
            type="primary",
            disabled=not ready_to_run if uploaded_files else False,
            use_container_width=True
        )

    if run_btn:
        if not uploaded_files or len(uploaded_files) < 2:
            st.error("请至少上传 2 个 Excel 文件（分别包含「答卷记录」和「项目人员」）")
        elif not has_src or not has_staff:
            st.error("上传的文件中未能同时找到「答卷记录」和「项目人员」工作表，请核验后重试")
        else:
            with st.status("正在沙盒中安全调用 Python 原生统计引擎...", expanded=True) as status:
                st.write("1. 搭建隔离沙盒...")
                with tempfile.TemporaryDirectory() as temp_dir:
                    for uf in uploaded_files:
                        file_path = os.path.join(temp_dir, uf.name)
                        with open(file_path, "wb") as f:
                            f.write(uf.getvalue())
                        st.write(f"   输入文件入库: `{uf.name}`")

                    py_src = os.path.join(ROOT_DIR, "上药报表统计", "统计报表.py")
                    py_dst = os.path.join(temp_dir, "统计报表.py")
                    shutil.copy2(py_src, py_dst)

                    env = os.environ.copy()
                    env["PYTHONIOENCODING"] = "utf-8"
                    env["PYTHONUTF8"] = "1"

                    st.write("2. 执行 Python 统计报表脚本...")
                    proc = subprocess.run(
                        [sys.executable, "统计报表.py"],
                        cwd=temp_dir,
                        env=env,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace"
                    )

                    output_file = os.path.join(temp_dir, "统计总表.xlsx")
                    if proc.returncode == 0 and os.path.exists(output_file):
                        status.update(label="统计总表生成完毕", state="complete")
                        with open(output_file, "rb") as f:
                            result_bytes = f.read()

                        st.success("上药雷允上进度表已成功生成")

                        xl = pd.ExcelFile(io.BytesIO(result_bytes))
                        df_sum = xl.parse('统计汇总')
                        df_staff = xl.parse('人员维度') if '人员维度' in xl.sheet_names else pd.DataFrame()
                        
                        total_tasks = 0
                        prov_count = len(df_sum) - 1 if len(df_sum) > 1 else len(df_sum)
                        if '任务数量' in df_sum.columns:
                            total_tasks = df_sum['任务数量'].dropna().iloc[-1]

                        # 自动归档至历史记录
                        kpi_summary = f"覆盖省份 {prov_count} 个 · 总任务 {int(total_tasks) if str(total_tasks).isdigit() else total_tasks} 份 · 人员规模 {len(df_staff)} 位专家"
                        history_manager.save_run(MODULE_SHANGYAO, {"统计总表.xlsx": result_bytes}, summary=kpi_summary)

                        st.markdown("##### 核心业务 KPI 概览")
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            st.metric("覆盖省份", f"{prov_count} 个")
                        with c2:
                            st.metric("总任务数量", f"{int(total_tasks) if str(total_tasks).isdigit() else total_tasks} 份")
                        with c3:
                            st.metric("项目人员规模", f"{len(df_staff)} 位专家")

                        if not df_sum.empty and '医院所在省' in df_sum.columns and '任务数量' in df_sum.columns:
                            df_chart = df_sum[df_sum['医院所在省'] != '总计'].copy()
                            if not df_chart.empty:
                                st.markdown("##### 各省份任务分布看板")
                                st.bar_chart(df_chart.set_index('医院所在省')['任务数量'])

                        st.markdown("##### 报表数据多维度预览")
                        tabs = st.tabs([f"工作表: {s}" for s in xl.sheet_names])
                        for idx, s in enumerate(xl.sheet_names):
                            with tabs[idx]:
                                df_sheet = xl.parse(s)
                                st.dataframe(df_sheet, use_container_width=True)

                        with st.expander("查看原始脚本终端执行输出", expanded=False):
                            st.code(proc.stdout)

                        st.markdown("---")
                        st.download_button(
                            label=f"下载【统计总表.xlsx】 ({format_size(len(result_bytes))})",
                            data=result_bytes,
                            file_name="统计总表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        status.update(label="报表生成失败", state="error")
                        st.error("执行脚本时出错，详情如下：")
                        st.code(proc.stderr or proc.stdout)


# =============================================================
# 模块三：北京整合-上药雷允上结算包
# =============================================================
elif current_module == MODULE_ZHENGHE:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-badge">
                {SVG_DOCTOR_SETTLE}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_ZHENGHE}</div>
                <div class="ios-hero-subtitle">自适应任意行数、表头同义词智能定位、卡号防科学计数失真、双向勾稽门禁。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            Python 审计内核
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_ZHENGHE)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="20" height="14" rx="3"/><line x1="2" y1="10" x2="22" y2="10"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">多表智能嗅探 · 自动对齐</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">结算明细文件 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">包含医生姓名、单价、语料/任务编号等核心数据的明细表格</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">辅助任务明细 / 确认单 <span class="req-tag-opt">可选</span></div>
                    <div class="bento-card-desc">第三方系统导出的核验单据，提供时自动进行双向交叉勾稽核验</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_zhenghe_files = st.file_uploader(
        "拖拽或批量选择上传源表格 (.xlsx / .xls，支持同时上传结算明细与辅助确认单)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_zhenghe"
    )

    # 智能实时预识别嗅探
    file_yl_obj = None
    file_task_obj = None

    if uploaded_zhenghe_files:
        # 第一轮：按文件名强特征优先匹配
        for uf in uploaded_zhenghe_files:
            fname = uf.name.lower()
            if any(k in fname for k in ["任务明细", "确认单", "js-", "核验", "对账"]):
                if not file_task_obj:
                    file_task_obj = uf
            elif any(k in fname for k in ["语料", "结算明细", "明细", "结算", "上药", "雷允上"]):
                if not file_yl_obj:
                    file_yl_obj = uf

        # 第二轮：表头探测
        for uf in uploaded_zhenghe_files:
            if uf in [file_yl_obj, file_task_obj]:
                continue
            try:
                df_head = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=2)
                h_str = "".join([str(c) for c in df_head.columns])
                if not file_task_obj and ("任务名称" in h_str or "课程名称" in h_str or "确认单" in h_str):
                    file_task_obj = uf
                elif not file_yl_obj and ("医生" in h_str or "姓名" in h_str or "单价" in h_str or "卡号" in h_str or "银行" in h_str):
                    file_yl_obj = uf
            except Exception:
                pass

        # 兜底匹配：若只上传了 1 个文件且未识别，默认为结算主明细
        if not file_yl_obj and len(uploaded_zhenghe_files) == 1:
            file_yl_obj = uploaded_zhenghe_files[0]
        elif not file_yl_obj and uploaded_zhenghe_files:
            for uf in uploaded_zhenghe_files:
                if uf != file_task_obj:
                    file_yl_obj = uf
                    break

        st.markdown("##### 实时文件嗅探匹配结果")
        c1, c2 = st.columns(2)
        with c1:
            if file_yl_obj:
                render_html(f'<div class="ios-status-card success"><div class="ios-status-card-title">01 结算明细文件 (*必选)</div><div class="ios-status-card-val">{file_yl_obj.name}</div></div>')
            else:
                render_html('<div class="ios-status-card warning"><div class="ios-status-card-title">01 结算明细文件 (*必选)</div><div class="ios-status-card-val">未识别 (需包含“明细/语料/结算”)</div></div>')
        with c2:
            if file_task_obj:
                render_html(f'<div class="ios-status-card success"><div class="ios-status-card-title">02 辅助任务明细/确认单 (可选)</div><div class="ios-status-card-val">{file_task_obj.name}</div></div>')
            else:
                render_html('<div class="ios-status-card neutral"><div class="ios-status-card-title">02 辅助任务明细/确认单 (可选)</div><div class="ios-status-card-val">未提供 (将直接根据主明细加总)</div></div>')

    _, col_btn, _ = st.columns([1, 1.8, 1])
    with col_btn:
        run_btn = st.button(f"开始生成：{MODULE_ZHENGHE}", type="primary", use_container_width=True)

    if run_btn:
        if not file_yl_obj:
            st.error("请先上传【01. 结算明细文件】（必选）后再点击开始生成。")
        else:
            with st.status("正在沙盒中执行北京整合-上药雷允上结算逻辑...", expanded=True) as status:
                st.write("1. 正在初始化沙盒...")
                with tempfile.TemporaryDirectory() as temp_dir:
                    yl_path = os.path.join(temp_dir, file_yl_obj.name)
                    with open(yl_path, "wb") as f:
                        f.write(file_yl_obj.getvalue())

                    task_path = None
                    if file_task_obj:
                        task_path = os.path.join(temp_dir, file_task_obj.name)
                        with open(task_path, "wb") as f:
                            f.write(file_task_obj.getvalue())

                    st.write("2. 调度原版结算审计引擎...")
                    script_dir = os.path.join(ROOT_DIR, "整合学会统计")
                    
                    env = os.environ.copy()
                    env["PYTHONIOENCODING"] = "utf-8"
                    env["PYTHONUTF8"] = "1"

                    runner_script = """
import sys
from pathlib import Path
script_dir = sys.argv[1]
yl = sys.argv[2]
task = sys.argv[3] if sys.argv[3] != "NONE" else None
out_dir = sys.argv[4]

sys.path.insert(0, script_dir)
from settlement_mac import process_settlement

out1, out2 = process_settlement(yl, task, out_dir, log_func=print)
print("SUCCESS_OUT1:" + str(out1))
print("SUCCESS_OUT2:" + str(out2))
"""
                    proc = subprocess.run(
                        [
                            sys.executable,
                            "-c",
                            runner_script,
                            script_dir,
                            yl_path,
                            task_path if task_path else "NONE",
                            temp_dir
                        ],
                        cwd=temp_dir,
                        env=env,
                        capture_output=True,
                        text=True,
                        encoding="utf-8"
                    )

                    gen_files = {}
                    ignore_names = [file_yl_obj.name]
                    if file_task_obj:
                        ignore_names.append(file_task_obj.name)
                    for p in Path(temp_dir).glob("*.xlsx"):
                        if p.name not in ignore_names:
                            with open(p, "rb") as f:
                                gen_files[p.name] = f.read()

                    if proc.returncode == 0 and gen_files:
                        status.update(label="结算表格生成完毕", state="complete")
                        st.success(f"成功完成对账结算，共生成 {len(gen_files)} 个结算报表")

                        doc_count_match = re.search(r'汇总完成[^\d]*(\d+)[^\d]*位医生', proc.stdout)
                        amount_match = re.search(r'劳务实发总额[^\d]*([0-9\.,]+)', proc.stdout)

                        # 自动归档至历史记录
                        doc_cnt = doc_count_match.group(1) if doc_count_match else f"{len(gen_files)}"
                        amt_str = amount_match.group(1) if amount_match else "已核算平账"
                        kpi_summary = f"本次结算专家数 {doc_cnt} 人 · 实发劳务总金额 ¥ {amt_str}"
                        history_manager.save_run(MODULE_ZHENGHE, gen_files, summary=kpi_summary)

                        st.markdown("##### 结算对账 KPI 看板")
                        k1, k2 = st.columns(2)
                        with k1:
                            st.metric("本次结算专家数", f"{doc_count_match.group(1)} 人" if doc_count_match else f"{len(gen_files)} 份表单")
                        with k2:
                            st.metric("实发劳务总金额", f"¥ {amount_match.group(1)}" if amount_match else "已核算平账")

                        st.markdown("##### 生成表格数据概览与下载")
                        tabs = st.tabs([f"{k} ({format_size(len(v))})" for k, v in gen_files.items()])
                        for idx, (fname, fbytes) in enumerate(gen_files.items()):
                            with tabs[idx]:
                                try:
                                    df_prev = pd.read_excel(io.BytesIO(fbytes))
                                    st.dataframe(df_prev.head(100), use_container_width=True)
                                except Exception:
                                    pass
                                st.download_button(
                                    label=f"单独下载【{fname}】 ({format_size(len(fbytes))})",
                                    data=fbytes,
                                    file_name=fname,
                                    key=f"dl_{fname}"
                                )

                        with st.expander("查看审计核验完整日志", expanded=False):
                            st.code(proc.stdout)

                        if len(gen_files) > 1:
                            st.markdown("---")
                            zip_bytes = create_zip_archive(gen_files)
                            st.download_button(
                                label=f"打包下载全部结算表格 (ZIP) ({format_size(len(zip_bytes))})",
                                data=zip_bytes,
                                file_name="医生劳务结算完整包.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                    else:
                        status.update(label="结算生成失败", state="error")
                        st.error("执行过程出现错误：")
                        st.code(proc.stderr or proc.stdout)


# =============================================================
# 模块二：语料库电签信息表
# =============================================================
elif current_module == MODULE_CORPUS:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-badge">
                {SVG_CORPUS_SIGN}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_CORPUS}</div>
                <div class="ios-hero-subtitle">智能嗅探 3 张关键表格，自动完成实发清单、银行发放表与对账明细的端到端核验。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            闭环交叉核验
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_CORPUS)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M9 15l2 2 4-4"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">3 表特征自动嗅探匹配</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">三方项目支付清单 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">包含本期要结算发放的语料编号、结算单价等</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">语料列表 (明文) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">包含题目内容、医生姓名、身份证号等业务数据</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">03</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">用户列表 (明文) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">包含医生银行卡号、开户行、预留手机号等金融信息</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_files = st.file_uploader(
        "拖拽或批量选择上传 3 个源表格 (.xlsx / .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_corpus"
    )

    # 智能实时预识别嗅探
    pay_name, corpus_name, user_name = None, None, None
    if uploaded_files:
        for uf in uploaded_files:
            fname = uf.name.lower()
            if ("支付" in fname or "清单" in fname or "结算" in fname) and ("语料列表" not in fname and "明文" not in fname):
                pay_name = uf.name
            elif ("语料列表" in fname or "语料" in fname) and ("支付" not in fname):
                corpus_name = uf.name
            elif ("用户列表" in fname or "用户" in fname or "医生列表" in fname):
                user_name = uf.name

        for uf in uploaded_files:
            if uf.name in [pay_name, corpus_name, user_name]:
                continue
            try:
                df_head = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=2)
                h_str = "".join([str(c) for c in df_head.columns])
                if not pay_name and ("单价" in h_str or "结算" in h_str) and "语料" in h_str:
                    pay_name = uf.name
                elif not corpus_name and ("题目" in h_str or "回答" in h_str or "疾病" in h_str):
                    corpus_name = uf.name
                elif not user_name and ("银行卡" in h_str or "开户行" in h_str):
                    user_name = uf.name
            except Exception:
                pass

        render_html('<div class="ios-precheck-box"><b>智能嗅探识别进度：</b><br>')
        chk_cols = st.columns(3)
        with chk_cols[0]:
            if pay_name:
                render_html(f'<span class="ios-badge-success">支付清单：已锁定</span><br><small style="opacity:0.8;">{pay_name}</small>', container=chk_cols[0])
            else:
                render_html('<span class="ios-badge-pending">待识别：三方项目支付清单</span>', container=chk_cols[0])
        with chk_cols[1]:
            if corpus_name:
                render_html(f'<span class="ios-badge-success">语料明文：已锁定</span><br><small style="opacity:0.8;">{corpus_name}</small>', container=chk_cols[1])
            else:
                render_html('<span class="ios-badge-pending">待识别：语料列表(明文)</span>', container=chk_cols[1])
        with chk_cols[2]:
            if user_name:
                render_html(f'<span class="ios-badge-success">用户明文：已锁定</span><br><small style="opacity:0.8;">{user_name}</small>', container=chk_cols[2])
            else:
                render_html('<span class="ios-badge-pending">待识别：用户列表(明文)</span>', container=chk_cols[2])
        render_html('</div>')

    _, col_btn, _ = st.columns([1, 1.8, 1])
    with col_btn:
        ready_corpus = bool(uploaded_files and len(uploaded_files) >= 3)
        run_btn = st.button(
            f"开始生成：{MODULE_CORPUS}",
            type="primary",
            disabled=not ready_corpus if uploaded_files else False,
            use_container_width=True
        )

    if run_btn:
        if not uploaded_files or len(uploaded_files) < 3:
            st.error("每月语料结算需要 3 个源表格（支付清单、语料明文、用户明文），请完整上传")
        else:
            with st.status("正在沙盒中执行每月语料结算闭环核验...", expanded=True) as status:
                st.write("1. 初始化沙盒环境并暂存表格...")
                with tempfile.TemporaryDirectory() as temp_dir:
                    original_names = set()
                    for uf in uploaded_files:
                        fpath = os.path.join(temp_dir, uf.name)
                        with open(fpath, "wb") as f:
                            f.write(uf.getvalue())
                        original_names.add(uf.name)

                    st.write("2. 执行原版结算核算引擎...")
                    script_path = os.path.join(ROOT_DIR, "语料库电签统计", "settlement_tool.py")

                    env = os.environ.copy()
                    env["PYTHONIOENCODING"] = "utf-8"
                    env["PYTHONUTF8"] = "1"

                    proc = subprocess.run(
                        [sys.executable, script_path, temp_dir],
                        cwd=temp_dir,
                        env=env,
                        capture_output=True,
                        text=True,
                        encoding="utf-8"
                    )

                    output_files = {}
                    warnings_text = None
                    for p in Path(temp_dir).iterdir():
                        if p.name not in original_names and not p.name.startswith("~$"):
                            if p.name.endswith(".txt") and "警告" in p.name:
                                with open(p, "r", encoding="utf-8") as wf:
                                    warnings_text = wf.read()
                            with open(p, "rb") as f:
                                output_files[p.name] = f.read()

                    if proc.returncode == 0 and output_files:
                        status.update(label="月度结算全量完成", state="complete")
                        st.success(f"成功完成结算，共产出 {len(output_files)} 个结算与对账文件")

                        doc_cnt = re.search(r'本期结算总人数\s*:\s*(\d+)', proc.stdout)
                        corp_cnt = re.search(r'本期结算语料数\s*:\s*(\d+)', proc.stdout)
                        amt_cnt = re.search(r'本期应付总金额\s*:\s*¥\s*([0-9\.,]+)', proc.stdout)

                        # 自动归档至历史记录
                        d_str = f"{doc_cnt.group(1)} 位医生" if doc_cnt else "已校验"
                        c_str = f"{corp_cnt.group(1)} 条" if corp_cnt else "已校验"
                        a_str = f"¥ {amt_cnt.group(1)}" if amt_cnt else "已校验"
                        kpi_summary = f"结算总人数 {d_str} · 结算语料数 {c_str} · 应付总金额 {a_str}"
                        history_manager.save_run(MODULE_CORPUS, output_files, summary=kpi_summary)

                        st.markdown("##### 本期结算核心 KPI 看板")
                        m1, m2, m3 = st.columns(3)
                        with m1:
                            st.metric("结算总人数", f"{doc_cnt.group(1)} 位医生" if doc_cnt else "已校验")
                        with m2:
                            st.metric("结算语料数", f"{corp_cnt.group(1)} 条" if corp_cnt else "已校验")
                        with m3:
                            st.metric("应付总金额", f"¥ {amt_cnt.group(1)}" if amt_cnt else "已校验")

                        if warnings_text:
                            st.warning("发现部分待核查医生（如银行卡缺失），已自动单独提取为警告名单：")
                            st.text_area("待核查名单详情", warnings_text, height=130)

                        st.markdown("##### 生成表格预览与下载")
                        tabs = st.tabs([f"{k} ({format_size(len(v))})" for k, v in output_files.items()])
                        for idx, (fname, fbytes) in enumerate(output_files.items()):
                            with tabs[idx]:
                                if fname.endswith(".xlsx"):
                                    try:
                                        df_p = pd.read_excel(io.BytesIO(fbytes))
                                        st.dataframe(df_p.head(100), use_container_width=True)
                                    except Exception:
                                        pass
                                elif fname.endswith(".txt"):
                                    st.code(fbytes.decode("utf-8", errors="ignore"))

                                st.download_button(
                                    label=f"单独下载【{fname}】 ({format_size(len(fbytes))})",
                                    data=fbytes,
                                    file_name=fname,
                                    key=f"dl_corp_{fname}"
                                )

                        with st.expander("查看月度结算审计核查完整日志", expanded=False):
                            st.code(proc.stdout)

                        st.markdown("---")
                        zip_data = create_zip_archive(output_files)
                        st.download_button(
                            label=f"打包下载全部月度结果文件 (ZIP) ({format_size(len(zip_data))})",
                            data=zip_data,
                            file_name="月度语料结算完整结果包.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                    else:
                        status.update(label="结算生成失败", state="error")
                        st.error("执行脚本时出错，详情如下：")
                        st.code(proc.stderr or proc.stdout)


# =============================================================
# 模块四：陈菊梅基金会-雷允上结算包
# =============================================================
elif current_module == MODULE_JUMEI:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-badge">
                {SVG_JUMEI_SETTLE}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_JUMEI}</div>
                <div class="ios-hero-subtitle">智能关联项目进度表、语料明文与用户明文，规范导出费用明细汇总、各项目明细与语料对账总表。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            Python 结算内核
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_JUMEI)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M12 8v8"/><path d="M8 12h8"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">2 表极简驱动 · 智能防乱序嗅探</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">医生底表 / 用户列表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「用户」或「医生」，包含身份证号、开户银行、支行、银行卡号、手机号</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">语料交付表 / 进度表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「语料」或「进度」，包含题目内容、结算单价、审核状态、项目名称、手机号</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">03</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">辅助进度表 <span class="req-tag-opt" style="background:#e0f2fe;color:#0369a1;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;">可选兼容</span></div>
                    <div class="bento-card-desc">支持传统 3 表同时上传，系统将智能优先选取完整语料交付数据进行全自适应核销</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_jumei_files = st.file_uploader(
        "拖拽或批量选择上传源表格 (.xlsx / .xls，支持 2 表或 3 表同时上传)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_jumei"
    )

    # 智能预识别嗅探（支持单项目及多项目交付表同时上传）
    doc_file_obj = None
    corpus_file_objs = []
    detected_project_tags = []

    if uploaded_jumei_files:
        # 1. 评分法高精度识别医生底表/用户明文表
        scored_files = []
        for uf in uploaded_jumei_files:
            fname = uf.name.lower()
            score = 0
            if any(k in fname for k in ["用户", "医生", "底表", "资质"]):
                score += 20
            if any(k in fname for k in ["语料", "交付", "进度", "题目"]):
                score -= 20

            # 表头特征探测
            try:
                df_head = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=3)
                h_str = " ".join([str(c) for c in df_head.columns])
                if any(k in h_str for k in ["开户行", "开户银行", "支行"]):
                    score += 25
                if any(k in h_str for k in ["银行卡", "卡号", "结算账号"]):
                    score += 25
                if any(k in h_str for k in ["身份证", "证件号", "证件号码"]):
                    score += 15
                if any(k in h_str for k in ["单价", "结算单价", "词条", "语料", "题目"]):
                    score -= 20
            except Exception:
                pass
            scored_files.append((score, uf))

        # 得分最高且 >= 20 的判定为医生底表
        scored_files.sort(key=lambda x: x[0], reverse=True)
        if scored_files and scored_files[0][0] >= 20:
            doc_file_obj = scored_files[0][1]
        elif scored_files:
            one_f = next((x[1] for x in scored_files if "1.xlsx" in x[1].name.lower()), None)
            doc_file_obj = one_f or scored_files[0][1]

        # 2. 其余所有文件自动归类为语料交付表 (支持单表或多项目交付表)
        for _, uf in scored_files:
            if doc_file_obj and uf.name == doc_file_obj.name:
                continue
            corpus_file_objs.append(uf)
            # 探测该语料表涉及的项目
            try:
                df_c_head = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=100)
                p_cols = [c for c in df_c_head.columns if any(k in str(c) for k in ["项目", "所属项目"])]
                if p_cols:
                    unique_p = df_c_head[p_cols[0]].dropna().unique()
                    for p in unique_p:
                        for tag in ["云南", "长春", "北京", "上海", "广州", "深圳", "四川", "山东"]:
                            if tag in str(p) and tag not in detected_project_tags:
                                detected_project_tags.append(tag)
                for tag in ["云南", "长春", "北京", "上海", "广州", "深圳", "四川", "山东"]:
                    if tag in uf.name and tag not in detected_project_tags:
                        detected_project_tags.append(tag)
            except Exception:
                pass

        # 智能匹配各卡槽表格
        def _get_corpus_tag(uf_obj):
            fname = uf_obj.name
            for tag in ["云南", "长春", "北京", "上海", "广州", "深圳", "四川", "山东", "浙江", "江苏"]:
                if tag in fname:
                    return tag
            try:
                df_peek = pd.read_excel(io.BytesIO(uf_obj.getvalue()), nrows=50)
                p_cols = [c for c in df_peek.columns if any(k in str(c) for k in ["项目", "所属项目"])]
                if p_cols:
                    for v in df_peek[p_cols[0]].dropna().astype(str):
                        for tag in ["云南", "长春", "北京", "上海", "广州", "深圳", "四川", "山东", "浙江", "江苏"]:
                            if tag in v:
                                return tag
            except Exception:
                pass
            return None

        c_yunnan = next((f for f in corpus_file_objs if _get_corpus_tag(f) == "云南"), None)
        c_changchun = next((f for f in corpus_file_objs if _get_corpus_tag(f) == "长春"), None)
        remaining_corpus = [f for f in corpus_file_objs if f not in [c_yunnan, c_changchun]]

        file_2_obj = c_yunnan or (corpus_file_objs[0] if corpus_file_objs else None)
        if file_2_obj == c_yunnan:
            file_3_obj = c_changchun or (remaining_corpus[0] if remaining_corpus else None)
        else:
            file_3_obj = corpus_file_objs[1] if len(corpus_file_objs) > 1 else None

        extra_corpus = [f for f in corpus_file_objs if f not in [file_2_obj, file_3_obj]]

        # 渲染识别状态指示条（与 Module 6 规范嗅探状态样式一致）
        render_html('<div class="ios-precheck-box"><b>智能多表嗅探识别状态：</b><br>')
        chk_cols = st.columns(3)
        with chk_cols[0]:
            if doc_file_obj:
                render_html(f'<span class="ios-badge-success">1. 医生底表 / 用户列表：已锁定</span><br><small style="opacity:0.8;">{doc_file_obj.name}</small>', container=chk_cols[0])
            else:
                render_html('<span class="ios-badge-pending">待识别：1. 医生底表 / 用户列表 (1.xlsx)</span>', container=chk_cols[0])

        with chk_cols[1]:
            if file_2_obj:
                t2 = _get_corpus_tag(file_2_obj)
                t2_str = f" ({t2})" if t2 else ""
                render_html(f'<span class="ios-badge-success">2. 语料交付表{t2_str}：已锁定</span><br><small style="opacity:0.8;">{file_2_obj.name}</small>', container=chk_cols[1])
            else:
                render_html('<span class="ios-badge-pending">待识别：2. 语料交付表 (2.xlsx)</span>', container=chk_cols[1])

        with chk_cols[2]:
            if file_3_obj:
                t3 = _get_corpus_tag(file_3_obj)
                t3_str = f" ({t3})" if t3 else ""
                extra_note = f"<br><small style='opacity:0.75;'>另有 {len(extra_corpus)} 个额外项目交付表已锁定</small>" if extra_corpus else ""
                render_html(f'<span class="ios-badge-success">3. 语料交付/辅助表{t3_str}：已锁定</span><br><small style="opacity:0.8;">{file_3_obj.name}</small>{extra_note}', container=chk_cols[2])
            elif len(corpus_file_objs) == 1:
                render_html('<span class="ios-badge-pending" style="background:rgba(2,132,199,0.08);color:#0284c7;border-color:rgba(2,132,199,0.25);">3. 辅助进度/新项目表：可选 (当前单表模式)</span>', container=chk_cols[2])
            else:
                render_html('<span class="ios-badge-pending">待识别：3. 语料交付/辅助进度表 (3.xlsx)</span>', container=chk_cols[2])

        # 智能状态说明条
        if detected_project_tags:
            ordered_tags = []
            if "长春" in detected_project_tags:
                ordered_tags.append("长春 (6.xlsx)")
            if "云南" in detected_project_tags:
                ordered_tags.append("云南 (7.xlsx)")
            next_idx = 8
            for t in detected_project_tags:
                if t not in ["长春", "云南"]:
                    ordered_tags.append(f"{t} ({next_idx}.xlsx)")
                    next_idx += 1
            split_desc = "、".join(ordered_tags)
            info_text = f"已自动嗅探识别到 <b>{len(detected_project_tags)} 个独立项目</b>（{split_desc}），将自动拆分为独立 Excel 并生成 ZIP 打包。"
        else:
            info_text = "系统将全自动读取【结算单价】档位、手机号精确匹配与个税核销，无需手动配置参数。"

        render_html(f"""
        <div style="margin: 12px 0 16px 0; padding: 12px 18px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; font-size: 13px; color: #166534; display: flex; align-items: center; justify-content: space-between;">
            <div style="display: flex; align-items: center; gap: 8px;">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                <span><b>自动智能嗅探与拆分</b>：{info_text}</span>
            </div>
            <div style="font-size: 12px; color: #15803d; opacity: 0.85;">多项目独立拆分引擎</div>
        </div>
        """)

    # 操作按钮黄金居中排布
    _, col_btn, _ = st.columns([1, 1.8, 1])
    with col_btn:
        start_jumei = st.button(f"开始生成：{MODULE_JUMEI}", type="primary", use_container_width=True)

    if start_jumei:
        if not uploaded_jumei_files or len(uploaded_jumei_files) < 2:
            st.error("请上传至少 2 个源数据文件（1 个医生底表 + 至少 1 个语料交付表）后再次点击生成。")
        elif not (doc_file_obj and corpus_file_objs):
            st.warning("系统未能自动匹配必要表格，请确认上传了「用户/医生」底表与「语料/进度」交付表。")
        else:
            with st.status("正在启动陈菊梅基金会雷允上劳务结算引擎...", expanded=True) as status:
                st.write("1. 正在初始化沙箱运行隔离环境...")
                temp_dir = tempfile.mkdtemp(prefix="jumei_settle_")
                try:
                    u_path = os.path.join(temp_dir, doc_file_obj.name)
                    with open(u_path, "wb") as f:
                        f.write(doc_file_obj.getvalue())

                    c_paths = []
                    for c_obj in corpus_file_objs:
                        c_p = os.path.join(temp_dir, c_obj.name)
                        with open(c_p, "wb") as f:
                            f.write(c_obj.getvalue())
                        c_paths.append(c_p)

                    today_str = datetime.datetime.now().strftime("%Y%m%d")

                    script_path = os.path.join(ROOT_DIR, "陈菊梅基金会-雷允上结算包", "generate_settlement.py")
                    st.write("2. 正在执行多项目独立拆分(6/7/8...)、全数据驱动单价动态提取、开户行智能清洗与劳务个税核销...")

                    cmd = [
                        sys.executable,
                        script_path,
                        "--doc", u_path,
                        "--corpus"
                    ] + c_paths + [
                        "--out_dir", temp_dir
                    ]

                    env = os.environ.copy()
                    env["PYTHONIOENCODING"] = "utf-8"
                    env["PYTHONUTF8"] = "1"

                    proc = subprocess.run(
                        cmd,
                        cwd=temp_dir,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        env=env
                    )

                    uploaded_names = set([doc_file_obj.name] + [c.name for c in corpus_file_objs])
                    all_xlsx = [f for f in os.listdir(temp_dir) if f.endswith(".xlsx") and f not in uploaded_names]
                    named_files = [f for f in all_xlsx if "-劳务费用明细表.xlsx" in f]
                    if not named_files:
                        named_files = [f for f in all_xlsx if re.match(r'^\d+\.xlsx$', f)] or all_xlsx
                    named_files.sort()

                    zip_fn = "劳务费用明细表_全部独立项目包.zip"
                    zip_fp = os.path.join(temp_dir, zip_fn)
                    has_zip = os.path.exists(zip_fp)

                    if proc.returncode == 0 and named_files:
                        status.update(label="独立项目结算表生成完成！", state="complete")
                        st.success(f"成功生成 {len(named_files)} 份独立项目劳务结算 Excel 文件（各文件均含专属的项目结算表与医生明细表）！")

                        # 提取 KPI 数据
                        total_items = re.search(r'共\s*(\d+)\s*条', proc.stdout)
                        total_amt = re.search(r'税后总金额:\s*([0-9\.,]+)\s*元', proc.stdout)
                        proj_matches = re.findall(r'项目【(.*?)】.*?(\d+)\s*位医生.*?(\d+)\s*条语料.*?金额合计:\s*([0-9\.,]+)\s*元', proc.stdout)
                        total_docs = sum(int(m[1]) for m in proj_matches) if proj_matches else None

                        # 自动归档至历史记录
                        jumei_files = {}
                        for fn in named_files:
                            fp = os.path.join(temp_dir, fn)
                            if os.path.exists(fp):
                                with open(fp, "rb") as ef:
                                    jumei_files[fn] = ef.read()
                        if has_zip and os.path.exists(zip_fp):
                            with open(zip_fp, "rb") as zf:
                                jumei_files[f"{today_str}-雷允上全项目劳务结算包.zip"] = zf.read()

                        kpi_summary = f"结算表 {len(named_files)} 个 · 专家 {total_docs if total_docs is not None else '多'} 位 · 语料 {total_items.group(1) if total_items else '多'} 条 · 金额 ¥ {total_amt.group(1) if total_amt else '平账'}"
                        history_manager.save_run(MODULE_JUMEI, jumei_files, summary=kpi_summary)

                        st.markdown("##### 本期结算核心 KPI 看板")
                        k1, k2, k3, k4 = st.columns(4)
                        with k1:
                            st.metric("独立结算表数", f"{len(named_files)} 个独立 Excel")
                        with k2:
                            st.metric("结算总人数", f"{total_docs} 位医生" if total_docs is not None else "已核算")
                        with k3:
                            st.metric("语料词条数", f"{total_items.group(1)} 条" if total_items else "已核算")
                        with k4:
                            st.metric("税后总金额", f"¥ {total_amt.group(1)}" if total_amt else "已核算")

                        # 下载专区
                        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                        if has_zip:
                            with open(zip_fp, "rb") as zf:
                                zip_bytes = zf.read()
                            st.markdown("##### 📦 一键打包下载全部独立表格")
                            _, col_dl_zip, _ = st.columns([1, 1.8, 1])
                            with col_dl_zip:
                                st.download_button(
                                    label=f"📦 一键打包下载全部项目独立结算表 (ZIP · {format_size(len(zip_bytes))})",
                                    data=zip_bytes,
                                    file_name=f"{today_str}-雷允上全项目劳务结算包.zip",
                                    mime="application/zip",
                                    type="primary",
                                    use_container_width=True
                                )

                        st.markdown("##### 📄 各项目独立结算表单独下载")
                        dl_cols = st.columns(min(len(named_files), 3))
                        for i, fn in enumerate(named_files):
                            c_idx = i % min(len(named_files), 3)
                            fp = os.path.join(temp_dir, fn)
                            with open(fp, "rb") as ef:
                                f_bytes = ef.read()
                            with dl_cols[c_idx]:
                                st.download_button(
                                    label=f"⬇️ 下载【{fn}】 ({format_size(len(f_bytes))})",
                                    data=f_bytes,
                                    file_name=fn,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key=f"dl_single_proj_{i}"
                                )

                        # 多独立工作簿在线预览
                        st.markdown("##### 📑 独立项目工作簿在线预览")
                        proj_tabs = st.tabs([f"📄 {fn}" for fn in named_files])
                        for p_idx, fn in enumerate(named_files):
                            with proj_tabs[p_idx]:
                                fp = os.path.join(temp_dir, fn)
                                with open(fp, "rb") as ef:
                                    f_bytes = ef.read()
                                xl_file = pd.ExcelFile(io.BytesIO(f_bytes))
                                inner_tabs = st.tabs([f"Sheet: {s}" for s in xl_file.sheet_names])
                                for s_idx, sname in enumerate(xl_file.sheet_names):
                                    with inner_tabs[s_idx]:
                                        df_sheet = pd.read_excel(xl_file, sheet_name=sname)
                                        st.caption(f"【{fn}】之工作表【{sname}】共 {len(df_sheet)} 行数据（展示前 100 行）：")
                                        st.dataframe(df_sheet.head(100), use_container_width=True)

                        with st.expander("查看数据映射与结算执行完整日志", expanded=False):
                            st.code(proc.stdout)
                    else:
                        status.update(label="结算包生成失败", state="error")
                        st.error("执行脚本时出错，详情如下：")
                        st.code(proc.stderr or proc.stdout)
                except Exception as e:
                    status.update(label="处理异常", state="error")
                    st.error(f"处理数据时发生异常: {str(e)}")
                finally:
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception:
                        pass


# =============================================================
# 模块五：老Saas医院导入模板
# =============================================================
elif current_module == MODULE_SAAS:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-badge">
                {SVG_HOSPITAL_SAAS}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_SAAS}</div>
                <div class="ios-hero-subtitle">全国 49.4 万全量医疗机构智能匹配、自动规整标准 10 级医院等级，支持批量处理与 300 条/批智能切分。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            49.4万 超级知识库
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_SAAS)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M12 8v8"/><path d="M8 12h8"/></svg>
                <span>导入规范与智能匹配标准</span>
            </div>
            <div class="bento-req-badge">全国49.4万机构 · 10项标准定级 · 300条智能分批</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">标准字段识别 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">表格首行需包含「机构名称」或「医院名称」，支持「省」、「市」辅助精确定位</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">10级标准自动定级 <span class="req-tag-must">规范</span></div>
                    <div class="bento-card-desc">严格符合系统下拉约束：无等级、一/二/三级医院、甲等/乙等</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">03</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">洗稿去重与 300条分批 <span class="req-tag-opt">自动化</span></div>
                    <div class="bento-card-desc">智能剔除知识库中已有机构与自身重复，仅对全新机构定级并切分导入</div>
                </div>
            </div>
        </div>
    </div>
    """)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    tab_batch, tab_single = st.tabs(["批量 Excel 导入补齐与切分", "单家机构等级即时检索"])

    with tab_single:
        st.markdown("##### 单家医院/医疗机构等级快速检索")
        st.caption("即时查询全国 49.4 万医疗机构等级数据库，支持全称、简称、去括号归一化及联网多级检索。")
        col_q1, col_q2, col_q3 = st.columns([2, 1, 1])
        with col_q1:
            q_hosp = st.text_input("输入医院或机构名称", placeholder="例如：成都市第三人民医院 / 四川大学华西医院", key="q_hosp_input")
        with col_q2:
            q_prov = st.text_input("省份 (可选)", placeholder="如：四川省", key="q_hosp_prov")
        with col_q3:
            q_city = st.text_input("城市 (可选)", placeholder="如：成都市", key="q_hosp_city")

        if st.button("立即查询机构等级", key="btn_single_query", type="secondary"):
            if not q_hosp.strip():
                st.warning("请输入机构或医院名称")
            else:
                with st.spinner("正在检索 49.4 万全国医疗知识库..."):
                    matcher = get_hospital_matcher()
                    grade, reason = matcher.get_grade(q_hosp, q_prov, q_city, enable_online_search=True)
                    st.success(f"检索完成！机构名称: **{q_hosp.strip()}**")
                    col_res1, col_res2 = st.columns(2)
                    with col_res1:
                        render_html(f"""
                        <div class="ios-status-card success">
                            <div class="ios-status-card-title">标准评定等级</div>
                            <div class="ios-status-card-val" style="font-size: 1.3rem; color: #0284c7;">{grade}</div>
                        </div>
                        """)
                    with col_res2:
                        render_html(f"""
                        <div class="ios-status-card success">
                            <div class="ios-status-card-title">命中规则与来源</div>
                            <div class="ios-status-card-val">{reason}</div>
                        </div>
                        """)

    with tab_batch:
        uploaded_saas_file = st.file_uploader(
            "拖拽或点击上传待处理的医院 Excel 文件 (.xlsx)",
            type=["xlsx"],
            key="upload_saas_excel"
        )

        # 洗稿去重配置面板 (以系统 49.4 万全量在库为唯一权威基准)
        with st.expander("🧼 洗稿与增量去重设置（凡系统 49.4 万在库医院直接剔除）", expanded=True):
            render_html("""
            <div style="background: rgba(2, 132, 199, 0.06); border: 1px solid rgba(2, 132, 199, 0.18); border-radius: 12px; padding: 12px 16px; margin-bottom: 12px; font-size: 0.88rem; line-height: 1.6; color: #0369a1;">
                <strong>💡 系统在库比对权威基准：</strong>
                全国 <strong>49.4 万系统已有全量医院库</strong>（已 100% 完整收录包含全部 494,022 条在库机构与历史导入模板）。<br/>
                凡是在这 <strong>49.4 万条系统已有医院</strong> 中有记录的机构，洗稿时<strong>一律自动直接剔除</strong>，仅保留全新未收录的医院，避免老 SaaS 重复导入冲突！
            </div>
            """)
            col_dd1, col_dd2 = st.columns(2)
            with col_dd1:
                enable_dedup = st.checkbox(
                    "开启「洗稿去重」功能（直接剔除 49.4 万在库已有医院，仅保留新机构）",
                    value=True,
                    help="凡是在系统 49.4 万全量库中已存在的机构，系统将自动剔除，仅为全新机构补齐等级并切分导入。"
                )
            with col_dd2:
                dedup_internal = st.checkbox(
                    "同时剔除上传表格内部的自身同名重复项",
                    value=True,
                    help="若上传的 Excel 内部有多行指向同一家医院，仅保留首行，后续行作为内部重复剔除。"
                )

        selected_dedup_mode = "49.4w"

        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            batch_split_size = st.slider(
                "单批次切分行数 (每批最大条数)",
                min_value=100,
                max_value=1000,
                value=300,
                step=50,
                help="老SaaS系统后台限制通常为 300 条/批。若无需切分可后续只下载全量表。"
            )
        with col_opt2:
            enable_online_chk = st.checkbox(
                "启用重点公立医院在线联网复核",
                value=True,
                help="若全国历史知识库缺失等级的公立医院，自动联网搜索核验并持久化学习"
            )

        # 操作按钮黄金居中排布（文字必须严格与模块名称完全一致！）
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        _, col_btn, _ = st.columns([1, 1.8, 1])
        with col_btn:
            start_saas = st.button(f"开始生成：{MODULE_SAAS}", type="primary", use_container_width=True)

        if start_saas:
            if not uploaded_saas_file:
                st.error("请先上传待处理的医院 Excel 表格后再点击开始生成。")
            else:
                with st.status(f"正在启动【{MODULE_SAAS}】智能处理引擎...", expanded=True) as status:
                    st.write("1. 正在准备沙箱隔离执行环境...")
                    temp_dir = tempfile.mkdtemp(prefix="saas_hosp_")
                    try:
                        in_file_path = os.path.join(temp_dir, uploaded_saas_file.name)
                        with open(in_file_path, "wb") as f:
                            f.write(uploaded_saas_file.getvalue())

                        st.write("2. 正在载入全国 49.4 万医疗机构知识库与双引擎索引...")
                        matcher = get_hospital_matcher()

                        st.write("3. 正在逐行进行多级智能定级、洗稿去重与标准下拉菜单规范化...")
                        prog_bar = st.progress(0, text="正在处理数据行...")

                        def update_progress(curr, total):
                            pct = min(curr / max(total, 1), 1.0)
                            prog_bar.progress(pct, text=f"已匹配完成 {curr}/{total} 条 ({(pct*100):.1f}%)")

                        out_file_name = f"已补充等级_{uploaded_saas_file.name}"
                        out_file_path = os.path.join(temp_dir, out_file_name)

                        res = matcher.process_excel(
                            input_excel=in_file_path,
                            output_excel=out_file_path,
                            split_size=batch_split_size if batch_split_size > 0 else None,
                            enable_online=enable_online_chk,
                            progress_callback=update_progress,
                            enable_dedup=enable_dedup,
                            dedup_mode=selected_dedup_mode,
                            dedup_internal=dedup_internal
                        )

                        if res:
                            dedup_stats = res.get("dedup_stats")
                            total_rows = res["total_rows"]
                            grade_counts = res["grade_counts"]
                            batch_files = res["batch_files"]
                            elapsed = res["elapsed"]

                            # 读取已剔除机构名单
                            excluded_bytes = None
                            df_excluded = None
                            if dedup_stats and dedup_stats.get("output_excluded_excel") and os.path.exists(dedup_stats["output_excluded_excel"]):
                                try:
                                    with open(dedup_stats["output_excluded_excel"], "rb") as ef:
                                        excluded_bytes = ef.read()
                                    df_excluded = pd.read_excel(io.BytesIO(excluded_bytes))
                                except Exception:
                                    pass

                            # 情况 A: 如果所有数据都被剔除了（无新数据导入）
                            if total_rows == 0:
                                prog_bar.progress(1.0, text="洗稿完成！")
                                status.update(label="洗稿完成：所有机构均已在系统 49.4 万在库中存在", state="complete")
                                st.info(f"💡 洗稿去重提示：您上传的 {dedup_stats['total_input']} 家机构已全部在系统 49.4 万在库医院中收录，未发现全新机构，无需生成导入批次。")

                                k1, k2, k3 = st.columns(3)
                                with k1:
                                    st.metric("原始上传总数", f"{dedup_stats['total_input']:,} 家")
                                with k2:
                                    st.metric("系统已有机构", f"{dedup_stats['excluded_kb_count']:,} 家", "49.4万库已在库", delta_color="inverse")
                                with k3:
                                    st.metric("表格自身重复", f"{dedup_stats['excluded_internal_count']:,} 家")

                                if excluded_bytes:
                                    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                                    _, col_dl_ex, _ = st.columns([1, 1.8, 1])
                                    with col_dl_ex:
                                        st.download_button(
                                            label=f"🚫 下载【已剔除系统已有机构名单.xlsx】 ({format_size(len(excluded_bytes))})",
                                            data=excluded_bytes,
                                            file_name=f"已剔除系统已有机构名单_{uploaded_saas_file.name}",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            use_container_width=True,
                                            type="primary"
                                        )
                                if df_excluded is not None and not df_excluded.empty:
                                    with st.expander(f"查看已剔除系统已有机构明细 (共 {len(df_excluded)} 家，展示前 100 行)", expanded=True):
                                        st.dataframe(df_excluded.head(100), use_container_width=True)

                            elif os.path.exists(out_file_path):
                                prog_bar.progress(1.0, text="处理完毕！")
                                status.update(label="洗稿过滤与医院等级补充全部完成！", state="complete")

                                with open(out_file_path, "rb") as f:
                                    filled_excel_bytes = f.read()

                                rated_count = sum(v for k, v in grade_counts.items() if k != '无等级')

                                if dedup_stats:
                                    st.success(f"🎉 成功完成洗稿去重与全新机构等级补充！原始上传 {dedup_stats['total_input']} 家，剔除系统已有/重复 {dedup_stats['total_excluded']} 家，保留并补充 {total_rows} 家全新待导入机构，耗时 {elapsed} 秒。")
                                    # 5 列 KPI 看板
                                    st.markdown("##### 核心洗稿与处理指标看板")
                                    k1, k2, k3, k4, k5 = st.columns(5)
                                    with k1:
                                        st.metric("原始上传行数", f"{dedup_stats['total_input']:,} 家")
                                    with k2:
                                        st.metric("剔除系统已有", f"{dedup_stats['excluded_kb_count']:,} 家", delta=f"-{dedup_stats['excluded_kb_count']} (49.4万库)", delta_color="inverse")
                                    with k3:
                                        st.metric("剔除自身重复", f"{dedup_stats['excluded_internal_count']:,} 家", delta=f"-{dedup_stats['excluded_internal_count']}", delta_color="inverse")
                                    with k4:
                                        st.metric("保留全新机构", f"{total_rows:,} 家", delta="待导入老SaaS")
                                    with k5:
                                        st.metric("切分导入批次", f"{len(batch_files)} 个批次包" if batch_files else "未切分")
                                else:
                                    st.success(f"🎉 成功完成 {total_rows} 家机构的等级智能识别与切分！耗时 {elapsed} 秒。")
                                    st.markdown("##### 核心处理指标看板")
                                    k1, k2, k3, k4 = st.columns(4)
                                    with k1:
                                        st.metric("处理机构总数", f"{total_rows:,} 家")
                                    with k2:
                                        st.metric("有等级机构数", f"{rated_count:,} 家", f"定级率 {rated_count/max(total_rows,1)*100:.1f}%")
                                    with k3:
                                        st.metric("切分批次数", f"{len(batch_files)} 个分批文件" if batch_files else "未切分")
                                    with k4:
                                        st.metric("引擎检索耗时", f"{elapsed} 秒")

                                # 等级分布展示
                                with st.expander("查看本次各等级评定分布统计", expanded=False):
                                    df_dist = pd.DataFrame([
                                        {"医院等级": k, "机构数量": v, "占比": f"{v/max(total_rows,1)*100:.2f}%"}
                                        for k, v in grade_counts.items() if v > 0
                                    ])
                                    st.dataframe(df_dist, use_container_width=True, hide_index=True)

                                # 打包 ZIP
                                zip_bytes = None
                                zip_name = f"老Saas医院导入_分批包_{len(batch_files)}批_{total_rows}条.zip"
                                if batch_files:
                                    zip_buffer = io.BytesIO()
                                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                        zf.write(out_file_path, arcname=out_file_name)
                                        for bf in batch_files:
                                            zf.write(bf, arcname=os.path.join(f"分批导入_每批{batch_split_size}条", os.path.basename(bf)))
                                        if excluded_bytes and dedup_stats and dedup_stats["total_excluded"] > 0:
                                            zf.write(dedup_stats["output_excluded_excel"], arcname="已剔除系统已有机构名单.xlsx")
                                    zip_bytes = zip_buffer.getvalue()

                                # 自动归档至历史记录
                                saas_files = {out_file_name: filled_excel_bytes}
                                if excluded_bytes and dedup_stats and dedup_stats.get("total_excluded", 0) > 0:
                                    saas_files[f"已剔除系统已有机构名单_{uploaded_saas_file.name}"] = excluded_bytes
                                if zip_bytes:
                                    saas_files[zip_name] = zip_bytes

                                if is_dedup_flow and dedup_stats:
                                    saas_summary = f"录入模板 {total_rows} 家 · 剔除系统已有 {dedup_stats['excluded_kb_count']} 家 · 自身去重 {dedup_stats['excluded_internal_count']} 家 · 分批 {len(batch_files)} 包"
                                else:
                                    saas_summary = f"录入模板 {total_rows} 家 · 有等级 {rated_count} 家 · 定级率 {rated_count/max(total_rows,1)*100:.1f}% · 分批 {len(batch_files)} 包"
                                history_manager.save_run(MODULE_SAAS, saas_files, summary=saas_summary)

                                # 居中下载按钮组
                                st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
                                if zip_bytes:
                                    _, col_dl_zip, _ = st.columns([1, 1.8, 1])
                                    with col_dl_zip:
                                        st.download_button(
                                            label=f"📦 下载全部切分批次完整压缩包 ({format_size(len(zip_bytes))})",
                                            data=zip_bytes,
                                            file_name=zip_name,
                                            mime="application/zip",
                                            use_container_width=True,
                                            type="primary"
                                        )

                                _, col_dl_excel, _ = st.columns([1, 1.8, 1])
                                with col_dl_excel:
                                    st.download_button(
                                        label=f"📄 下载洗稿后全新机构已定级表格 (.xlsx) ({format_size(len(filled_excel_bytes))})",
                                        data=filled_excel_bytes,
                                        file_name=out_file_name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )

                                if excluded_bytes and dedup_stats and dedup_stats["total_excluded"] > 0:
                                    _, col_dl_ex, _ = st.columns([1, 1.8, 1])
                                    with col_dl_ex:
                                        st.download_button(
                                            label=f"🚫 下载【已剔除系统已有机构名单.xlsx】 ({format_size(len(excluded_bytes))})",
                                            data=excluded_bytes,
                                            file_name=f"已剔除系统已有机构名单_{uploaded_saas_file.name}",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            use_container_width=True
                                        )

                                # 在线数据预览
                                st.markdown("##### 洗稿后全新机构数据预览 (前 100 行)")
                                try:
                                    df_preview = pd.read_excel(io.BytesIO(filled_excel_bytes), nrows=100)
                                    st.dataframe(df_preview, use_container_width=True)
                                except Exception as e:
                                    st.caption(f"预览加载失败: {e}")

                                if df_excluded is not None and not df_excluded.empty:
                                    with st.expander(f"查看已剔除系统已有机构名单 (共 {len(df_excluded)} 行，展示前 100 行)", expanded=False):
                                        st.dataframe(df_excluded.head(100), use_container_width=True)

                        else:
                            status.update(label="处理失败", state="error")
                            st.error("处理 Excel 表格失败，请检查文件首行是否包含「机构名称」或「医院名称」列。")

                    except Exception as e:
                        status.update(label="处理异常", state="error")
                        st.error(f"处理数据时发生异常: {str(e)}")
                    finally:
                        try:
                            shutil.rmtree(temp_dir)
                        except Exception:
                            pass


# =============================================================
# 模块六：北检&康恩贝结算表
# =============================================================
elif current_module == MODULE_BEIJIAN:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-badge">
                {SVG_BEIJIAN_SETTLE}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_BEIJIAN}</div>
                <div class="ios-hero-subtitle">智能关联待结算语料、专家银行资质与全量语料库，一键合规导出专家劳务报酬表、已结算语料明细表与作品结算总表。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            三表智能联动
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_BEIJIAN)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M9 15l2 2 4-4"/></svg>
                <span>源文件规范与处理说明</span>
            </div>
            <div class="bento-req-badge">3 张源表关联核算</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">待结算词条列表 (1.xlsx) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">提供本次结算的目标【语料词条编号】清单（如 133 条待结算语料）</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">专家信息与银行卡表 (2.xlsx) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">包含专家姓名、身份证号、银行卡号、开户银行及支行信息</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">03</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">语料库全量明细表 (3.xlsx) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">包含语料题目内容、回答记录、医院、科室、职称、审核状态等全量数据</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_files = st.file_uploader(
        "拖拽或批量选择上传全部源表格（支持一次性将 1.xlsx、2.xlsx、3.xlsx 多表同时拖入）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_beijian_files"
    )

    # 智能实时预识别嗅探
    file_1_obj, file_2_obj, file_3_obj = None, None, None
    if uploaded_files:
        # 第一轮：按文件名强特征优先匹配
        for uf in uploaded_files:
            fname = uf.name.lower()
            if fname in ['1.xlsx', '1.xls'] or ('1' in fname and '待结算' in fname) or ('待结算' in fname or '结算名单' in fname or '词条名单' in fname or '语料库项目' in fname):
                if not file_1_obj:
                    file_1_obj = uf
            elif fname in ['2.xlsx', '2.xls'] or ('2' in fname and ('专家' in fname or '银行' in fname)) or ('银行' in fname or '卡号' in fname or '专家' in fname or '资质' in fname or '用户' in fname):
                if not file_2_obj:
                    file_2_obj = uf
            elif fname in ['3.xlsx', '3.xls'] or ('3' in fname and '明细' in fname) or ('语料明细' in fname or '全量' in fname or '原始明细' in fname):
                if not file_3_obj:
                    file_3_obj = uf

        # 第二轮：若仍有未匹配表格，深入读取前 2 行表头内容智能识别
        for uf in uploaded_files:
            if uf in [file_1_obj, file_2_obj, file_3_obj]:
                continue
            try:
                df_peek = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=2)
                h_str = "".join([str(c) for c in df_peek.columns])
                if not file_2_obj and ("银行卡" in h_str or "开户行" in h_str or "支行" in h_str):
                    file_2_obj = uf
                elif not file_3_obj and ("题目内容" in h_str or "回答记录" in h_str or "首次提交时间" in h_str):
                    file_3_obj = uf
                elif not file_1_obj and ("是否结算" in h_str or "结算单价" in h_str or "语料词条编号" in h_str):
                    file_1_obj = uf
            except Exception:
                pass

        # 渲染识别状态指示条
        render_html('<div class="ios-precheck-box"><b>智能多表嗅探识别状态：</b><br>')
        chk_cols = st.columns(3)
        with chk_cols[0]:
            if file_1_obj:
                render_html(f'<span class="ios-badge-success">1. 待结算词条列表：已锁定</span><br><small style="opacity:0.8;">{file_1_obj.name}</small>', container=chk_cols[0])
            else:
                render_html('<span class="ios-badge-pending">待识别：1. 待结算词条列表 (1.xlsx)</span>', container=chk_cols[0])
        with chk_cols[1]:
            if file_2_obj:
                render_html(f'<span class="ios-badge-success">2. 专家银行信息表：已锁定</span><br><small style="opacity:0.8;">{file_2_obj.name}</small>', container=chk_cols[1])
            else:
                render_html('<span class="ios-badge-pending">待识别：2. 专家银行卡信息表 (2.xlsx)</span>', container=chk_cols[1])
        with chk_cols[2]:
            if file_3_obj:
                render_html(f'<span class="ios-badge-success">3. 语料全量明细表：已锁定</span><br><small style="opacity:0.8;">{file_3_obj.name}</small>', container=chk_cols[2])
            else:
                render_html('<span class="ios-badge-pending">待识别：3. 语料库全量明细表 (3.xlsx)</span>', container=chk_cols[2])

    render_html("""
    <div style="margin: 12px 0 16px 0; padding: 12px 18px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; font-size: 13px; color: #166534; display: flex; align-items: center; justify-content: space-between;">
        <div style="display: flex; align-items: center; gap: 8px;">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
            <span><b>自动智能核算</b>：系统将直接从上传的源文件中自动抓取各篇语料的<b>【结算单价】</b>并按待结算目标自动对齐，无需手动配置参数。</span>
        </div>
        <div style="font-size: 12px; color: #15803d; opacity: 0.85;">纯 Python 原生渲染标准交付表格</div>
    </div>
    """)

    # 开始生成大按钮（与侧边栏标题 100% 一致）
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    btn_start_beijian = st.button(
        f"开始生成：{MODULE_BEIJIAN}",
        type="primary",
        use_container_width=True
    )

    if btn_start_beijian:
        if not file_1_obj or not file_2_obj or not file_3_obj:
            st.error("未识别齐全部 3 张关键表格！请在上方上传区拖入或选择包含【1. 待结算词条列表】、【2. 专家银行卡信息表】及【3. 语料库全量明细表】的文件。")
        else:
            with st.status(f"正在沙盒环境中执行【{MODULE_BEIJIAN}】核心结算逻辑...", expanded=True) as status:
                st.write("1. 正在初始化沙盒与暂存上传文件...")
                temp_dir = tempfile.mkdtemp()
                try:
                    f1_path = os.path.join(temp_dir, file_1_obj.name)
                    with open(f1_path, "wb") as f:
                        f.write(file_1_obj.getvalue())

                    f2_path = os.path.join(temp_dir, file_2_obj.name)
                    with open(f2_path, "wb") as f:
                        f.write(file_2_obj.getvalue())

                    f3_path = os.path.join(temp_dir, file_3_obj.name)
                    with open(f3_path, "wb") as f:
                        f.write(file_3_obj.getvalue())

                    st.write("2. 调度北检&康恩贝三表智能核算引擎...")
                    import importlib
                    beijian_engine = importlib.import_module("北检&康恩贝结算表.generate_settlement")

                    logs_list = []
                    def log_collector(msg):
                        clean_msg = str(msg).strip()
                        logs_list.append(clean_msg)
                        st.write(clean_msg)

                    res = beijian_engine.process_beijian_kangbei(
                        file_1=f1_path,
                        file_2=f2_path,
                        file_3=f3_path,
                        output_dir=temp_dir,
                        include_all_133=True,
                        log_func=log_collector
                    )

                    if res.get("success"):
                        status.update(label=f"【{MODULE_BEIJIAN}】处理完成！", state="complete")
                        st.success("对账与结算明细表单生成完毕！")

                        # KPI 指标展示
                        st.markdown("##### 结算核对核心指标看板")
                        k1, k2, k3, k4 = st.columns(4)
                        with k1:
                            st.metric("本次结算词条数", f"{res['total_items']} 篇")
                        with k2:
                            st.metric("本次结算专家数", f"{res['total_doctors']} 位")
                        with k3:
                            st.metric("实发劳务总金额", f"¥ {res['total_net']:,.2f}")
                        with k4:
                            st.metric("应发税前总额", f"¥ {res['total_gross']:,.2f}", delta=f"-¥ {res['total_tax']:,.2f} 税金")

                        # 读取生成文件
                        files_map = res.get("files", {})
                        f7_path = files_map.get("7_专家劳务报酬明细表.xlsx")
                        f8_path = files_map.get("8_已结算语料词条明细表.xlsx")
                        f9_path = files_map.get("9_作品劳务结算总表.xlsx")

                        f7_bytes = open(f7_path, "rb").read() if f7_path and os.path.exists(f7_path) else None
                        f8_bytes = open(f8_path, "rb").read() if f8_path and os.path.exists(f8_path) else None
                        f9_bytes = open(f9_path, "rb").read() if f9_path and os.path.exists(f9_path) else None

                        # 打包 ZIP
                        all_outputs = {}
                        if f7_bytes:
                            all_outputs["7_专家劳务报酬明细表.xlsx"] = f7_bytes
                        if f8_bytes:
                            all_outputs["8_已结算语料词条明细表.xlsx"] = f8_bytes
                        if f9_bytes:
                            all_outputs["9_作品劳务结算总表.xlsx"] = f9_bytes

                        zip_bytes = create_zip_archive(all_outputs)

                        # 自动归档至历史记录
                        beijian_hist_files = dict(all_outputs)
                        if zip_bytes:
                            beijian_hist_files[f"北检康恩贝结算全套表单_{res['total_items']}条_{res['total_doctors']}人.zip"] = zip_bytes
                        beijian_summary = f"结算专家 {res.get('total_doctors', 0)} 位 · 语料词条 {res.get('total_items', 0)} 篇 · 实发总额 ¥ {res.get('total_net', 0):,.2f} · 税前 ¥ {res.get('total_gross', 0):,.2f}"
                        history_manager.save_run(MODULE_BEIJIAN, beijian_hist_files, summary=beijian_summary)

                        # 下载按钮组
                        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
                        _, col_dl_zip, _ = st.columns([1, 1.8, 1])
                        with col_dl_zip:
                            st.download_button(
                                label=f"📦 一键打包下载全部结算表单 (ZIP) ({format_size(len(zip_bytes))})",
                                data=zip_bytes,
                                file_name=f"北检康恩贝结算全套表单_{res['total_items']}条_{res['total_doctors']}人.zip",
                                mime="application/zip",
                                use_container_width=True,
                                type="primary"
                            )

                        cd1, cd2, cd3 = st.columns(3)
                        with cd1:
                            if f7_bytes:
                                st.download_button(
                                    label=f"📄 下载【7. 专家劳务报酬明细表】 ({format_size(len(f7_bytes))})",
                                    data=f7_bytes,
                                    file_name="7_专家劳务报酬明细表.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        with cd2:
                            if f8_bytes:
                                st.download_button(
                                    label=f"📄 下载【8. 语料词条明细表】 ({format_size(len(f8_bytes))})",
                                    data=f8_bytes,
                                    file_name="8_已结算语料词条明细表.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        with cd3:
                            if f9_bytes:
                                st.download_button(
                                    label=f"📄 下载【9. 项目作品结算表】 ({format_size(len(f9_bytes))})",
                                    data=f9_bytes,
                                    file_name="9_作品劳务结算总表.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )

                        # 数据预览 Tabs
                        st.markdown("##### 生成结果数据在线预览")
                        tab7, tab8, tab9 = st.tabs(["7. 专家劳务报酬明细表", "8. 语料词条明细表", "9. 项目作品结算表"])
                        with tab7:
                            if f7_bytes:
                                try:
                                    df7_prev = pd.read_excel(io.BytesIO(f7_bytes), skiprows=3)
                                    st.dataframe(df7_prev.head(100), use_container_width=True)
                                except Exception as e:
                                    st.caption(f"预览加载失败: {e}")
                        with tab8:
                            if f8_bytes:
                                try:
                                    df8_prev = pd.read_excel(io.BytesIO(f8_bytes))
                                    st.dataframe(df8_prev.head(100), use_container_width=True)
                                except Exception as e:
                                    st.caption(f"预览加载失败: {e}")
                        with tab9:
                            if f9_bytes:
                                try:
                                    df9_prev = pd.read_excel(io.BytesIO(f9_bytes), skiprows=3)
                                    st.dataframe(df9_prev.head(100), use_container_width=True)
                                except Exception as e:
                                    st.caption(f"预览加载失败: {e}")

                    else:
                        status.update(label="处理失败", state="error")
                        st.error("执行过程出现错误，请检查输入表格格式。")

                except Exception as e:
                    status.update(label="处理异常", state="error")
                    st.error(f"处理数据时发生异常: {str(e)}")
                finally:
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception:
                        pass


# =============================================================
# 模块七：北检&华东科普视频电签表
# =============================================================
elif current_module == MODULE_KOPU:
    render_html(f"""
    <div class="ios-hero-banner">
        <div class="ios-hero-left">
            <div class="ios-hero-icon-box">
                {SVG_KOPU_VIDEO}
            </div>
            <div>
                <div class="ios-hero-title">{MODULE_KOPU}</div>
                <div class="ios-hero-subtitle">智能关联科普视频任务编号、视频明细与专家银行信息，严格按照第三方系统规范生成纯净版电签表。</div>
            </div>
        </div>
        <div class="ios-hero-pill">
            <span class="ios-hero-pill-dot"></span>
            Python 电签结算内核
        </div>
    </div>
    """)

    # 历史归档抽屉 (位于顶部横幅下方)
    history_manager.render_history_ui(MODULE_KOPU)

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M12 8v8"/><path d="M8 12h8"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">3 表全自动智能嗅探勾稽</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">待结算任务编号表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「任务编号」或「任务」，包含待结算任务的【任务明细编号】清单</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">科普视频明细表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「明细」或「视频」，包含【作品编号】、【积分】、【用户名称】、【身份证号】</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">03</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">专家用户信息表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「用户信息」或「用户」，包含【银行卡号】、【开户行名称】、【所在医院】、【职称】</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_kopu_files = st.file_uploader(
        "拖拽或批量选择上传源表格 (.xlsx / .xls，支持同时选择上传 3 个文件)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_kopu"
    )

    # 智能实时预识别嗅探
    task_file_obj, detail_file_obj, user_file_obj = None, None, None
    if uploaded_kopu_files:
        # 第一轮：按文件名强特征优先匹配
        for uf in uploaded_kopu_files:
            fname = uf.name.lower()
            if any(k in fname for k in ['任务编号', '任务', 'task']):
                if not task_file_obj:
                    task_file_obj = uf
            elif any(k in fname for k in ['明细表', '明细', '作品', 'detail']):
                if not detail_file_obj:
                    detail_file_obj = uf
            elif any(k in fname for k in ['用户信息', '用户', '专家', 'user', 'info']):
                if not user_file_obj:
                    user_file_obj = uf

        # 第二轮：表头特征智能探测
        for uf in uploaded_kopu_files:
            if uf in [task_file_obj, detail_file_obj, user_file_obj]:
                continue
            try:
                df_peek = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=2)
                h_str = "".join([str(c) for c in df_peek.columns])
                if not task_file_obj and ("任务明细编号" in h_str or "任务编号" in h_str):
                    task_file_obj = uf
                elif not detail_file_obj and ("作品编号" in h_str or "科普课件" in h_str or "课件链接" in h_str or "作品编码" in h_str):
                    detail_file_obj = uf
                elif not user_file_obj and ("银行卡号" in h_str or "开户行" in h_str or "支行名称" in h_str or "电签银行卡号" in h_str):
                    user_file_obj = uf
            except Exception:
                pass

        # 渲染识别状态指示条（与规范样式完全一致的药丸卡片）
        render_html('<div class="ios-precheck-box"><b>智能多表嗅探识别状态：</b><br>')
        chk_cols = st.columns(3)
        with chk_cols[0]:
            if task_file_obj:
                render_html(f'<span class="ios-badge-success">1. 待结算任务编号表：已锁定</span><br><small style="opacity:0.8;">{task_file_obj.name}</small>', container=chk_cols[0])
            else:
                render_html('<span class="ios-badge-pending">待识别：1. 待结算任务编号表 (任务/编号)</span>', container=chk_cols[0])
        with chk_cols[1]:
            if detail_file_obj:
                render_html(f'<span class="ios-badge-success">2. 科普视频明细表：已锁定</span><br><small style="opacity:0.8;">{detail_file_obj.name}</small>', container=chk_cols[1])
            else:
                render_html('<span class="ios-badge-pending">待识别：2. 科普视频明细表 (明细/视频)</span>', container=chk_cols[1])
        with chk_cols[2]:
            if user_file_obj:
                render_html(f'<span class="ios-badge-success">3. 专家用户信息表：已锁定</span><br><small style="opacity:0.8;">{user_file_obj.name}</small>', container=chk_cols[2])
            else:
                render_html('<span class="ios-badge-pending">待识别：3. 专家用户信息表 (用户/专家)</span>', container=chk_cols[2])

        render_html("""
        <div style="margin: 12px 0 16px 0; padding: 12px 18px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; font-size: 13px; color: #166534; display: flex; align-items: center; justify-content: space-between;">
            <div style="display: flex; align-items: center; gap: 8px;">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                <span><b>自动智能核算</b>：系统将直接以【任务明细编号】为准，精准勾稽视频明细与专家银行信息，输出零样式纯净版《科普电签表.xlsx》。</span>
            </div>
            <div style="font-size: 12px; color: #15803d; opacity: 0.85;">纯 Python 原生渲染标准导入表格</div>
        </div>
        """)

    # 黄金比例操作按钮
    _, col_btn, _ = st.columns([1, 1.8, 1])
    with col_btn:
        start_kopu = st.button(f"开始生成：{MODULE_KOPU}", type="primary", use_container_width=True)

    if start_kopu:
        if not (task_file_obj and detail_file_obj and user_file_obj):
            st.error("未识别齐全 3 张关键表格！请确认已上传：1. 任务编号表、2. 视频明细表、3. 用户信息表。")
        else:
            with st.status("正在启动北检&华东科普视频电签核算引擎...", expanded=True) as status:
                st.write("1. 正在初始化沙箱运行隔离环境...")
                temp_dir = tempfile.mkdtemp(prefix="kopu_sign_")
                try:
                    t_path = os.path.join(temp_dir, task_file_obj.name)
                    d_path = os.path.join(temp_dir, detail_file_obj.name)
                    u_path = os.path.join(temp_dir, user_file_obj.name)
                    with open(t_path, "wb") as f: f.write(task_file_obj.getvalue())
                    with open(d_path, "wb") as f: f.write(detail_file_obj.getvalue())
                    with open(u_path, "wb") as f: f.write(user_file_obj.getvalue())

                    out_path = os.path.join(temp_dir, "科普电签表.xlsx")

                    st.write("2. 正在执行三方勾稽对账、积分汇总、智能省市匹配与纯净表单生成...")
                    sys.path.append(os.path.join(ROOT_DIR, "北检&华东科普视频电签表"))
                    from generate_sign_table import generate_kopu_sign_workbook
                    
                    res_path, stats = generate_kopu_sign_workbook(t_path, d_path, u_path, out_path)

                    if os.path.exists(out_path):
                        status.update(label="科普电签表生成完成！", state="complete")
                        st.success(f"成功生成纯净版《科普电签表.xlsx》！共汇总 {stats['total_doctors']} 位专家、{stats['total_tasks']} 条视频，劳务总额 ¥ {stats['total_amount']:,} 元。")

                        # 核心 KPI 看板
                        st.markdown("##### 本期结算核心 KPI 看板")
                        k1, k2, k3, k4 = st.columns(4)
                        with k1:
                            st.metric("结算总人数", f"{stats['total_doctors']} 位专家")
                        with k2:
                            st.metric("结算任务数", f"{stats['total_tasks']} 条视频")
                        with k3:
                            st.metric("劳务总金额", f"¥ {stats['total_amount']:,} 元")
                        with k4:
                            st.metric("信息匹配率", f"{stats['matched_rate']}")

                        # 下载专区
                        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                        with open(out_path, "rb") as ef:
                            out_bytes = ef.read()

                        # 自动归档至历史记录
                        kopu_summary = f"{stats['total_doctors']} 位专家 · {stats['total_tasks']} 条视频 · 劳务总额 ¥ {stats['total_amount']:,} 元 · 匹配率 {stats['matched_rate']}"
                        history_manager.save_run(MODULE_KOPU, {"科普电签表.xlsx": out_bytes}, summary=kopu_summary)

                        _, col_dl, _ = st.columns([1, 1.8, 1])
                        with col_dl:
                            st.download_button(
                                label=f"⬇️ 一键下载【科普电签表.xlsx】 ({format_size(len(out_bytes))})",
                                data=out_bytes,
                                file_name="科普电签表.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )

                        # 在线数据预览
                        st.markdown("##### 📑 科普电签表数据在线预览")
                        df_preview = pd.read_excel(io.BytesIO(out_bytes))
                        st.caption(f"共 {len(df_preview)} 行数据（展示前 100 行）：")
                        st.dataframe(df_preview.head(100), use_container_width=True)

                    else:
                        status.update(label="生成失败", state="error")
                        st.error("执行过程出现错误，未能在沙箱中生成目标表格。")

                except Exception as e:
                    status.update(label="处理异常", state="error")
                    st.error(f"处理数据时发生异常: {str(e)}")
                finally:
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception:
                        pass
