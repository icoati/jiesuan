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


if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

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

MODULE_OPTIONS = [
    MODULE_SHANGYAO,
    MODULE_CORPUS,
    MODULE_ZHENGHE,
    MODULE_JUMEI
]

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
    render_html('<span class="ios-badge-success">7天免密保护中</span>', container=st.sidebar)
    render_html('<span class="ios-badge-success">4大业务模块就绪</span>', container=st.sidebar)


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

    _, col_btn, _ = st.columns([1, 1.4, 1])
    with col_btn:
        ready_to_run = bool(has_src and has_staff)
        run_btn = st.button(
            "开始生成报表",
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

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="20" height="14" rx="3"/><line x1="2" y1="10" x2="22" y2="10"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">列序自适应 · 防格式失真</div>
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

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        file_yl = st.file_uploader(
            "1. 结算明细文件 (*必选)",
            type=["xlsx", "xls"],
            key="upload_yl"
        )
    with col_f2:
        file_task = st.file_uploader(
            "2. 辅助任务明细/确认单 (可选)",
            type=["xlsx", "xls"],
            key="upload_task"
        )

    if file_yl:
        try:
            xl_yl = pd.ExcelFile(io.BytesIO(file_yl.getvalue()))
            render_html(f'<span class="ios-badge-success">结算主明细已加载（包含工作表：{", ".join(xl_yl.sheet_names[:3])}）</span>')
        except Exception:
            pass

    _, col_btn, _ = st.columns([1, 1.4, 1])
    with col_btn:
        run_btn = st.button("一键生成结算表格", type="primary", use_container_width=True)

    if run_btn:
        if not file_yl:
            st.error("请先上传【1. 结算明细文件】")
        else:
            with st.status("正在沙盒中执行北京整合-上药雷允上结算逻辑...", expanded=True) as status:
                st.write("1. 正在初始化沙盒...")
                with tempfile.TemporaryDirectory() as temp_dir:
                    yl_path = os.path.join(temp_dir, file_yl.name)
                    with open(yl_path, "wb") as f:
                        f.write(file_yl.getvalue())

                    task_path = None
                    if file_task:
                        task_path = os.path.join(temp_dir, file_task.name)
                        with open(task_path, "wb") as f:
                            f.write(file_task.getvalue())

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
                    for p in Path(temp_dir).glob("*.xlsx"):
                        if p.name not in [file_yl.name, getattr(file_task, 'name', '')]:
                            with open(p, "rb") as f:
                                gen_files[p.name] = f.read()

                    if proc.returncode == 0 and gen_files:
                        status.update(label="结算表格生成完毕", state="complete")
                        st.success(f"成功完成对账结算，共生成 {len(gen_files)} 个结算报表")

                        doc_count_match = re.search(r'汇总完成[^\d]*(\d+)[^\d]*位医生', proc.stdout)
                        amount_match = re.search(r'劳务实发总额[^\d]*([0-9\.,]+)', proc.stdout)

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

    _, col_btn, _ = st.columns([1, 1.4, 1])
    with col_btn:
        ready_corpus = bool(uploaded_files and len(uploaded_files) >= 3)
        run_btn = st.button(
            "开始生成每月结算发放表",
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

    render_html("""
    <div class="bento-req-container">
        <div class="bento-req-header">
            <div class="bento-req-title">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0284c7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M12 8v8"/><path d="M8 12h8"/></svg>
                <span>源文件规范与自动识别要求</span>
            </div>
            <div class="bento-req-badge">3 表特征智能嗅探 · 自动对齐</div>
        </div>
        <div class="bento-req-grid">
            <div class="bento-card">
                <div class="bento-card-num">01</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">项目进度表 <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「进度」或「雷允上」，包含语料编号、医生姓名、项目名称、结算单价等</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">02</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">语料列表 (明文) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「语料」，包含语料词条编号、手机号、身份证号等业务明细</div>
                </div>
            </div>
            <div class="bento-card">
                <div class="bento-card-num">03</div>
                <div class="bento-card-content">
                    <div class="bento-card-title">用户列表 (明文) <span class="req-tag-must">必须</span></div>
                    <div class="bento-card-desc">文件名含「用户」，包含身份证号、开户银行、支行、银行卡号</div>
                </div>
            </div>
        </div>
    </div>
    """)

    uploaded_jumei_files = st.file_uploader(
        "拖拽或批量选择上传 3 个源表格 (.xlsx / .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="upload_jumei"
    )

    # 智能预识别嗅探
    prog_name, corpus_name, user_name = None, None, None
    if uploaded_jumei_files:
        for uf in uploaded_jumei_files:
            fname = uf.name.lower()
            if ("进度" in fname or "雷允上" in fname) and ("明文" not in fname and "费用" not in fname and "对账" not in fname and "备份" not in fname and "用户" not in fname):
                prog_name = uf.name
            elif ("语料" in fname) and ("进度" not in fname and "费用" not in fname and "用户" not in fname):
                corpus_name = uf.name
            elif ("用户" in fname or "医生" in fname) and ("进度" not in fname and "语料" not in fname):
                user_name = uf.name

        # 内容兜底启发式探测
        for uf in uploaded_jumei_files:
            if uf.name in [prog_name, corpus_name, user_name]:
                continue
            try:
                df_head = pd.read_excel(io.BytesIO(uf.getvalue()), nrows=2)
                h_str = "".join([str(c) for c in df_head.columns])
                if not prog_name and ("单价" in h_str or "结算单价" in h_str) and ("项目" in h_str or "医院" in h_str):
                    prog_name = uf.name
                elif not corpus_name and ("语料" in h_str or "词条" in h_str) and ("题目" in h_str or "手机" in h_str):
                    corpus_name = uf.name
                elif not user_name and ("开户" in h_str or "支行" in h_str or "银行卡" in h_str or "卡号" in h_str):
                    user_name = uf.name
            except Exception:
                pass

        st.markdown("##### 实时文件嗅探匹配结果")
        c1, c2, c3 = st.columns(3)
        with c1:
            if prog_name:
                render_html(f'<div class="ios-status-card success"><div class="ios-status-card-title">01 项目进度表</div><div class="ios-status-card-val">{prog_name}</div></div>')
            else:
                render_html('<div class="ios-status-card warning"><div class="ios-status-card-title">01 项目进度表</div><div class="ios-status-card-val">未识别 (需包含“进度”)</div></div>')
        with c2:
            if corpus_name:
                render_html(f'<div class="ios-status-card success"><div class="ios-status-card-title">02 语料列表(明文)</div><div class="ios-status-card-val">{corpus_name}</div></div>')
            else:
                render_html('<div class="ios-status-card warning"><div class="ios-status-card-title">02 语料列表(明文)</div><div class="ios-status-card-val">未识别 (需包含“语料”)</div></div>')
        with c3:
            if user_name:
                render_html(f'<div class="ios-status-card success"><div class="ios-status-card-title">03 用户列表(明文)</div><div class="ios-status-card-val">{user_name}</div></div>')
            else:
                render_html('<div class="ios-status-card warning"><div class="ios-status-card-title">03 用户列表(明文)</div><div class="ios-status-card-val">未识别 (需包含“用户”)</div></div>')

    # 操作按钮黄金居中排布
    _, col_btn, _ = st.columns([1, 1.4, 1])
    with col_btn:
        start_jumei = st.button("开始生成结算包", type="primary", use_container_width=True)

    if start_jumei:
        if not uploaded_jumei_files or len(uploaded_jumei_files) < 3:
            st.error("请上传全部 3 个源数据文件（项目进度表、语料明文表、用户明文表）后再次点击生成。")
        elif not (prog_name and corpus_name and user_name):
            st.warning("系统未能自动匹配全部 3 张必要表格，请确认文件名分别包含「进度/雷允上」、「语料」、「用户」关键字。")
        else:
            with st.status("正在启动陈菊梅基金会雷允上劳务结算引擎...", expanded=True) as status:
                st.write("1. 正在初始化沙箱运行隔离环境...")
                temp_dir = tempfile.mkdtemp(prefix="jumei_settle_")
                try:
                    prog_file_obj = next(f for f in uploaded_jumei_files if f.name == prog_name)
                    corpus_file_obj = next(f for f in uploaded_jumei_files if f.name == corpus_name)
                    user_file_obj = next(f for f in uploaded_jumei_files if f.name == user_name)

                    p_path = os.path.join(temp_dir, prog_file_obj.name)
                    c_path = os.path.join(temp_dir, corpus_file_obj.name)
                    u_path = os.path.join(temp_dir, user_file_obj.name)

                    with open(p_path, "wb") as f:
                        f.write(prog_file_obj.getvalue())
                    with open(c_path, "wb") as f:
                        f.write(corpus_file_obj.getvalue())
                    with open(u_path, "wb") as f:
                        f.write(user_file_obj.getvalue())

                    today_str = datetime.datetime.now().strftime("%Y%m%d")
                    out_name = f"{today_str}-劳务费用明细表.xlsx"
                    out_path = os.path.join(temp_dir, out_name)

                    script_path = os.path.join(ROOT_DIR, "陈菊梅基金会-雷允上结算包", "generate_settlement.py")
                    st.write("2. 正在执行多表模糊映射、劳务个税反算与分项目汇总...")

                    cmd = [
                        sys.executable,
                        script_path,
                        "--prog", p_path,
                        "--corpus", c_path,
                        "--user", u_path,
                        "--output", out_path
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

                    if proc.returncode == 0 and os.path.exists(out_path):
                        status.update(label="结算包生成完成！", state="complete")
                        with open(out_path, "rb") as f:
                            excel_bytes = f.read()

                        st.success("成功生成劳务费用明细表，包含汇总表、各项目明细与语料对账总表！")

                        # 提取 KPI 数据
                        total_items = re.search(r'共\s*(\d+)\s*条', proc.stdout)
                        total_amt = re.search(r'税后总金额:\s*([0-9\.,]+)\s*元', proc.stdout)
                        proj_matches = re.findall(r'项目【(.*?)】:\s*(\d+)\s*位医生，金额合计:\s*([0-9\.,]+)\s*元', proc.stdout)

                        total_docs = sum(int(m[1]) for m in proj_matches) if proj_matches else None

                        st.markdown("##### 本期结算核心 KPI 看板")
                        k1, k2, k3, k4 = st.columns(4)
                        with k1:
                            st.metric("结算总人数", f"{total_docs} 位医生" if total_docs is not None else "已核算")
                        with k2:
                            st.metric("语料词条数", f"{total_items.group(1)} 条" if total_items else "已核算")
                        with k3:
                            st.metric("税后总金额", f"¥ {total_amt.group(1)}" if total_amt else "已核算")
                        with k4:
                            st.metric("涉及项目数", f"{len(proj_matches)} 个项目" if proj_matches else "已核算")

                        # 居中下载大按钮
                        st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
                        _, col_dl, _ = st.columns([1, 1.6, 1])
                        with col_dl:
                            st.download_button(
                                label=f"下载【{out_name}】 ({format_size(len(excel_bytes))})",
                                data=excel_bytes,
                                file_name=out_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                        # 数据多 Sheet 在线预览
                        st.markdown("##### 报表工作表 (Sheet) 在线预览")
                        try:
                            xl_file = pd.ExcelFile(io.BytesIO(excel_bytes))
                            sheet_tabs = st.tabs(xl_file.sheet_names)
                            for idx, sname in enumerate(xl_file.sheet_names):
                                with sheet_tabs[idx]:
                                    df_sheet = pd.read_excel(xl_file, sheet_name=sname)
                                    st.caption(f"工作表【{sname}】共 {len(df_sheet)} 行数据（展示前 100 行）：")
                                    st.dataframe(df_sheet.head(100), use_container_width=True)
                        except Exception as e:
                            st.info("数据预览生成完毕，您可以直接点击上方按钮下载 Excel 文件查看完整格式与公式。")

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

