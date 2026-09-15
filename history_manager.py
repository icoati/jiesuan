# -*- coding: utf-8 -*-
"""
======================================================================
医疗数据统计与劳务结算云平台 - 历史生成结果管理器 (History Manager)
特性：
1. 零外部数据库依赖：纯文件系统结构化存储，跨平台 (Windows / VPS Linux) 100% 兼容
2. 大容量安全归档：默认每个模块保留最近 100 次生成结果，空间占用极小 (< 500MB)
3. 自动滚动淘汰：超额时按时间戳自动清理最早的历史记录目录，保护磁盘空间
4. 完备元数据快照：包含生成时间戳、业务 KPI 摘要、文件列表及大小
5. 优雅 UI 组件：折叠抽屉、核心指标卡片、一键回溯下载、在线快速预览
======================================================================
"""

import os
import io
import re
import json
import time
import shutil
import zipfile
import datetime
from pathlib import Path
import pandas as pd
import streamlit as st

# 历史记录根目录
ROOT_DIR = os.path.abspath(os.path.dirname(__file__))
HISTORY_BASE_DIR = os.path.join(ROOT_DIR, "history_records")

# 每个模块最大保留记录数 (VPS 40G 空间充裕，默认保留 200 条)
MAX_HISTORY_PER_MODULE = 200

# 强制统一为中国北京时间 (Asia/Shanghai, UTC+8)
# 在 Linux / VPS 上自动同步系统时区，彻底消除 VPS 默认为 UTC 导致的时间慢 8 小时问题
if hasattr(time, 'tzset'):
    try:
        os.environ['TZ'] = 'Asia/Shanghai'
        time.tzset()
    except Exception:
        pass

BEIJING_TZ = datetime.timezone(datetime.timedelta(hours=8))

def get_beijing_now() -> datetime.datetime:
    """获取标准的中国北京时间 (UTC+8)"""
    return datetime.datetime.now(BEIJING_TZ)

def format_record_timestamp(rec: dict) -> str:
    """
    统一将历史记录时间格式化为中国北京时间 (UTC+8)
    自动识别旧版在 UTC VPS 环境下生成的无时区标记记录，智能换算 +8 小时
    """
    ts_str = rec.get("timestamp", "")
    if not ts_str or ts_str == "未知时间":
        return "未知时间"
    # 已标记为 UTC+8 的新记录，直接使用
    if rec.get("timezone") in ["UTC+8", "UTC+8 (北京时间)"]:
        return ts_str
    # 兼容处理未标记时区的历史记录（若在 UTC VPS 上生成，自动换算为北京时间）
    try:
        dt = datetime.datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
        return (dt + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ts_str


def _format_size(num_bytes: int) -> str:
    """人性化文件大小格式化"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} TB"


def _sanitize_name(name: str) -> str:
    """清理模块名称，确保其在 Windows / Linux 作为目录名合法"""
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()


def _ensure_base_dir():
    """确保历史记录根目录存在"""
    if not os.path.exists(HISTORY_BASE_DIR):
        os.makedirs(HISTORY_BASE_DIR, exist_ok=True)


def get_module_history_dir(module_name: str) -> str:
    """获取指定模块的历史存储目录路径"""
    _ensure_base_dir()
    safe_name = _sanitize_name(module_name)
    m_dir = os.path.join(HISTORY_BASE_DIR, safe_name)
    if not os.path.exists(m_dir):
        os.makedirs(m_dir, exist_ok=True)
    return m_dir


def save_run(module_name: str, files_dict: dict, summary: str = "") -> str:
    """
    归档保存一次成功的生成结果
    :param module_name: 业务模块名称
    :param files_dict: 文件名字典 { "文件名.xlsx": bytes_data, ... }
    :param summary: 本次结算的核心 KPI 摘要文本
    :return: 归档目录路径
    """
    if not files_dict:
        return ""

    m_dir = get_module_history_dir(module_name)
    now = get_beijing_now()
    # 采用 时间戳 + 微秒四位 保证严格唯一且天然时间序
    run_id = f"{now.strftime('%Y%m%d_%H%M%S')}_{int(time.time() * 1000) % 10000:04d}"
    run_dir = os.path.join(m_dir, run_id)
    os.makedirs(run_dir, exist_ok=True)

    file_meta_list = []

    # 1. 写入各个生成的文件
    for fname, fbytes in files_dict.items():
        if not fbytes:
            continue
        clean_fname = os.path.basename(fname)
        file_path = os.path.join(run_dir, clean_fname)
        with open(file_path, "wb") as f:
            f.write(fbytes)
        file_meta_list.append({
            "name": clean_fname,
            "size": len(fbytes),
            "size_str": _format_size(len(fbytes))
        })

    # 2. 如果包含多个文件且其中没有包含全量 zip，则自动额外打包一个整合 zip 供一键下载
    has_zip = any(item["name"].lower().endswith(".zip") for item in file_meta_list)
    if len(file_meta_list) > 1 and not has_zip:
        zip_name = "全部生成结果完整包.zip"
        zip_path = os.path.join(run_dir, zip_name)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in file_meta_list:
                fp = os.path.join(run_dir, item["name"])
                zf.write(fp, arcname=item["name"])
        zip_size = os.path.getsize(zip_path)
        file_meta_list.append({
            "name": zip_name,
            "size": zip_size,
            "size_str": _format_size(zip_size),
            "is_auto_zip": True
        })

    # 3. 写入元数据 meta.json (显式记录中国北京时间与时区)
    meta = {
        "run_id": run_id,
        "module": module_name,
        "timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
        "timezone": "UTC+8 (北京时间)",
        "summary": summary.strip(),
        "files": file_meta_list
    }
    with open(os.path.join(run_dir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    # 4. 执行自动滚动淘汰机制 (超过 MAX_HISTORY_PER_MODULE 时删除最早的记录)
    _prune_history(m_dir, max_keep=MAX_HISTORY_PER_MODULE)

    return run_dir


def _prune_history(module_dir: str, max_keep: int = MAX_HISTORY_PER_MODULE):
    """自动清理超出保留上限的最早历史记录"""
    try:
        entries = [
            d for d in os.listdir(module_dir)
            if os.path.isdir(os.path.join(module_dir, d))
        ]
        # 按目录名排序（因采用 YYYYMMDD_HHMMSS 格式，字典序即为时间先后）
        entries.sort()
        if len(entries) > max_keep:
            to_delete = entries[:len(entries) - max_keep]
            for d in to_delete:
                shutil.rmtree(os.path.join(module_dir, d), ignore_errors=True)
    except Exception:
        pass


def get_module_history(module_name: str) -> list:
    """
    获取指定模块的历史记录列表（最新排在最前）
    """
    m_dir = get_module_history_dir(module_name)
    if not os.path.exists(m_dir):
        return []

    records = []
    entries = [
        d for d in os.listdir(m_dir)
        if os.path.isdir(os.path.join(m_dir, d))
    ]
    # 倒序，最新的在前
    entries.sort(reverse=True)

    for entry in entries:
        r_dir = os.path.join(m_dir, entry)
        meta_file = os.path.join(r_dir, "meta.json")
        if os.path.exists(meta_file):
            try:
                with open(meta_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    data["dir_path"] = r_dir
                    records.append(data)
            except Exception:
                pass
    return records


def get_storage_stats(module_name: str) -> dict:
    """获取该模块历史占用统计"""
    records = get_module_history(module_name)
    total_bytes = 0
    for r in records:
        for f in r.get("files", []):
            total_bytes += f.get("size", 0)
    return {
        "count": len(records),
        "total_bytes": total_bytes,
        "total_size_str": _format_size(total_bytes),
        "max_limit": MAX_HISTORY_PER_MODULE
    }


def clear_module_history(module_name: str):
    """清空指定模块的全部历史记录"""
    m_dir = get_module_history_dir(module_name)
    if os.path.exists(m_dir):
        shutil.rmtree(m_dir, ignore_errors=True)
        os.makedirs(m_dir, exist_ok=True)


def render_history_ui(module_name: str):
    """
    在 Streamlit 页面渲染模块历史归档面板（折叠抽屉）
    """
    records = get_module_history(module_name)
    count = len(records)
    stats = get_storage_stats(module_name)

    title = f"🕒 历史生成记录与往期回溯 ({count} 条)"
    if count == 0:
        title = "🕒 历史生成记录与往期回溯 (暂无记录)"

    with st.expander(title, expanded=False):
        if count == 0:
            st.info("💡 暂无历史生成记录。在此模块点击生成后，系统将自动对生成表格与核心指标进行快照归档，随时可在此回溯查看与下载。")
            return

        # 顶部轻量状态条
        st.markdown(
            f"""
            <div style="display:flex; justify-content:space-between; align-items:center; background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:8px 14px; margin-bottom:14px; font-size:13px; color:#475569;">
                <div>
                    📦 已归档 <b>{count}</b> 次生成结果 · 占用空间 <b>{stats['total_size_str']}</b> · 最多保留 <b>{MAX_HISTORY_PER_MODULE}</b> 条
                </div>
                <div style="color:#16a34a; font-size:12px; font-weight:500;">
                    ✅ 随时回溯下载 · 重启不丢失
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # 遍历每条历史记录
        for r_idx, rec in enumerate(records):
            run_id = rec.get("run_id", f"r_{r_idx}")
            timestamp = format_record_timestamp(rec)
            summary = rec.get("summary", "")
            files = rec.get("files", [])
            dir_path = rec.get("dir_path", "")

            # 卡片容器
            with st.container():
                st.markdown(
                    f"""
                    <div style="background:#ffffff; border:1px solid #e5e7eb; border-left:4px solid #0284c7; border-radius:6px; padding:10px 14px; margin-bottom:10px;">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                            <div style="display:flex; align-items:center; gap:8px;">
                                <span style="font-size:14px; font-weight:600; color:#1e293b;">📅 {timestamp}</span>
                                <span style="font-size:11px; color:#0284c7; background:#f0f9ff; border:1px solid #bae6fd; padding:1px 6px; border-radius:4px; font-weight:500;">北京时间 (UTC+8)</span>
                            </div>
                            <span style="font-size:12px; background:#eff6ff; color:#1d4ed8; padding:2px 8px; border-radius:4px; border:1px solid #dbeafe;">
                                {f'包含 {len(files)} 个文件' if len(files) > 1 else '单文件结果'}
                            </span>
                        </div>
                        {f'<div style="font-size:13px; color:#0369a1; margin-top:2px;">📊 <b>指标快照</b>：{summary}</div>' if summary else ''}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # 下载按钮区
                if files:
                    # 如果有多个文件，排版紧凑
                    cols = st.columns(min(len(files), 3))
                    for f_idx, f_item in enumerate(files):
                        c_idx = f_idx % min(len(files), 3)
                        fname = f_item.get("name", "")
                        fsize_str = f_item.get("size_str", "")
                        fpath = os.path.join(dir_path, fname)

                        if os.path.exists(fpath):
                            with open(fpath, "rb") as bf:
                                b_data = bf.read()

                            btn_label = f"⬇️ 下载【{fname}】 ({fsize_str})"
                            if f_item.get("is_auto_zip"):
                                btn_label = f"📦 打包下载全部 (ZIP) ({fsize_str})"

                            with cols[c_idx]:
                                st.download_button(
                                    label=btn_label,
                                    data=b_data,
                                    file_name=fname,
                                    key=f"hdl_{_sanitize_name(module_name)}_{run_id}_{f_idx}",
                                    use_container_width=True
                                )

                # 快速数据预览 (仅对 Excel 文件)
                excel_files = [f for f in files if f.get("name", "").endswith((".xlsx", ".xls"))]
                if excel_files:
                    with st.expander(f"👁️ 在线预览数据 ({timestamp})", expanded=False):
                        if len(excel_files) == 1:
                            target_f = excel_files[0]
                            fp = os.path.join(dir_path, target_f["name"])
                            if os.path.exists(fp):
                                try:
                                    xl = pd.ExcelFile(fp)
                                    if len(xl.sheet_names) == 1:
                                        df_p = xl.parse(xl.sheet_names[0])
                                        st.caption(f"工作表【{xl.sheet_names[0]}】共 {len(df_p)} 行（展示前 30 行）：")
                                        st.dataframe(df_p.head(30), use_container_width=True)
                                    else:
                                        p_tabs = st.tabs([f"Sheet: {s}" for s in xl.sheet_names])
                                        for s_i, s_name in enumerate(xl.sheet_names):
                                            with p_tabs[s_i]:
                                                df_p = xl.parse(s_name)
                                                st.caption(f"工作表【{s_name}】共 {len(df_p)} 行（展示前 30 行）：")
                                                st.dataframe(df_p.head(30), use_container_width=True)
                                except Exception as err:
                                    st.caption(f"预览加载失败: {err}")
                        else:
                            f_tabs = st.tabs([f["name"] for f in excel_files])
                            for ef_i, target_f in enumerate(excel_files):
                                with f_tabs[ef_i]:
                                    fp = os.path.join(dir_path, target_f["name"])
                                    if os.path.exists(fp):
                                        try:
                                            xl = pd.ExcelFile(fp)
                                            df_p = xl.parse(xl.sheet_names[0])
                                            st.caption(f"工作表【{xl.sheet_names[0]}】共 {len(df_p)} 行（展示前 30 行）：")
                                            st.dataframe(df_p.head(30), use_container_width=True)
                                        except Exception as err:
                                            st.caption(f"预览加载失败: {err}")

                st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    # 抽屉下方保留微距，与下方规范卡片自然隔离
    st.markdown("<div style='margin-bottom: 16px;'></div>", unsafe_allow_html=True)
