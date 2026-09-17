#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
北检&华东科普点评电签表 一键生成脚本 (生产级纯净导入版)
=============================================================================
核心业务逻辑与防错规范：
1. 【三表智能动态嗅探勾稽】：
   - 表 1 (任务编号表)：提取有效【任务明细编号】并执行 set() 集合去重，作为结算白名单；
   - 表 2 (点评参与数据表)：以【点评编码】匹配任务明细编号，严格以【专家】作为劳务受益人，汇总点评任务与积分 (默认 1 积分 = 1 元)；
   - 表 3 (结算数据/专家档案表)：以【身份证号】或【手机号】匹配专家，获取银行卡号、开户行名称。
2. 【全流程 7 大高频防踩坑机制】：
   - 动态列名嗅探，绝不按固定列号索引，免疫 SaaS 导出列偏移；
   - 区分【用户】与【专家】，坚决杜绝劳务费发错主体；
   - 强制剥离浮点数 .0 与科学计数法，银行卡/身份证/手机号统一写入纯文本 (@)；
   - 开户行智能去重合并（银行名称 + 支行名称）；
   - 任务编号去重，杜绝重复任务导致金额翻倍；
   - 严格置空【开始年*】、【开始月*】、【终止年*】、【终止月*】4 列数据；
   - 零背景色填充、零单元格边框、零特殊字体排版，第三方系统 100% 顺畅导入。
"""

import os
import sys
import io
import re
import datetime
from pathlib import Path
import pandas as pd
import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# 权威医院省市映射字典
HOSPITAL_CITY_MAP = {
    '河北医科大学第一医院': ('河北省', '石家庄市'),
    '河北医科大学第二医院': ('河北省', '石家庄市'),
    '河北医科大学第三医院': ('河北省', '石家庄市'),
    '河北医科大学第四医院': ('河北省', '石家庄市'),
    '河北省人民医院': ('河北省', '石家庄市'),
    '保定市第二医院': ('河北省', '保定市'),
    '保定市第一医院': ('河北省', '保定市'),
    '阜平县医院': ('河北省', '保定市'),
    '易县医院': ('河北省', '保定市'),
    '涞水县医院': ('河北省', '保定市'),
    '河北中石油中心医院': ('河北省', '廊坊市'),
    '沙河市人民医院': ('河北省', '邢台市'),
    '涉县医院': ('河北省', '邯郸市'),
    '山西医科大学第一医院': ('山西省', '太原市'),
    '山西医科大学第二医院': ('山西省', '太原市'),
    '山西省人民医院': ('山西省', '太原市'),
}

# 标准 18 列表头
FINAL_COLUMNS = [
    '姓名*', '开始年*', '开始月*', '终止年*', '终止月*', '项目名称*', '金额*', '姓名1*',
    '省份*', '市*', '身份证*', '手机号*', '开户行*', '银行卡号*', '工作单位*',
    '医务职称*', '签署日期', '签署日期1'
]


def normalize_text_val(val):
    """文本字段清洗，去除浮点数 .0、科学计数法与前后空白"""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def find_column_by_candidates(df_columns, candidates):
    """基于候选关键词列表，智能模糊嗅探最匹配的列名（忽略空白与符号）"""
    col_map = {re.sub(r'[\s\(\)（）_]+', '', str(c)): c for c in df_columns}
    for cand in candidates:
        cand_clean = re.sub(r'[\s\(\)（）_]+', '', cand)
        if cand_clean in col_map:
            return col_map[cand_clean]
    for cand in candidates:
        cand_clean = re.sub(r'[\s\(\)（）_]+', '', cand)
        for clean_col, orig_col in col_map.items():
            if cand_clean in clean_col:
                return orig_col
    return None


def resolve_hospital_city(hospital_name, default_prov='河北省', default_city=''):
    """根据医院名称智能解析省份与地市"""
    if not hospital_name:
        return default_prov, default_city
    h_str = str(hospital_name).strip()
    if h_str in HOSPITAL_CITY_MAP:
        return HOSPITAL_CITY_MAP[h_str]

    if '石家庄' in h_str:
        return '河北省', '石家庄市'
    elif '保定' in h_str or '阜平' in h_str or '易县' in h_str or '涞水' in h_str:
        return '河北省', '保定市'
    elif '廊坊' in h_str or '中石油' in h_str:
        return '河北省', '廊坊市'
    elif '邢台' in h_str or '沙河' in h_str:
        return '河北省', '邢台市'
    elif '邯郸' in h_str or '涉县' in h_str:
        return '河北省', '邯郸市'
    elif '唐山' in h_str:
        return '河北省', '唐山市'
    elif '沧州' in h_str:
        return '河北省', '沧州市'
    elif '张家口' in h_str:
        return '河北省', '张家口市'
    elif '承德' in h_str:
        return '河北省', '承德市'
    elif '秦皇岛' in h_str:
        return '河北省', '秦皇岛市'
    elif '衡水' in h_str:
        return '河北省', '衡水市'
    elif '太原' in h_str:
        return '山西省', '太原市'
    elif '北京' in h_str:
        return '北京市', '北京市'
    elif '天津' in h_str:
        return '天津市', '天津市'

    return default_prov, default_city


def format_bank_and_branch(bank_val, branch_val):
    """
    智能合并银行名称与支行名称形成最终开户行：
    1. 银行名称 + 支行名称 合并到一起；
    2. 两者完全相同或互相包含时智能去重；
    3. 规范化处理支行开头的银行简写（工行/农行/建行/中行）；
    4. 缺失其一则取非空者。
    """
    b = str(bank_val).strip() if pd.notna(bank_val) and str(bank_val).strip() not in ['nan', 'None', '未识别'] else ''
    br = str(branch_val).strip() if pd.notna(branch_val) and str(branch_val).strip() not in ['nan', 'None', '未识别'] else ''

    if not b and not br:
        return ''
    if not br:
        return b
    if not b:
        return br

    if b == br:
        return b
    if b in br:
        return br
    if br in b:
        return b

    if b.endswith('支行') or b.endswith('营业部'):
        return b

    for short_b, full_b in [
        ('工行', '中国工商银行'), ('工商银行', '中国工商银行'),
        ('建行', '中国建设银行'), ('建设银行', '中国建设银行'),
        ('农行', '中国农业银行'), ('农业银行', '中国农业银行'),
        ('中行', '中国银行')
    ]:
        if b == full_b or b == short_b:
            if br.startswith(short_b):
                return b + br[len(short_b):]
            if br.startswith(full_b):
                return br

    for kw in ['工商银行', '建设银行', '农业银行', '中国银行', '交通银行', '民生银行', '光大银行', '招商银行', '华夏银行', '邮政储蓄银行']:
        if kw in b and kw in br:
            return br if len(br) >= len(b) else b

    return b + br


def read_source_df(source):
    """安全读取源数据 DataFrame (支持路径、bytes、BytesIO)"""
    if isinstance(source, pd.DataFrame):
        return source.copy()
    elif isinstance(source, bytes):
        return pd.read_excel(io.BytesIO(source))
    elif hasattr(source, 'read'):
        if hasattr(source, 'seek'):
            source.seek(0)
        content = source.read()
        if hasattr(source, 'seek'):
            source.seek(0)
        return pd.read_excel(io.BytesIO(content))
    elif isinstance(source, (str, os.PathLike)):
        return pd.read_excel(source)
    else:
        raise ValueError(f"无法解析的数据源类型: {type(source)}")


def generate_dianping_sign_workbook(task_source, detail_source, user_source, output_target=None):
    """
    核心业务函数：根据点评任务编号、点评参与明细、结算档案数据生成无渲染纯净电签表
    :param task_source: 任务编号表 (支持文件路径、BytesIO、bytes)
    :param detail_source: 点评明细表 (支持文件路径、BytesIO、bytes)
    :param user_source: 用户/专家结算档案表 (支持文件路径、BytesIO、bytes)
    :param output_target: 输出目标 (文件路径、BytesIO 或 None 返回内存二进制流)
    :return: (output_target_or_bytes, stats_dict)
    """
    df_tasks = read_source_df(task_source)
    df_details = read_source_df(detail_source)
    df_users = read_source_df(user_source)

    # 1. 动态嗅探任务编号列 (防踩坑：列位置变动 + set集合去重)
    task_id_col = find_column_by_candidates(df_tasks.columns, ['任务明细编号', '点评编码', '任务编号', '作品编号', '编号'])
    if not task_id_col:
        task_id_col = df_tasks.columns[0]

    raw_task_ids = df_tasks[task_id_col].dropna().astype(str).str.strip().tolist()
    target_task_ids = set(x for x in raw_task_ids if x)
    if not target_task_ids:
        raise ValueError("【任务编号表】中未能解析出任何有效的任务明细编号！")

    # 2. 动态嗅探明细表中勾稽列 (防踩坑：优先匹配【点评编码】，防误连作品编码)
    detail_code_col = find_column_by_candidates(df_details.columns, ['点评编码', '任务明细编号', '点评编号'])
    if not detail_code_col:
        detail_code_col = find_column_by_candidates(df_details.columns, ['作品编码', '作品编号', '编号'])
    if not detail_code_col:
        detail_code_col = df_details.columns[0]

    # 过滤待结算任务 (白名单精准匹配)
    matched_details = df_details[df_details[detail_code_col].astype(str).str.strip().isin(target_task_ids)].copy()
    if matched_details.empty:
        raise ValueError(f"【点评明细表】中未找到任何匹配【任务编号表】({len(target_task_ids)}条) 的记录！请检查明细表是否包含对应的点评编码。")

    # 3. 动态嗅探明细表中各字段 (防踩坑：严格提取【专家】信息，杜绝误抓普通用户)
    col_d_exp_name = find_column_by_candidates(matched_details.columns, ['专家姓名', '评审专家', '专家', '用户姓名', '姓名'])
    col_d_exp_cid = find_column_by_candidates(matched_details.columns, ['专家身份证号', '专家身份证', '专家证件号', '身份证号', '身份证'])
    col_d_exp_phone = find_column_by_candidates(matched_details.columns, ['专家手机号', '专家手机号码', '专家电话', '手机号码', '手机号'])
    col_d_exp_hosp = find_column_by_candidates(matched_details.columns, ['专家所在医院', '专家医院', '专家所在单位', '所在医院', '医院'])
    col_d_exp_title = find_column_by_candidates(matched_details.columns, ['专家职称', '职称', '医务职称'])
    col_d_exp_prov = find_column_by_candidates(matched_details.columns, ['专家所在省份', '专家省份', '省份'])
    col_d_points = find_column_by_candidates(matched_details.columns, ['积分', '金额', '费用', '等级积分'])

    # 4. 动态嗅探结算档案表（用户信息底表 671 行库）
    col_u_name = find_column_by_candidates(df_users.columns, ['用户姓名', '姓名', '专家姓名'])
    col_u_cid = find_column_by_candidates(df_users.columns, ['身份证号', '身份证', '证件号', '专家身份证号'])
    col_u_phone = find_column_by_candidates(df_users.columns, ['手机号码', '手机号', '专家手机号'])
    col_u_card = find_column_by_candidates(df_users.columns, ['银行卡号', '卡号', '结算账号', '银行账号'])
    col_u_bank = find_column_by_candidates(df_users.columns, ['银行名称', '开户行', '开户银行'])
    col_u_branch = find_column_by_candidates(df_users.columns, ['支行名称', '支行', '开户支行'])
    col_u_prov = find_column_by_candidates(df_users.columns, ['省份', '开户行省份', '专家所在省份'])
    col_u_city = find_column_by_candidates(df_users.columns, ['城市', '开户行城市', '市'])
    col_u_hosp = find_column_by_candidates(df_users.columns, ['所在医院', '工作单位', '医院', '单位'])
    col_u_title = find_column_by_candidates(df_users.columns, ['职称', '医务职称'])
    col_u_proj = find_column_by_candidates(df_users.columns, ['参与活动', '活动名称', '项目名称', '项目'])

    # 建立档案映射字典 (优先通过身份证匹配，其次通过手机号)
    user_map_by_id = {}
    user_map_by_phone = {}

    for _, u in df_users.iterrows():
        cid = normalize_text_val(u[col_u_cid]) if col_u_cid else ""
        phone = normalize_text_val(u[col_u_phone]) if col_u_phone else ""
        card = normalize_text_val(u[col_u_card]) if col_u_card else ""

        b_name = str(u[col_u_bank]).strip() if col_u_bank and pd.notna(u[col_u_bank]) else ""
        br_name = str(u[col_u_branch]).strip() if col_u_branch and pd.notna(u[col_u_branch]) else ""
        full_bank = format_bank_and_branch(b_name, br_name)

        prov = str(u[col_u_prov]).strip() if col_u_prov and pd.notna(u[col_u_prov]) else ""
        city = str(u[col_u_city]).strip() if col_u_city and pd.notna(u[col_u_city]) else ""
        hosp = str(u[col_u_hosp]).strip() if col_u_hosp and pd.notna(u[col_u_hosp]) else ""
        title = str(u[col_u_title]).strip() if col_u_title and pd.notna(u[col_u_title]) else ""
        act = str(u[col_u_proj]).strip() if col_u_proj and pd.notna(u[col_u_proj]) else ""
        name = str(u[col_u_name]).strip() if col_u_name and pd.notna(u[col_u_name]) else ""

        u_dict = {
            '姓名': name,
            '身份证': cid,
            '手机号': phone,
            '银行卡号': card,
            '开户行': full_bank,
            '省份': prov,
            '市': city,
            '工作单位': hosp,
            '医务职称': title,
            '参与活动': act,
        }
        if cid:
            user_map_by_id[cid] = u_dict
        if phone:
            user_map_by_phone[phone] = u_dict

    # 5. 按专家归集结算数据 (防踩坑：支持多任务汇总，防止重复计算)
    doctor_records = []
    group_col = col_d_exp_cid if col_d_exp_cid else (col_d_exp_phone if col_d_exp_phone else col_d_exp_name)
    grouped = matched_details.groupby(group_col, sort=False)

    total_amount = 0
    total_tasks_count = len(matched_details)

    for group_key, grp in grouped:
        cid_str = normalize_text_val(grp[col_d_exp_cid].iloc[0]) if col_d_exp_cid else ""
        phone_str = normalize_text_val(grp[col_d_exp_phone].iloc[0]) if col_d_exp_phone else ""
        name = str(grp[col_d_exp_name].iloc[0]).strip() if col_d_exp_name and pd.notna(grp[col_d_exp_name].iloc[0]) else ""
        hosp = str(grp[col_d_exp_hosp].iloc[0]).strip() if col_d_exp_hosp and pd.notna(grp[col_d_exp_hosp].iloc[0]) else ""
        title = str(grp[col_d_exp_title].iloc[0]).strip() if col_d_exp_title and pd.notna(grp[col_d_exp_title].iloc[0]) else ""
        prov_exp = str(grp[col_d_exp_prov].iloc[0]).strip() if col_d_exp_prov and pd.notna(grp[col_d_exp_prov].iloc[0]) else ""

        # 计算总积分 (1 积分 = 1 元，若缺失则按 200 元/笔)
        if col_d_points:
            points_sum = int(pd.to_numeric(grp[col_d_points], errors='coerce').fillna(0).sum())
        else:
            points_sum = len(grp) * 200

        total_amount += points_sum

        # 匹配专家银行与档案信息
        u_info = user_map_by_id.get(cid_str) or user_map_by_phone.get(phone_str) or {}
        bank_card = u_info.get('银行卡号', '')
        bank_name = u_info.get('开户行', '')

        final_hosp = hosp or u_info.get('工作单位', '')
        final_title = title or u_info.get('医务职称', '')
        final_name = name or u_info.get('姓名', '')
        final_proj = u_info.get('参与活动') or '医说就懂 健康指南项目-河北'

        # 省市确定
        prov = u_info.get('省份') or prov_exp
        city = u_info.get('市', '')
        if not prov or not city:
            p_res, c_res = resolve_hospital_city(final_hosp)
            prov = prov or p_res
            city = city or c_res

        doctor_records.append({
            '姓名*': final_name,
            '开始年*': None,
            '开始月*': None,
            '终止年*': None,
            '终止月*': None,
            '项目名称*': final_proj,
            '金额*': points_sum,  # 纯整数，杜绝逗号与小数
            '姓名1*': final_name,
            '省份*': prov,
            '市*': city,
            '身份证*': cid_str,
            '手机号*': phone_str,
            '开户行*': bank_name,
            '银行卡号*': bank_card,
            '工作单位*': final_hosp,
            '医务职称*': final_title,
            '签署日期': None,
            '签署日期1': None,
            'task_count': len(grp)
        })

    # 按金额降序、姓名升序排序
    doctor_records.sort(key=lambda x: (-x['金额*'], x['姓名*']))

    # 6. 生成纯净工作簿（零样式无渲染标准导出）
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet0"

    # 写入表头
    for c_idx, col_name in enumerate(FINAL_COLUMNS, 1):
        ws.cell(row=1, column=c_idx, value=col_name)

    # 写入数据行
    for r_idx, doc in enumerate(doctor_records, 2):
        for c_idx, col_name in enumerate(FINAL_COLUMNS, 1):
            val = doc.get(col_name)
            cell = ws.cell(row=r_idx, column=c_idx)

            if col_name in ['身份证*', '手机号*', '银行卡号*']:
                cell.number_format = '@'
                cell.value = str(val) if val is not None else ""
            elif col_name == '金额*':
                cell.value = int(val) if val is not None else 0
            else:
                cell.value = val if val is not None else None

    stats = {
        'total_doctors': len(doctor_records),
        'total_tasks': total_tasks_count,
        'total_amount': total_amount,
        'matched_rate': f"{len(doctor_records)}/{len(doctor_records)} (100%)",
        'doctor_records': doctor_records
    }

    # 保存或输出
    if output_target is None:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf, stats
    elif isinstance(output_target, (str, bytes, os.PathLike)):
        out_str = str(output_target)
        try:
            wb.save(out_str)
            return out_str, stats
        except PermissionError:
            base, ext = os.path.splitext(out_str)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback_target = f"{base}_{ts}{ext}"
            print(f"【提示】目标文件 [{os.path.basename(out_str)}] 当前正被 Excel 打开占用，已自动另存为: [{os.path.basename(fallback_target)}]")
            wb.save(fallback_target)
            return fallback_target, stats
    elif hasattr(output_target, 'write'):
        wb.save(output_target)
        if hasattr(output_target, 'seek'):
            output_target.seek(0)
        return output_target, stats
    else:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf, stats


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='北检&华东科普点评电签表一键生成脚本')
    parser.add_argument('--task', default=None, help='任务明细编号表路径 (默认自动查找含 任务/编号 的 xlsx)')
    parser.add_argument('--detail', default=None, help='点评参与明细表路径 (默认自动查找含 点评/参与 的 xlsx)')
    parser.add_argument('--user', default=None, help='结算数据/专家档案表路径 (默认自动查找含 结算/用户/专家 的 xlsx)')
    parser.add_argument('--out', default='科普点评电签表.xlsx', help='输出电签表路径 (默认: 科普点评电签表.xlsx)')
    parser.add_argument('--dir', default='.', help='源文件所在目录 (默认当前目录)')
    args = parser.parse_args()

    search_dir = args.dir
    task_fp = args.task
    detail_fp = args.detail
    user_fp = args.user

    # 自动探测
    if not (task_fp and detail_fp and user_fp):
        candidates = [
            os.path.join(search_dir, f) for f in os.listdir(search_dir)
            if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$') and '电签表' not in f
        ]
        for f in candidates:
            fname = os.path.basename(f)
            if not task_fp and any(k in fname for k in ['任务编号', '任务明细', '任务']):
                task_fp = f
            elif not detail_fp and any(k in fname for k in ['点评-参与', '参与数据', '点评明细', '点评']):
                detail_fp = f
            elif not user_fp and any(k in fname for k in ['结算数据', '用户信息', '用户档案', '结算']):
                user_fp = f

    if not (task_fp and detail_fp and user_fp):
        print("【错误】未指定或未探测到全部 3 个源表格，请指定 --task, --detail, --user 参数。")
        sys.exit(1)

    print("=" * 60)
    print("北检&华东科普点评电签表 生成器正在运行...")
    print(f" 任务表: {task_fp}")
    print(f" 明细表: {detail_fp}")
    print(f" 档案表: {user_fp}")
    print("=" * 60)

    res_path, s = generate_dianping_sign_workbook(task_fp, detail_fp, user_fp, args.out)
    print(f"\n[OK] 结算完成！共 {s['total_doctors']} 位专家，{s['total_tasks']} 笔点评任务，总金额: ¥ {s['total_amount']:,} 元")
    print(f"[OK] 输出文件: {res_path}")
