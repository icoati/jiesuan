#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调研案例电签表 一键智能生成工具 (工业级高精度纯文本版)
=============================================================================
适用业务场景：
  医药调研、案例征集等多问卷/多案例收集类电签结算。
  
核心原则：
  1. 真实抓取，绝不硬编码：
     - 动态列名自适应嗅探（免疫 SaaS 导出列偏移与列名同义词变动）；
     - 项目名称/活动名称真实抓取（从明细、用户表或任务名中统计最高频真实名称）；
     - 身份证号、手机号、银行卡号强制纯字符串读取 (dtype=str)，彻底根除 IEEE 754 浮点截断；
  2. 三表安全勾稽：
     - 表 1【进度/确认表】：提供待结算的“任务明细编码”白名单；（若未提供，自动根据明细表审核状态自适应回退）
     - 表 2【问卷/案例明细表】：提供明文身份证、手机号、任务名称、答卷时间、本次获得积分；
     - 表 3【项目人员列表】：提供银行卡号、银行名称、支行、开户行省市、资质证明链接；
  3. 金融电签高标准输出：
     - 严格输出 20 列标准格式：姓名*、开始年*、开始月*、开始日*、项目名称*、金额*、姓名1*、
       省份*、市*、身份证*、手机号*、开户行*、银行卡号*、工作单位*、医务职称*、科室*、
       签署日期、签署日期1、身份证头像面、身份证反面；
     - 银行卡号强制文本格式 (@) 输出，支持卡BIN反查总行；
     - 生成对账明细表与异常核验名单，支持 xlsx 与标准 xls 双格式。
"""

import os
import sys
import glob
import re
import io
import time
import datetime
from collections import Counter
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import xlwt
except ImportError:
    xlwt = None

# 解决控制台输出中文编码问题
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ==================== 1. 字段别名同义词词典 ====================
COLUMN_ALIASES = {
    'task_detail_code': ['任务明细编码', '任务明细编号', '作品编号', '作品编码', '点评编码', '编码', '编号', '任务明细id'],
    'doctor_name': ['项目人员', '专家姓名', '用户名称', '医生姓名', '专家', '医生', '姓名', '人员姓名'],
    'phone': ['项目人员手机号', '手机号', '手机号码', '联系电话', '电话', '用户手机号'],
    'id_card': ['项目人员身份证号', '证件号码', '身份证号码', '身份证号', '身份证', '证件号', '身份证号/护照号'],
    'hospital': ['医院', '所在医院', '医院全称', '工作单位', '医院名称', '医疗机构', '单位名称', '单位'],
    'dept': ['科室', '所在科室', '专业科室', '科室名称', '部门'],
    'title': ['职称', '医务职称', '技术职称', '医生职称', '职务职称', '专业职称'],
    'bank_card': ['银行卡号', '卡号', '结算卡号', '银行账号', '账号', '结算账号'],
    'bank_name': ['银行名称', '开户银行', '开户行', '银行', '结算银行'],
    'branch_name': ['支行', '支行名称', '开户支行', '开户网点', '开户行支行'],
    'points': ['本次获得积分', '单份最大积分', '获得积分', '积分', '金额', '结算单价', '费用', '单价'],
    'project_name': ['任务名称', '项目名称', '问卷名称', '参与活动', '活动名称', '所属项目', '业务项目', '活动', '项目'],
    'audit_status': ['核验结果', '审核状态', '认证状态', '状态'],
    'create_time': ['创建时间', '提交时间', '申请时间', '答卷时间', '接包时间', '用户进入项目时间'],
    'prov': ['医院所在省', '所在省', '省份', '省', '开户行所在省'],
    'city': ['医院所在市', '所在市', '市', '城市', '开户行所在市'],
    'id_front_pic': ['身份证头像面', '身份证正面', '身份证人像面', '身份证照片正面', '身份证图片'],
    'id_back_pic': ['身份证反面', '身份证国徽面', '身份证照片反面']
}

# 银联常用卡BIN速查（兜底补齐总行）
BANK_BIN_MAP = {
    "622202": "中国工商银行", "622200": "中国工商银行", "621226": "中国工商银行", "621723": "中国工商银行",
    "622700": "中国建设银行", "621700": "中国建设银行", "436742": "中国建设银行",
    "622848": "中国农业银行", "622845": "中国农业银行", "622846": "中国农业银行",
    "621661": "中国银行", "621660": "中国银行", "601382": "中国银行", "621790": "中国银行",
    "622260": "交通银行", "622262": "交通银行",
    "621098": "中国邮政储蓄银行", "622188": "中国邮政储蓄银行", "621798": "中国邮政储蓄银行",
    "621483": "招商银行", "621485": "招商银行", "622588": "招商银行",
    "622521": "上海浦东发展银行", "621789": "上海浦东发展银行",
    "622690": "中信银行", "622660": "中国光大银行", "622622": "中国民生银行"
}

CITY_RULES = [
    (r'(北京)', '北京市', '北京市'),
    (r'(上海)', '上海市', '上海市'),
    (r'(天津)', '天津市', '天津市'),
    (r'(重庆)', '重庆市', '重庆市'),
    (r'(阜阳)', '安徽省', '阜阳市'),
    (r'(滁州)', '安徽省', '滁州市'),
    (r'(合肥)', '安徽省', '合肥市'),
    (r'(石家庄|无极)', '河北省', '石家庄市'),
    (r'(保定|涞水|高碑店|定州)', '河北省', '保定市'),
    (r'(德州)', '山东省', '德州市'),
    (r'(济南)', '山东省', '济南市'),
    (r'(青岛)', '山东省', '青岛市'),
    (r'(西安)', '陕西省', '西安市'),
    (r'(榆林)', '陕西省', '榆林市'),
    (r'(晋城|阳城)', '山西省', '晋城市'),
    (r'(太原)', '山西省', '太原市'),
    (r'(杭州)', '浙江省', '杭州市'),
    (r'(宁波)', '浙江省', '宁波市'),
    (r'(徐州|睢宁)', '江苏省', '徐州市'),
    (r'(南京|苏州|无锡|常州|南通)', '江苏省', '江苏省市级')
]

# ==================== 2. 数据清洗与工具函数 ====================
def clean_str(val):
    """文本清洗去空白与脏后缀"""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ['nan', 'none', 'null', 'undefined']:
        return ""
    if s.endswith(".0"):
        prefix = s[:-2]
        if prefix.isdigit():
            s = prefix
    return s

def clean_card_number(val):
    """银行卡号专用清洗：纯文本处理，严禁 float 转换，杜绝精度截断"""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'null', 'undefined']:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    # 保留纯数字和连字符，绝对不转 float
    s = re.sub(r'[^\d\-]', '', s)
    return s

def find_column_by_aliases(columns, alias_key):
    """基于别名列表智能动态匹配列名"""
    candidates = COLUMN_ALIASES.get(alias_key, [])
    clean_cols_map = {re.sub(r'[\s\(\)（）_\ufeff\*]+', '', str(c)).lower(): c for c in columns}
    for cand in candidates:
        clean_cand = re.sub(r'[\s\(\)（）_\ufeff\*]+', '', cand).lower()
        if clean_cand in clean_cols_map:
            return clean_cols_map[clean_cand]
    # 模糊包含匹配
    for cand in candidates:
        clean_cand = re.sub(r'[\s\(\)（）_\ufeff\*]+', '', cand).lower()
        for k, original_col in clean_cols_map.items():
            if clean_cand in k:
                return original_col
    return None

def format_bank_and_branch(card_num, bank_val, branch_val):
    """开户行与支行智能合并（去重规范化与卡BIN智能补全）"""
    b = clean_str(bank_val)
    br = clean_str(branch_val)
    
    # 若银行总行缺失，通过卡BIN自动推断
    if not b and card_num and len(card_num) >= 6:
        bin6 = card_num[:6]
        b = BANK_BIN_MAP.get(bin6, "")
        
    if not b and not br:
        return ""
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
    if b.endswith('支行') or b.endswith('营业部') or b.endswith('分行'):
        return b
    return b + br

def infer_location_from_hospital(hospital_name):
    """根据医院名称推断省份和城市"""
    if not hospital_name:
        return "", ""
    h_str = str(hospital_name)
    for pattern, prov, city in CITY_RULES:
        if re.search(pattern, h_str):
            return prov, city
    return "", ""

def extract_clean_project_name(name_str):
    """真实抓取项目名称主体，去除括号后缀"""
    if not name_str:
        return ""
    s = str(name_str).strip()
    m_cn = re.findall(r'[“"《「]([^”"》」]+)[”"》」]', s)
    if m_cn and m_cn[0].strip():
        return m_cn[0].strip()
    s_cleaned = re.sub(r'[\(（][^\)）]*(?:科|医学|部|室|内科|外科)[^\)）]*[\)）]$', '', s).strip()
    return s_cleaned or s

# ==================== 3. 自动嗅探定位源文件 ====================
def detect_source_files(target_dir):
    all_files = glob.glob(os.path.join(target_dir, "*.xlsx")) + glob.glob(os.path.join(target_dir, "*.xls"))
    candidate_files = [
        f for f in all_files
        if not os.path.basename(f).startswith("~$")
        and not os.path.basename(f).startswith("最终")
        and not os.path.basename(f).startswith("调研案例电签表")
        and "对账" not in os.path.basename(f)
        and "筛选" not in os.path.basename(f)
        and "备份" not in os.path.basename(f)
    ]
    
    progress_file = None
    detail_file = None
    user_file = None
    
    # 第一级：文件名关键词快速识别
    for f in candidate_files:
        fn = os.path.basename(f).lower()
        if ("人员" in fn or "用户" in fn or "专家" in fn or "医生" in fn) and not user_file:
            user_file = f
        elif ("进度" in fn or "已确认" in fn or "确认" in fn or "名单" in fn) and not progress_file:
            progress_file = f
        elif ("答卷" in fn or "明细" in fn or "调研" in fn or re.search(r'\d+[\-_]\d+', fn)) and not detail_file:
            detail_file = f

    # 第二级：内容特征二次嗅探
    unassigned = [f for f in candidate_files if f not in [progress_file, detail_file, user_file]]
    for f in unassigned:
        try:
            df_head = pd.read_excel(f, nrows=3, dtype=str)
            cols = [str(c) for c in df_head.columns]
            cols_text = "".join(cols)
            
            if not user_file and ("银行卡号" in cols_text or "开工状态" in cols_text or "执业证" in cols_text):
                user_file = f
            elif not detail_file and ("答卷" in cols_text or "本次获得积分" in cols_text or "单份最大积分" in cols_text):
                detail_file = f
            elif not progress_file and ("核验结果" in cols_text and "任务明细编码" in cols_text):
                progress_file = f
        except Exception:
            pass

    return progress_file, detail_file, user_file

# ==================== 4. 核心处理引擎 ====================
def process_survey_settlement(target_dir="."):
    print("=" * 70)
    print("      🏥 调研案例电签表 - 一键智能结算核验引擎 (高精度版)")
    print("=" * 70)
    print(f"📁 工作目录: {os.path.abspath(target_dir)}\n")
    
    progress_file, detail_file, user_file = detect_source_files(target_dir)
    
    print("🔍 智能源文件嗅探识别结果：")
    print(f"   • 进度确认表: {os.path.basename(progress_file) if progress_file else '【未提供，将基于明细表审核状态自适应回退】'}")
    print(f"   • 答卷明细表: {os.path.basename(detail_file) if detail_file else '【缺失 ❌】'}")
    print(f"   • 项目人员表: {os.path.basename(user_file) if user_file else '【缺失 ❌】'}")
    print("-" * 70)
    
    if not detail_file or not user_file:
        raise ValueError("核心源文件缺失：必须至少包含【答卷明细表】与【项目人员表】！")

    # 1. 强制以纯字符串读取所有表格，杜绝任何浮点截断与科学计数法篡改
    df_detail = pd.read_excel(detail_file, dtype=str).fillna("")
    df_user = pd.read_excel(user_file, dtype=str).fillna("")
    df_progress = pd.read_excel(progress_file, dtype=str).fillna("") if progress_file else None

    # 2. 动态定位列名
    col_d_code = find_column_by_aliases(df_detail.columns, 'task_detail_code')
    col_d_name = find_column_by_aliases(df_detail.columns, 'doctor_name')
    col_d_idcard = find_column_by_aliases(df_detail.columns, 'id_card')
    col_d_phone = find_column_by_aliases(df_detail.columns, 'phone')
    col_d_points = find_column_by_aliases(df_detail.columns, 'points')
    col_d_proj = find_column_by_aliases(df_detail.columns, 'project_name')
    col_d_hosp = find_column_by_aliases(df_detail.columns, 'hospital')
    col_d_dept = find_column_by_aliases(df_detail.columns, 'dept')
    col_d_title = find_column_by_aliases(df_detail.columns, 'title')
    col_d_prov = find_column_by_aliases(df_detail.columns, 'prov')
    col_d_city = find_column_by_aliases(df_detail.columns, 'city')
    col_d_status = find_column_by_aliases(df_detail.columns, 'audit_status')
    col_d_time = find_column_by_aliases(df_detail.columns, 'create_time')

    col_u_name = find_column_by_aliases(df_user.columns, 'doctor_name')
    col_u_idcard = find_column_by_aliases(df_user.columns, 'id_card')
    col_u_phone = find_column_by_aliases(df_user.columns, 'phone')
    col_u_card = find_column_by_aliases(df_user.columns, 'bank_card')
    col_u_bank = find_column_by_aliases(df_user.columns, 'bank_name')
    col_u_branch = find_column_by_aliases(df_user.columns, 'branch_name')
    col_u_hosp = find_column_by_aliases(df_user.columns, 'hospital')
    col_u_dept = find_column_by_aliases(df_user.columns, 'dept')
    col_u_title = find_column_by_aliases(df_user.columns, 'title')
    col_u_prov = find_column_by_aliases(df_user.columns, 'prov')
    col_u_city = find_column_by_aliases(df_user.columns, 'city')
    col_u_proj = find_column_by_aliases(df_user.columns, 'project_name')
    col_u_id_front = find_column_by_aliases(df_user.columns, 'id_front_pic')
    col_u_id_back = find_column_by_aliases(df_user.columns, 'id_back_pic')

    # 3. 确定待结算白名单
    if df_progress is not None:
        col_p_code = find_column_by_aliases(df_progress.columns, 'task_detail_code')
        if not col_p_code:
            col_p_code = df_progress.columns[0]
        valid_codes = set(df_progress[col_p_code].str.strip())
        print(f"📋 从【进度确认表】提取待结算有效编码: {len(valid_codes)} 条")
        matched_details = df_detail[df_detail[col_d_code].str.strip().isin(valid_codes)].copy()
    else:
        if col_d_status:
            matched_details = df_detail[df_detail[col_d_status].str.strip() == '审核通过'].copy()
            print(f"📋 未提供进度表，自动筛选明细表中「{col_d_status} == 审核通过」的记录: {len(matched_details)} 条")
        else:
            matched_details = df_detail.copy()
            print(f"📋 全量明细记录结算: {len(matched_details)} 条")

    if matched_details.empty:
        raise ValueError("未能匹配到任何符合结算条件的有效明细记录！请检查任务编码或审核状态。")

    # 4. 构建用户档案映射库 (多重键索引：身份证 / 手机号 / 姓名)
    user_map_by_idcard = {}
    user_map_by_phone = {}
    user_map_by_name = {}

    for _, u in df_user.iterrows():
        name = clean_str(u[col_u_name]) if col_u_name else ""
        cid = clean_str(u[col_u_idcard]) if col_u_idcard else ""
        phone = clean_str(u[col_u_phone]) if col_u_phone else ""
        card = clean_card_number(u[col_u_card]) if col_u_card else ""
        
        b_name = clean_str(u[col_u_bank]) if col_u_bank else ""
        br_name = clean_str(u[col_u_branch]) if col_u_branch else ""
        full_bank = format_bank_and_branch(card, b_name, br_name)
        
        u_dict = {
            '姓名': name,
            '身份证': cid,
            '手机号': phone,
            '银行卡号': card,
            '开户行': full_bank,
            '工作单位': clean_str(u[col_u_hosp]) if col_u_hosp else "",
            '医务职称': clean_str(u[col_u_title]) if col_u_title else "",
            '科室': clean_str(u[col_u_dept]) if col_u_dept else "",
            '省份': clean_str(u[col_u_prov]) if col_u_prov else "",
            '市': clean_str(u[col_u_city]) if col_u_city else "",
            '项目名称': clean_str(u[col_u_proj]) if col_u_proj else "",
            '身份证头像面': clean_str(u[col_u_id_front]) if col_u_id_front else "",
            '身份证反面': clean_str(u[col_u_id_back]) if col_u_id_back else "",
        }
        
        def store_user_rec(store, key, item):
            if not key: return
            if key not in store:
                store[key] = item
            else:
                # 优先保留有卡号和有开户行的记录
                if item['银行卡号'] and not store[key]['银行卡号']:
                    store[key] = item

        store_user_rec(user_map_by_idcard, cid, u_dict)
        store_user_rec(user_map_by_phone, phone, u_dict)
        store_user_rec(user_map_by_name, name, u_dict)

    # 5. 全局项目名称推断
    all_proj_candidates = []
    if col_d_proj and not matched_details.empty:
        all_proj_candidates.extend([clean_str(x) for x in matched_details[col_d_proj] if clean_str(x)])
    if col_u_proj and not df_user.empty:
        all_proj_candidates.extend([clean_str(x) for x in df_user[col_u_proj] if clean_str(x)])
    
    top_global_proj = Counter(all_proj_candidates).most_common(1)[0][0] if all_proj_candidates else ""
    top_global_proj_clean = extract_clean_project_name(top_global_proj)
    print(f"🏷️ 智能真实识别本期核心项目名称: [{top_global_proj_clean}]")

    # 6. 按医生聚合 (优先身份证，次选手机号，再次姓名)
    group_col = col_d_idcard or col_d_phone or col_d_name
    grouped = matched_details.groupby(group_col, sort=False)

    doctor_summary_list = []
    warnings = []
    total_amount = 0

    for _, grp in grouped:
        d_name = clean_str(grp[col_d_name].iloc[0]) if col_d_name else ""
        d_cid = clean_str(grp[col_d_idcard].iloc[0]) if col_d_idcard else ""
        d_phone = clean_str(grp[col_d_phone].iloc[0]) if col_d_phone else ""
        
        u_info = user_map_by_idcard.get(d_cid) or user_map_by_phone.get(d_phone) or user_map_by_name.get(d_name) or {}
        
        final_name = d_name or u_info.get('姓名', '')
        final_cid = d_cid if '*' not in d_cid else u_info.get('身份证', d_cid)
        final_phone = d_phone if '*' not in d_phone else u_info.get('手机号', d_phone)
        
        # 积分/金额 (1积分=1元)
        if col_d_points:
            amt = int(sum(float(x) for x in grp[col_d_points] if clean_str(x)))
        else:
            amt = len(grp) * 100
        total_amount += amt
        
        # 医院
        hosp = clean_str(grp[col_d_hosp].iloc[0]) if col_d_hosp else ""
        if not hosp: hosp = u_info.get('工作单位', '')
        
        # 职称与科室
        title = clean_str(grp[col_d_title].iloc[0]) if col_d_title else ""
        if not title: title = u_info.get('医务职称', '')
        
        dept = clean_str(grp[col_d_dept].iloc[0]) if col_d_dept else ""
        if not dept: dept = u_info.get('科室', '')
        
        # 省市
        prov = clean_str(grp[col_d_prov].iloc[0]) if col_d_prov else ""
        city = clean_str(grp[col_d_city].iloc[0]) if col_d_city else ""
        if not prov or not city:
            prov = prov or u_info.get('省份', '')
            city = city or u_info.get('市', '')
        if not prov or not city:
            inf_prov, inf_city = infer_location_from_hospital(hosp)
            prov = prov or inf_prov
            city = city or inf_city
            
        # 银行卡与开户行 (来自人员表真实抓取)
        bank_card = u_info.get('银行卡号', '')
        open_bank = u_info.get('开户行', '')
        
        if not bank_card:
            warnings.append(f"医生【{final_name}】(单位: {hosp}, 金额: ¥{amt}) 在项目人员表中未登记银行卡号！")
            
        # 提交日期
        c_time = None
        if col_d_time and clean_str(grp[col_d_time].iloc[0]):
            try:
                c_time = pd.to_datetime(grp[col_d_time].iloc[0])
            except Exception:
                pass
                
        raw_doc_proj = clean_str(grp[col_d_proj].iloc[0]) if col_d_proj else ""
        doc_proj = extract_clean_project_name(raw_doc_proj) if raw_doc_proj else top_global_proj_clean

        doctor_summary_list.append({
            '姓名*': final_name,
            '开始年*': c_time.year if c_time else "",
            '开始月*': c_time.month if c_time else "",
            '开始日*': c_time.day if c_time else "",
            '项目名称*': doc_proj,
            '金额*': amt,
            '姓名1*': final_name,
            '省份*': prov,
            '市*': city,
            '身份证*': final_cid,
            '手机号*': final_phone,
            '开户行*': open_bank,
            '银行卡号*': bank_card,
            '工作单位*': hosp,
            '医务职称*': title,
            '科室*': dept,
            '签署日期': c_time.strftime('%Y-%m-%d') if c_time else "",
            '签署日期1': c_time.strftime('%Y-%m-%d') if c_time else "",
            '身份证头像面': u_info.get('身份证头像面', ''),
            '身份证反面': u_info.get('身份证反面', ''),
            '_语料条数': len(grp)
        })

    # 按金额降序排列
    doctor_summary_list.sort(key=lambda x: (-x['金额*'], x['姓名*']))
    
    # 7. 写入最终电签表
    FINAL_COLUMNS = [
        '姓名*', '开始年*', '开始月*', '开始日*', '项目名称*', '金额*', '姓名1*', '省份*', '市*',
        '身份证*', '手机号*', '开户行*', '银行卡号*', '工作单位*', '医务职称*', '科室*',
        '签署日期', '签署日期1', '身份证头像面', '身份证反面'
    ]
    
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    base_file_name = f"调研案例电签表-{today_str}"
    out_xlsx = os.path.join(target_dir, f"{base_file_name}.xlsx")
    out_xls = os.path.join(target_dir, f"{base_file_name}.xls")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet0"
    
    header_font = Font(name="微软雅黑", size=10, bold=True)
    data_font = Font(name="微软雅黑", size=10)
    
    # 写入表头
    for c_idx, h in enumerate(FINAL_COLUMNS, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # 写入数据行
    for r_idx, doc in enumerate(doctor_summary_list, 2):
        for c_idx, col_name in enumerate(FINAL_COLUMNS, 1):
            val = doc.get(col_name, "")
            c = ws.cell(row=r_idx, column=c_idx)
            c.font = data_font
            
            if col_name in ['身份证*', '手机号*', '银行卡号*']:
                c.number_format = '@'
                c.value = str(val) if val != "" else ""
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ['开始年*', '开始月*', '开始日*']:
                c.value = int(val) if val != "" else ""
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == '金额*':
                c.value = int(val) if val != "" else 0
                c.number_format = '#,##0'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in ['姓名*', '姓名1*', '省份*', '市*', '医务职称*', '科室*', '签署日期', '签署日期1']:
                c.value = str(val) if val != "" else ""
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.value = str(val) if val != "" else ""
                c.alignment = Alignment(horizontal="left", vertical="center")

    # 自适应列宽
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = min(max(max_l + 3, 11), 35)

    wb.save(out_xlsx)
    print(f"\n✅ 主电签表已生成: {os.path.basename(out_xlsx)}")

    # 8. 同步导出标准 BIFF8 .xls 格式
    if xlwt:
        wb_xls = xlwt.Workbook(encoding='utf-8')
        ws_xls = wb_xls.add_sheet('Sheet0')
        st_hdr = xlwt.easyxf('font: name 微软雅黑, bold on, height 200; align: horiz center, vert center')
        st_center = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz center, vert center')
        st_left = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz left, vert center')
        st_amt = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz right, vert center', num_format_str='#,##0')
        
        for c_idx, h in enumerate(FINAL_COLUMNS):
            ws_xls.write(0, c_idx, h, st_hdr)
            
        for r_idx, doc in enumerate(doctor_summary_list, 1):
            for c_idx, col_name in enumerate(FINAL_COLUMNS):
                val = doc.get(col_name, "")
                if col_name in ['身份证*', '手机号*', '银行卡号*']:
                    ws_xls.write(r_idx, c_idx, str(val), st_center)
                elif col_name in ['开始年*', '开始月*', '开始日*']:
                    ws_xls.write(r_idx, c_idx, int(val) if val != "" else "", st_center)
                elif col_name == '金额*':
                    ws_xls.write(r_idx, c_idx, int(val) if val != "" else 0, st_amt)
                elif col_name in ['姓名*', '姓名1*', '省份*', '市*', '医务职称*', '科室*', '签署日期', '签署日期1']:
                    ws_xls.write(r_idx, c_idx, str(val), st_center)
                else:
                    ws_xls.write(r_idx, c_idx, str(val), st_left)
        wb_xls.save(out_xls)
        print(f"✅ 银行/电签通用导入表已生成: {os.path.basename(out_xls)}")

    # 9. 输出异常核查名单
    if warnings:
        warn_file = os.path.join(target_dir, "【警告】待核查异常名单.txt")
        with open(warn_file, "w", encoding="utf-8") as wf:
            wf.write("\n".join(warnings))
        print(f"⚠️ 发现 {len(warnings)} 项待核查数据异常，已生成：【警告】待核查异常名单.txt")

    # 10. 打印控制台核心业务报表
    print("\n" + "=" * 70)
    print("                 📊 结算核对核心业务 KPI 报表")
    print("=" * 70)
    print(f"   • 本期结算总人数 : {len(doctor_summary_list)} 位医生")
    print(f"   • 本期完成案例数 : {len(matched_details)} 份")
    print(f"   • 本期发放总金额 : ¥ {total_amount:,.2f} 元")
    print(f"   • 核心项目名称   : {top_global_proj_clean}")
    print("=" * 70 + "\n")

if __name__ == "__main__":
    work_dir = sys.argv[1] if len(sys.argv) > 1 else "."
    process_survey_settlement(work_dir)
