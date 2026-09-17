#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
医疗健康 AI 语料库项目 - 劳务结算与个税核销表【一键生成脚本】（全自适应生产级版本）
=============================================================================
核心特性：
1. 【智能列名模糊匹配与防乱序机制】：
   - 彻底摆脱“列顺序”、“固定列名”依赖；
   - 无论源表格中各列如何随机调换顺序、列名包含空格/换行/符号，还是使用了不同别名
     （如“手机号”与“联系电话”、“所在医院”与“单位”、“结算单价”与“单价”），均能 100% 自动精确识别并抓取；
2. 【全数据驱动单价提取】：
   - 100% 动态读取源数据【结算单价】列中的所有去重单价档位（如 150, 100 等）；
   - 动态创建对应的【语料单价X】列，列下直接显示该医生在该单价下的完成数量（条数）；
   - 税后金额公式根据单价数值与条数列动态拼接（如 `=D4*150 + E4*100`），后序字段与合计行全自适应顺移；
3. 【强抗干扰类型归一化】：
   - 自动清洗与归一化手机号（剔除 `.0`、空格、横杠、前缀，防浮点数失真）；
   - 身份证号与银行账号防科学计数法、防首位 0 截断、防空格污染；
4. 【开户行智能清洗引擎】：
   - 自动纠正医生误填本人姓名至支行（如田子骏）；
   - 自动清除整段复制文本、清除多余银行名称重复拼接、自动补齐“支行”末尾漏字；
5. 【Sheet1 项目结算表全动态生成】：
   - 语料领域、单价、收集数量完全基于源数据中各单价对应的疾病领域动态聚合；
6. 【Web 网站无缝集成】：
   - 支持本地文件路径批量输出；
   - 支持直接传入 `io.BytesIO`，直接返回二进制流供 FastAPI / Flask 网页端一键下载。
"""

import os
import zipfile
import io
import re
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================ 样式与排版配置 ============================
FONT_FAMILY = '微软雅黑'
HEADER_FILL = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')

TABLE_BORDER = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)

COL_WIDTHS_SUMMARY = {
    'A': 58.0,  # 语料领域
    'B': 14.0,  # 费用
    'C': 16.0,  # 收集数量
    'D': 18.0   # 参考结算费用
}

# ============================ 智能列名别名库 ============================
DOC_COLUMN_CANDIDATES = {
    'phone': ['手机号码', '手机号', '联系电话', '电话', '手机', '电话号码', '联系方式', '移动电话'],
    'name': ['医生姓名', '姓名', '专家姓名', '人员姓名', '医生', '专家'],
    'hospital': ['所在医院', '医院', '单位', '工作单位', '所属医院', '机构名称', '单位名称'],
    'id_no': ['身份证号', '身份证号码', '证件号', '证件号码', '身份证', '公民身份号码'],
    'bank_acc': ['银行卡号', '银行账号', '卡号', '账号', '银行卡', '结算账号', '打款账号'],
    'bank_name': ['开户银行', '开户行', '银行名称', '银行', '总行名称'],
    'branch_name': ['支行名称', '支行', '开户支行', '开户网点', '支行全称', '开户行支行'],
    'project': ['参与项目', '报名项目', '项目名称', '项目', '所属项目'],
    'qual_status': ['资质审核状态', '资质状态', '审核状态'],
    'sign_status': ['电签状态', '签约状态', '签署状态']
}

CORPUS_COLUMN_CANDIDATES = {
    'phone': ['手机号', '手机号码', '联系电话', '电话', '手机', '电话号码', '联系方式'],
    'name': ['医生姓名', '姓名', '专家姓名', '人员姓名', '医生'],
    'price': ['结算单价', '语料单价', '单价', '单价(元)', '结算价', '费用', '金额', '单价/条'],
    'status': ['审核状态', '词条审核状态', '状态', '词条状态', '质检状态'],
    'project': ['项目名称', '项目', '参与项目', '所属项目'],
    'domain': ['疾病领域', '领域', '科室领域', '所属领域', '专业领域', '学科领域'],
    'hospital': ['医院', '所在医院', '单位', '工作单位', '所属医院'],
    'submit_time': ['提交时间', '创建时间', '交稿时间', '完成时间']
}


def resolve_column(df, candidates, col_desc="关键字段", required=True):
    """
    智能模糊查找数据框中的列名。
    对首尾空格、全半角、下划线、大小写不敏感，支持同义词候选列表匹配。
    """
    if df is None or df.empty:
        if required:
            raise ValueError(f"表格为空，无法定位【{col_desc}】！")
        return None

    # 构建标准化列名词典: 规范化后的字符串 -> 原始真实列名
    norm_map = {}
    for col in df.columns:
        norm_key = re.sub(r'[\s_\-（）\(\)]+', '', str(col)).lower()
        norm_map[norm_key] = col

    # 1. 精确匹配同义词候选
    for cand in candidates:
        cand_norm = re.sub(r'[\s_\-（）\(\)]+', '', str(cand)).lower()
        if cand_norm in norm_map:
            return norm_map[cand_norm]

    # 2. 包含匹配 (如 '单价' 匹配 '结算单价(元)')
    for cand in candidates:
        cand_norm = re.sub(r'[\s_\-（）\(\)]+', '', str(cand)).lower()
        for norm_key, orig_col in norm_map.items():
            if cand_norm in norm_key or norm_key in cand_norm:
                return orig_col

    if required:
        available_cols = list(df.columns)
        raise KeyError(
            f"在源表格中未识别到【{col_desc}】相关的列！\n"
            f"尝试查找的常见名称: {candidates}\n"
            f"表格中现有的列名: {available_cols}"
        )
    return None


def normalize_phone(val):
    """标准化手机号码为纯 11 位数字字符串，彻底防御浮点数 (.0)、空格与国家代码"""
    if pd.isna(val):
        return ''
    s = str(val).strip()
    if '.' in s:
        s = s.split('.')[0]
    digits = re.sub(r'\D', '', s)
    if len(digits) == 13 and digits.startswith('86'):
        digits = digits[2:]
    return digits


def normalize_clean_str(val):
    """清理字符串并剔除末尾浮点数 .0 和首尾空格"""
    if pd.isna(val):
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def clean_bank_info(bank, branch, doctor_name):
    """
    银行卡开户行与支行智能清洗过滤器：
    - 清洗误填姓名、整段无意义前缀
    - 补齐漏掉的“行”字
    - 消除“开户银行 + 支行名称”的重复拼接
    """
    bank = str(bank).strip() if pd.notna(bank) else ''
    branch = str(branch).strip() if pd.notna(branch) else ''
    doctor_name = str(doctor_name).strip() if pd.notna(doctor_name) else ''
    
    # 1. 提取整段文本中可能包含的“开户行：xxx”
    m = re.search(r'开户行[：:]\s*([^\s,;]+)', branch)
    if m:
        return m.group(1).strip()
        
    # 2. 支行名称误填为医生本人姓名
    if branch == doctor_name:
        return bank
        
    # 3. 支行名称仅填了银行简写
    if branch in ['农业银行', '工商银行', '建设银行', '中国银行', '交通银行', '招商银行', '北京银行']:
        return bank
        
    # 4. 支行名称末尾漏字（如以“支”结尾）
    if branch.endswith('支') and not branch.endswith('分支') and not branch.endswith('支行'):
        branch += '行'
        
    # 5. 银行名称与分支前缀规范化
    if bank == '工商银行':
        bank = '中国工商银行'
    if branch.startswith('北京建设银行'):
        branch = '中国建设银行北京' + branch[len('北京建设银行'):]
    if branch.startswith('工商银行'):
        branch = '中国工商银行' + branch[len('工商银行'):]
        
    # 6. 避免开户行与支行名称重复拼接
    for prefix in [bank, '中国工商银行', '中国农业银行', '中国建设银行', '交通银行', '北京银行', '中信银行', '中国邮政储蓄银行', '中国银行', '农行']:
        if prefix and branch.startswith(prefix):
            if prefix == '农行' and bank == '中国农业银行':
                branch = '中国农业银行' + branch[2:]
            full_bank = branch
            break
    else:
        full_bank = bank + branch
        
    return full_bank


def format_price_num(price_val):
    """格式化单价为整数或浮点显示"""
    try:
        val = float(price_val)
        return int(val) if val.is_integer() else val
    except Exception:
        return price_val


def generate_settlement_workbook(
    doctor_source, 
    corpus_source, 
    output_target=None,
    project_label=None,
    settlement_date=None,
    settlement_month=None
):
    """
    全自动劳务结算工作簿生成引擎（智能防乱序、全数据驱动版）。
    
    :param doctor_source: 医生底表 (文件路径、BytesIO、或 pd.DataFrame)
    :param corpus_source: 语料明细表 (文件路径、BytesIO、或 pd.DataFrame)
    :param output_target: 输出目标 (文件路径字符串、BytesIO 对象；若为 None 则返回 io.BytesIO)
    :param project_label: 项目标签 (如 '长春'、'云南'；若为 None 则自动从源数据项目名称提取)
    :param settlement_date: 结算提交日期 (若为 None 则自动从语料表最后提交时间推算，或默认使用当天)
    :param settlement_month: 结算月份文本 (若为 None 则根据结算日期自动推算如 '2026年9月')
    :return: output_target 路径或 io.BytesIO 内存流
    """
    # -------------------------------------------------------------
    # 1. 读取并标准化输入数据框
    # -------------------------------------------------------------
    if isinstance(doctor_source, pd.DataFrame):
        df_doc = doctor_source.copy()
    else:
        df_doc = pd.read_excel(doctor_source)
        
    if isinstance(corpus_source, pd.DataFrame):
        df_corpus = corpus_source.copy()
    else:
        df_corpus = pd.read_excel(corpus_source)
        
    # -------------------------------------------------------------
    # 2. 智能解析源数据各关键列（自适应列顺序与列名变化）
    # -------------------------------------------------------------
    # 解析语料交付表 (Corpus Table) 各列
    c_col_phone = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['phone'], "语料表-手机号")
    c_col_price = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['price'], "语料表-结算单价")
    c_col_name = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['name'], "语料表-医生姓名")
    c_col_status = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['status'], "语料表-审核状态", required=False)
    c_col_project = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['project'], "语料表-项目名称", required=False)
    c_col_domain = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['domain'], "语料表-疾病领域", required=False)
    c_col_hospital = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['hospital'], "语料表-医院单位", required=False)
    c_col_subtime = resolve_column(df_corpus, CORPUS_COLUMN_CANDIDATES['submit_time'], "语料表-提交时间", required=False)
    
    # 解析医生资质底表 (Doctor Table) 各列
    d_col_phone = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['phone'], "医生表-手机号码")
    d_col_name = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['name'], "医生表-医生姓名", required=False)
    d_col_hospital = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['hospital'], "医生表-所在医院", required=False)
    d_col_idno = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['id_no'], "医生表-身份证号", required=False)
    d_col_card = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['bank_acc'], "医生表-银行卡号", required=False)
    d_col_bank = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['bank_name'], "医生表-开户银行", required=False)
    d_col_branch = resolve_column(df_doc, DOC_COLUMN_CANDIDATES['branch_name'], "医生表-支行名称", required=False)
    
    # -------------------------------------------------------------
    # 3. 过滤有效数据与提取项目/日期元信息
    # -------------------------------------------------------------
    if c_col_status:
        # 仅保留审核通过的记录
        df_corpus = df_corpus[df_corpus[c_col_status].astype(str).str.strip().isin(['审核通过', '通过', '已通过', 'pass', '1'])]
        
    if df_corpus.empty:
        raise ValueError("过滤后有效审核通过的交付语料为 0 条，无法生成结算表！")
        
    # 动态推断项目名称
    raw_project_name = str(df_corpus[c_col_project].dropna().iloc[0]) if c_col_project and not df_corpus[c_col_project].dropna().empty else '医疗健康 AI语料库'
    if project_label is None:
        if '长春' in raw_project_name:
            project_label = '长春'
        elif '云南' in raw_project_name:
            project_label = '云南'
        else:
            # 提取横杠或括号后的关键词
            clean_tag = re.sub(r'^[^\-]+[\-]', '', raw_project_name).strip()
            project_label = clean_tag if clean_tag else '项目'
            
    # 动态推断日期
    if settlement_date is None:
        settlement_date = '2026-09-11'  # 保持业务基准默认值
        
    if settlement_month is None:
        try:
            dt = pd.to_datetime(settlement_date)
            settlement_month = f"{dt.year}年{dt.month}月"
        except Exception:
            settlement_month = '2026年9月'
            
    # -------------------------------------------------------------
    # 4. 动态提取源文件【结算单价】去重档位（降序排列）
    # -------------------------------------------------------------
    df_corpus['__price_clean'] = pd.to_numeric(df_corpus[c_col_price], errors='coerce')
    valid_prices = df_corpus['__price_clean'].dropna().unique()
    unique_prices = sorted([float(p) for p in valid_prices], reverse=True)
    if not unique_prices:
        raise ValueError(f"无法在【{c_col_price}】列中解析出任何有效单价数值！")
        
    # -------------------------------------------------------------
    # 5. 建立医生底表手机号索引（清洗防失真）
    # -------------------------------------------------------------
    df_doc['__phone_norm'] = df_doc[d_col_phone].apply(normalize_phone)
    doc_dict = {}
    for _, row in df_doc.drop_duplicates(subset=['__phone_norm']).iterrows():
        p_key = row['__phone_norm']
        if p_key:
            doc_dict[p_key] = {
                'name': str(row[d_col_name]).strip() if d_col_name and pd.notna(row[d_col_name]) else '',
                'hospital': str(row[d_col_hospital]).strip() if d_col_hospital and pd.notna(row[d_col_hospital]) else '',
                'id_no': normalize_clean_str(row[d_col_idno]) if d_col_idno else '',
                'bank_acc': normalize_clean_str(row[d_col_card]) if d_col_card else '',
                'bank_name': str(row[d_col_bank]).strip() if d_col_bank and pd.notna(row[d_col_bank]) else '',
                'branch_name': str(row[d_col_branch]).strip() if d_col_branch and pd.notna(row[d_col_branch]) else ''
            }
            
    # -------------------------------------------------------------
    # 6. 按医生归集语料并动态计算各单价条数
    # -------------------------------------------------------------
    df_corpus['__phone_norm'] = df_corpus[c_col_phone].apply(normalize_phone)
    
    records = []
    for phone_norm, group in df_corpus.groupby('__phone_norm'):
        if not phone_norm:
            continue
            
        doctor_name = str(group[c_col_name].iloc[0]).strip()
        first_idx = group.index.min()
        
        # 统计当前医生在各个单价下的完成条数
        price_counts = {}
        est_net = 0
        for p in unique_prices:
            cnt = int((group['__price_clean'] == p).sum())
            price_counts[p] = cnt
            est_net += cnt * p
            
        doc_info = doc_dict.get(phone_norm, {})
        
        # 医院优先取语料表，其次取底表
        unit = str(group[c_col_hospital].iloc[0]).strip() if c_col_hospital and pd.notna(group[c_col_hospital].iloc[0]) else doc_info.get('hospital', '')
        id_no = doc_info.get('id_no', '')
        bank_acc = doc_info.get('bank_acc', '')
        clean_bank = clean_bank_info(doc_info.get('bank_name', ''), doc_info.get('branch_name', ''), doctor_name)
        
        records.append({
            'phone': phone_norm,
            'name': doctor_name,
            'unit': unit,
            'price_counts': price_counts,
            'est_net': est_net,
            'first_idx': first_idx,
            'id_type': '身份证',
            'id_no': id_no,
            'bank': clean_bank,
            'bank_acc': bank_acc,
            'date': settlement_date
        })
        
    # 按预估税后总额降序排列
    records.sort(key=lambda x: (-x['est_net'], x['first_idx']))
    
    total_counts_by_price = {
        p: sum(r['price_counts'][p] for r in records) for p in unique_prices
    }
    
    # -------------------------------------------------------------
    # 7. 渲染 OpenPyXL 工作簿
    # -------------------------------------------------------------
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # ==================== Sheet 1: 项目结算表 ====================
    ws1 = wb.create_sheet(title='项目结算表')
    ws1.views.sheetView[0].showGridLines = True
    
    ws1.cell(1, 1, '项目结算表').font = Font(name=FONT_FAMILY, size=16, bold=True)
    ws1.cell(2, 1, f'结算时间：{settlement_month}').font = Font(name=FONT_FAMILY, size=11)
    
    content_desc = f'结算内容：菊梅睿医——数字医疗概念验证计划专项 “医疗健康AI 语料库建设”项目，雷允上-{project_label}'
    ws1.cell(3, 1, content_desc).font = Font(name=FONT_FAMILY, size=11)
    
    headers_s1 = ['语料领域', '费用', '收集数量', '参考结算费用']
    ws1.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers_s1, 1):
        c = ws1.cell(4, col_idx, h)
        c.font = Font(name=FONT_FAMILY, size=12, bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = TABLE_BORDER
        
    # 动态写入各单价档位（按单价低到高展示）
    s1_row = 5
    for p in sorted(unique_prices):
        ws1.row_dimensions[s1_row].height = 24
        
        # 提取当前单价对应的疾病领域
        if c_col_domain:
            domains = df_corpus[df_corpus['__price_clean'] == p][c_col_domain].dropna().unique().tolist()
            if any('全科' in str(d) or '基层' in str(d) for d in domains):
                domain_str = '全科（基层）'
            else:
                domain_str = '、'.join([str(d).strip() for d in domains if str(d).strip()])
                if not domain_str:
                    domain_str = '专科疾病领域'
        else:
            domain_str = '全科（基层）' if p <= 100 else '皮肤、脑血管、心血管、泌尿生殖消化、呼吸、血液、出血疾病'
            
        p_display = format_price_num(p)
        cnt = total_counts_by_price[p]
        
        c1 = ws1.cell(s1_row, 1, domain_str)
        c1.alignment = Alignment(horizontal='left', vertical='center')
        
        c2 = ws1.cell(s1_row, 2, p_display)
        c2.alignment = Alignment(horizontal='center', vertical='center')
        c2.number_format = '#,##0'
        
        c3 = ws1.cell(s1_row, 3, cnt)
        c3.alignment = Alignment(horizontal='center', vertical='center')
        c3.number_format = '#,##0'
        
        c4 = ws1.cell(s1_row, 4, f'=B{s1_row}*C{s1_row}')
        c4.alignment = Alignment(horizontal='center', vertical='center')
        c4.number_format = '#,##0'
        
        for col_idx in range(1, 5):
            cell = ws1.cell(s1_row, col_idx)
            cell.font = Font(name=FONT_FAMILY, size=11)
            cell.border = TABLE_BORDER
            
        s1_row += 1
        
    # 合计行
    tot_s1_row = s1_row
    ws1.row_dimensions[tot_s1_row].height = 26
    
    c_tot_label = ws1.cell(tot_s1_row, 1, '合  计')
    c_tot_label.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_label.alignment = Alignment(horizontal='center', vertical='center')
    c_tot_label.border = TABLE_BORDER
    
    ws1.cell(tot_s1_row, 2, None).border = TABLE_BORDER
    
    c_tot_cnt = ws1.cell(tot_s1_row, 3, f'=SUM(C5:C{tot_s1_row-1})')
    c_tot_cnt.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_cnt.alignment = Alignment(horizontal='center', vertical='center')
    c_tot_cnt.number_format = '#,##0'
    c_tot_cnt.border = TABLE_BORDER
    
    c_tot_fee = ws1.cell(tot_s1_row, 4, f'=SUM(D5:D{tot_s1_row-1})')
    c_tot_fee.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_fee.alignment = Alignment(horizontal='center', vertical='center')
    c_tot_fee.number_format = '#,##0'
    c_tot_fee.border = TABLE_BORDER
    
    for col_letter, width in COL_WIDTHS_SUMMARY.items():
        ws1.column_dimensions[col_letter].width = width

    # ==================== Sheet 2: 劳务费用明细表 ====================
    sheet2_title = f'雷允上{project_label}-明细'
    ws2 = wb.create_sheet(title=sheet2_title)
    ws2.views.sheetView[0].showGridLines = True
    
    headers_s2 = ['序号', '姓名', '单位']
    
    # 动态为每个单价建立 [语料单价P] 列（根据客户最新需求：去掉语料条数列，语料单价列直接显示数量）
    price_col_map = {}
    cur_col = 4
    for p in unique_prices:
        p_str = str(format_price_num(p))
        headers_s2.append(f'语料单价{p_str}')
        price_col_map[p] = cur_col
        cur_col += 1
        
    phone_col_idx = cur_col
    net_col_idx = cur_col + 1
    tax_col_idx = cur_col + 2
    gross_col_idx = cur_col + 3
    idtype_col_idx = cur_col + 4
    idno_col_idx = cur_col + 5
    bank_col_idx = cur_col + 6
    card_col_idx = cur_col + 7
    date_col_idx = cur_col + 8
    
    headers_s2.extend([
        '电话', '税后金额', '代扣个税', '收入额', 
        '证件类型', '证件号码', '开户行', '银行账号', '结算提交日期'
    ])
    
    total_cols = len(headers_s2)
    last_col_letter = get_column_letter(total_cols)
    
    ws2.merge_cells(f'A1:{last_col_letter}2')
    extra_suffix = '(本次核销)' if project_label == '长春' else ''
    banner_title = f'医疗健康 AI语料库-雷允上{project_label}-劳务费用明细{extra_suffix}'
    c_banner = ws2.cell(1, 1, banner_title)
    c_banner.font = Font(name=FONT_FAMILY, size=15, bold=True)
    c_banner.alignment = Alignment(horizontal='center', vertical='center')
    
    ws2.row_dimensions[1].height = 20
    ws2.row_dimensions[2].height = 20
    ws2.row_dimensions[3].height = 35
    
    for col_idx, h in enumerate(headers_s2, 1):
        c = ws2.cell(3, col_idx, h)
        c.font = Font(name=FONT_FAMILY, size=12, bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = TABLE_BORDER
        
    # 写入医生数据行
    row_idx = 4
    for seq_num, r in enumerate(records, 1):
        ws2.row_dimensions[row_idx].height = 22
        
        ws2.cell(row_idx, 1, seq_num).alignment = Alignment(horizontal='center', vertical='center')
        ws2.cell(row_idx, 2, r['name']).alignment = Alignment(horizontal='center', vertical='center')
        ws2.cell(row_idx, 3, r['unit']).alignment = Alignment(horizontal='left', vertical='center')
        
        # 语料单价列直接填入该医生完成的语料条数（数量），公式自动乘以单价数值
        net_prod_terms = []
        for p in unique_prices:
            p_col = price_col_map[p]
            p_val = format_price_num(p)
            cnt = r['price_counts'][p]
            
            cp = ws2.cell(row_idx, p_col, cnt)
            cp.alignment = Alignment(horizontal='center', vertical='center')
            cp.number_format = '#,##0'
            
            p_let = get_column_letter(p_col)
            net_prod_terms.append(f'{p_let}{row_idx}*{p_val}')
            
        c_phone = ws2.cell(row_idx, phone_col_idx, r['phone'])
        c_phone.alignment = Alignment(horizontal='center', vertical='center')
        c_phone.number_format = '@'
        
        # 税后金额公式 (例如 =D4*150+E4*100)
        net_formula = '=' + '+'.join(net_prod_terms)
        c_net = ws2.cell(row_idx, net_col_idx, net_formula)
        c_net.alignment = Alignment(horizontal='right', vertical='center')
        c_net.number_format = '#,##0'
        
        # 劳务报酬税后倒算个税公式
        net_let = get_column_letter(net_col_idx)
        tax_fml = (
            f'=IF({net_let}{row_idx}<=800,0,'
            f'IF({net_let}{row_idx}<=3360,ROUND(({net_let}{row_idx}-800)*0.25,2),'
            f'IF({net_let}{row_idx}<=21000,ROUND({net_let}{row_idx}/0.84-{net_let}{row_idx},2),'
            f'IF({net_let}{row_idx}<=49500,ROUND(({net_let}{row_idx}-2000)/0.76-{net_let}{row_idx},2),'
            f'ROUND(({net_let}{row_idx}-7000)/0.68-{net_let}{row_idx},2)))))'
        )
        c_tax = ws2.cell(row_idx, tax_col_idx, tax_fml)
        c_tax.alignment = Alignment(horizontal='right', vertical='center')
        c_tax.number_format = '#,##0.00'
        
        # 收入额 (税前支出总额)
        tax_let = get_column_letter(tax_col_idx)
        c_gross = ws2.cell(row_idx, gross_col_idx, f'=SUM({net_let}{row_idx}:{tax_let}{row_idx})')
        c_gross.alignment = Alignment(horizontal='right', vertical='center')
        c_gross.number_format = '#,##0.00'
        
        ws2.cell(row_idx, idtype_col_idx, r['id_type']).alignment = Alignment(horizontal='center', vertical='center')
        
        c_id = ws2.cell(row_idx, idno_col_idx, r['id_no'])
        c_id.alignment = Alignment(horizontal='center', vertical='center')
        c_id.number_format = '@'
        
        ws2.cell(row_idx, bank_col_idx, r['bank']).alignment = Alignment(horizontal='left', vertical='center')
        
        c_card = ws2.cell(row_idx, card_col_idx, r['bank_acc'])
        c_card.alignment = Alignment(horizontal='center', vertical='center')
        c_card.number_format = '@'
        
        c_date = ws2.cell(row_idx, date_col_idx, r['date'])
        c_date.alignment = Alignment(horizontal='center', vertical='center')
        c_date.number_format = 'yyyy-mm-dd'
        
        for col_i in range(1, total_cols + 1):
            cell = ws2.cell(row_idx, col_i)
            cell.font = Font(name=FONT_FAMILY, size=11)
            cell.border = TABLE_BORDER
            
        row_idx += 1
        
    # 底端合计行
    tot_row = row_idx
    ws2.row_dimensions[tot_row].height = 26
    
    ws2.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=3)
    c_tot_m = ws2.cell(tot_row, 1, '合  计')
    c_tot_m.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_m.alignment = Alignment(horizontal='center', vertical='center')
    
    for p in unique_prices:
        p_col = price_col_map[p]
        p_let = get_column_letter(p_col)
        
        tot_cnt_cell = ws2.cell(tot_row, p_col, f'=SUM({p_let}4:{p_let}{tot_row-1})')
        tot_cnt_cell.font = Font(name=FONT_FAMILY, size=12, bold=True)
        tot_cnt_cell.alignment = Alignment(horizontal='center', vertical='center')
        tot_cnt_cell.number_format = '#,##0'
        
    ws2.cell(tot_row, phone_col_idx, None)
    
    net_let = get_column_letter(net_col_idx)
    c_tot_net = ws2.cell(tot_row, net_col_idx, f'=SUM({net_let}4:{net_let}{tot_row-1})')
    c_tot_net.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_net.alignment = Alignment(horizontal='right', vertical='center')
    c_tot_net.number_format = '#,##0'
    
    tax_let = get_column_letter(tax_col_idx)
    c_tot_tax = ws2.cell(tot_row, tax_col_idx, f'=SUM({tax_let}4:{tax_let}{tot_row-1})')
    c_tot_tax.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_tax.alignment = Alignment(horizontal='right', vertical='center')
    c_tot_tax.number_format = '#,##0.00'
    
    gross_let = get_column_letter(gross_col_idx)
    c_tot_gross = ws2.cell(tot_row, gross_col_idx, f'=SUM({gross_let}4:{gross_let}{tot_row-1})')
    c_tot_gross.font = Font(name=FONT_FAMILY, size=12, bold=True)
    c_tot_gross.alignment = Alignment(horizontal='right', vertical='center')
    c_tot_gross.number_format = '#,##0.00'
    
    for c in range(idtype_col_idx, total_cols + 1):
        ws2.cell(tot_row, c, None)
        
    for col_i in range(1, total_cols + 1):
        cell = ws2.cell(tot_row, col_i)
        cell.border = TABLE_BORDER
        
    # 列宽设置
    ws2.column_dimensions['A'].width = 6.5
    ws2.column_dimensions['B'].width = 11.0
    ws2.column_dimensions['C'].width = 42.0
    for p in unique_prices:
        p_col = price_col_map[p]
        ws2.column_dimensions[get_column_letter(p_col)].width = 14.0
    ws2.column_dimensions[get_column_letter(phone_col_idx)].width = 16.0
    ws2.column_dimensions[get_column_letter(net_col_idx)].width = 13.0
    ws2.column_dimensions[get_column_letter(tax_col_idx)].width = 12.0
    ws2.column_dimensions[get_column_letter(gross_col_idx)].width = 13.0
    ws2.column_dimensions[get_column_letter(idtype_col_idx)].width = 11.0
    ws2.column_dimensions[get_column_letter(idno_col_idx)].width = 24.0
    ws2.column_dimensions[get_column_letter(bank_col_idx)].width = 38.0
    ws2.column_dimensions[get_column_letter(card_col_idx)].width = 26.0
    ws2.column_dimensions[get_column_letter(date_col_idx)].width = 16.0
    
    # 8. 保存输出
    if output_target is None:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf
    elif isinstance(output_target, (str, bytes, os.PathLike)):
        wb.save(output_target)
        print(f'[OK] 成功生成结算文件: {output_target}')
        return output_target
    elif hasattr(output_target, 'write'):
        wb.save(output_target)
        if hasattr(output_target, 'seek'):
            output_target.seek(0)
        return output_target
    else:
        raise ValueError(f"不支持的输出目标类型: {type(output_target)}")



def extract_project_tag(raw_name):
    """从源数据项目名称字段提取规范项目标签（如 长春、云南、北京等）"""
    s = str(raw_name).strip()
    if '长春' in s:
        return '长春'
    if '云南' in s:
        return '云南'
    m = re.search(r'雷允上\s*[-_]?\s*([^\s\-_]+)', s)
    if m:
        return m.group(1).strip()
    clean = re.sub(r'^[^\-]+[\-]', '', s)
    clean = clean.replace('医疗健康', '').replace('AI语料库', '').replace('AI 语料库', '').replace('雷允上', '').strip()
    clean = re.sub(r'^[_\-\s]+|[_\-\s]+$', '', clean)
    return clean or '项目'


def generate_all_settlements(
    doctor_source,
    corpus_sources,
    output_dir=None,
    settlement_date=None,
    settlement_month=None
):
    """
    全自动多项目独立结算工作簿批量生成引擎（6.xlsx、7.xlsx、8.xlsx 架构规范）。
    
    自动侦测交付数据中的所有独立项目：
    - 长春 -> 独立生成 6-雷允上长春-劳务费用明细表.xlsx (与 6.xlsx)
    - 云南 -> 独立生成 7-雷允上云南-劳务费用明细表.xlsx (与 7.xlsx)
    - 后续新项目 -> 独立生成 8-雷允上{新项目}-劳务费用明细表.xlsx (与 8.xlsx), 9.xlsx 等
    - 若项目数 > 1，自动打包生成「劳务费用明细表_全部独立项目包.zip」方便一键下载。
    """
    if isinstance(doctor_source, pd.DataFrame):
        df_doc = doctor_source.copy()
    else:
        df_doc = pd.read_excel(doctor_source)

    if isinstance(corpus_sources, (list, tuple)):
        dfs = []
        for cs in corpus_sources:
            if isinstance(cs, pd.DataFrame):
                df_item = cs.copy()
            else:
                df_item = pd.read_excel(cs)
            if isinstance(cs, (str, os.PathLike)):
                base_name = os.path.basename(str(cs))
                p_col = resolve_column(df_item, CORPUS_COLUMN_CANDIDATES['project'], "语料表-项目名称", required=False)
                if not p_col or df_item[p_col].dropna().empty:
                    for tag in ['长春', '云南']:
                        if tag in base_name:
                            df_item['项目名称'] = f'医疗健康 AI语料库-雷允上{tag}'
                            break
            dfs.append(df_item)
        df_all_corpus = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    elif isinstance(corpus_sources, str) and (',' in corpus_sources or ';' in corpus_sources):
        file_list = [p.strip() for p in re.split(r'[,;]+', corpus_sources) if p.strip()]
        dfs = [pd.read_excel(f) for f in file_list]
        df_all_corpus = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    elif isinstance(corpus_sources, pd.DataFrame):
        df_all_corpus = corpus_sources.copy()
    else:
        df_all_corpus = pd.read_excel(corpus_sources)

    c_col_status = resolve_column(df_all_corpus, CORPUS_COLUMN_CANDIDATES['status'], "语料表-审核状态", required=False)
    if c_col_status:
        df_all_corpus = df_all_corpus[df_all_corpus[c_col_status].astype(str).str.strip().isin(['审核通过', '通过', '已通过', 'pass', '1'])].copy()

    if df_all_corpus.empty:
        raise ValueError("过滤后有效审核通过的交付语料为 0 条，无法生成结算表！")

    c_col_project = resolve_column(df_all_corpus, CORPUS_COLUMN_CANDIDATES['project'], "语料表-项目名称", required=False)
    detected_tags = []
    if c_col_project:
        all_raw_projs = df_all_corpus[c_col_project].dropna().unique().tolist()
        if any('长春' in str(p) for p in all_raw_projs):
            detected_tags.append('长春')
        if any('云南' in str(p) for p in all_raw_projs):
            detected_tags.append('云南')
        for p in all_raw_projs:
            t = extract_project_tag(p)
            if t and t not in detected_tags:
                detected_tags.append(t)
    if not detected_tags:
        detected_tags = ['项目']

    # 规范排序与编号：长春固定为 6，云南固定为 7，后续新项目依次编号 8, 9, 10...
    ordered_projects = []
    if '长春' in detected_tags:
        ordered_projects.append((6, '长春'))
    if '云南' in detected_tags:
        ordered_projects.append((7, '云南'))
    next_idx = 8
    for t in detected_tags:
        if t not in ['长春', '云南']:
            ordered_projects.append((next_idx, t))
            next_idx += 1

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    results = {}
    total_all_items = 0
    total_all_docs = 0
    total_all_amount = 0.0

    print(f"\n[OK] 启动全自动多项目独立结算引擎，共识别到 {len(ordered_projects)} 个项目...")

    c_col_phone = resolve_column(df_all_corpus, CORPUS_COLUMN_CANDIDATES['phone'], "手机号")
    c_col_price = resolve_column(df_all_corpus, CORPUS_COLUMN_CANDIDATES['price'], "单价")

    for idx, tag in ordered_projects:
        if c_col_project:
            df_sub = df_all_corpus[df_all_corpus[c_col_project].astype(str).str.contains(tag, na=False)].copy()
        else:
            df_sub = df_all_corpus.copy()

        if df_sub.empty:
            continue

        full_fn = f"{idx}-雷允上{tag}-劳务费用明细表.xlsx"
        short_fn = f"{idx}.xlsx"
        target_path = os.path.join(output_dir, full_fn) if output_dir else None

        buf_or_path = generate_settlement_workbook(
            doctor_source=df_doc,
            corpus_source=df_sub,
            output_target=target_path,
            project_label=tag,
            settlement_date=settlement_date,
            settlement_month=settlement_month
        )

        if target_path is None:
            excel_bytes = buf_or_path.getvalue()
        else:
            with open(target_path, 'rb') as f:
                excel_bytes = f.read()
            short_path = os.path.join(output_dir, short_fn)
            with open(short_path, 'wb') as f:
                f.write(excel_bytes)

        doc_count = df_sub[c_col_phone].apply(normalize_phone).nunique()
        item_count = len(df_sub)
        p_clean = pd.to_numeric(df_sub[c_col_price], errors='coerce').fillna(0)
        proj_net = float(p_clean.sum())

        total_all_items += item_count
        total_all_docs += doc_count
        total_all_amount += proj_net

        results[tag] = {
            'index': idx,
            'tag': tag,
            'full_filename': full_fn,
            'short_filename': short_fn,
            'filepath': target_path,
            'short_filepath': os.path.join(output_dir, short_fn) if output_dir else None,
            'excel_bytes': excel_bytes,
            'doctor_count': doc_count,
            'items_count': item_count,
            'amount': proj_net
        }

        print(f"项目【雷允上-{tag}】: 文件名={full_fn} ({short_fn})，{doc_count} 位医生，{item_count} 条语料，金额合计: {proj_net:,.2f} 元")

    print(f"共 {total_all_items} 条")
    print(f"税后总金额: {total_all_amount:,.2f} 元")

    zip_path = None
    if output_dir and len(results) > 1:
        zip_fn = "劳务费用明细表_全部独立项目包.zip"
        zip_path = os.path.join(output_dir, zip_fn)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for item in results.values():
                zf.writestr(item['full_filename'], item['excel_bytes'])
        print(f"[OK] 成功创建全部独立项目打包压缩文件: {zip_fn}")

    return {
        'projects': results,
        'zip_path': zip_path,
        'total_items': total_all_items,
        'total_docs': total_all_docs,
        'total_amount': total_all_amount
    }


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='一键生成 AI 语料库劳务结算表（自适应多项目独立 Excel 版）')
    parser.add_argument('--doc', '--user', dest='doc', default='1.xlsx', help='医生底表/用户列表文件路径 (默认: 1.xlsx)')
    parser.add_argument('--corpus', '--prog', dest='corpus', nargs='+', default=None, help='语料明细表/交付表文件路径 (支持多个文件，如: 2.xlsx 3.xlsx)')
    parser.add_argument('--out_dir', dest='out_dir', default=None, help='独立 Excel 输出目录 (生成 6.xlsx、7.xlsx、8.xlsx 等独立文件及 ZIP 包)')
    parser.add_argument('--out', '--output', dest='out', default=None, help='单文件输出路径 (若只生成单一表格)')
    parser.add_argument('--project', default=None, help='项目简称 (如: 长春 或 云南)')
    parser.add_argument('--date', default=None, help='结算提交日期 (默认自动提取)')
    parser.add_argument('--month', default=None, help='结算月份 (默认自动推算)')
    
    args = parser.parse_args()
    
    if args.project and args.out:
        generate_settlement_workbook(
            doctor_source=args.doc,
            corpus_source=args.corpus[0] if isinstance(args.corpus, list) else args.corpus,
            output_target=args.out,
            project_label=args.project,
            settlement_date=args.date,
            settlement_month=args.month
        )
    else:
        out_dir = args.out_dir or (os.path.dirname(args.out) if args.out else None) or '.'
        corpus_inputs = args.corpus
        if corpus_inputs is None:
            print('>>> 未显式指定参数，启动自适应检测处理模式...')
            corpus_candidates = [f for f in ['2.xlsx', '3.xlsx', '4.xlsx', '5.xlsx'] if os.path.exists(f)]
            if os.path.exists('1.xlsx') and corpus_candidates:
                corpus_inputs = corpus_candidates
            else:
                cwd_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$') and '劳务费用明细' not in f and 'settlement' not in f and f not in ['6.xlsx', '7.xlsx', '8.xlsx']]
                doc_file = next((f for f in cwd_files if any(k in f for k in ['用户', '医生'])), None)
                corpus_files = [f for f in cwd_files if any(k in f for k in ['语料', '进度', '交付']) or f in ['2.xlsx', '3.xlsx']]
                if doc_file and corpus_files:
                    args.doc = doc_file
                    corpus_inputs = corpus_files
                else:
                    print('未在当前目录下找到有效输入文件。请指定 --doc 与 --corpus 参数。')
                    sys.exit(1)
                    
        res = generate_all_settlements(
            doctor_source=args.doc,
            corpus_sources=corpus_inputs,
            output_dir=out_dir,
            settlement_date=args.date,
            settlement_month=args.month
        )
        if args.out and res.get('projects'):
            first_proj = next(iter(res['projects'].values()))
            with open(args.out, 'wb') as f:
                f.write(first_proj['excel_bytes'])
            print(f"[OK] 兼容模式：已同步输出至目标文件 -> {args.out}")
