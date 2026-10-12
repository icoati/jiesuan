# -*- coding: utf-8 -*-
"""
=============================================================================
北检结算表生成引擎 - HCP自行承担税费模式 (Beijian Settlement Engine - HCP Tax)
=============================================================================
功能说明：
1. 模式：HCP自行承担税费（正算扣税 / 劳务报酬法定预扣预缴）
   - 任务金额直接作为税前应发金额 (Gross Pay)
   - 专家依法自行承担个税（应发 <= 800 免税；800~4000 扣除 800 费用按 20% 计税；>4000 扣除 20% 费用按超额累进计税）
   - 税后实发金额 (Net Pay) = 税前应发 - 代扣税金
2. 接收 3 张核心源表：
   - 1.xlsx (待结算词条编号名单)
   - 2.xlsx (专家银行卡与身份资质表)
   - 3.xlsx (语料库全量审核明细表)
3. 关联匹配与规范化输出 3 张标准化交付表单：
   - HCP自行承担税费-YYYY-MM-DD.xlsx (基于 4.xlsx 模板格式，正算税前应发、代扣个税、税后实发)
   - 8_已结算语料词条明细表.xlsx (基于 5.xlsx 模板格式，按结算单价标记结算状态并保留原始题目/回答记录)
   - 9_作品劳务结算总表.xlsx (基于 6.xlsx 模板格式，作品维度单价与乘法关联核算)
=============================================================================
"""

import os
import sys
import argparse
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

thin_border = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF')
)

def find_col(df, candidates, default=None):
    """在 DataFrame 列中模糊匹配候选字段名"""
    cols = [str(c).strip() for c in df.columns]
    for cand in candidates:
        for c in cols:
            if cand == c:
                return c
    for cand in candidates:
        for c in cols:
            if cand in c:
                return c
    return default

def calculate_tax(gross_pay):
    """
    HCP自行承担税费模式（正算扣税 / 劳务报酬法定预扣预缴）：
    任务金额作为税前应发金额 (Gross Pay)
    专家自行承担税金，税金从应发中扣除，实际到手为税后实发 (Net Pay)

    规则：
    1. 收入 <= 800 元：免征个税，税额 = 0，实发 = 收入
    2. 800 < 收入 <= 4,000 元：应纳税所得额 = 收入 - 800，预扣率 20%，速算扣除数 0
       例如：收入 1000 元，应纳税所得额 = 200 元，税金 = 40 元，实发 = 960 元
    3. 收入 > 4,000 元：应纳税所得额 = 收入 * (1 - 20%) = 收入 * 0.8
       - 应纳税所得额 <= 20,000 元：预扣率 20%，速算扣除数 0
       - 20,000 < 应纳税所得额 <= 50,000 元：预扣率 30%，速算扣除数 2,000
       - 应纳税所得额 > 50,000 元：预扣率 40%，速算扣除数 7,000
    返回: gross_pay, tax, net_pay
    """
    gross = round(float(gross_pay), 2)
    if gross <= 800.0:
        tax = 0.0
    elif gross <= 4000.0:
        tax = (gross - 800.0) * 0.20
    else:
        taxable = gross * 0.80
        if taxable <= 20000.0:
            tax = taxable * 0.20
        elif taxable <= 50000.0:
            tax = taxable * 0.30 - 2000.0
        else:
            tax = taxable * 0.40 - 7000.0
    tax = round(tax, 2)
    net_pay = round(gross - tax, 2)
    return gross, tax, net_pay

def process_beijian_kangbei(
    file_1,                 # 1.xlsx (待结算词条列表)
    file_2,                 # 2.xlsx (专家银行卡信息)
    file_3,                 # 3.xlsx (语料库明细数据)
    tpl_4=None,             # 4.xlsx 模板路径 (可选，默认使用内置)
    tpl_5=None,             # 5.xlsx 模板路径 (可选，默认使用内置)
    tpl_6=None,             # 6.xlsx 模板路径 (可选，默认使用内置)
    output_dir=".",         # 输出目录
    include_all_133=True,   # True: 以1.xlsx全部词条为准; False: 仅3.xlsx中结算状态=='已加入结算单'
    price_per_item=None,    # 若为 None，则直接从源文件（3.xlsx 或 1.xlsx）的「结算单价」列自动抓取；若无则兜底 100
    log_func=print
):
    """
    北检结算表完整处理主函数 (HCP自行承担税费 / 正算扣税模式)
    """
    if log_func is None:
        log_func = print

    log_func("[提示] 启动北检结算表生成流水线（模式：💼 HCP自行承担税费 / 正算扣税）...")
    os.makedirs(output_dir, exist_ok=True)

    # 1. 采用 100% 纯原生 Python 动态生成规范化交付报表，零外部 Excel 模板依赖
    log_func("[环境] 采用 100% 纯原生 Python 引擎生成交付表格，零外部文件依赖")

    # 2. 读取原始 Excel 数据表
    log_func("[读取] 正在读取 1. 待结算词条表、2. 专家银行信息表、3. 语料明细全库...")
    df1 = pd.read_excel(file_1)
    df2 = pd.read_excel(file_2)
    df3 = pd.read_excel(file_3)

    # 动态匹配关键字段
    col_id_1 = find_col(df1, ['语料词条编号', '语料编号', '词条编号', '编号'], default='语料词条编号')
    col_id_3 = find_col(df3, ['语料词条编号', '语料编号', '词条编号', '编号'], default='语料词条编号')
    col_status_3 = find_col(df3, ['结算状态', '状态'], default='结算状态')

    if not col_id_1 or not col_id_3:
        raise ValueError(f"无法识别核心编号列：1.xlsx 识别为 {col_id_1}，3.xlsx 识别为 {col_id_3}")

    if col_id_1 not in df1.columns:
        raise ValueError(f"1. 待结算表中未找到词条编号列，当前列包含: {df1.columns.tolist()}")
    if col_id_3 not in df3.columns:
        raise ValueError(f"3. 语料明细表中未找到词条编号列，当前列包含: {df3.columns.tolist()}")

    # 确定目标词条编号范围
    if include_all_133:
        target_ids = list(df1[col_id_1].dropna().astype(str).str.strip())
        log_func(f"[锁定] 采用全量模式：以 1.xlsx 待结算表为准，共锁定 {len(target_ids)} 条目标词条")
    else:
        if col_status_3 in df3.columns:
            target_ids = list(df3[df3[col_status_3].astype(str).str.strip() == '已加入结算单'][col_id_3].dropna().astype(str).str.strip())
            log_func(f"[锁定] 采用状态过滤模式：仅提取结算状态为'已加入结算单'的词条，共锁定 {len(target_ids)} 条")
        else:
            target_ids = list(df1[col_id_1].dropna().astype(str).str.strip())
            log_func(f"[警告] 3.xlsx 中未发现结算状态列，自动回退为按 1.xlsx 全量锁定 {len(target_ids)} 条词条")

    # 3. 筛选并按 3.xlsx 出现顺序保留词条记录
    df3_copy = df3.copy()
    df3_copy['clean_id'] = df3_copy[col_id_3].astype(str).str.strip()
    df3_filtered = df3_copy[df3_copy['clean_id'].isin(target_ids)].copy()

    # 保持原始出现顺序
    original_id_order = df3_copy['clean_id'].tolist()
    df3_filtered['sort_order'] = df3_filtered['clean_id'].map(lambda x: original_id_order.index(x) if x in original_id_order else 999999)
    df3_filtered = df3_filtered.sort_values('sort_order').drop(columns=['sort_order', 'clean_id'])

    log_func(f"[匹配] 语料明细总库成功匹配到 {len(df3_filtered)} 条待结算数据")

    # 动态抓取每条语料的结算单价（直接从源文件 3.xlsx 或 1.xlsx 提取，兜底 100）
    col_price_3 = find_col(df3, ['结算单价', '单篇单价', '单价', '费用'])
    col_price_1 = find_col(df1, ['结算单价', '单篇单价', '单价', '费用'])

    if col_price_3 and col_price_3 in df3_filtered.columns:
        df3_filtered['结算单价'] = pd.to_numeric(df3_filtered[col_price_3], errors='coerce')
    else:
        df3_filtered['结算单价'] = None

    if col_price_1 and col_price_1 in df1.columns:
        price_map_1 = dict(zip(df1[col_id_1].astype(str).str.strip(), pd.to_numeric(df1[col_price_1], errors='coerce')))
        df3_filtered['结算单价'] = df3_filtered['结算单价'].fillna(df3_filtered[col_id_3].astype(str).str.strip().map(price_map_1))

    fallback_price = float(price_per_item) if price_per_item is not None else 100.0
    df3_filtered['结算单价'] = df3_filtered['结算单价'].fillna(fallback_price)
    df3_filtered['结算状态'] = '已加入结算单'

    min_p = float(df3_filtered['结算单价'].min())
    max_p = float(df3_filtered['结算单价'].max())
    if min_p == max_p:
        log_func(f"[单价] 成功从源文件自动抓取结算单价：统一为 {min_p:g} 元/篇")
    else:
        log_func(f"[单价] 成功从源文件自动抓取结算单价：单价区间为 {min_p:g} ~ {max_p:g} 元/篇")

    # =========================================================================
    # 生成表 8: 语料明细表
    # =========================================================================
    log_func("[生成] 正在生成【8. 语料词条明细表】...")
    cols_5 = ['语料词条编号', '医生姓名', '手机号', '身份证号码', '医院', '科室', '职称',
              '项目名称', '疾病领域', '题目内容', '回答记录', '首次提交时间', '审核时间',
              '审核状态', '结算状态', '结算单价']

    wb8 = openpyxl.Workbook()
    ws8 = wb8.active
    ws8.title = "语料明细"

    widths_8 = [28, 12, 16, 24, 28, 12, 16, 32, 15, 45, 60, 20, 20, 12, 14, 12]
    for idx, w in enumerate(widths_8, start=1):
        ws8.column_dimensions[get_column_letter(idx)].width = w

    ws8.row_dimensions[1].height = 24
    header_font_8 = Font(name="宋体", size=11, bold=True)
    header_fill_8 = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c_idx, col_name in enumerate(cols_5, start=1):
        cell = ws8.cell(1, c_idx, col_name)
        cell.font = header_font_8
        cell.fill = header_fill_8
        cell.alignment = center_align
        cell.border = thin_border

    data_font_8 = Font(name="Calibri", size=10)
    for r_idx, (_, row_data) in enumerate(df3_filtered.iterrows(), start=2):
        ws8.row_dimensions[r_idx].height = 20
        for c_idx, col_name in enumerate(cols_5, start=1):
            cell = ws8.cell(r_idx, c_idx)
            matched_col = find_col(df3_filtered, [col_name])
            val = row_data[matched_col] if matched_col else row_data.get(col_name, "")
            if pd.isna(val):
                val = ""
            cell.value = str(val) if col_name in ['身份证号码', '手机号', '语料词条编号'] else val
            cell.font = data_font_8
            cell.border = thin_border
            if col_name in ['题目内容', '回答记录']:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = center_align

    out_path_8 = os.path.join(output_dir, "8_已结算语料词条明细表.xlsx")
    wb8.save(out_path_8)
    wb8.save(os.path.join(output_dir, "8.xlsx"))
    wb8.close()
    log_func(f"[完成] 【8. 语料词条明细表】已生成 (共 {len(df3_filtered)} 行)")

    # =========================================================================
    # 生成表 9: 项目作品结算表
    # =========================================================================
    log_func("[生成] 正在生成【9. 项目作品结算表】...")
    wb9 = openpyxl.Workbook()
    ws9 = wb9.active
    ws9.title = "作品结算"

    ws9.views.sheetView[0].showGridLines = True

    ws9.row_dimensions[1].height = 30
    ws9.row_dimensions[2].height = 24
    ws9.row_dimensions[3].height = 20
    ws9.row_dimensions[4].height = 24

    ws9.merge_cells("A1:K1")
    title_cell = ws9.cell(1, 1, "专家劳务明细表")
    title_cell.font = Font(name="微软雅黑", size=16, bold=True)
    title_cell.alignment = center_align

    ws9.cell(2, 1, "开票单位：北京中卫医药卫生检验所")
    ws9.cell(2, 8, "发放部门：临床评价中心")
    meta_font = Font(name="微软雅黑", size=10)
    ws9.cell(2, 1).font = meta_font
    ws9.cell(2, 8).font = meta_font

    ws9.merge_cells("A3:K3")
    sec_cell = ws9.cell(3, 1, "劳务发放明细：")
    sec_cell.font = Font(name="微软雅黑", size=10, bold=True)
    sec_cell.alignment = Alignment(horizontal="left", vertical="center")

    headers_9 = ['序号', '专家', '作品类型', '作品数量', '计酬标准（元/篇）',
                 '应发金额', '税金', '实发金额', '开户行', '卡号', '身份证号']
    widths_9 = [8, 14, 18, 14, 20, 16, 14, 16, 32, 26, 24]
    for idx, w in enumerate(widths_9, start=1):
        ws9.column_dimensions[get_column_letter(idx)].width = w

    for c_idx, h_name in enumerate(headers_9, start=1):
        cell = ws9.cell(4, c_idx, h_name)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    doc_name_col_9 = find_col(df3_filtered, ['医生姓名', '姓名', '专家'], default='医生姓名')
    id_card_col_9 = find_col(df3_filtered, ['身份证号码', '身份证号', '证件号'], default='身份证号码')
    phone_col_9 = find_col(df3_filtered, ['手机号', '手机号码', '电话'], default='手机号')

    wb2 = openpyxl.load_workbook(file_2, data_only=True)
    ws2 = wb2.active
    header2 = [str(ws2.cell(1, c).value or '').strip() for c in range(1, ws2.max_column + 1)]

    def get_col_idx(cand_list):
        for cand in cand_list:
            for idx, h in enumerate(header2, start=1):
                if cand in h:
                    return idx
        return None

    col_idx_id = get_col_idx(['身份证', '证件']) or 10
    col_idx_bank = get_col_idx(['银行名称', '开户银行']) or 7
    col_idx_card = get_col_idx(['卡号', '账号']) or 8
    col_idx_branch = get_col_idx(['开户行', '支行']) or 9

    bank_info_map = {}
    for r in range(2, ws2.max_row + 1):
        id_val = ws2.cell(r, col_idx_id).value
        if id_val:
            id_str = str(id_val).strip().upper()
            bank_info_map[id_str] = {
                'bank': ws2.cell(r, col_idx_bank).value or '',
                'card': str(ws2.cell(r, col_idx_card).value or '').strip(),
                'branch': ws2.cell(r, col_idx_branch).value or ''
            }

    doc_summary_9 = df3_filtered.groupby([doc_name_col_9, id_card_col_9], sort=False).agg(
        item_count=(col_id_3, 'count'),
        total_pay=('结算单价', 'sum')
    ).reset_index()

    for idx, rdata in doc_summary_9.iterrows():
        r_num = idx + 5
        ws9.row_dimensions[r_num].height = 20
        d_name = str(rdata[doc_name_col_9]).strip()
        i_card = str(rdata[id_card_col_9]).strip().upper()
        i_cnt = int(rdata['item_count'])
        t_gross = float(rdata['total_pay'])

        # HCP自行承担税费模式：t_gross 是税前应发金额，从应发金额扣税
        gross_pay_9, tax_9, net_pay_9 = calculate_tax(t_gross)

        unit_p = round(t_gross / i_cnt, 2) if i_cnt > 0 else 0

        b_data = bank_info_map.get(i_card, {})
        row_vals_9 = [
            idx + 1, d_name, '问答语料审核', i_cnt, unit_p,
            gross_pay_9, tax_9, net_pay_9,
            b_data.get('branch', ''), b_data.get('card', ''), i_card
        ]

        for c_idx, val in enumerate(row_vals_9, start=1):
            cell = ws9.cell(r_num, c_idx, val)
            cell.font = Font(name="微软雅黑", size=10)
            cell.border = thin_border
            cell.alignment = center_align
            if c_idx in [10, 11]:
                cell.number_format = '@'

    out_path_9 = os.path.join(output_dir, "9_作品劳务结算总表.xlsx")
    wb9.save(out_path_9)
    wb9.save(os.path.join(output_dir, "6.xlsx"))
    wb9.save(os.path.join(output_dir, "9.xlsx"))
    wb9.close()
    log_func(f"[完成] 【9. 项目作品结算表】已生成 (共 {len(doc_summary_9)} 位专家作品结算)")

    # =========================================================================
    # 生成表 7: 专家劳务报酬明细表 (HCP自行承担税费主交付表)
    # =========================================================================
    log_func("[生成] 正在生成【HCP自行承担税费结算表】...")
    wb7 = openpyxl.Workbook()
    ws7 = wb7.active
    ws7.title = "劳务发放表"

    ws7.views.sheetView[0].showGridLines = True

    ws7.row_dimensions[1].height = 30
    ws7.row_dimensions[2].height = 24
    ws7.row_dimensions[3].height = 20
    ws7.row_dimensions[4].height = 24

    ws7.merge_cells("A1:M1")
    tcell = ws7.cell(1, 1, "专家劳务明细表（HCP自行承担税费）")
    tcell.font = Font(name="微软雅黑", size=16, bold=True)
    tcell.alignment = center_align

    ws7.cell(2, 1, "开票单位：北京中卫医药卫生检验所")
    ws7.cell(2, 9, "发放部门：临床评价中心")
    ws7.cell(2, 1).font = meta_font
    ws7.cell(2, 9).font = meta_font

    ws7.merge_cells("A3:M3")
    scell = ws7.cell(3, 1, "劳务发放明细：")
    scell.font = Font(name="微软雅黑", size=10, bold=True)
    scell.alignment = Alignment(horizontal="left", vertical="center")

    headers_7 = ['序号', '专家', '医院', '科室', '职称', '电话',
                 '银行', '银行卡号', '开户行', '身份证号', '应发金额', '税金', '实发金额']
    widths_7 = [8, 12, 28, 14, 14, 16, 16, 26, 32, 24, 15, 14, 15]
    for idx, w in enumerate(widths_7, start=1):
        ws7.column_dimensions[get_column_letter(idx)].width = w

    for c_idx, h_name in enumerate(headers_7, start=1):
        cell = ws7.cell(4, c_idx, h_name)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    doc_name_col = find_col(df3_filtered, ['医生姓名', '姓名', '专家'], default='医生姓名')
    id_card_col = find_col(df3_filtered, ['身份证号码', '身份证号', '证件号'], default='身份证号码')
    hosp_col = find_col(df3_filtered, ['医院', '所在医院', '单位'], default='医院')
    dept_col = find_col(df3_filtered, ['科室', '所在科室'], default='科室')
    title_col_doc = find_col(df3_filtered, ['职称', '职务'], default='职称')
    phone_col = find_col(df3_filtered, ['手机号', '手机号码', '电话'], default='手机号')

    doc_counts = df3_filtered.groupby([doc_name_col, id_card_col, hosp_col, dept_col, title_col_doc], sort=False).agg(
        count=(col_id_3, 'count'),
        pay_amt=('结算单价', 'sum'),
        phone=(phone_col, 'first')
    ).reset_index()

    data_font_7 = Font(name="微软雅黑", size=10)
    total_net_all = 0.0
    total_tax_all = 0.0
    total_gross_all = 0.0

    for idx, rdata in doc_counts.iterrows():
        row_num = idx + 5
        ws7.row_dimensions[row_num].height = 22
        doc_name = str(rdata[doc_name_col]).strip()
        id_card = str(rdata[id_card_col]).strip().upper()
        hosp = str(rdata[hosp_col] or '')
        dept = str(rdata[dept_col] or '')
        title = str(rdata[title_col_doc] or '')
        phone = str(rdata['phone'] or '').replace('.0', '')
        count = int(rdata['count'])

        # 匹配银行信息
        binfo = bank_info_map.get(id_card, {})
        bank = binfo.get('bank', '')
        card = binfo.get('card', '')
        branch = binfo.get('branch', '')

        # HCP自行承担税费模式：任务单价累计作为税前应发金额，依法扣除个税后为实发金额
        pay_amt = float(rdata['pay_amt'])
        gross_pay, tax, net_pay = calculate_tax(pay_amt)

        total_net_all += net_pay
        total_tax_all += tax
        total_gross_all += gross_pay

        row_vals = [
            idx + 1,        # 序号
            doc_name,       # 医生
            hosp,           # 医院
            dept,           # 科室
            title,          # 职称
            phone,          # 电话
            bank,           # 银行
            card,           # 银行卡号
            branch,         # 开户行
            id_card,        # 身份证号
            gross_pay,      # 应发金额
            tax,            # 税金
            net_pay         # 实发金额
        ]

        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws7.cell(row_num, c_idx, val)
            cell.font = data_font_7
            cell.border = thin_border
            cell.alignment = center_align
            if c_idx in [8, 10]:
                cell.number_format = '@'

    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    final_filename = f"HCP自行承担税费-{today_str}.xlsx"
    out_path_final = os.path.join(output_dir, final_filename)
    wb7.save(out_path_final)
    out_path_7 = os.path.join(output_dir, "7_专家劳务报酬明细表.xlsx")
    wb7.save(out_path_7)
    wb7.save(os.path.join(output_dir, "7.xlsx"))
    wb7.close()
    log_func(f"[完成] 最终表格【{final_filename}】生成成功 (共 {len(doc_counts)} 位医生)")

    log_func("[成功] 全部 3 张结算表单生成完毕！")

    return {
        "success": True,
        "mode": "HCP自行承担税费",
        "final_table_name": final_filename,
        "total_items": len(df3_filtered),
        "total_doctors": len(doc_counts),
        "total_net": round(total_net_all, 2),
        "total_tax": round(total_tax_all, 2),
        "total_gross": round(total_gross_all, 2),
        "files": {
            final_filename: out_path_final,
            "7_专家劳务报酬明细表.xlsx": out_path_final,
            "8_已结算语料词条明细表.xlsx": out_path_8,
            "9_作品劳务结算总表.xlsx": out_path_9,
            "7.xlsx": os.path.join(output_dir, "7.xlsx"),
            "8.xlsx": os.path.join(output_dir, "8.xlsx"),
            "9.xlsx": os.path.join(output_dir, "9.xlsx"),
        }
    }

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="北检&康恩贝结算表生成器 (HCP自行承担税费模式)")
    parser.add_argument("-1", "--file1", default="1.xlsx", help="1.xlsx 待结算词条名单")
    parser.add_argument("-2", "--file2", default="2.xlsx", help="2.xlsx 专家银行卡信息")
    parser.add_argument("-3", "--file3", default="3.xlsx", help="3.xlsx 语料库明细数据")
    parser.add_argument("-4", "--tpl4", default=None, help="4.xlsx 模板")
    parser.add_argument("-5", "--tpl5", default=None, help="5.xlsx 模板")
    parser.add_argument("-6", "--tpl6", default=None, help="6.xlsx 模板")
    parser.add_argument("-o", "--outdir", default=".", help="输出文件夹")
    parser.add_argument("--price", type=float, default=100.0, help="结算单价 (元)")
    parser.add_argument("--filtered-only", action="store_true", help="仅按已加入结算单结算")
    args = parser.parse_args()

    res = process_beijian_kangbei(
        file_1=args.file1,
        file_2=args.file2,
        file_3=args.file3,
        tpl_4=args.tpl4,
        tpl_5=args.tpl5,
        tpl_6=args.tpl6,
        output_dir=args.outdir,
        include_all_133=not args.filtered_only,
        price_per_item=args.price
    )
    print(f"\n处理结果: 词条数: {res['total_items']}, 医生数: {res['total_doctors']}, 应发总额: ¥{res['total_gross']}, 实发总额: ¥{res['total_net']}, 代扣税金: ¥{res['total_tax']}")
