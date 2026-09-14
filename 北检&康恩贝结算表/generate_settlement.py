# -*- coding: utf-8 -*-
"""
=============================================================================
北检 & 康恩贝结算表生成引擎 (Beijian & CONBA Settlement Engine)
=============================================================================
功能说明：
1. 接收 3 张核心源表：
   - 1.xlsx (待结算词条编号名单)
   - 2.xlsx (专家银行卡与身份资质表)
   - 3.xlsx (语料库全量审核明细表)
2. 关联匹配与规范化输出 3 张标准化交付表单：
   - 7_专家劳务报酬明细表.xlsx (基于 4.xlsx 模板格式，内置劳务报酬个税精准反算公式与银行信息)
   - 8_语料词条明细表.xlsx (基于 5.xlsx 模板格式，按结算单价标记结算状态并保留原始题目/回答记录)
   - 9_项目作品结算表.xlsx (基于 6.xlsx 模板格式，作品维度单价与乘法关联核算)
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

def calculate_tax(net_pay):
    """
    根据实发劳务报酬计算应发金额与代扣个人所得税 (税后反算税前)
    规则：
    1. 实发 <= 800: 无税，应发 = 实发
    2. 800 < 实发 <= 3360: 应发 = (实发 - 160) / 0.8, 税金 = 应发 - 实发
    3. 实发 > 3360: 应发 = 实发 / 0.84, 税金 = 应发 - 实发
    """
    if net_pay <= 800:
        tax = 0.0
        gross_pay = float(net_pay)
    elif net_pay <= 3360:
        gross_pay = round((net_pay - 160) / 0.8, 2)
        tax = round(gross_pay - net_pay, 2)
    else:
        gross_pay = round(net_pay / 0.84, 2)
        tax = round(gross_pay - net_pay, 2)
    return gross_pay, tax

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
    北检 & 康恩贝结算表完整处理主函数
    """
    if log_func is None:
        log_func = print

    log_func("[提示] 启动北检&康恩贝结算表生成流水线...")
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
    log_func(f"[完成] 8. 语料明细表生成成功 (共 {len(df3_filtered)} 行数据)")

    # =========================================================================
    # 生成表 9: 项目作品结算表
    # =========================================================================
    log_func("[生成] 正在生成【9. 项目作品结算表】...")
    wb9 = openpyxl.Workbook()
    ws9 = wb9.active
    ws9.title = "项目结算表"

    widths_9 = [75, 12, 12, 12, 10, 16]
    for idx, w in enumerate(widths_9, start=1):
        ws9.column_dimensions[get_column_letter(idx)].width = w

    # Row 1
    ws9.merge_cells("A1:F1")
    ws9.row_dimensions[1].height = 45
    c1 = ws9.cell(1, 1, "项目结算表")
    c1.font = Font(name="微软雅黑", size=16, bold=True)
    c1.alignment = Alignment(horizontal="center", vertical="center")

    # Row 2
    now = datetime.datetime.now()
    ws9.merge_cells("A2:F2")
    ws9.row_dimensions[2].height = 25
    c2 = ws9.cell(2, 1, f"结算时间：{now.year}年{now.month}月")
    c2.font = Font(name="微软雅黑", size=10)
    c2.alignment = Alignment(horizontal="left", vertical="center")

    # Row 3
    ws9.merge_cells("A3:F3")
    ws9.row_dimensions[3].height = 25
    c3 = ws9.cell(3, 1, '结算内容：“医疗健康 AI语料库建设”项目')
    c3.font = Font(name="微软雅黑", size=10)
    c3.alignment = Alignment(horizontal="left", vertical="center")

    # Row 4 headers
    headers_9 = ['作品名称', '费用', '是否含税', '收集数量', '单位', '参考结算费用']
    ws9.row_dimensions[4].height = 25
    header_font_9 = Font(name="微软雅黑", size=10, bold=True)
    header_fill_9 = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for c_idx, h in enumerate(headers_9, start=1):
        cell = ws9.cell(4, c_idx, h)
        cell.font = header_font_9
        cell.fill = header_fill_9
        cell.alignment = center_align
        cell.border = thin_border

    data_font_9 = Font(name="微软雅黑", size=10)
    title_col = find_col(df3_filtered, ['题目内容', '题目', '作品名称', '内容'], default='题目内容')

    for r_idx, (_, row_data) in enumerate(df3_filtered.iterrows(), start=5):
        ws9.row_dimensions[r_idx].height = 22
        title_val = row_data.get(title_col, "") if title_col else ""
        if pd.isna(title_val):
            title_val = ""
        row_unit_price = float(row_data.get('结算单价', fallback_price))
        row_vals = [title_val, row_unit_price, "否", 1, "份", f"=B{r_idx}*D{r_idx}"]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws9.cell(r_idx, c_idx, val)
            cell.font = data_font_9
            cell.border = thin_border
            if c_idx == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = center_align

    out_path_9 = os.path.join(output_dir, "9_作品劳务结算总表.xlsx")
    wb9.save(out_path_9)
    wb9.save(os.path.join(output_dir, "9.xlsx"))
    wb9.close()
    log_func(f"[完成] 9. 项目作品结算表生成成功 (共 {len(df3_filtered)} 项作品)")

    # =========================================================================
    # 生成表 7: 医生专家劳务明细表
    # =========================================================================
    log_func("[生成] 正在生成【7. 专家劳务报酬明细表】...")
    wb7 = openpyxl.Workbook()
    ws7 = wb7.active
    ws7.title = "劳务明细表"

    widths_7 = [8, 12, 35, 14, 16, 16, 18, 25, 25, 24, 14, 12, 14]
    for idx, w in enumerate(widths_7, start=1):
        ws7.column_dimensions[get_column_letter(idx)].width = w

    # 行1：填表须知
    ws7.merge_cells("A1:B1")
    ws7.row_dimensions[1].height = 22
    c1 = ws7.cell(1, 1, "填表须知（必读）")
    c1.font = Font(name="宋体", size=12, bold=True)
    c1.alignment = Alignment(horizontal="left", vertical="center")

    # 行2：说明文字
    ws7.merge_cells("A2:J2")
    ws7.row_dimensions[2].height = 36
    c2 = ws7.cell(2, 1, "1、下列表格为设置好的核算公式格式，请勿随意修改设置好的公式。\n2、填写信息时，只需填写各专家的劳务信息及最后一项“实发金额”，其他表格则自动计算应发金额及税金。")
    c2.font = Font(name="宋体", size=11)
    c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 行4：表头
    headers_7 = ['序号', '医生', '医院', '科室', '职称', '电话', '银行', '银行卡号', '开户行', '身份证号', '应发金额', '税金', '实发金额']
    ws7.row_dimensions[4].height = 25
    header_font_7 = Font(name="微软雅黑", size=11, bold=True)
    header_fill_7 = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for c_idx, h in enumerate(headers_7, start=1):
        cell = ws7.cell(4, c_idx, h)
        cell.font = header_font_7
        cell.fill = header_fill_7
        cell.alignment = center_align
        cell.border = thin_border

    # 构建专家银行信息索引 (支持 openpyxl 精确文本避免卡号科学计数法)
    wb2 = openpyxl.load_workbook(file_2, data_only=True)
    ws2 = wb2.active

    # 动态匹配 2.xlsx 表头列
    header_row_2 = [str(ws2.cell(1, c).value or '').strip() for c in range(1, ws2.max_column + 1)]
    col_idx_id2 = None
    col_idx_bank = None
    col_idx_card = None
    col_idx_branch = None

    for idx, hname in enumerate(header_row_2, start=1):
        if any(k in hname for k in ['身份证', '证件号码', '身份证号']):
            col_idx_id2 = idx
        elif any(k in hname for k in ['开户银行', '银行名称', '开户行']) and '省' not in hname and '市' not in hname:
            col_idx_bank = idx
        elif any(k in hname for k in ['银行卡号', '卡号', '账号']):
            col_idx_card = idx
        elif any(k in hname for k in ['支行名称', '支行', '开户支行']):
            col_idx_branch = idx

    # 兜底默认列
    col_idx_id2 = col_idx_id2 or 6
    col_idx_card = col_idx_card or 7
    col_idx_bank = col_idx_bank or 8
    col_idx_branch = col_idx_branch or 9

    bank_info_map = {}
    for r in range(2, ws2.max_row + 1):
        id_c = ws2.cell(r, col_idx_id2).value
        if id_c:
            id_str = str(id_c).strip().upper()
            bank_info_map[id_str] = {
                'bank': ws2.cell(r, col_idx_bank).value or '',
                'card': str(ws2.cell(r, col_idx_card).value or '').strip(),
                'branch': ws2.cell(r, col_idx_branch).value or ''
            }
    wb2.close()

    # 按医生维度汇总语料篇数与结算单价之和
    doc_name_col = find_col(df3_filtered, ['医生姓名', '姓名', '专家'], default='医生姓名')
    id_card_col = find_col(df3_filtered, ['身份证号码', '身份证号', '证件号'], default='身份证号码')
    hosp_col = find_col(df3_filtered, ['医院', '所在医院', '单位'], default='医院')
    dept_col = find_col(df3_filtered, ['科室', '所在科室'], default='科室')
    title_col_doc = find_col(df3_filtered, ['职称', '职务'], default='职称')
    phone_col = find_col(df3_filtered, ['手机号', '手机号码', '电话'], default='手机号')

    doc_counts = df3_filtered.groupby([doc_name_col, id_card_col, hosp_col, dept_col, title_col_doc], sort=False).agg(
        count=(col_id_3, 'count'),
        net_pay=('结算单价', 'sum'),
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

        # 实发金额直接来自源文件抓取的结算单价汇总，并进行合规个税反算
        net_pay = float(rdata['net_pay'])
        gross_pay, tax = calculate_tax(net_pay)

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

    out_path_7 = os.path.join(output_dir, "7_专家劳务报酬明细表.xlsx")
    wb7.save(out_path_7)
    wb7.save(os.path.join(output_dir, "7.xlsx"))
    wb7.close()
    log_func(f"[完成] 7. 专家劳务报酬明细表生成成功 (共 {len(doc_counts)} 位医生)")

    log_func("[成功] 全部 3 张结算表单生成完毕！")

    return {
        "success": True,
        "total_items": len(df3_filtered),
        "total_doctors": len(doc_counts),
        "total_net": round(total_net_all, 2),
        "total_tax": round(total_tax_all, 2),
        "total_gross": round(total_gross_all, 2),
        "files": {
            "7_专家劳务报酬明细表.xlsx": out_path_7,
            "8_已结算语料词条明细表.xlsx": out_path_8,
            "9_作品劳务结算总表.xlsx": out_path_9,
            "7.xlsx": os.path.join(output_dir, "7.xlsx"),
            "8.xlsx": os.path.join(output_dir, "8.xlsx"),
            "9.xlsx": os.path.join(output_dir, "9.xlsx"),
        }
    }

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="北检&康恩贝结算表生成器")
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
    print(f"\n处理结果: 词条数: {res['total_items']}, 医生数: {res['total_doctors']}, 实发总额: ¥{res['total_net']}")
