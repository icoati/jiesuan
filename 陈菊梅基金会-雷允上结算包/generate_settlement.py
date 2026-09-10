#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
雷允上语料库劳务费用明细表一键生成脚本 (支持 Mac / Windows / Linux)
功能特点：
1. 智能查找文件：自动识别目录下的“进度表”、“语料明文表”、“用户明文表”；
2. 智能列名映射：即使不同批次的Excel列名、顺序有细微差异，通过关键词模糊匹配自动对齐；
3. 严格遵循结算标准：以进度表作为最终人员和金额基准，通过语料编号关联明文与银行信息；
4. 规范导出：自动生成“汇总-明细”、“各项目明细”及“语料结算明细对账表”，内置个税反算与汇总公式。
"""

import os
import sys
import glob
import re
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 列名同义词映射表（支持模糊匹配）
COLUMN_SYNONYMS = {
    '语料编号': ['语料词条编号', '语料编号', '词条编号', '语料id', '编号', 'corpus_id', 'id'],
    '医生姓名': ['医生姓名', '姓名', '医生', '专家姓名', '专家', 'name'],
    '手机号': ['手机号', '手机号码', '电话', '联系电话', '手机', 'phone', 'mobile'],
    '身份证号': ['身份证号码', '身份证号', '证件号码', '证件号', '身份证', 'id_card', 'idcard'],
    '医院': ['医院', '所在医院', '单位', '单位名称', '就职医院', 'hospital'],
    '科室': ['科室', '所在科室', '部门', 'department'],
    '职称': ['职称', '职务', 'title'],
    '项目名称': ['项目名称', '项目', '所属项目', '参与项目', 'project'],
    '结算单价': ['结算单价', '标签', '单价', '金额', '费用', 'price', 'amount'],
    '疾病领域': ['疾病领域', '领域', '学科', '科别', 'field'],
    '题目内容': ['题目内容', '题目', '问题', '语料内容', 'content', 'title'],
    '提交时间': ['提交时间', '时间', '提交日期', '日期', 'submit_time', 'time'],
    '审核状态': ['审核状态', '状态', '审核结果', 'status'],
    '开户银行': ['开户银行', '银行名称', '开户行', '银行', 'bank'],
    '支行名称': ['支行名称', '开户支行', '支行', 'branch'],
    '银行卡号': ['银行卡号', '银行账号', '卡号', '账号', 'card_number', 'account']
}

def match_column(df, field_key, required=True):
    """根据同义词列表和模糊匹配在 DataFrame 中查找匹配的列名"""
    cols = df.columns.tolist()
    # 1. 精确匹配
    synonyms = COLUMN_SYNONYMS.get(field_key, [])
    for syn in synonyms:
        for c in cols:
            if str(c).strip().lower() == syn.lower():
                return c
                
    # 2. 包含匹配
    for syn in synonyms:
        for c in cols:
            c_clean = str(c).strip().lower()
            if syn.lower() in c_clean or c_clean in syn.lower():
                return c
                
    if required:
        raise ValueError(f"无法在表格列中识别出【{field_key}】列！当前表格的列为: {cols}")
    return None

def find_file_by_keywords(target_dir, must_have, must_not=None):
    """根据关键字智能查找目录下的文件"""
    must_not = must_not or []
    candidates = []
    for f in os.listdir(target_dir):
        if not f.endswith(('.xlsx', '.xls')) or f.startswith('~$'):
            continue
        fname_lower = f.lower()
        has_all = all(kw.lower() in fname_lower for kw in must_have)
        has_not = any(kw.lower() in fname_lower for kw in must_not)
        if has_all and not has_not:
            candidates.append(os.path.join(target_dir, f))
    return candidates

def auto_detect_files(work_dir):
    """自动侦测 进度表、语料列表、用户列表"""
    # 1. 进度表
    prog_files = find_file_by_keywords(work_dir, ['进度'], ['费用', '明细', '对账', '备份'])
    if not prog_files:
        prog_files = find_file_by_keywords(work_dir, ['雷允上'], ['费用', '明细', '对账', '用户', '语料列表', '备份'])
        
    # 2. 语料明文表
    corpus_files = find_file_by_keywords(work_dir, ['语料'], ['进度', '费用', '明细', '对账', '备份'])
    
    # 3. 用户明文表
    user_files = find_file_by_keywords(work_dir, ['用户'], ['进度', '费用', '明细', '对账', '备份'])
    
    return prog_files, corpus_files, user_files

def format_bank_name(bank_name, branch_name):
    """智能清洗并格式化开户行信息"""
    b = str(bank_name).strip() if pd.notnull(bank_name) else ''
    br = str(branch_name).strip() if pd.notnull(branch_name) else ''
    
    # 特殊文本处理（如用户填入含有“户名”、“开户行”前缀）
    if '开户行：' in br:
        br = br.split('开户行：')[-1].strip()
    elif '开户行' in br:
        br = br.split('开户行')[-1].strip()
        
    if not br or br == b:
        return b
    if not b:
        return br
    if br.startswith(b):
        return br
    return f"{b}{br}"

def generate_settlement(work_dir=None, prog_path=None, corpus_path=None, user_path=None, output_path=None):
    if work_dir is None:
        work_dir = os.path.abspath(os.path.dirname(__file__)) if '__file__' in globals() else os.getcwd()
        
    print(f"工作目录: {work_dir}")
    
    # 自动查找文件
    if not (prog_path and corpus_path and user_path):
        p_cands, c_cands, u_cands = auto_detect_files(work_dir)
        prog_path = prog_path or (p_cands[0] if p_cands else None)
        corpus_path = corpus_path or (c_cands[0] if c_cands else None)
        user_path = user_path or (u_cands[0] if u_cands else None)
        
    print(f"-> 进度表: {os.path.basename(prog_path) if prog_path else '未找到'}")
    print(f"-> 语料表(明文): {os.path.basename(corpus_path) if corpus_path else '未找到'}")
    print(f"-> 用户表(明文): {os.path.basename(user_path) if user_path else '未找到'}")
    
    if not (prog_path and corpus_path and user_path):
        print("\n【错误】未集齐所需的三张数据表，请检查当前文件夹中的文件名！")
        return False
        
    # 读取各表
    print("\n正在读取数据源...")
    df_prog = pd.read_excel(prog_path)
    df_corpus = pd.read_excel(corpus_path)
    df_user = pd.read_excel(user_path)
    
    # 列名匹配与映射
    prog_id_col = match_column(df_prog, '语料编号')
    prog_name_col = match_column(df_prog, '医生姓名')
    prog_proj_col = match_column(df_prog, '项目名称')
    prog_price_col = match_column(df_prog, '结算单价')
    prog_hosp_col = match_column(df_prog, '医院')
    prog_dept_col = match_column(df_prog, '科室', required=False)
    prog_title_col = match_column(df_prog, '职称', required=False)
    prog_field_col = match_column(df_prog, '疾病领域', required=False)
    prog_content_col = match_column(df_prog, '题目内容', required=False)
    prog_time_col = match_column(df_prog, '提交时间', required=False)
    prog_status_col = match_column(df_prog, '审核状态', required=False)
    
    corpus_id_col = match_column(df_corpus, '语料编号')
    corpus_phone_col = match_column(df_corpus, '手机号')
    corpus_idcard_col = match_column(df_corpus, '身份证号')
    
    user_idcard_col = match_column(df_user, '身份证号')
    user_bank_col = match_column(df_user, '开户银行', required=False)
    user_branch_col = match_column(df_user, '支行名称', required=False)
    user_card_col = match_column(df_user, '银行卡号', required=False)
    user_name_col = match_column(df_user, '医生姓名', required=False)
    
    print(f"进度表共 {len(df_prog)} 条语料记录。")
    
    # 建立映射字典
    corpus_map = {}
    for _, r in df_corpus.iterrows():
        cid = str(r[corpus_id_col]).strip()
        corpus_map[cid] = {
            'phone': str(r[corpus_phone_col]).strip() if pd.notnull(r[corpus_phone_col]) else '',
            'idcard': str(r[corpus_idcard_col]).strip() if pd.notnull(r[corpus_idcard_col]) else ''
        }
        
    user_id_map = {}
    user_name_map = {}
    for _, r in df_user.iterrows():
        idc = str(r[user_idcard_col]).strip() if pd.notnull(r[user_idcard_col]) else ''
        name = str(r[user_name_col]).strip() if (user_name_col and pd.notnull(r[user_name_col])) else ''
        b = r[user_bank_col] if user_bank_col else ''
        br = r[user_branch_col] if user_branch_col else ''
        card = str(r[user_card_col]).strip() if (user_card_col and pd.notnull(r[user_card_col])) else ''
        
        bank_formatted = format_bank_name(b, br)
        item_info = {'bank': bank_formatted, 'card': card}
        if idc:
            user_id_map[idc] = item_info
        if name:
            user_name_map[name] = item_info

    # 处理进度表记录
    log_rows = []
    for _, r in df_prog.iterrows():
        cid = str(r[prog_id_col]).strip()
        c_info = corpus_map.get(cid, {'phone': '', 'idcard': ''})
        
        doc_name = str(r[prog_name_col]).strip()
        proj_name = str(r[prog_proj_col]).strip()
        price = float(r[prog_price_col]) if pd.notnull(r[prog_price_col]) else 0.0
        hosp = str(r[prog_hosp_col]).strip() if pd.notnull(r[prog_hosp_col]) else ''
        dept = str(r[prog_dept_col]).strip() if (prog_dept_col and pd.notnull(r[prog_dept_col])) else ''
        title = str(r[prog_title_col]).strip() if (prog_title_col and pd.notnull(r[prog_title_col])) else ''
        field = str(r[prog_field_col]).strip() if (prog_field_col and pd.notnull(r[prog_field_col])) else ''
        content = str(r[prog_content_col]).strip() if (prog_content_col and pd.notnull(r[prog_content_col])) else ''
        sub_time = str(r[prog_time_col]) if (prog_time_col and pd.notnull(r[prog_time_col])) else ''
        status = str(r[prog_status_col]).strip() if (prog_status_col and pd.notnull(r[prog_status_col])) else '审核通过'
        
        log_rows.append({
            '语料词条编号': cid,
            '医生姓名': doc_name,
            '手机号': c_info['phone'],
            '身份证号码': c_info['idcard'],
            '医院': hosp,
            '科室': dept,
            '职称': title,
            '项目名称': proj_name,
            '结算单价': price,
            '疾病领域': field,
            '题目内容': content,
            '提交时间': sub_time,
            '审核状态': status
        })
        
    df_log = pd.DataFrame(log_rows)
    total_amount = df_log['结算单价'].sum()
    print(f"语料对账表明细整合完成，共 {len(df_log)} 条，税后总金额: {total_amount:,.2f} 元。")

    # 按项目分组汇总各医生
    projects = df_log['项目名称'].unique().tolist()
    def proj_sort_key(p):
        if '云南' in p:
            return (0, p)
        elif '长春' in p:
            return (1, p)
        return (2, p)
    projects = sorted(projects, key=proj_sort_key)
    
    project_dfs = {}
    all_summary_rows = []
    
    for proj in projects:
        sub_log = df_log[df_log['项目名称'] == proj]
        grp = sub_log.groupby('医生姓名', sort=False).agg({
            '结算单价': 'sum',
            '医院': 'first',
            '手机号': 'first',
            '身份证号码': 'first',
            '项目名称': 'first'
        }).reset_index()
        
        grp = grp.sort_values(by=['结算单价'], ascending=False).reset_index(drop=True)
        
        rows = []
        for idx, r in grp.iterrows():
            name = r['医生姓名']
            idc = r['身份证号码']
            
            u_info = user_id_map.get(idc) or user_name_map.get(name, {'bank': '', 'card': ''})
            
            item = {
                '序号': idx + 1,
                '姓名': name,
                '单位': r['医院'],
                '电话': r['手机号'],
                '税后金额': r['结算单价'],
                '证件类型': '身份证',
                '证件号码': idc,
                '开户行': u_info['bank'],
                '银行账号': u_info['card'],
                '项目名称': proj
            }
            rows.append(item)
            all_summary_rows.append(item)
            
        project_dfs[proj] = pd.DataFrame(rows)
        print(f"-> 项目【{proj}】: {len(rows)} 位医生，金额合计: {sum(r['税后金额'] for r in rows):,.2f} 元。")
        
    df_all_sum = pd.DataFrame(all_summary_rows)
    df_all_sum['序号'] = range(1, len(df_all_sum) + 1)
    
    # 排序对账表
    doc_rank_map = {r['姓名']: i for i, r in enumerate(all_summary_rows)}
    df_log['proj_rank'] = df_log['项目名称'].map(lambda x: proj_sort_key(x)[0])
    df_log['doc_rank'] = df_log['医生姓名'].map(lambda x: doc_rank_map.get(x, 9999))
    df_log = df_log.sort_values(by=['proj_rank', 'doc_rank', '提交时间'], ascending=[True, True, False]).reset_index(drop=True)
    df_log = df_log.drop(columns=['proj_rank', 'doc_rank'])

    # 导出 Excel 工作簿
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    font_title = Font(name='微软雅黑', size=18, bold=True)
    font_header = Font(name='微软雅黑', size=12, bold=True)
    font_data = Font(name='微软雅黑', size=11, bold=False)
    font_total = Font(name='微软雅黑', size=12, bold=True)
    align_center = Alignment(horizontal='center', vertical='center')
    align_header = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')

    def write_fee_sheet(ws, title_text, df_data, is_summary=False):
        max_c = 12 if is_summary else 11
        col_letter_max = get_column_letter(max_c)
        
        ws.row_dimensions[1].height = 24.0
        ws.row_dimensions[2].height = 24.0
        ws.row_dimensions[3].height = 34.8
        
        ws.merge_cells(f'A1:{col_letter_max}2')
        t_cell = ws.cell(1, 1, title_text)
        t_cell.font = font_title
        t_cell.alignment = align_center
        
        headers = ['序号', '姓名', '单位', '电话', '税后金额', '代扣个税', '收入额', '证件类型', '证件号码', '开户行', '银行账号']
        if is_summary:
            headers.append('项目名称')
            
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(3, c_idx, h)
            cell.font = font_header
            cell.alignment = align_header
            cell.border = thin_border
            
        num_rows = len(df_data)
        for i, r in df_data.iterrows():
            row_num = 4 + i
            ws.row_dimensions[row_num].height = 22.0
            
            vals = [
                r['序号'],
                r['姓名'],
                r['单位'],
                str(r['电话']),
                r['税后金额'],
                f'=IF(E{row_num}<=800,0,IF(E{row_num}<=3360,ROUND((E{row_num}-800)*0.25,2),IF(E{row_num}<=21000,ROUND(E{row_num}/0.84-E{row_num},2),IF(E{row_num}<=49500,ROUND((E{row_num}-2000)/0.76-E{row_num},2),ROUND((E{row_num}-7000)/0.68-E{row_num},2)))))',
                f'=SUM(E{row_num}:F{row_num})',
                r['证件类型'],
                str(r['证件号码']),
                r['开户行'],
                str(r['银行账号'])
            ]
            if is_summary:
                vals.append(r['项目名称'])
                
            for c_idx, v in enumerate(vals, start=1):
                cell = ws.cell(row_num, c_idx, v)
                cell.font = font_data
                cell.alignment = align_center
                cell.border = thin_border
                
                if c_idx in [4, 9, 11]:
                    cell.number_format = '@'
                elif c_idx == 5:
                    cell.number_format = '0'
                elif c_idx in [6, 7]:
                    cell.number_format = '#,##0.00'
                else:
                    cell.number_format = 'General'
                    
        tot_row = 4 + num_rows
        ws.row_dimensions[tot_row].height = 22.0
        ws.merge_cells(f'A{tot_row}:D{tot_row}')
        tot_cell = ws.cell(tot_row, 1, '合  计')
        tot_cell.font = font_total
        tot_cell.alignment = align_center
        
        for c_idx in range(1, 5):
            ws.cell(tot_row, c_idx).border = thin_border
            
        c_e = ws.cell(tot_row, 5, f'=SUM(E4:E{tot_row-1})')
        c_f = ws.cell(tot_row, 6, f'=SUM(F4:F{tot_row-1})')
        c_g = ws.cell(tot_row, 7, f'=SUM(G4:G{tot_row-1})')
        for c_obj in [c_e, c_f, c_g]:
            c_obj.font = font_total
            c_obj.alignment = align_center
            c_obj.border = thin_border
            c_obj.number_format = '#,##0.00'
            
        ws.merge_cells(f'H{tot_row}:{col_letter_max}{tot_row}')
        for c_idx in range(8, max_c + 1):
            ws.cell(tot_row, c_idx).border = thin_border
            
        widths = {'A': 6.0, 'B': 12.0, 'C': 38.0, 'D': 20.0, 'E': 15.0, 'F': 15.0, 'G': 15.0, 'H': 10.0, 'I': 26.0, 'J': 45.0, 'K': 28.0}
        if is_summary:
            widths['L'] = 32.0
        for col_l, w in widths.items():
            ws.column_dimensions[col_l].width = w
            
        ws.views.sheetView[0].showGridLines = True

    # 1. 汇总表
    ws_sum = wb.create_sheet(title='汇总-劳务费用明细')
    write_fee_sheet(ws_sum, '劳务费用明细汇总表', df_all_sum, is_summary=True)
    
    # 2. 各项目明细表
    for proj in projects:
        short_name = proj.replace('医疗健康 AI语料库-', '')
        sheet_title = f"{short_name}-明细"
        ws_p = wb.create_sheet(title=sheet_title)
        write_fee_sheet(ws_p, f"{proj}-劳务费用明细", project_dfs[proj], is_summary=False)
        
    # 3. 语料结算明细对账表
    ws_log = wb.create_sheet(title='语料结算明细对账表')
    ws_log.row_dimensions[1].height = 26.0
    headers_log = ['语料词条编号', '医生姓名', '手机号', '身份证号码', '医院', '科室', '职称', '项目名称', '结算单价', '疾病领域', '题目内容', '提交时间', '审核状态']
    for c_idx, h in enumerate(headers_log, start=1):
        cell = ws_log.cell(1, c_idx, h)
        cell.font = font_header
        cell.alignment = align_header
        cell.border = thin_border
        
    for i, r in df_log.iterrows():
        row_num = 2 + i
        ws_log.row_dimensions[row_num].height = 20.0
        vals = [
            str(r['语料词条编号']), r['医生姓名'], str(r['手机号']), str(r['身份证号码']),
            r['医院'], r['科室'], r['职称'], r['项目名称'], r['结算单价'],
            r['疾病领域'], r['题目内容'], r['提交时间'], r['审核状态']
        ]
        for c_idx, v in enumerate(vals, start=1):
            cell = ws_log.cell(row_num, c_idx, v)
            cell.font = font_data
            cell.border = thin_border
            if c_idx in [10, 11]:
                cell.alignment = align_left
            else:
                cell.alignment = align_center
            if c_idx in [1, 3, 4]:
                cell.number_format = '@'
            elif c_idx == 9:
                cell.number_format = '#,##0.00'
            else:
                cell.number_format = 'General'
                
    widths_log = {'A': 24.0, 'B': 12.0, 'C': 16.0, 'D': 22.0, 'E': 28.0, 'F': 14.0, 'G': 14.0, 'H': 28.0, 'I': 12.0, 'J': 16.0, 'K': 40.0, 'L': 20.0, 'M': 12.0}
    for col_l, w in widths_log.items():
        ws_log.column_dimensions[col_l].width = w
    ws_log.views.sheetView[0].showGridLines = True
    
    # 确定输出文件名
    if not output_path:
        today_str = datetime.datetime.now().strftime('%Y%m%d')
        output_path = os.path.join(work_dir, f"{today_str}-劳务费用明细表.xlsx")
        
    wb.save(output_path)
    print(f"\n【成功】费用明细表已顺利导出至: {output_path}")
    return True

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='雷允上语料库劳务费用明细表一键生成工具')
    parser.add_argument('--dir', help='工作目录（默认当前目录）', default=None)
    parser.add_argument('--prog', help='进度表文件路径', default=None)
    parser.add_argument('--corpus', help='语料明文表文件路径', default=None)
    parser.add_argument('--user', help='用户明文表文件路径', default=None)
    parser.add_argument('--output', help='输出结果文件路径', default=None)
    args = parser.parse_args()
    
    generate_settlement(
        work_dir=args.dir,
        prog_path=args.prog,
        corpus_path=args.corpus,
        user_path=args.user,
        output_path=args.output
    )
