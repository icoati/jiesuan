# -*- coding: utf-8 -*-
"""
====================================================
心血管内科报表生成脚本（纯 Python 原生版）
与原版 统计报表.js 业务逻辑、字段映射与表单格式 100% 一致

使用方法：
  1. 把两个系统导出的 xlsx 文件放到本脚本同一文件夹
  2. 在终端运行：python 统计报表.py
  3. 生成的报表会出现在同一文件夹：统计总表.xlsx
====================================================
"""

import os
import sys
import argparse
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))
OUTPUT_NAME = '统计总表.xlsx'


def convert_age_group(age) -> str:
    """年龄转年龄段（5档区间分类，与原版 JS 严格对齐）"""
    if age is None or age == '':
        return ''
    try:
        n = int(float(age))
    except (ValueError, TypeError):
        return ''
    if n <= 0:
        return ''
    if n <= 40:
        return '40岁以下'
    if n <= 50:
        return '41-50岁'
    if n <= 60:
        return '51-60岁'
    if n <= 70:
        return '61-70岁'
    return '71岁以上'


def find_source_files(target_dir: str):
    """自动从目标文件夹查找包含「答卷记录」和「项目人员」工作表的 xlsx/xls 文件"""
    candidates = [
        f for f in os.listdir(target_dir)
        if (f.endswith('.xlsx') or f.endswith('.xls'))
        and f != OUTPUT_NAME
        and not f.startswith('~$')
    ]

    src_file = None
    staff_file = None

    for file in candidates:
        file_path = os.path.join(target_dir, file)
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            names = wb.sheetnames
            wb.close()

            if not src_file and '答卷记录' in names:
                src_file = {'path': file_path, 'name': file}
            if not staff_file and '项目人员' in names:
                staff_file = {'path': file_path, 'name': file}
        except Exception:
            pass

        if src_file and staff_file:
            break

    return src_file, staff_file


def read_sheet_rows(file_path: str, sheet_name: str) -> list:
    """读取指定工作表全部数据为二维列表（保留原始排布）"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"文件中未找到「{sheet_name}」工作表: {file_path}")
    ws = wb[sheet_name]
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append(list(r))
    wb.close()
    return rows


def build_summary(staff_data: list) -> list:
    """构建 Sheet1: 统计汇总"""
    prov_map = {}
    total = 0

    for i in range(1, len(staff_data)):
        row = staff_data[i]
        if not row:
            continue
        task_name = str(row[0] or '') if len(row) > 0 else ''
        province = str(row[5] or '') if len(row) > 5 else ''
        try:
            count = int(float(row[7])) if len(row) > 7 and row[7] is not None else 0
        except (ValueError, TypeError):
            count = 0

        if '心血管内科' in task_name:
            if province not in prov_map:
                prov_map[province] = 0
            prov_map[province] += count
            total += count

    provinces = sorted(prov_map.keys())
    result = [['医院所在省', '药品规格', '任务数量']]
    for prov in provinces:
        spec = '片剂' if prov == '上海市' else '滴丸'
        result.push([prov, spec, prov_map[prov]]) if hasattr(result, 'push') else result.append([prov, spec, prov_map[prov]])

    result.append(['总计', None, total])
    return result


def build_staff_sheet(staff_data: list) -> list:
    """构建 Sheet2: 人员维度"""
    staff = []
    for i in range(1, len(staff_data)):
        row = staff_data[i]
        if not row:
            continue
        task_name = str(row[0] or '') if len(row) > 0 else ''
        if '心血管内科' in task_name:
            try:
                task_done = int(float(row[7])) if len(row) > 7 and row[7] is not None else 0
            except (ValueError, TypeError):
                task_done = 0

            staff.append({
                '任务名称': task_name,
                '姓名': str(row[1] or '') if len(row) > 1 else '',
                '手机号': '',
                '医院': str(row[2] or '') if len(row) > 2 else '',
                '科室': str(row[3] or '') if len(row) > 3 else '',
                '职称': str(row[4] or '') if len(row) > 4 else '',
                '医院所在省': str(row[5] or '') if len(row) > 5 else '',
                '医院所在市': str(row[6] or '') if len(row) > 6 else '',
                '开工状态': '已开工',
                '资质是否通过': '是',
                '任务完成情况': task_done,
                '更新时间': str(row[8] or '') if len(row) > 8 else ''
            })

    # 按更新时间倒序排序（与 JS 保持严格一致）
    staff.sort(key=lambda x: str(x['更新时间']), reverse=True)

    header = ['任务名称', '姓名', '手机号', '医院', '科室', '职称',
              '医院所在省', '医院所在市', '开工状态', '资质是否通过', '任务完成情况', '更新时间']

    result = [header]
    for s in staff:
        result.append([
            s['任务名称'], s['姓名'], s['手机号'], s['医院'], s['科室'], s['职称'],
            s['医院所在省'], s['医院所在市'], s['开工状态'], s['资质是否通过'],
            s['任务完成情况'], s['更新时间']
        ])
    return result


def build_case_sheet(source_data: list, staff_data: list) -> list:
    """构建 Sheet3: 案例维度"""
    staff_set = set()
    for i in range(1, len(staff_data)):
        row = staff_data[i]
        if not row:
            continue
        task_name = str(row[0] or '') if len(row) > 0 else ''
        if '心血管内科' in task_name:
            name = str(row[1] or '').strip() if len(row) > 1 else ''
            if name:
                staff_set.add(name)

    header = [
        '序号', '任务名称', '项目人员', '手机号', '医院', '医院所在省份',
        '医院所在市', '结算单号', '答卷标题', '审核状态', '审核未通过原因',
        '提交答卷时间', '1、姓名（首字母缩写）：', '2、性别：', '3、年龄：（岁）',
        '4、诊断日期:', '5、初诊 / 复诊', '6、临床表现（可多选）：',
        '7、疾病诊断（可多选）：', '8、既往史（可多选）：', '33、其他改善：'
    ]

    result = [header]
    seq_num = 1

    for i in range(1, len(source_data)):
        row = source_data[i]
        if not row:
            continue
        staff_name = str(row[2] or '').strip() if len(row) > 2 else ''

        def get_val(idx):
            return str(row[idx]) if (len(row) > idx and row[idx] is not None) else ''

        if staff_name in staff_set:
            result.append([
                seq_num,
                get_val(1),
                staff_name,
                '',
                get_val(3),
                get_val(4),
                get_val(5),
                '',
                get_val(8),
                get_val(12),
                get_val(13),
                get_val(14),
                get_val(15),
                get_val(16),
                convert_age_group(get_val(17)),
                get_val(18),
                get_val(19),
                get_val(20),
                get_val(21),
                get_val(22),
                get_val(33)
            ])
            seq_num += 1

    return result


def apply_style(ws, data: list):
    """设置工作表高质感商务浅蓝样式与自适应列宽"""
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    header_font = Font(name='微软雅黑', size=10, bold=True)
    header_fill = PatternFill(start_color='C8DCF0', end_color='C8DCF0', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    data_font = Font(name='微软雅黑', size=10, bold=False)
    data_align = Alignment(horizontal='center', vertical='center')

    col_count = len(data[0]) if data else 0

    # 填充表头样式
    ws.row_dimensions[1].height = 26.0
    for c in range(1, col_count + 1):
        cell = ws.cell(1, c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # 数据行排版
    for r_idx in range(2, len(data) + 1):
        ws.row_dimensions[r_idx].height = 20.0
        for c_idx in range(1, col_count + 1):
            cell = ws.cell(r_idx, c_idx)
            cell.font = data_font
            cell.alignment = data_align
            cell.border = thin_border

    # 自适应列宽（中文2字符，英文1字符）
    for c in range(col_count):
        max_len = 0
        for r in range(min(len(data), 60)):
            val = str(data[r][c] or '')
            w = 0
            for ch in val:
                w += 2 if ord(ch) > 127 else 1
            if w > max_len:
                max_len = w
        col_letter = get_column_letter(c + 1)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 42)


def generate_shangyao_report(src_path: str, staff_path: str, output_path: str = None) -> bool:
    """执行报表生成完整业务逻辑"""
    if not output_path:
        output_path = os.path.join(SCRIPT_DIR, OUTPUT_NAME)

    print("[*] 正在读取数据...")
    source_data = read_sheet_rows(src_path, '答卷记录')
    print(f"   [OK] 答卷记录: {os.path.basename(src_path)} ({len(source_data) - 1} 条记录)")

    staff_data = read_sheet_rows(staff_path, '项目人员')
    print(f"   [OK] 项目人员: {os.path.basename(staff_path)} ({len(staff_data) - 1} 位人员)")

    print("\n[*] 正在生成报表...")
    summary_data = build_summary(staff_data)
    staff_sheet_data = build_staff_sheet(staff_data)
    case_sheet_data = build_case_sheet(source_data, staff_data)

    print(f"   [OK] 统计汇总: {len(summary_data) - 2} 个省份")
    print(f"   [OK] 人员维度: {len(staff_sheet_data) - 1} 位人员")
    print(f"   [OK] 案例维度: {len(case_sheet_data) - 1} 条案例")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 移除默认工作表

    # 1. 统计汇总
    ws_sum = wb.create_sheet(title='统计汇总')
    for row in summary_data:
        ws_sum.append(row)
    apply_style(ws_sum, summary_data)

    # 统计汇总总计行合并 A、B 列（总计横跨 A、B 列，总任务数量规范居于 C 列）
    total_row = len(summary_data)
    ws_sum.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)

    # 2. 人员维度
    ws_staff = wb.create_sheet(title='人员维度')
    for row in staff_sheet_data:
        ws_staff.append(row)
    apply_style(ws_staff, staff_sheet_data)

    # 3. 案例维度
    ws_case = wb.create_sheet(title='案例维度')
    for row in case_sheet_data:
        ws_case.append(row)
    apply_style(ws_case, case_sheet_data)

    wb.save(output_path)
    print(f"\n[OK] 报表已生成: {output_path}")
    print("   包含工作表: 统计汇总 / 人员维度 / 案例维度")
    return True


def main():
    print("=" * 50)
    print("  心血管内科报表生成工具 (Python 纯净版)")
    print("=" * 50)
    print(f"[*] 工作目录: {SCRIPT_DIR}")

    parser = argparse.ArgumentParser(description='心血管内科报表生成工具')
    parser.add_argument('--src', help='答卷记录文件路径', default=None)
    parser.add_argument('--staff', help='项目人员文件路径', default=None)
    parser.add_argument('--output', help='输出结果文件路径', default=None)
    args = parser.parse_args()

    src_path = args.src
    staff_path = args.staff

    if not (src_path and staff_path):
        print("\n[*] 正在查找文件夹中的源文件...")
        src_info, staff_info = find_source_files(SCRIPT_DIR)
        if not src_info:
            print("[ERR] 未找到包含「答卷记录」工作表的 xlsx 文件")
            print("   请将系统导出的答卷记录表放到本文件夹后重试。")
            sys.exit(1)
        if not staff_info:
            print("[ERR] 未找到包含「项目人员」工作表的 xlsx 文件")
            print("   请将系统导出的项目人员列表放到本文件夹后重试。")
            sys.exit(1)
        src_path = src_info['path']
        staff_path = staff_info['path']

    out_path = args.output or os.path.join(SCRIPT_DIR, OUTPUT_NAME)
    generate_shangyao_report(src_path, staff_path, out_path)


if __name__ == '__main__':
    main()
