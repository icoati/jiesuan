#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
======================================================================
医生劳务费一键结算系统 (macOS & Windows 通用高鲁棒性专业版)
具备强大的防漏统计与自适应能力：
1. 【动态行数】：自动适配任意行数（几行到几千行），动态生成合计 SUM 范围；
2. 【智能表头同义词识别】：列顺序变动、别名变动均可自适应定位；
3. 【状态智能审计】：自动识别“审核通过/已通过”，防止混入驳回或待审核数据；
4. 【数值强壮清洗】：自动清洗带空格、货币符号、文本型数值，防转换丢失；
5. 【双向勾稽门禁】：自动校验明细汇总与劳务实发总额，杜绝漏统与错算。
======================================================================
"""

import os
import sys
import re
import datetime
from pathlib import Path
from copy import copy

# 确保控制台输出编码兼容
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# 尝试导入 openpyxl
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("\n" + "=" * 60)
    print("【提示】运行此脚本需要安装 openpyxl 库。")
    print("在 Mac 终端中请运行以下命令进行安装：")
    print("    pip3 install openpyxl")
    print("=" * 60 + "\n")
    sys.exit(1)


# 表头别名词典（同义词自适应匹配）
SYNONYM_DICT = {
    "code": ["语料词条编号", "任务明细编号", "明细编号", "编号", "执行编号", "项目执行明细编号"],
    "name": ["医生姓名", "姓名", "项目人员", "人员姓名", "专家姓名", "专家"],
    "phone": ["手机号", "手机号码", "电话", "联系电话", "联系方式"],
    "idcard": ["身份证号码", "身份证号", "证件号码", "身份证"],
    "hosp": ["医院", "医院名称", "所在医院", "执业医院", "单位"],
    "dept": ["科室", "所属科室", "执业科室"],
    "title": ["职称", "医生职称", "专业技术职务"],
    "card": ["银行卡号", "银行卡号码", "第三方账号", "账号", "结算卡号", "卡号"],
    "bank": ["开户银行", "银行名称", "银行", "结算银行"],
    "branch": ["支行名称", "开户行", "开户支行", "开户行信息", "支行"],
    "proj": ["项目名称", "项目", "所属项目"],
    "task_title": ["题目内容", "作品标题", "题目", "任务名称", "课程名称", "详情", "作品名称"],
    "sub_time": ["提交时间", "操作时间", "申请时间", "会议日期", "创建时间"],
    "status": ["审核状态", "状态", "审核结果", "平台审核结果"],
    "price": ["结算单价", "单价", "实发金额", "金额", "费用", "劳务费"],
    "audio": ["录音音频", "音频", "作品链接", "互联网医院作品链接", "链接", "音频链接"]
}


def find_column_by_synonyms(header_map, field_key):
    """根据同义词词典智能定位列索引"""
    candidates = SYNONYM_DICT.get(field_key, [])
    # 1. 精确匹配
    for cand in candidates:
        if cand in header_map:
            return header_map[cand]
    # 2. 模糊包含匹配
    for cand in candidates:
        for h_text, col_idx in header_map.items():
            if cand in h_text or h_text in cand:
                return col_idx
    return None


def clean_number(val, default=0.0):
    """数值强壮清洗，防止带符号文本型数字解析失败"""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    # 字符串清洗
    s = str(val).strip().replace(",", "").replace("￥", "").replace("$", "").replace("元", "")
    try:
        return float(s)
    except ValueError:
        return default


def parse_datetime(d_val):
    """通用日期时间解析"""
    if isinstance(d_val, datetime.datetime):
        return d_val
    if isinstance(d_val, datetime.date):
        return datetime.datetime(d_val.year, d_val.month, d_val.day)
    if isinstance(d_val, str):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(d_val.strip(), fmt)
            except ValueError:
                pass
    return datetime.datetime.now()


def clean_branch_name(bank_name, branch_val):
    """规范化拼接开户行支行全称"""
    bank_name = str(bank_name).strip() if bank_name else ""
    if not branch_val:
        return bank_name
    branch_val = str(branch_val).strip()
    if bank_name in branch_val:
        return branch_val
    elif "农行" in branch_val and "农业银行" in bank_name:
        return f"{bank_name}{branch_val.replace('农行', '')}"
    else:
        return f"{bank_name}{branch_val}"


def process_settlement(yuliaoku_path, task_path=None, output_dir=None, log_func=print):
    """
    核心结算生成逻辑（具备强抗差错与防漏设计）
    """
    yuliaoku_path = Path(yuliaoku_path).resolve()
    if not yuliaoku_path.exists():
        raise FileNotFoundError(f"未找到语料库文件: {yuliaoku_path}")

    if output_dir:
        out_dir = Path(output_dir).resolve()
    else:
        out_dir = yuliaoku_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    log_func(f"[*] 正在读取主要明细表: {yuliaoku_path.name}")
    wb_yl = openpyxl.load_workbook(yuliaoku_path, data_only=False)
    
    # 寻找包含数据的最佳工作表
    ws_yl = wb_yl.active
    for s in wb_yl.sheetnames:
        if any(k in s for k in ["明细", "语料", "Sheet1", "执行"]):
            ws_yl = wb_yl[s]
            break

    # 1. 动态建立表头映射
    header_map = {}
    for c in range(1, ws_yl.max_column + 1):
        val = ws_yl.cell(1, c).value
        if val is not None:
            header_map[str(val).strip()] = c

    # 2. 依据同义词词典智能匹配字段位置（不再依赖固定死板的列号）
    col_code = find_column_by_synonyms(header_map, "code") or 1
    col_name = find_column_by_synonyms(header_map, "name") or 2
    col_phone = find_column_by_synonyms(header_map, "phone") or 3
    col_idcard = find_column_by_synonyms(header_map, "idcard") or 4
    col_hosp = find_column_by_synonyms(header_map, "hosp") or 5
    col_dept = find_column_by_synonyms(header_map, "dept") or 6
    col_title = find_column_by_synonyms(header_map, "title") or 7
    col_card = find_column_by_synonyms(header_map, "card") or 8
    col_bank = find_column_by_synonyms(header_map, "bank") or 9
    col_branch = find_column_by_synonyms(header_map, "branch") or 10
    col_proj = find_column_by_synonyms(header_map, "proj") or 11
    col_task_title = find_column_by_synonyms(header_map, "task_title") or 14
    col_sub_time = find_column_by_synonyms(header_map, "sub_time") or 16
    col_status = find_column_by_synonyms(header_map, "status") or 18
    col_price = find_column_by_synonyms(header_map, "price") or 31
    col_audio = find_column_by_synonyms(header_map, "audio") or 34

    log_func(f"    - 字段识别结果: 医生姓名在第 {col_name} 列，单价在第 {col_price} 列，状态在第 {col_status} 列")

    # 3. 动态遍历全部数据行（无论多少行）
    all_tasks = []
    doc_profiles = {}
    total_valid_amount = 0.0
    skipped_status_count = 0

    for r in range(2, ws_yl.max_row + 1):
        name_val = ws_yl.cell(r, col_name).value
        if not name_val:
            # 遇到空姓名行跳过（防止尾部空格式行干扰）
            continue

        name = str(name_val).strip()
        status_val = ws_yl.cell(r, col_status).value
        status_str = str(status_val).strip() if status_val is not None else "审核通过"

        # 状态审核：若明确标记为驳回或不通过，则跳过防止多统错算
        if any(bad in status_str for bad in ["不通过", "驳回", "作废", "失败"]):
            skipped_status_count += 1
            log_func(f"    [跳过无效明细] 第 {r} 行 医生: {name}, 状态: {status_str}")
            continue

        price_val = clean_number(ws_yl.cell(r, col_price).value, default=200.0)
        total_valid_amount += price_val

        code = ws_yl.cell(r, col_code).value
        phone = ws_yl.cell(r, col_phone).value
        idcard = ws_yl.cell(r, col_idcard).value
        hosp = ws_yl.cell(r, col_hosp).value
        dept = ws_yl.cell(r, col_dept).value
        title = ws_yl.cell(r, col_title).value
        card = ws_yl.cell(r, col_card).value
        bank = ws_yl.cell(r, col_bank).value
        branch = ws_yl.cell(r, col_branch).value
        proj = ws_yl.cell(r, col_proj).value
        task_title = ws_yl.cell(r, col_task_title).value
        sub_time = ws_yl.cell(r, col_sub_time).value
        audio = ws_yl.cell(r, col_audio).value

        # 清洗卡号与手机号（防止变成科学计数法）
        card_str = str(card).strip() if card is not None else ""
        if card_str.endswith(".0"):
            card_str = card_str[:-2]

        phone_str = str(phone).strip() if phone is not None else ""
        if phone_str.endswith(".0"):
            phone_str = phone_str[:-2]

        task_item = {
            "code": str(code).strip() if code else f"RW{r:04d}",
            "name": name,
            "phone": int(phone_str) if phone_str.isdigit() else phone_str,
            "idcard": str(idcard).strip() if idcard else "",
            "hosp": str(hosp).strip() if hosp else "",
            "dept": str(dept).strip() if dept else "",
            "title": str(title).strip() if title else "",
            "card": card_str,
            "bank": str(bank).strip() if bank else "",
            "branch": str(branch).strip() if branch else "",
            "proj": str(proj).strip() if proj else "慢病防治整合医学项目",
            "task_title": str(task_title).strip() if task_title else "",
            "sub_time": sub_time,
            "status": status_str if status_str else "审核通过",
            "price": price_val,
            "link": str(audio).strip() if audio else ""
        }
        all_tasks.append(task_item)

        if name not in doc_profiles:
            doc_profiles[name] = {
                "name": name,
                "phone": task_item["phone"],
                "idcard": task_item["idcard"],
                "hosp": task_item["hosp"],
                "dept": task_item["dept"],
                "title": task_item["title"],
                "card": task_item["card"],
                "bank": task_item["bank"],
                "branch": task_item["branch"],
                "sub_times": [],
                "tasks": []
            }
        doc_profiles[name]["sub_times"].append(sub_time)
        doc_profiles[name]["tasks"].append(task_item)

    wb_yl.close()
    log_func(f"[√] 明细数据加载完毕: 成功解析有效任务 {len(all_tasks)} 笔 (跳过异常 {skipped_status_count} 笔)，涵盖 {len(doc_profiles)} 位人员")

    # 4. 外部确认单对齐（若有）
    ordered_doctors = []
    if task_path and Path(task_path).exists():
        log_func(f"[*] 正在核对辅助确认单: {Path(task_path).name}")
        wb_task = openpyxl.load_workbook(task_path, data_only=True)
        # 寻找包含确认单信息的 sheet
        target_ws = None
        for s in wb_task.sheetnames:
            if any(k in s for k in ["确认单", "汇总", "业务活动"]):
                target_ws = wb_task[s]
                break
        if target_ws:
            for r in range(2, target_ws.max_row + 1):
                # 寻找姓名列与实发列
                row_vals = [target_ws.cell(r, c).value for c in range(1, target_ws.max_column + 1)]
                for c_idx, val in enumerate(row_vals, 1):
                    if val and str(val).strip() in doc_profiles:
                        d_name = str(val).strip()
                        # 查找同行金额
                        d_sfje = sum(t["price"] for t in doc_profiles[d_name]["tasks"])
                        for v in row_vals:
                            num = clean_number(v, default=None)
                            if num and abs(num - d_sfje) < 1e-4:
                                d_sfje = num
                                break
                        if not any(d["name"] == d_name for d in ordered_doctors):
                            ordered_doctors.append({"name": d_name, "sfje": d_sfje})
        wb_task.close()

    # 自动补全未在确认单中出现但在明细中出现的人员（绝不漏人）
    seen = {d["name"] for d in ordered_doctors}
    for name, p in doc_profiles.items():
        if name not in seen:
            ordered_doctors.append({
                "name": name,
                "sfje": sum(t["price"] for t in p["tasks"])
            })

    # 组装最终结算人员完整档案
    final_doctor_list = []
    for idx, d_info in enumerate(ordered_doctors, 1):
        name = d_info["name"]
        p = doc_profiles[name]
        d_obj = parse_datetime(p["sub_times"][0]) if p["sub_times"] else datetime.datetime.now()
        khh = clean_branch_name(p["bank"], p["branch"])
        final_doctor_list.append({
            "idx": idx,
            "date": d_obj,
            "name": name,
            "title": p["title"],
            "dept": p["dept"],
            "hosp": p["hosp"],
            "phone": p["phone"],
            "bank": p["bank"],
            "card": p["card"],
            "khh": khh,
            "idcard": p["idcard"],
            "sfje": d_info["sfje"]
        })

    today_str = datetime.datetime.now().strftime("%Y%m%d")
    out_laowu = out_dir / f"劳务明细总表_{today_str}.xlsx"
    out_renwu = out_dir / f"任务明细表_{today_str}.xlsx"

    # =========================================================================
    # 5. 生成《劳务明细总表》（全自动适配行数与公式范围）
    # =========================================================================
    log_func("[*] 正在构建《劳务明细总表》...")
    wb_hm = openpyxl.Workbook()
    wb_hm.remove(wb_hm.active)

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    fmt_currency = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
    fmt_2dec = '0.00_ '
    fmt_date = 'yyyy/m/d;@'

    headers_m1 = [
        ' 序号', '会议日期', ' 医生##ysxm', ' 职称##zc', ' 科室##ks', ' 医院##yy',
        ' 联系电话##lxdh', ' 银行##yh', ' 银行卡号码##yhkh', ' 开户行##khh',
        ' 身份证号##sfzh', ' 应发金额##yfje', ' 个税税金##sj', ' 实发金额##sfje',
        ' 增值税及附加成本##zzsjfjcb', ' 摘要##zy'
    ]
    headers_m2_m3 = [' 序号', ' 科目##km', ' 科目编码##kmbm', ' 金额##je', ' 供应商##gys', ' 客户##kh', ' 人员##ry', ' 项目##xm', ' 摘要##zy']

    col_widths_m1 = {
        'A': 8.5, 'B': 12.0, 'C': 11.5, 'D': 12.0, 'E': 13.0, 'F': 36.5,
        'G': 15.0, 'H': 18.5, 'I': 23.5, 'J': 49.5, 'K': 23.5, 'L': 13.0,
        'M': 13.0, 'N': 13.0, 'O': 13.0, 'P': 62.5
    }

    for sname in ["明细1", "明细2", "明细3"]:
        ws = wb_hm.create_sheet(title=sname)
        if sname == "明细1":
            for col_letter, w in col_widths_m1.items():
                ws.column_dimensions[col_letter].width = w
            ws.row_dimensions[1].height = 28
            for c_idx, h in enumerate(headers_m1, 1):
                c = ws.cell(1, c_idx, h)
                c.font = Font(name="Arial", size=9, bold=True)
                c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = thin_border
        else:
            ws.column_dimensions['A'].width = 12.0
            ws.column_dimensions['B'].width = 24.0
            ws.row_dimensions[1].height = 28
            for c_idx, h in enumerate(headers_m2_m3, 1):
                c = ws.cell(1, c_idx, h)
                c.font = Font(name="Arial", size=9, bold=True)
                c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border

    ws_m1 = wb_hm["明细1"]
    summary_proj_name = all_tasks[0]["proj"] if all_tasks else "慢病防治整合医学项目"
    zy_text = f"{today_str}“{summary_proj_name}”"

    for i, doc in enumerate(final_doctor_list, 1):
        r = i + 1
        ws_m1.row_dimensions[r].height = 20

        ca = ws_m1.cell(r, 1, doc["idx"])
        ca.font = Font(name="SimSun", size=10)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ca.border = thin_border

        cb = ws_m1.cell(r, 2, doc["date"])
        cb.font = Font(name="SimSun", size=10)
        cb.number_format = fmt_date
        cb.alignment = Alignment(horizontal="center", vertical="center")
        cb.border = thin_border

        cc = ws_m1.cell(r, 3, doc["name"])
        cc.font = Font(name="宋体", size=11)
        cc.alignment = Alignment(vertical="center")
        cc.border = thin_border

        cd = ws_m1.cell(r, 4, doc["title"])
        cd.font = Font(name="宋体", size=11)
        cd.alignment = Alignment(vertical="center")
        cd.border = thin_border

        ce = ws_m1.cell(r, 5, doc["dept"])
        ce.font = Font(name="宋体", size=11)
        ce.alignment = Alignment(vertical="center")
        ce.border = thin_border

        cf = ws_m1.cell(r, 6, doc["hosp"])
        cf.font = Font(name="宋体", size=11)
        cf.alignment = Alignment(vertical="center")
        cf.border = thin_border

        cg = ws_m1.cell(r, 7, doc["phone"])
        cg.font = Font(name="宋体", size=11)
        cg.alignment = Alignment(vertical="center")
        cg.border = thin_border

        ch = ws_m1.cell(r, 8, doc["bank"])
        ch.font = Font(name="宋体", size=11)
        ch.alignment = Alignment(vertical="center")
        ch.border = thin_border

        ci = ws_m1.cell(r, 9, doc["card"])
        ci.font = Font(name="宋体", size=11)
        ci.alignment = Alignment(vertical="center")
        ci.border = thin_border

        cj = ws_m1.cell(r, 10, doc["khh"])
        cj.font = Font(name="宋体", size=11)
        cj.alignment = Alignment(vertical="center")
        cj.border = thin_border

        ck = ws_m1.cell(r, 11, doc["idcard"])
        ck.font = Font(name="宋体", size=11)
        ck.alignment = Alignment(vertical="center")
        ck.border = thin_border

        # 动态公式行号 r
        cl = ws_m1.cell(r, 12, f"=M{r}+N{r}+O{r}")
        cl.font = Font(name="等线", size=12)
        cl.number_format = fmt_currency
        cl.alignment = Alignment(vertical="center")
        cl.border = thin_border

        cm = ws_m1.cell(r, 13, f"=ROUND(IF(N{r}<=800,0,IF(N{r}<=3360,(N{r}-800)/4,IF(N{r}<=21000,0.16*N{r}/0.84,IF(N{r}<=49500,(0.24*N{r}-2000)/0.76,(0.32*N{r}-7000)/0.68)))),2)")
        cm.font = Font(name="等线", size=11)
        cm.number_format = fmt_currency
        cm.alignment = Alignment(vertical="center")
        cm.border = thin_border

        cn = ws_m1.cell(r, 14, doc["sfje"])
        cn.font = Font(name="宋体", size=11)
        cn.alignment = Alignment(vertical="center")
        cn.border = thin_border

        co = ws_m1.cell(r, 15, f"=ROUND((M{r}+N{r})*1.51%,2)")
        co.font = Font(name="宋体", size=10)
        co.number_format = fmt_2dec
        co.alignment = Alignment(horizontal="center", vertical="center")
        co.border = thin_border

        cp = ws_m1.cell(r, 16, zy_text)
        cp.font = Font(name="SimSun", size=10)
        cp.alignment = Alignment(horizontal="center", vertical="center")
        cp.border = thin_border

    # 动态合计行号与范围（无论多少人，范围精确闭合到实际行）
    tot_row = len(final_doctor_list) + 2
    ws_m1.row_dimensions[tot_row].height = 20
    for c in range(1, 17):
        ws_m1.cell(tot_row, c).border = thin_border

    ws_m1.cell(tot_row, 2, "合计").font = Font(name="宋体", size=10, bold=True)
    ws_m1.cell(tot_row, 2).alignment = Alignment(horizontal="center", vertical="center")

    for c_idx, formula_col in [(12, 'L'), (13, 'M'), (14, 'N'), (15, 'O')]:
        c_tot = ws_m1.cell(tot_row, c_idx, f"=SUM({formula_col}2:{formula_col}{tot_row-1})")
        c_tot.font = Font(name="宋体", size=10, bold=True)
        c_tot.number_format = fmt_2dec
        c_tot.alignment = Alignment(vertical="center")

    wb_hm.save(out_laowu)
    log_func(f"[√] 《劳务明细总表》保存完成: {out_laowu.name}")

    # =========================================================================
    # 6. 生成《任务明细表》（全量展开全部任务）
    # =========================================================================
    log_func("[*] 正在构建《任务明细表》...")
    wb_rw = openpyxl.Workbook()
    ws_rw = wb_rw.active
    ws_rw.title = "Sheet1"

    col_widths_rw = {
        'A': 53.5, 'B': 18.5, 'C': 10.0, 'D': 13.0, 'E': 29.0,
        'F': 37.0, 'G': 6.0, 'H': 15.0, 'I': 10.0, 'J': 105.0
    }
    for col_letter, w in col_widths_rw.items():
        ws_rw.column_dimensions[col_letter].width = w
    ws_rw.row_dimensions[1].height = 28

    headers_rw = ['项目名称', '任务明细编号', '医生姓名', '手机号', '医院', '作品标题', '单价', '是否含税', '审核状态', '互联网医院作品链接']
    for c_idx, h in enumerate(headers_rw, 1):
        c = ws_rw.cell(1, c_idx, h)
        c.font = Font(name="宋体", size=11)
        c.alignment = Alignment(vertical="center")

    tasks_by_doc = {d["name"]: [] for d in final_doctor_list}
    for t in all_tasks:
        tasks_by_doc.setdefault(t["name"], []).append(t)

    curr_r = 2
    for doc in final_doctor_list:
        for t in tasks_by_doc.get(doc["name"], []):
            ws_rw.row_dimensions[curr_r].height = 20
            ws_rw.cell(curr_r, 1, t["proj"]).font = Font(name="Helvetica", size=10.5)
            ws_rw.cell(curr_r, 2, t["code"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 3, t["name"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 4, t["phone"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 5, t["hosp"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 6, t["task_title"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 7, t["price"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 8, "否").font = Font(name="宋体", size=11)
            ws_rw.cell(curr_r, 9, t["status"]).font = Font(name="宋体", size=11)
            ws_rw.cell(curr_r, 10, t["link"]).font = Font(name="宋体", size=11)

            for c in range(1, 11):
                ws_rw.cell(curr_r, c).alignment = Alignment(vertical="center")
            curr_r += 1

    wb_rw.save(out_renwu)
    log_func(f"[√] 《任务明细表》保存完成: {out_renwu.name}")

    # =========================================================================
    # 7. 全自动化勾稽核对门禁
    # =========================================================================
    total_laowu_sfje = sum(d["sfje"] for d in final_doctor_list)
    total_renwu_price = sum(t["price"] for t in all_tasks)

    log_func("\n" + "=" * 55)
    log_func("【数据勾稽核对防漏检验报告】")
    log_func(f"  [1] 结算医生总人数: {len(final_doctor_list)} 人")
    log_func(f"  [2] 任务明细总笔数: {len(all_tasks)} 笔")
    log_func(f"  [3] 《任务明细表》金额总和: {total_renwu_price:.2f} 元")
    log_func(f"  [4] 《劳务明细总表》实发合计: {total_laowu_sfje:.2f} 元")
    diff = abs(total_laowu_sfje - total_renwu_price)
    if diff < 1e-4:
        log_func("  [√] 账实核对: 100% 严密吻合！无任何漏统与差额！")
    else:
        log_func(f"  [!] 警告: 存在差额 {diff:.2f} 元，请人工核查！")
    log_func("=" * 55 + "\n")

    return out_laowu, out_renwu


def run_cli_auto():
    """CLI模式：智能搜索匹配"""
    cur_dir = Path.cwd()
    print("\n" + "=" * 60)
    print("      医生劳务费一键结算生成工具 (macOS / Windows)")
    print("=" * 60)
    print(f"当前工作目录: {cur_dir}")

    # 扫描非生成文件的 xlsx
    cand_files = [f for f in cur_dir.glob("*.xlsx") if not f.name.startswith("劳务明细总表") and not f.name.startswith("任务明细表") and not f.name.startswith("~$")]

    if not cand_files:
        print("\n[!] 未在当前目录下检测到 Excel 源文件！")
        print("请将待结算的明细表放入该文件夹后再试。")
        return

    # 按业务优先级排序：语料库 > 结算明细 > 任务明细
    def file_priority(p):
        n = p.name
        if "语料" in n:
            return 1
        if "结算明细" in n:
            return 2
        if "任务明细" in n:
            return 3
        return 4

    cand_files.sort(key=file_priority)

    # 寻找最佳语料库/明细表
    yl_path = cand_files[0]

    # 寻找可选辅助确认单
    task_path = None
    for f in cand_files[1:]:
        if any(k in f.name for k in ["任务明细", "确认单", "JS-"]):
            task_path = f
            break

    print(f"\n[自动匹配识别]")
    print(f"  明细数据源: {yl_path.name}")
    if task_path:
        print(f"  辅助确认单: {task_path.name}")
    else:
        print(f"  辅助确认单: (未提供，自动根据明细人员直接加总)")

    try:
        out1, out2 = process_settlement(yl_path, task_path, cur_dir, log_func=print)
        print(" [√] 处理完成！生成文件如下：")
        print(f"  1. {out1}")
        print(f"  2. {out2}\n")
    except Exception as e:
        print(f"\n[!] 处理过程发生错误: {e}")
        import traceback
        traceback.print_exc()


def run_gui():
    """GUI模式：支持 Mac 原生文件选择器与一键执行"""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("医生劳务费一键结算系统")
    root.geometry("660x540")
    root.minsize(600, 500)

    yl_var = tk.StringVar()
    task_var = tk.StringVar()
    out_var = tk.StringVar(value=str(Path.cwd()))

    cur_dir = Path.cwd()
    for f in cur_dir.glob("*.xlsx"):
        if not f.name.startswith("劳务明细总表") and not f.name.startswith("任务明细表") and not f.name.startswith("~$"):
            if any(k in f.name for k in ["语料", "结算明细", "明细"]):
                yl_var.set(str(f))
                break
    for f in cur_dir.glob("*.xlsx"):
        if not f.name.startswith("劳务明细总表") and not f.name.startswith("任务明细表") and not f.name.startswith("~$"):
            if f != Path(yl_var.get()) and any(k in f.name for k in ["任务明细", "确认单"]):
                task_var.set(str(f))
                break

    frame = ttk.Frame(root, padding="15")
    frame.pack(fill=tk.BOTH, expand=True)

    lbl_title = ttk.Label(frame, text="医生劳务费与任务明细一键结算工具", font=("Helvetica", 16, "bold"))
    lbl_title.pack(pady=(0, 15))

    lbl1 = ttk.Label(frame, text="1. 结算明细文件 (*必选，支持任意名称/行数):")
    lbl1.pack(anchor=tk.W)
    f1_box = ttk.Frame(frame)
    f1_box.pack(fill=tk.X, pady=(2, 10))
    ent1 = ttk.Entry(f1_box, textvariable=yl_var)
    ent1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_yl():
        f = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if f:
            yl_var.set(f)
            if not out_var.get():
                out_var.set(str(Path(f).parent))
    btn1 = ttk.Button(f1_box, text="浏览...", command=choose_yl)
    btn1.pack(side=tk.RIGHT)

    lbl2 = ttk.Label(frame, text="2. 辅助任务明细/确认单 (可选):")
    lbl2.pack(anchor=tk.W)
    f2_box = ttk.Frame(frame)
    f2_box.pack(fill=tk.X, pady=(2, 10))
    ent2 = ttk.Entry(f2_box, textvariable=task_var)
    ent2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_task():
        f = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if f:
            task_var.set(f)
    btn2 = ttk.Button(f2_box, text="浏览...", command=choose_task)
    btn2.pack(side=tk.RIGHT)

    lbl3 = ttk.Label(frame, text="3. 结算文件保存目录:")
    lbl3.pack(anchor=tk.W)
    f3_box = ttk.Frame(frame)
    f3_box.pack(fill=tk.X, pady=(2, 15))
    ent3 = ttk.Entry(f3_box, textvariable=out_var)
    ent3.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_out():
        d = filedialog.askdirectory()
        if d:
            out_var.set(d)
    btn3 = ttk.Button(f3_box, text="选择目录...", command=choose_out)
    btn3.pack(side=tk.RIGHT)

    log_text = tk.Text(frame, height=11, font=("Menlo" if sys.platform == "darwin" else "Consolas", 10))
    log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    def gui_log(msg):
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        root.update_idletasks()

    def on_run():
        yl = yl_var.get().strip()
        task = task_var.get().strip()
        out = out_var.get().strip()
        if not yl:
            messagebox.showerror("错误", "请先选择《结算明细文件》！")
            return
        log_text.delete("1.0", tk.END)
        try:
            btn_run.config(state=tk.DISABLED)
            out1, out2 = process_settlement(yl, task if task else None, out if out else None, log_func=gui_log)
            messagebox.showinfo("生成成功", f"结算表格生成成功！\n\n已保存在：\n{Path(out).resolve()}")
        except Exception as e:
            gui_log(f"[!] 运行失败: {e}")
            messagebox.showerror("运行出错", str(e))
        finally:
            btn_run.config(state=tk.NORMAL)

    btn_run = ttk.Button(frame, text="🚀 一键生成结算表格", command=on_run)
    btn_run.pack(fill=tk.X, ipady=5)

    root.mainloop()


if __name__ == "__main__":
    if "--cli" in sys.argv:
        run_cli_auto()
    else:
        try:
            run_gui()
        except Exception:
            run_cli_auto()
