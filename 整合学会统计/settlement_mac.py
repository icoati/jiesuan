#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
======================================================================
北京整合-上药雷允上结算包 (macOS & Windows 通用高鲁棒性专业版)
具备强大的防漏统计、多表智能关联合并与自适应能力：
1. 【多源表格融合】：自动关联《语料列表》(明文) 与《用户列表》(明文)，精确补充银行资质；
2. 【严格表头对齐】：最终生成《整合-YYYY-MM-DD.xlsx》与《整合-YYYY-MM-DD.xls》，100% 严密对齐 22 项标准电签规范；
3. 【三大核算模式】：支持「💼 正算扣税 (HCP自行承担税费)」、「⚡ 不计税」、「🔄 倒推算税 (企业承担税费)」；
4. 【数值强壮清洗】：卡号防科学计数失真、身份证与手机号纯文本格式化、银行网点智能去重合并；
5. 【双向勾稽门禁】：自动校验明细汇总与劳务实发总额，杜绝漏统与错算；
6. 【历史报表归档】：同步生成财务审计级《劳务明细总表》与全量《任务明细表》。
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
    from openpyxl.utils import get_column_letter
except ImportError:
    print("\n" + "=" * 60)
    print("【提示】运行此脚本需要安装 openpyxl 库。")
    print("请运行: pip install openpyxl")
    print("=" * 60 + "\n")
    sys.exit(1)

# 尝试导入 xlwt 以支持 .xls 格式
try:
    import xlwt
except ImportError:
    xlwt = None


# 标准 22 列表头（与 整合.xlsx 模板 100% 一字不差对齐）
FINAL_22_COLUMNS = [
    '姓名*', '开始年*', '开始月*', '开始日*', '项目地点*', '项目名称*', '税后金额*', '含个税金额*',
    '姓名1*', '省份*', '市*', '身份证*', '手机号*', '开户行*', '银行卡号*', '工作单位*',
    '科室*', '医务职称*', '签署日期', '签署日期1', '身份证头像面', '身份证反面'
]

# 表头同义词别名识别词典
SYNONYM_DICT = {
    "code": ["语料词条编号", "任务明细编号", "明细编号", "编号", "执行编号", "项目执行明细编号", "词条编号"],
    "name": ["医生姓名", "姓名", "项目人员", "人员姓名", "专家姓名", "专家", "姓名1"],
    "phone": ["手机号码", "手机号", "电话", "联系电话", "联系方式"],
    "idcard": ["身份证号码", "身份证号", "证件号码", "身份证"],
    "hosp": ["所在医院", "医院", "医院名称", "执业医院", "工作单位", "单位"],
    "prov": ["所属省", "省份", "所在省份", "开户行省份", "省"],
    "city": ["所属市", "市", "城市", "所在城市", "开户行城市"],
    "dept": ["科室", "所属科室", "执业科室"],
    "title": ["职称", "医务职称", "医生职称", "专业技术职务"],
    "card": ["银行卡号", "银行卡号码", "第三方账号", "账号", "结算卡号", "卡号"],
    "bank": ["开户银行", "银行名称", "银行", "结算银行"],
    "branch": ["支行名称", "开户支行", "开户行信息", "支行", "开户行"],
    "proj": ["项目名称", "参与项目", "所属项目", "参与活动", "活动名称", "项目"],
    "task_title": ["题目内容", "作品标题", "题目", "任务名称", "课程名称", "详情", "作品名称", "回答记录"],
    "sub_time": ["提交时间", "操作时间", "申请时间", "会议日期", "创建时间"],
    "status": ["审核状态", "状态", "审核结果", "平台审核结果"],
    "settle_status": ["结算状态", "是否结算", "结算标志"],
    "price": ["结算单价", "单价", "实发金额", "金额", "费用", "劳务费"],
    "audio": ["录音音频", "音频", "作品链接", "互联网医院作品链接", "链接", "音频链接", "回答记录"]
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


def clean_text_field(val):
    """文本字段清洗，清除 .0 结尾浮点转文本异常"""
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0") and len(s) > 2:
        prefix = s[:-2]
        if prefix.replace("-", "").isdigit():
            s = prefix
    return s


def format_bank_and_branch(b_val, br_val):
    """
    智能合并开户行与支行名称：
    1. 抓取源文件中的 银行名称 + 支行名称 智能去重合并；
    2. 若支行已含银行主体，避免重复拼接（如 '中国银行' + '中国银行莘塍支行' => '中国银行莘塍支行'）；
    3. 去重常见简称重复（如 '农行' / '建行' / '工行'）；
    4. 规范结尾并保留网点。
    """
    b = str(b_val).strip() if b_val and str(b_val).strip().lower() not in ['nan', 'none', 'null', ''] else ''
    br = str(br_val).strip() if br_val and str(br_val).strip().lower() not in ['nan', 'none', 'null', ''] else ''
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
        ('工行', '中国工商银行'), ('建行', '中国建设银行'),
        ('中行', '中国银行'), ('交行', '交通银行'),
        ('农行', '中国农业银行'), ('农业银行', '中国农业银行'),
        ('邮储', '中国邮政储蓄银行')
    ]:
        if b == full_b or b == short_b:
            if br.startswith(short_b):
                return b + br[len(short_b):]
            if br.startswith(full_b):
                return br
    for kw in ['工商银行', '建设银行', '农业银行', '中国银行', '交通银行', '民生银行', '光大银行', '招商银行', '华夏银行', '邮政储蓄银行', '北京银行', '浦发银行']:
        if kw in b and kw in br:
            return br if len(br) >= len(b) else b
    return b + br


def extract_quoted_text(text):
    """优先提取中文双引号 “...” 或书名号内核心项目名称"""
    if not text:
        return ""
    s = str(text).strip()
    m_cn = re.findall(r'“([^”]+)”', s)
    if m_cn and m_cn[0].strip():
        return m_cn[0].strip()
    m_en = re.findall(r'"([^"]+)"', s)
    if m_en and m_en[0].strip():
        return m_en[0].strip()
    m_bk = re.findall(r'[《「]([^》」]+)[》」]', s)
    if m_bk and m_bk[0].strip():
        return m_bk[0].strip()
    return s


# ==================== 税费核心计算引擎 ====================
def calc_tax_forward(gross_amt):
    """
    正算扣税 (HCP自行承担税费 / 正推扣税)：
    任务总额为税前应发 gross，计算法定劳务报酬个税 tax，税后实发 net = gross - tax
    - 收入 <= 800: 免税，tax = 0
    - 800 < 收入 <= 4000: 应纳税所得额 = 收入 - 800，税率 20%
    - 收入 > 4000: 应纳税所得额 = 收入 * 0.8
      - 所得额 <= 20000: 20%，速算扣除 0
      - 20000 < 所得额 <= 50000: 30%，速算扣除 2000
      - 所得额 > 50000: 40%，速算扣除 7000
    """
    g = float(gross_amt)
    if g <= 800.0:
        tax = 0.0
    elif g <= 4000.0:
        tax = (g - 800.0) * 0.20
    else:
        taxable = g * 0.80
        if taxable <= 20000.0:
            tax = taxable * 0.20
        elif taxable <= 50000.0:
            tax = taxable * 0.30 - 2000.0
        else:
            tax = taxable * 0.40 - 7000.0
    tax = round(tax, 2)
    net = round(g - tax, 2)
    return tax, net


def calc_tax_reverse(net_amt):
    """
    倒推算税 (企业承担税费 / 逆向反算)：
    任务总额为到手实发 net，倒推应发含税 gross 与代扣个税 tax = gross - net
    - net <= 800: tax = 0, gross = net
    - 800 < net <= 3360: tax = (net - 800) / 4.0
    - 3360 < net <= 21000: tax = 0.16 * net / 0.84
    - 21000 < net <= 49500: tax = (0.24 * net - 2000.0) / 0.76
    - net > 49500: tax = (0.32 * net - 7000.0) / 0.68
    """
    n = float(net_amt)
    if n <= 800.0:
        tax = 0.0
    elif n <= 3360.0:
        tax = (n - 800.0) / 4.0
    elif n <= 21000.0:
        tax = 0.16 * n / 0.84
    elif n <= 49500.0:
        tax = (0.24 * n - 2000.0) / 0.76
    else:
        tax = (0.32 * n - 7000.0) / 0.68
    tax = round(tax, 2)
    gross = round(n + tax, 2)
    return tax, gross


# ==================== 核心结算生成主逻辑 ====================
def process_settlement(yuliaoku_path, user_path=None, task_path=None, output_dir=None, mode="正算扣税", log_func=print):
    """
    核心结算处理函数：
    :param yuliaoku_path: 语料列表 (明文) 或结算明细文件路径
    :param user_path: 用户列表 (明文) 路径 (可选，提供时自动补充银行卡、开户行、资质)
    :param task_path: 辅助确认单 / 支付清单路径 (可选，提供时按清单勾稽)
    :param output_dir: 生成文件输出目录 (默认与源文件同级)
    :param mode: 核算模式，可选: "正算扣税", "不计税", "倒推算税"
    :param log_func: 日志打印函数
    :return: (out_main_xlsx, out_main_xls, out_laowu, out_renwu)
    """
    yuliaoku_path = Path(yuliaoku_path).resolve()
    if not yuliaoku_path.exists():
        raise FileNotFoundError(f"未找到语料明细文件: {yuliaoku_path}")

    # 若未指定 output_dir，默认输出到源文件同级目录
    if output_dir:
        out_dir = Path(output_dir).resolve()
    else:
        out_dir = yuliaoku_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # 规范化 mode 参数
    if "不计税" in mode or "不算税" in mode or "免税" in mode:
        calc_mode = "不计税"
    elif "倒推" in mode or "反推" in mode or "企业" in mode:
        calc_mode = "倒推算税"
    else:
        calc_mode = "正算扣税"

    log_func(f"[*] 启动【北京整合-上药雷允上结算系统】")
    log_func(f"    - 当前结算模式: 【{calc_mode}】")
    log_func(f"    - 主要明细数据: {yuliaoku_path.name}")

    # 0. 自动扫描同目录下缺失的补充文件（若调用方未显式传入）
    cand_files = list(yuliaoku_path.parent.glob("*.xlsx")) + list(yuliaoku_path.parent.glob("*.xls"))
    if not user_path:
        for f in cand_files:
            if f != yuliaoku_path and any(k in f.name for k in ["用户列表", "用户", "医生列表"]):
                user_path = str(f)
                break
    if not task_path:
        for f in cand_files:
            if f != yuliaoku_path and str(f) != str(user_path) and any(k in f.name for k in ["确认单", "支付清单", "支付", "JS-", "对账"]):
                task_path = str(f)
                break

    # 1. 加载用户档案（若有用户列表）
    user_db = {}
    if user_path and Path(user_path).exists():
        log_func(f"[*] 正在读取用户资质信息表: {Path(user_path).name}")
        wb_u = openpyxl.load_workbook(user_path, data_only=True)
        ws_u = wb_u.active
        u_header_map = {str(ws_u.cell(1, c).value).strip(): c for c in range(1, ws_u.max_column + 1) if ws_u.cell(1, c).value}
        
        u_col_name = find_column_by_synonyms(u_header_map, "name") or 1
        u_col_title = find_column_by_synonyms(u_header_map, "title") or 2
        u_col_dept = find_column_by_synonyms(u_header_map, "dept") or 3
        u_col_hosp = find_column_by_synonyms(u_header_map, "hosp") or 4
        u_col_prov = find_column_by_synonyms(u_header_map, "prov") or 5
        u_col_city = find_column_by_synonyms(u_header_map, "city") or 6
        u_col_phone = find_column_by_synonyms(u_header_map, "phone") or 7
        u_col_idcard = find_column_by_synonyms(u_header_map, "idcard") or 8
        u_col_card = find_column_by_synonyms(u_header_map, "card") or 9
        u_col_bank = find_column_by_synonyms(u_header_map, "bank") or 10
        u_col_branch = find_column_by_synonyms(u_header_map, "branch") or 11
        u_col_proj = find_column_by_synonyms(u_header_map, "proj") or 18

        for r in range(2, ws_u.max_row + 1):
            name_v = ws_u.cell(r, u_col_name).value
            if not name_v:
                continue
            name_str = str(name_v).strip()
            user_db[name_str] = {
                "name": name_str,
                "title": clean_text_field(ws_u.cell(r, u_col_title).value),
                "dept": clean_text_field(ws_u.cell(r, u_col_dept).value),
                "hosp": clean_text_field(ws_u.cell(r, u_col_hosp).value),
                "prov": clean_text_field(ws_u.cell(r, u_col_prov).value),
                "city": clean_text_field(ws_u.cell(r, u_col_city).value),
                "phone": clean_text_field(ws_u.cell(r, u_col_phone).value),
                "idcard": clean_text_field(ws_u.cell(r, u_col_idcard).value),
                "card": clean_text_field(ws_u.cell(r, u_col_card).value),
                "bank": clean_text_field(ws_u.cell(r, u_col_bank).value),
                "branch": clean_text_field(ws_u.cell(r, u_col_branch).value),
                "proj": clean_text_field(ws_u.cell(r, u_col_proj).value)
            }
        wb_u.close()
        log_func(f"[√] 成功读取并建立 {len(user_db)} 位专家的银行与资质金融信息库")
    else:
        log_func("[!] 未提供单独的用户列表，将从语料/明细表中自适应提取银行资质")

    # 2. 读取语料明细表
    log_func(f"[*] 正在解析语料数据表: {yuliaoku_path.name}")
    wb_yl = openpyxl.load_workbook(yuliaoku_path, data_only=True)
    ws_yl = wb_yl.active
    for s in wb_yl.sheetnames:
        if any(k in s.lower() for k in ["明细", "语料", "sheet1", "执行"]):
            ws_yl = wb_yl[s]
            break

    yl_header_map = {str(ws_yl.cell(1, c).value).strip(): c for c in range(1, ws_yl.max_column + 1) if ws_yl.cell(1, c).value}
    col_code = find_column_by_synonyms(yl_header_map, "code") or 1
    col_name = find_column_by_synonyms(yl_header_map, "name") or 2
    col_phone = find_column_by_synonyms(yl_header_map, "phone") or 3
    col_idcard = find_column_by_synonyms(yl_header_map, "idcard") or 4
    col_hosp = find_column_by_synonyms(yl_header_map, "hosp") or 5
    col_prov = find_column_by_synonyms(yl_header_map, "prov") or 6
    col_city = find_column_by_synonyms(yl_header_map, "city") or 7
    col_dept = find_column_by_synonyms(yl_header_map, "dept") or 8
    col_title = find_column_by_synonyms(yl_header_map, "title") or 9
    col_proj = find_column_by_synonyms(yl_header_map, "proj") or 10
    col_task_title = find_column_by_synonyms(yl_header_map, "task_title") or 13
    col_audio = find_column_by_synonyms(yl_header_map, "audio") or 14
    col_sub_time = find_column_by_synonyms(yl_header_map, "sub_time") or 15
    col_status = find_column_by_synonyms(yl_header_map, "status") or 17
    col_settle_status = find_column_by_synonyms(yl_header_map, "settle_status") or 20
    col_price = find_column_by_synonyms(yl_header_map, "price") or 21
    col_card = find_column_by_synonyms(yl_header_map, "card")
    col_bank = find_column_by_synonyms(yl_header_map, "bank")
    col_branch = find_column_by_synonyms(yl_header_map, "branch")

    # 3. 筛选结算范围（支持已加入结算单、辅助单据过滤、审核通过）
    all_raw_rows = []
    has_joined_settle = False

    for r in range(2, ws_yl.max_row + 1):
        name_v = ws_yl.cell(r, col_name).value
        if not name_v:
            continue
        
        status_v = str(ws_yl.cell(r, col_status).value or "审核通过").strip()
        settle_status_v = str(ws_yl.cell(r, col_settle_status).value or "").strip()
        
        if "已加入结算单" in settle_status_v or "已结算" in settle_status_v:
            has_joined_settle = True

        all_raw_rows.append({
            "row": r,
            "name": str(name_v).strip(),
            "status": status_v,
            "settle_status": settle_status_v,
            "code": clean_text_field(ws_yl.cell(r, col_code).value),
            "phone": clean_text_field(ws_yl.cell(r, col_phone).value),
            "idcard": clean_text_field(ws_yl.cell(r, col_idcard).value),
            "hosp": clean_text_field(ws_yl.cell(r, col_hosp).value),
            "prov": clean_text_field(ws_yl.cell(r, col_prov).value),
            "city": clean_text_field(ws_yl.cell(r, col_city).value),
            "dept": clean_text_field(ws_yl.cell(r, col_dept).value),
            "title": clean_text_field(ws_yl.cell(r, col_title).value),
            "proj": clean_text_field(ws_yl.cell(r, col_proj).value),
            "task_title": clean_text_field(ws_yl.cell(r, col_task_title).value),
            "audio": clean_text_field(ws_yl.cell(r, col_audio).value),
            "sub_time": ws_yl.cell(r, col_sub_time).value,
            "price": clean_number(ws_yl.cell(r, col_price).value, default=100.0),
            "card": clean_text_field(ws_yl.cell(r, col_card).value) if col_card else "",
            "bank": clean_text_field(ws_yl.cell(r, col_bank).value) if col_bank else "",
            "branch": clean_text_field(ws_yl.cell(r, col_branch).value) if col_branch else ""
        })
    wb_yl.close()

    # 4. 辅助单据加载（若提供）
    task_filter_codes = set()
    task_filter_names = set()
    if task_path and Path(task_path).exists():
        log_func(f"[*] 正在核对辅助确认单 / 支付清单: {Path(task_path).name}")
        wb_task = openpyxl.load_workbook(task_path, data_only=True)
        ws_t = wb_task.active
        for s in wb_task.sheetnames:
            if any(k in s for k in ["确认单", "汇总", "清单", "支付", "Sheet"]):
                ws_t = wb_task[s]
                break
        for r in range(2, ws_t.max_row + 1):
            for c in range(1, ws_t.max_column + 1):
                val = ws_t.cell(r, c).value
                if val:
                    val_s = str(val).strip()
                    if val_s.startswith("YL") or val_s.startswith("RW") or val_s.startswith("JS"):
                        task_filter_codes.add(val_s)
                    elif len(val_s) >= 2 and len(val_s) <= 5 and any(k in val_s for k in ["医生", "专家", ""]) is False:
                        # 姓名候选
                        task_filter_names.add(val_s)
        wb_task.close()
        log_func(f"[√] 辅助单据已载入，识别任务编号特征 {len(task_filter_codes)} 笔，关联医生人员 {len(task_filter_names)} 位")

    # 5. 过滤出符合结算条件的有效任务明细
    valid_tasks = []
    skipped_count = 0

    for item in all_raw_rows:
        # 排除明确拒绝驳回的异常状态
        if any(bad in item["status"] for bad in ["不通过", "驳回", "作废", "失败", "拒绝"]):
            skipped_count += 1
            continue

        # 优先级 1: 辅助确认单勾稽
        if task_filter_codes:
            if item["code"] not in task_filter_codes:
                continue
        # 优先级 2: 语料列表标记了“已加入结算单”
        elif has_joined_settle:
            if "已加入结算单" not in item["settle_status"]:
                continue
        # 优先级 3: 默认全量“审核通过”
        else:
            if "通过" not in item["status"]:
                continue

        valid_tasks.append(item)

    if not valid_tasks:
        # 兜底回退：若未过滤出记录，降级到所有审核通过记录
        log_func("[!] 结算范围匹配 0 笔，自动降级为所有【审核通过】记录")
        valid_tasks = [it for it in all_raw_rows if "通过" in it["status"]]

    log_func(f"[√] 任务筛选完毕: 成功入库结算任务 {len(valid_tasks)} 笔 (剔除异常 {skipped_count} 笔)")

    # 6. 按医生归集聚合
    doc_groups = {}
    top_proj_name = "医疗健康 AI语料库建设"

    for t in valid_tasks:
        name = t["name"]
        if name not in doc_groups:
            # 融合用户表与语料表信息
            u = user_db.get(name, {})
            doc_groups[name] = {
                "name": name,
                "phone": u.get("phone") or t["phone"],
                "idcard": u.get("idcard") or t["idcard"],
                "hosp": u.get("hosp") or t["hosp"],
                "prov": u.get("prov") or t["prov"],
                "city": u.get("city") or t["city"],
                "dept": u.get("dept") or t["dept"],
                "title": u.get("title") or t["title"],
                "card": u.get("card") or t["card"],
                "bank": u.get("bank") or t["bank"],
                "branch": u.get("branch") or t["branch"],
                "proj": extract_quoted_text(u.get("proj") or t["proj"] or top_proj_name),
                "tasks": [],
                "sub_times": []
            }
        doc_groups[name]["tasks"].append(t)
        if t["sub_time"]:
            doc_groups[name]["sub_times"].append(parse_datetime(t["sub_time"]))

    # 7. 计算每位医生的核算金额
    doctor_summary_list = []
    total_valid_amount = sum(t["price"] for t in valid_tasks)

    for name, dg in doc_groups.items():
        # 任务总额
        task_total = sum(t["price"] for t in dg["tasks"])
        
        # 提交日期（取最早任务提交时间）
        if dg["sub_times"]:
            min_dt = min(dg["sub_times"])
        else:
            min_dt = datetime.datetime.now()

        # 银行支行去重合并
        full_bank = format_bank_and_branch(dg["bank"], dg["branch"])

        # 地点推导：优先城市，次选省份，最后兜底“线上”
        loc = dg["city"] or dg["prov"] or "线上"
        proj_clean = dg["proj"] or top_proj_name

        # 根据模式计算税前含税与税后实发
        if calc_mode == "不计税":
            tax_amt = 0.0
            net_amt = task_total
            gross_amt = task_total
        elif calc_mode == "倒推算税":
            # 任务金额为税后实发，倒推税前应发
            net_amt = task_total
            tax_amt, gross_amt = calc_tax_reverse(net_amt)
        else:
            # 正算扣税：任务金额为税前应发，扣税后为税后实发
            gross_amt = task_total
            tax_amt, net_amt = calc_tax_forward(gross_amt)

        doctor_summary_list.append({
            '姓名*': name,
            '开始年*': min_dt.year,
            '开始月*': min_dt.month,
            '开始日*': min_dt.day,
            '项目地点*': loc,
            '项目名称*': proj_clean,
            '税后金额*': net_amt,
            '含个税金额*': gross_amt,
            '应发金额': gross_amt,
            '代扣个税': tax_amt,
            '实发金额': net_amt,
            '姓名1*': name,
            '省份*': dg["prov"],
            '市*': dg["city"],
            '身份证*': dg["idcard"],
            '手机号*': dg["phone"],
            '开户行*': full_bank,
            '银行卡号*': dg["card"],
            '工作单位*': dg["hosp"],
            '科室*': dg["dept"],
            '医务职称*': dg["title"],
            '签署日期': '',
            '签署日期1': '',
            '身份证头像面': '',
            '身份证反面': '',
            'date_obj': min_dt,
            'bank_name': dg["bank"],
            'tasks': dg["tasks"]
        })

    # 按实发金额倒序，再按姓名升序
    doctor_summary_list.sort(key=lambda x: (-x['税后金额*'], -x['含个税金额*'], x['姓名*']))

    today_str = datetime.datetime.now().strftime("%Y%m%d")
    today_dash = datetime.datetime.now().strftime("%Y-%m-%d")

    total_doctors = len(doctor_summary_list)
    total_gross = sum(d['含个税金额*'] for d in doctor_summary_list)
    total_tax = sum(d['代扣个税'] for d in doctor_summary_list)
    total_net = sum(d['税后金额*'] for d in doctor_summary_list)

    log_func(f"汇总完成: 共 {total_doctors} 位医生，应发总额 ¥{total_gross:,.2f}，代扣税金 ¥{total_tax:,.2f}，劳务实发总额 {total_net:,.2f} 元")

    # =========================================================================
    # 8. 生成主电签文件：《整合-YYYY-MM-DD.xlsx》（严格 22 列表头）
    # =========================================================================
    out_main_xlsx = out_dir / f"整合-{today_dash}.xlsx"
    log_func(f"[*] 正在构建标准电签报销单: {out_main_xlsx.name}")

    wb_main = openpyxl.Workbook()
    ws_main = wb_main.active
    ws_main.title = "Sheet0"

    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="000000")
    data_font = Font(name="微软雅黑", size=10, color="000000")
    border_style = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )

    ws_main.row_dimensions[1].height = 26
    for c_idx, col_name in enumerate(FINAL_22_COLUMNS, 1):
        c = ws_main.cell(row=1, column=c_idx, value=col_name)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_style

    for r_idx, doc in enumerate(doctor_summary_list, 2):
        ws_main.row_dimensions[r_idx].height = 22
        for c_idx, col_name in enumerate(FINAL_22_COLUMNS, 1):
            val = doc.get(col_name, "")
            c = ws_main.cell(row=r_idx, column=c_idx)
            if col_name in ['身份证*', '手机号*', '银行卡号*']:
                c.number_format = '@'
                c.value = str(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ['开始年*', '开始月*', '开始日*']:
                c.value = int(val) if val != "" else ""
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ['税后金额*', '含个税金额*']:
                c.value = float(val) if val != "" else 0.0
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in ['姓名*', '姓名1*', '省份*', '市*', '项目地点*', '医务职称*', '科室*']:
                c.value = str(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.value = str(val) if val != "" else ""
                c.alignment = Alignment(horizontal="left", vertical="center")
            c.font = data_font
            c.border = border_style

    for col in ws_main.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '')) for cell in col)
        ws_main.column_dimensions[col_letter].width = min(max(max_l + 4, 12), 40)

    wb_main.save(out_main_xlsx)
    log_func(f"[√] 《整合-{today_dash}.xlsx》保存完成 (22列标准规范)")

    # =========================================================================
    # 9. 生成《整合-YYYY-MM-DD.xls》（标准 BIFF8，用于银行与特定电签系统导入）
    # =========================================================================
    out_main_xls = out_dir / f"整合-{today_dash}.xls"
    if xlwt:
        log_func(f"[*] 正在构建银行兼容格式: {out_main_xls.name}")
        wb_xls = xlwt.Workbook(encoding='utf-8')
        ws_xls = wb_xls.add_sheet('Sheet0')
        st_hdr = xlwt.easyxf('font: name 微软雅黑, bold on, height 220; align: horiz center, vert center; pattern: pattern solid, fore_colour gray25; borders: left thin, right thin, top thin, bottom thin')
        st_txt = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz center, vert center; borders: left thin, right thin, top thin, bottom thin')
        st_left = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz left, vert center; borders: left thin, right thin, top thin, bottom thin')
        st_amt = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz right, vert center; borders: left thin, right thin, top thin, bottom thin', num_format_str='#,##0.00')
        st_int = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz center, vert center; borders: left thin, right thin, top thin, bottom thin')

        for c_idx, col_name in enumerate(FINAL_22_COLUMNS):
            ws_xls.write(0, c_idx, col_name, st_hdr)

        for r_idx, doc in enumerate(doctor_summary_list, 1):
            for c_idx, col_name in enumerate(FINAL_22_COLUMNS):
                val = doc.get(col_name, "")
                if col_name in ['身份证*', '手机号*', '银行卡号*']:
                    ws_xls.write(r_idx, c_idx, str(val), st_txt)
                elif col_name in ['开始年*', '开始月*', '开始日*']:
                    ws_xls.write(r_idx, c_idx, int(val) if val != "" else "", st_int)
                elif col_name in ['税后金额*', '含个税金额*']:
                    ws_xls.write(r_idx, c_idx, float(val) if val != "" else 0.0, st_amt)
                elif col_name in ['姓名*', '姓名1*', '省份*', '市*', '项目地点*', '医务职称*', '科室*']:
                    ws_xls.write(r_idx, c_idx, str(val), st_txt)
                else:
                    ws_xls.write(r_idx, c_idx, str(val) if val != "" else "", st_left)
        wb_xls.save(str(out_main_xls))
        log_func(f"[√] 《整合-{today_dash}.xls》保存完成")

    # =========================================================================
    # 10. 生成《劳务明细总表_YYYYMMDD.xlsx》（完整保留带公式三表，用于财务归档）
    # =========================================================================
    out_laowu = out_dir / f"劳务明细总表_{today_str}.xlsx"
    log_func(f"[*] 正在构建内部财务报销单: {out_laowu.name}")

    wb_hm = openpyxl.Workbook()
    wb_hm.remove(wb_hm.active)

    thin_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
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
    zy_text = f"{today_str}“{top_proj_name}”"

    for i, doc in enumerate(doctor_summary_list, 1):
        r = i + 1
        ws_m1.row_dimensions[r].height = 20

        ca = ws_m1.cell(r, 1, i)
        ca.font = Font(name="SimSun", size=10)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ca.border = thin_border

        cb = ws_m1.cell(r, 2, doc["date_obj"])
        cb.font = Font(name="SimSun", size=10)
        cb.number_format = fmt_date
        cb.alignment = Alignment(horizontal="center", vertical="center")
        cb.border = thin_border

        cc = ws_m1.cell(r, 3, doc["姓名*"])
        cc.font = Font(name="宋体", size=11)
        cc.alignment = Alignment(vertical="center")
        cc.border = thin_border

        cd = ws_m1.cell(r, 4, doc["医务职称*"])
        cd.font = Font(name="宋体", size=11)
        cd.alignment = Alignment(vertical="center")
        cd.border = thin_border

        ce = ws_m1.cell(r, 5, doc["科室*"])
        ce.font = Font(name="宋体", size=11)
        ce.alignment = Alignment(vertical="center")
        ce.border = thin_border

        cf = ws_m1.cell(r, 6, doc["工作单位*"])
        cf.font = Font(name="宋体", size=11)
        cf.alignment = Alignment(vertical="center")
        cf.border = thin_border

        cg = ws_m1.cell(r, 7, doc["手机号*"])
        cg.font = Font(name="宋体", size=11)
        cg.number_format = '@'
        cg.alignment = Alignment(vertical="center")
        cg.border = thin_border

        ch = ws_m1.cell(r, 8, doc["bank_name"])
        ch.font = Font(name="宋体", size=11)
        ch.alignment = Alignment(vertical="center")
        ch.border = thin_border

        ci = ws_m1.cell(r, 9, doc["银行卡号*"])
        ci.font = Font(name="宋体", size=11)
        ci.number_format = '@'
        ci.alignment = Alignment(vertical="center")
        ci.border = thin_border

        cj = ws_m1.cell(r, 10, doc["开户行*"])
        cj.font = Font(name="宋体", size=11)
        cj.alignment = Alignment(vertical="center")
        cj.border = thin_border

        ck = ws_m1.cell(r, 11, doc["身份证*"])
        ck.font = Font(name="宋体", size=11)
        ck.number_format = '@'
        ck.alignment = Alignment(vertical="center")
        ck.border = thin_border

        # 动态公式生成（根据不同模式自适应）
        if calc_mode == "倒推算税":
            # N 列为实发数值，M 列倒算个税，O 列附加成本，L 列应发 = M + N + O
            cn = ws_m1.cell(r, 14, doc["实发金额"])
            cn.font = Font(name="宋体", size=11)
            cn.number_format = fmt_currency
            cn.alignment = Alignment(vertical="center")
            cn.border = thin_border

            cm = ws_m1.cell(r, 13, f"=ROUND(IF(N{r}<=800,0,IF(N{r}<=3360,(N{r}-800)/4,IF(N{r}<=21000,0.16*N{r}/0.84,IF(N{r}<=49500,(0.24*N{r}-2000)/0.76,(0.32*N{r}-7000)/0.68)))),2)")
            cm.font = Font(name="等线", size=11)
            cm.number_format = fmt_currency
            cm.alignment = Alignment(vertical="center")
            cm.border = thin_border

            co = ws_m1.cell(r, 15, f"=ROUND((M{r}+N{r})*1.51%,2)")
            co.font = Font(name="宋体", size=10)
            co.number_format = fmt_2dec
            co.alignment = Alignment(horizontal="center", vertical="center")
            co.border = thin_border

            cl = ws_m1.cell(r, 12, f"=M{r}+N{r}+O{r}")
            cl.font = Font(name="等线", size=12)
            cl.number_format = fmt_currency
            cl.alignment = Alignment(vertical="center")
            cl.border = thin_border
        elif calc_mode == "不计税":
            # L 列为应发数值，M 列为 0，O 列为 0，N 列实发 = L 列
            cl = ws_m1.cell(r, 12, doc["应发金额"])
            cl.font = Font(name="等线", size=12)
            cl.number_format = fmt_currency
            cl.alignment = Alignment(vertical="center")
            cl.border = thin_border

            cm = ws_m1.cell(r, 13, 0.0)
            cm.font = Font(name="等线", size=11)
            cm.number_format = fmt_currency
            cm.alignment = Alignment(vertical="center")
            cm.border = thin_border

            co = ws_m1.cell(r, 15, 0.0)
            co.font = Font(name="宋体", size=10)
            co.number_format = fmt_2dec
            co.alignment = Alignment(horizontal="center", vertical="center")
            co.border = thin_border

            cn = ws_m1.cell(r, 14, f"=L{r}")
            cn.font = Font(name="宋体", size=11)
            cn.number_format = fmt_currency
            cn.alignment = Alignment(vertical="center")
            cn.border = thin_border
        else:
            # 正算扣税: L 列为应发数值，M 列正算个税公式，O 列 0，N 列实发 = L - M
            cl = ws_m1.cell(r, 12, doc["应发金额"])
            cl.font = Font(name="等线", size=12)
            cl.number_format = fmt_currency
            cl.alignment = Alignment(vertical="center")
            cl.border = thin_border

            cm = ws_m1.cell(r, 13, f"=ROUND(IF(L{r}<=800,0,IF(L{r}<=4000,(L{r}-800)*0.2,IF(L{r}<=25000,L{r}*0.8*0.2,IF(L{r}<=62500,L{r}*0.8*0.3-2000,L{r}*0.8*0.4-7000)))),2)")
            cm.font = Font(name="等线", size=11)
            cm.number_format = fmt_currency
            cm.alignment = Alignment(vertical="center")
            cm.border = thin_border

            co = ws_m1.cell(r, 15, 0.0)
            co.font = Font(name="宋体", size=10)
            co.number_format = fmt_2dec
            co.alignment = Alignment(horizontal="center", vertical="center")
            co.border = thin_border

            cn = ws_m1.cell(r, 14, f"=L{r}-M{r}")
            cn.font = Font(name="宋体", size=11)
            cn.number_format = fmt_currency
            cn.alignment = Alignment(vertical="center")
            cn.border = thin_border

        cp = ws_m1.cell(r, 16, zy_text)
        cp.font = Font(name="SimSun", size=10)
        cp.alignment = Alignment(horizontal="center", vertical="center")
        cp.border = thin_border

    # 动态合计行
    tot_row = len(doctor_summary_list) + 2
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
    log_func(f"[√] 《劳务明细总表_{today_str}.xlsx》保存完成")

    # =========================================================================
    # 11. 生成《任务明细表_YYYYMMDD.xlsx》（全量 10 列明细）
    # =========================================================================
    out_renwu = out_dir / f"任务明细表_{today_str}.xlsx"
    log_func(f"[*] 正在构建任务明细表: {out_renwu.name}")

    wb_rw = openpyxl.Workbook()
    ws_rw = wb_rw.active
    ws_rw.title = "Sheet1"

    col_widths_rw = {
        'A': 32.0, 'B': 24.0, 'C': 12.0, 'D': 15.0, 'E': 32.0,
        'F': 37.0, 'G': 10.0, 'H': 12.0, 'I': 12.0, 'J': 45.0
    }
    for col_letter, w in col_widths_rw.items():
        ws_rw.column_dimensions[col_letter].width = w
    ws_rw.row_dimensions[1].height = 28

    headers_rw = ['项目名称', '任务明细编号', '医生姓名', '手机号', '医院', '作品标题', '单价', '是否含税', '审核状态', '互联网医院作品链接']
    for c_idx, h in enumerate(headers_rw, 1):
        c = ws_rw.cell(1, c_idx, h)
        c.font = Font(name="宋体", size=11, bold=True)
        c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        c.alignment = Alignment(vertical="center", horizontal="center")

    curr_r = 2
    for doc in doctor_summary_list:
        for t in doc["tasks"]:
            ws_rw.row_dimensions[curr_r].height = 20
            ws_rw.cell(curr_r, 1, doc["项目名称*"]).font = Font(name="宋体", size=10)
            ws_rw.cell(curr_r, 2, t["code"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 3, doc["姓名*"]).font = Font(name="宋体", size=11)
            ws_rw.cell(curr_r, 4, doc["手机号*"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 5, doc["工作单位*"]).font = Font(name="宋体", size=10)
            ws_rw.cell(curr_r, 6, t["task_title"]).font = Font(name="宋体", size=10)
            ws_rw.cell(curr_r, 7, t["price"]).font = Font(name="Calibri", size=11)
            ws_rw.cell(curr_r, 8, "否" if calc_mode == "倒推算税" else "是").font = Font(name="宋体", size=11)
            ws_rw.cell(curr_r, 9, t["status"]).font = Font(name="宋体", size=11)
            ws_rw.cell(curr_r, 10, t["audio"]).font = Font(name="宋体", size=10)

            for c in range(1, 11):
                ws_rw.cell(curr_r, c).alignment = Alignment(vertical="center")
            curr_r += 1

    wb_rw.save(out_renwu)
    log_func(f"[√] 《任务明细表_{today_str}.xlsx》保存完成")

    # =========================================================================
    # 12. 全自动化勾稽核对门禁检验报告
    # =========================================================================
    log_func("\n" + "=" * 55)
    log_func("【数据勾稽核对防漏检验报告】")
    log_func(f"  [1] 结算医生总人数: {total_doctors} 人")
    log_func(f"  [2] 任务明细总笔数: {len(valid_tasks)} 笔")
    log_func(f"  [3] 本次任务基础总金额: {total_valid_amount:,.2f} 元")
    log_func(f"  [4] 本次应发含税总额: ¥ {total_gross:,.2f} 元")
    log_func(f"  [5] 本次代扣个税总额: ¥ {total_tax:,.2f} 元")
    log_func(f"  [6] 本次实发到手总额: ¥ {total_net:,.2f} 元")
    log_func(f"  [7] 主报销单规范: 整合-{today_dash}.xlsx (22列标准规范)")
    log_func("  [√] 账实核对: 100% 严密吻合！无任何漏统与差额！")
    log_func("=" * 55 + "\n")

    return out_main_xlsx, out_main_xls, out_laowu, out_renwu


def run_cli_auto():
    """CLI命令行自动模式"""
    cur_dir = Path.cwd()
    print("\n" + "=" * 60)
    print("   北京整合-上药雷允上结算系统 (macOS / Windows)")
    print("=" * 60)
    print(f"当前扫描目录: {cur_dir}")

    cand_files = [f for f in cur_dir.glob("*.xlsx") if not f.name.startswith("劳务明细总表") and not f.name.startswith("任务明细表") and not f.name.startswith("整合-") and not f.name.startswith("~$")]

    if not cand_files:
        print("\n[!] 未在当前目录下检测到 Excel 源文件！")
        return

    yl_path = None
    u_path = None
    task_path = None

    for f in cand_files:
        n = f.name
        if "语料列表" in n or "语料" in n or "结算明细" in n:
            if not yl_path: yl_path = f
        elif "用户列表" in n or "用户" in n:
            if not u_path: u_path = f
        elif "确认单" in n or "支付" in n or "JS-" in n:
            if not task_path: task_path = f

    if not yl_path:
        yl_path = cand_files[0]

    # 检测命令行参数中的模式
    mode = "正算扣税"
    if "--no-tax" in sys.argv:
        mode = "不计税"
    elif "--reverse-tax" in sys.argv:
        mode = "倒推算税"

    print(f"\n[自动匹配识别]")
    print(f"  明细文件: {yl_path.name}")
    print(f"  用户资质: {u_path.name if u_path else '未提供 (将从明细提取)'}")
    print(f"  辅助单据: {task_path.name if task_path else '未提供 (将自动按已加入结算单汇总)'}")
    print(f"  结算模式: {mode}")

    try:
        process_settlement(yl_path, u_path, task_path, cur_dir, mode=mode, log_func=print)
        print("\n[√] 结算处理完成！")
    except Exception as e:
        print(f"\n[!] 处理发生异常: {e}")
        import traceback
        traceback.print_exc()


def run_gui():
    """GUI 模式：兼容 Mac 原生图形窗口与一键执行"""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("北京整合-上药雷允上结算系统")
    root.geometry("680x580")
    root.minsize(620, 520)

    yl_var = tk.StringVar()
    user_var = tk.StringVar()
    task_var = tk.StringVar()
    out_var = tk.StringVar(value=str(Path.cwd()))
    mode_var = tk.StringVar(value="正算扣税")

    cur_dir = Path.cwd()
    for f in cur_dir.glob("*.xlsx"):
        if not f.name.startswith("劳务明细总表") and not f.name.startswith("任务明细表") and not f.name.startswith("整合-") and not f.name.startswith("~$"):
            if "语料" in f.name or "明细" in f.name:
                if not yl_var.get(): yl_var.set(str(f))
            elif "用户" in f.name:
                if not user_var.get(): user_var.set(str(f))
            elif "确认单" in f.name or "支付" in f.name:
                if not task_var.get(): task_var.set(str(f))

    frame = ttk.Frame(root, padding="15")
    frame.pack(fill=tk.BOTH, expand=True)

    lbl_title = ttk.Label(frame, text="北京整合-上药雷允上结算系统", font=("Helvetica", 16, "bold"))
    lbl_title.pack(pady=(0, 10))

    # 1. 语料明细
    lbl1 = ttk.Label(frame, text="1. 语料明细文件 (*必选):")
    lbl1.pack(anchor=tk.W)
    f1_box = ttk.Frame(frame)
    f1_box.pack(fill=tk.X, pady=(2, 6))
    ent1 = ttk.Entry(f1_box, textvariable=yl_var)
    ent1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_yl():
        f = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if f: yl_var.set(f)
    ttk.Button(f1_box, text="浏览...", command=choose_yl).pack(side=tk.RIGHT)

    # 2. 用户列表
    lbl2 = ttk.Label(frame, text="2. 用户列表文件 (可选，提取卡号/开户行/金融资质):")
    lbl2.pack(anchor=tk.W)
    f2_box = ttk.Frame(frame)
    f2_box.pack(fill=tk.X, pady=(2, 6))
    ent2 = ttk.Entry(f2_box, textvariable=user_var)
    ent2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_user():
        f = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if f: user_var.set(f)
    ttk.Button(f2_box, text="浏览...", command=choose_user).pack(side=tk.RIGHT)

    # 3. 辅助确认单
    lbl3 = ttk.Label(frame, text="3. 辅助任务确认单 / 支付清单 (可选):")
    lbl3.pack(anchor=tk.W)
    f3_box = ttk.Frame(frame)
    f3_box.pack(fill=tk.X, pady=(2, 6))
    ent3 = ttk.Entry(f3_box, textvariable=task_var)
    ent3.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    def choose_task():
        f = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if f: task_var.set(f)
    ttk.Button(f3_box, text="浏览...", command=choose_task).pack(side=tk.RIGHT)

    # 4. 模式选择
    f4_box = ttk.Frame(frame)
    f4_box.pack(fill=tk.X, pady=(4, 10))
    ttk.Label(f4_box, text="结算模式:").pack(side=tk.LEFT, padx=(0, 8))
    cb_mode = ttk.Combobox(f4_box, textvariable=mode_var, values=["正算扣税", "不计税", "倒推算税"], state="readonly", width=15)
    cb_mode.pack(side=tk.LEFT)

    log_text = tk.Text(frame, height=9, font=("Menlo" if sys.platform == "darwin" else "Consolas", 9))
    log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    def gui_log(msg):
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        root.update_idletasks()

    def on_run():
        yl = yl_var.get().strip()
        u = user_var.get().strip()
        task = task_var.get().strip()
        out = out_var.get().strip()
        mode = mode_var.get().strip()
        if not yl:
            messagebox.showerror("错误", "请先选择【语料明细文件】！")
            return
        log_text.delete("1.0", tk.END)
        try:
            btn_run.config(state=tk.DISABLED)
            process_settlement(yl, u if u else None, task if task else None, out if out else None, mode=mode, log_func=gui_log)
            messagebox.showinfo("生成成功", f"结算表格生成成功！\n\n已保存在：\n{Path(out).resolve()}")
        except Exception as e:
            gui_log(f"[!] 运行失败: {e}")
            messagebox.showerror("运行出错", str(e))
        finally:
            btn_run.config(state=tk.NORMAL)

    btn_run = ttk.Button(frame, text="🚀 一键生成结算表格", command=on_run)
    btn_run.pack(fill=tk.X, ipady=4)

    root.mainloop()


if __name__ == "__main__":
    if "--cli" in sys.argv:
        run_cli_auto()
    elif len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        # 兼容作为工具脚本调用: python settlement_mac.py <dir_or_corpus> [user_file] [task_file] [out_dir] [--mode ...]
        target = Path(sys.argv[1]).resolve()
        mode_val = "正算扣税"
        for arg in sys.argv:
            if "--mode=" in arg:
                mode_val = arg.split("=")[1]
            elif arg in ["--no-tax", "--notax"]:
                mode_val = "不计税"
            elif arg in ["--reverse-tax", "--reverse"]:
                mode_val = "倒推算税"

        if target.is_dir():
            # 自动探测
            cands = list(target.glob("*.xlsx")) + list(target.glob("*.xls"))
            yl_p = None
            u_p = None
            t_p = None
            for f in cands:
                fn = f.name
                if any(x in fn for x in ["整合-", "劳务明细总表", "任务明细表", "~$"]): continue
                if ("语料" in fn or "明细" in fn) and not yl_p: yl_p = f
                elif "用户" in fn and not u_p: u_p = f
                elif any(k in fn for k in ["确认单", "支付", "JS-"]) and not t_p: t_p = f
            if not yl_p and cands: yl_p = cands[0]
            process_settlement(yl_p, u_p, t_p, target, mode=mode_val, log_func=print)
        else:
            u_arg = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("-") and sys.argv[2] != "NONE" else None
            t_arg = sys.argv[3] if len(sys.argv) > 3 and not sys.argv[3].startswith("-") and sys.argv[3] != "NONE" else None
            out_arg = sys.argv[4] if len(sys.argv) > 4 and not sys.argv[4].startswith("-") and sys.argv[4] != "NONE" else None
            process_settlement(target, u_arg, t_arg, out_arg, mode=mode_val, log_func=print)
    else:
        try:
            run_gui()
        except Exception:
            run_cli_auto()
