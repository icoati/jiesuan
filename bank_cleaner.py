#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HCP银行开户行智能清洗与质检引擎 (Bank Information Cleaner & Auditor)
专为解决 HCP 手填开户行常见问题：
1. 重复字样清洗（如“中国建设银行建设银行...” -> “中国建设银行...”）
2. 缺少银行只有支行时，通过卡号前6位(BIN码)自动反查总行并智能补全
3. 缺少支行/营业室/信用社等网点层级的预警与质检
4. 结尾漏字修补（如“太原解放路支” -> “太原解放路支行”）
5. 卡号识别银行与填报银行冲突核验
"""

import os
import re
import sys
from typing import Dict, List, Tuple, Optional
import pandas as pd

# 适配 Windows 控制台输出 UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ==============================================================================
# 1. 全国主流银行卡 BIN 码（卡号前6位/前8位）精选映射库 (覆盖常见大行、股份制及城商行)
# ==============================================================================
BANK_BIN_MAP = {
    # 中国工商银行
    "622202": "中国工商银行", "622200": "中国工商银行", "622208": "中国工商银行",
    "621225": "中国工商银行", "621226": "中国工商银行", "621288": "中国工商银行",
    "621558": "中国工商银行", "621559": "中国工商银行", "620058": "中国工商银行",
    "621722": "中国工商银行", "621723": "中国工商银行", "955880": "中国工商银行",
    "955881": "中国工商银行", "955882": "中国工商银行", "955888": "中国工商银行",
    "45806": "中国工商银行", "53098": "中国工商银行",

    # 中国建设银行
    "622700": "中国建设银行", "621700": "中国建设银行", "622280": "中国建设银行",
    "622707": "中国建设银行", "621284": "中国建设银行", "621598": "中国建设银行",
    "622166": "中国建设银行", "622168": "中国建设银行", "436742": "中国建设银行",
    "436745": "中国建设银行", "436748": "中国建设银行", "453242": "中国建设银行",
    "552801": "中国建设银行", "552802": "中国建设银行", "622725": "中国建设银行",

    # 中国农业银行
    "622848": "中国农业银行", "622845": "中国农业银行", "622846": "中国农业银行",
    "622849": "中国农业银行", "622844": "中国农业银行", "622847": "中国农业银行",
    "622820": "中国农业银行", "622821": "中国农业银行", "622822": "中国农业银行",
    "622823": "中国农业银行", "622824": "中国农业银行", "622825": "中国农业银行",
    "622826": "中国农业银行", "622827": "中国农业银行", "622836": "中国农业银行",
    "622837": "中国农业银行", "622838": "中国农业银行", "621682": "中国农业银行",
    "621683": "中国农业银行", "95599": "中国农业银行",

    # 中国银行
    "621661": "中国银行", "621660": "中国银行", "621662": "中国银行",
    "621663": "中国银行", "621667": "中国银行", "621668": "中国银行",
    "621669": "中国银行", "601382": "中国银行", "409666": "中国银行",
    "409667": "中国银行", "409668": "中国银行", "409669": "中国银行",
    "456351": "中国银行", "456352": "中国银行", "456353": "中国银行",

    # 交通银行
    "622260": "交通银行", "622261": "交通银行", "622262": "交通银行",
    "601428": "交通银行", "520169": "交通银行", "521899": "交通银行",
    "458123": "交通银行", "458124": "交通银行", "621245": "交通银行",

    # 中国邮政储蓄银行
    "621098": "中国邮政储蓄银行", "622188": "中国邮政储蓄银行", "621096": "中国邮政储蓄银行",
    "621095": "中国邮政储蓄银行", "621799": "中国邮政储蓄银行", "621798": "中国邮政储蓄银行",
    "621577": "中国邮政储蓄银行", "621599": "中国邮政储蓄银行", "601426": "中国邮政储蓄银行",

    # 招商银行
    "621483": "招商银行", "621485": "招商银行", "621486": "招商银行",
    "622588": "招商银行", "622575": "招商银行", "622576": "招商银行",
    "439225": "招商银行", "439226": "招商银行", "439227": "招商银行",
    "621286": "招商银行", "621792": "招商银行",

    # 浦发银行 (上海浦东发展银行)
    "622521": "浦发银行", "622522": "浦发银行", "622523": "浦发银行",
    "621789": "浦发银行", "621790": "浦发银行", "621370": "浦发银行",
    "984301": "浦发银行", "984302": "浦发银行",

    # 中信银行
    "622690": "中信银行", "622691": "中信银行", "622692": "中信银行",
    "622696": "中信银行", "622698": "中信银行", "433670": "中信银行",
    "621768": "中信银行", "621769": "中信银行",

    # 中国光大银行
    "622660": "中国光大银行", "622661": "中国光大银行", "622662": "中国光大银行",
    "622663": "中国光大银行", "622664": "中国光大银行", "622665": "中国光大银行",
    "621796": "中国光大银行", "621797": "中国光大银行",

    # 华夏银行
    "622630": "华夏银行", "622632": "华夏银行", "622633": "华夏银行",
    "621222": "华夏银行", "621282": "华夏银行", "623020": "华夏银行",

    # 中国民生银行
    "622622": "中国民生银行", "622615": "中国民生银行", "622617": "中国民生银行",
    "622619": "中国民生银行", "621691": "中国民生银行", "621692": "中国民生银行",
    "421317": "中国民生银行", "421865": "中国民生银行",

    # 广发银行
    "622556": "广发银行", "622558": "广发银行", "622559": "广发银行",
    "621462": "广发银行", "621463": "广发银行", "685800": "广发银行",

    # 平安银行 (含原深发展)
    "622155": "平安银行", "622156": "平安银行", "622298": "平安银行",
    "621626": "平安银行", "621627": "平安银行", "998801": "平安银行",

    # 兴业银行
    "622908": "兴业银行", "622909": "兴业银行", "622901": "兴业银行",
    "621297": "兴业银行", "621298": "兴业银行", "437466": "兴业银行",

    # 浙商银行
    "622365": "浙商银行", "622366": "浙商银行", "621466": "浙商银行",

    # 渤海银行
    "622877": "渤海银行", "622878": "渤海银行", "621453": "渤海银行",

    # 恒丰银行
    "622381": "恒丰银行", "622382": "恒丰银行", "621430": "恒丰银行",

    # 常见主要城商行/农商行
    "622812": "北京银行", "621468": "北京银行", "601428": "北京银行",
    "622268": "上海银行", "621293": "上海银行",
    "622516": "江苏银行", "621295": "江苏银行",
    "622510": "南京银行", "621296": "南京银行",
    "622284": "宁波银行", "621299": "宁波银行",
    "622853": "杭州银行", "621481": "杭州银行",
    "622580": "广州银行", "621477": "广州银行",
    "622384": "成都银行", "621479": "成都银行",
    "622928": "重庆银行", "621488": "重庆银行",
    "622394": "长沙银行", "621490": "长沙银行",
    "622881": "山西省农村信用社", "621471": "山西省农村信用社",
    "622880": "晋商银行", "621470": "晋商银行"
}

# 常见核心银行标准名称与其别名对应
KNOWN_BANK_KEYWORDS = [
    ("中国工商银行", ["工商银行", "工行"]),
    ("中国建设银行", ["建设银行", "建行"]),
    ("中国农业银行", ["农业银行", "农行"]),
    ("中国银行", ["中行"]),
    ("交通银行", ["交行"]),
    ("中国邮政储蓄银行", ["邮政储蓄银行", "邮储银行", "邮政储蓄", "邮政银行", "邮储"]),
    ("招商银行", ["招行"]),
    ("上海浦东发展银行", ["浦东发展银行", "浦发银行", "浦发"]),
    ("中信银行", ["中信"]),
    ("中国光大银行", ["光大银行", "光大"]),
    ("华夏银行", ["华夏"]),
    ("中国民生银行", ["民生银行", "民生"]),
    ("广发银行", ["广东发展银行"]),
    ("平安银行", ["平安"]),
    ("兴业银行", ["兴业"]),
    ("浙商银行", ["浙商"]),
    ("渤海银行", ["渤海"]),
    ("恒丰银行", ["恒丰"]),
    ("北京银行", []),
    ("上海银行", []),
    ("江苏银行", []),
    ("南京银行", []),
    ("宁波银行", []),
    ("杭州银行", []),
    ("广州银行", []),
    ("成都银行", []),
    ("重庆银行", []),
    ("长沙银行", []),
    ("晋商银行", []),
    ("农村信用社", ["农信社", "信用社", "农村合作银行", "农商行", "农村商业银行"])
]

# 合法网点层级关键词
BRANCH_KEYWORDS = [
    "支行", "分行", "营业部", "营业室", "分理处",
    "信用社", "农村信用社", "合作社", "储蓄所", "村镇银行"
]


# ==============================================================================
# 2. 核心清洗与质检逻辑
# ==============================================================================

def identify_bank_from_card(card_no: str) -> Optional[str]:
    """根据银行卡号前 6 位 (BIN 码) 智能识别所属银行"""
    if not card_no:
        return None
    # 过滤卡号中的空格与非数字
    clean_card = re.sub(r'\D', '', str(card_no))
    if len(clean_card) < 6:
        return None
    
    # 优先匹配 6 位
    prefix6 = clean_card[:6]
    if prefix6 in BANK_BIN_MAP:
        return BANK_BIN_MAP[prefix6]
        
    # 其次匹配 5 位
    prefix5 = clean_card[:5]
    if prefix5 in BANK_BIN_MAP:
        return BANK_BIN_MAP[prefix5]
        
    return None


def remove_redundant_bank_names(text: str) -> Tuple[str, bool]:
    """
    智能消除开户行中因拼接导致的重复银行名称
    例如：
    - 中国建设银行建设银行太原住房支行 -> 中国建设银行太原住房支行
    - 中国民生银行民生银行太原市体育南路支行 -> 中国民生银行太原市体育南路支行
    - 中国农业银行农行太原迎泽支行 -> 中国农业银行太原迎泽支行
    - 招商银行招商银行体育西路支行 -> 招商银行体育西路支行
    返回: (清洗后文本, 是否发生了去重修复)
    """
    if not text:
        return "", False
        
    original = text.strip()
    res = original
    changed = False

    # 1. 处理“中国XX银行 + XX银行”模式（最典型高频的重复）
    # 例：中国建设银行建设银行 -> 中国建设银行
    pattern_zg = r'(中国[^\s]+?银行)\s*(?:中国)?([^\s]+?银行)'
    m = re.search(pattern_zg, res)
    if m:
        p1, p2 = m.group(1), m.group(2)
        # 如果 p2 是 p1 的子串（如建设银行是中国建设银行的子串），或者两者相同
        if p2 in p1 or p1 in p2:
            res = res[:m.start()] + p1 + res[m.end():]
            changed = True

    # 2. 处理紧邻完全重复的银行词（例：民生银行民生银行 / 建设银行建设银行）
    pattern_dup = r'([^\s]+?银行)\s*\1'
    if re.search(pattern_dup, res):
        res = re.sub(pattern_dup, r'\1', res)
        changed = True

    # 3. 处理“全称 + 常见简称”（例：中国农业银行农行 -> 中国农业银行）
    short_map = [
        ("中国农业银行", "农行"),
        ("中国工商银行", "工行"),
        ("中国建设银行", "建行"),
        ("中国银行", "中行"),
        ("交通银行", "交行"),
        ("招商银行", "招行"),
        ("中国邮政储蓄银行", "邮储"),
        ("中国民生银行", "民生银行"),
        ("中国光大银行", "光大银行"),
        ("上海浦东发展银行", "浦发银行"),
    ]
    for full_n, short_n in short_map:
        dup_pat = f"{full_n}\\s*{short_n}"
        if re.search(dup_pat, res):
            res = re.sub(dup_pat, full_n, res)
            changed = True

    return res.strip(), changed


def clean_single_bank_record(
    raw_bank: str,
    card_no: Optional[str] = None,
    doctor_name: Optional[str] = None
) -> Dict[str, str]:
    """
    清洗并质检单条开户行记录
    返回字典结构：
    {
        "raw_bank": 原始填写,
        "cleaned_bank": 清洗后开户行,
        "status": "✅ 正常" | "⚠️ 需核实" | "❌ 异常",
        "detail": 质检说明,
        "is_repaired": 是否自动修复了内容,
        "repair_note": 修复说明
    }
    """
    raw_str = str(raw_bank).strip() if pd.notna(raw_bank) else ""
    if raw_str in ["nan", "None", "null", "NULL"]:
        raw_str = ""

    # 初始化结果
    result = {
        "raw_bank": raw_str,
        "cleaned_bank": raw_str,
        "status": "✅ 正常",
        "detail": "格式合规",
        "is_repaired": False,
        "repair_note": ""
    }

    if not raw_str:
        result["status"] = "❌ 缺失"
        result["detail"] = "开户行信息完全为空"
        return result

    curr = raw_str

    # 1. 过滤开户行文本中常见的无意义前导文本（如“开户行：”、“支行名称:”、“银行:”）
    prefix_clean = re.sub(r'^(?:开户行|开户银行|银行|网点|支行|开户网点)[：:\s]+', '', curr)
    if prefix_clean != curr:
        curr = prefix_clean
        result["is_repaired"] = True
        result["repair_note"] += "去除引导前缀; "

    # 2. 检查支行是否误填为了医生本人姓名
    if doctor_name and str(doctor_name).strip() and curr == str(doctor_name).strip():
        result["status"] = "❌ 误填姓名"
        result["detail"] = f"误填为了医生本人姓名【{doctor_name}】"
        result["cleaned_bank"] = ""
        return result

    # 3. 执行银行重复名称智能剔除
    cleaned_dup, has_dedup = remove_redundant_bank_names(curr)
    if has_dedup:
        curr = cleaned_dup
        result["is_repaired"] = True
        result["repair_note"] += "剔除重复银行名称; "

    # 4. 修补末尾漏字（如“太原迎泽支” -> “太原迎泽支行”）
    if curr.endswith("支") and not curr.endswith("分支") and not curr.endswith("支行"):
        curr += "行"
        result["is_repaired"] = True
        result["repair_note"] += "补齐末尾缺失的【行】字; "
    elif curr.endswith("分") and not curr.endswith("分行") and len(curr) >= 4:
        curr += "行"
        result["is_repaired"] = True
        result["repair_note"] += "补齐末尾缺失的【行】字; "

    # 5. 检查是否包含银行总行信息
    # 提取卡号对应银行（如有）
    inferred_bank = identify_bank_from_card(card_no) if card_no else None

    # 判断当前文本中是否有银行名称
    has_bank_name = False
    found_main_bank = ""
    for full_b, aliases in KNOWN_BANK_KEYWORDS:
        if full_b in curr:
            has_bank_name = True
            found_main_bank = full_b
            break
        for a in aliases:
            if a in curr:
                has_bank_name = True
                found_main_bank = full_b
                break
        if has_bank_name:
            break

    if not has_bank_name:
        # 宽松检查是否包含“银行”或“信用社”
        if "银行" in curr or "信用社" in curr or "合作社" in curr:
            has_bank_name = True

    # 6. 处理缺少银行总行（只有支行）的情况
    if not has_bank_name:
        if inferred_bank:
            # 奇迹补全：通过卡号 BIN 码识别出总行，自动拼接到支行前面！
            curr = f"{inferred_bank}{curr}"
            result["is_repaired"] = True
            result["repair_note"] += f"依据卡号BIN码自动补齐总行【{inferred_bank}】; "
            has_bank_name = True
            found_main_bank = inferred_bank
        else:
            result["status"] = "⚠️ 缺少银行"
            result["detail"] = "缺少银行总行名称(仅填写了支行/网点)，建议提供卡号或人工补全"

    # 7. 检查是否包含支行/分行等网点层级
    has_branch_level = any(bk in curr for bk in BRANCH_KEYWORDS)
    if not has_branch_level:
        if result["status"] == "✅ 正常":
            result["status"] = "⚠️ 缺少网点"
            result["detail"] = "缺少支行/分行/营业室等具体网点层级，转账极易退票"

    # 8. 卡号与填写银行冲突核验
    if inferred_bank and found_main_bank:
        # 核对两者是否一致
        # 如 inferred_bank 为“中国工商银行”，found_main_bank 为“中国建设银行”
        if (inferred_bank not in found_main_bank) and (found_main_bank not in inferred_bank):
            # 特殊别名判断
            is_mismatch = True
            for full_b, aliases in KNOWN_BANK_KEYWORDS:
                if (full_b == inferred_bank and found_main_bank in aliases) or \
                   (full_b == found_main_bank and inferred_bank in aliases):
                    is_mismatch = False
                    break
            if is_mismatch:
                result["status"] = "⚠️ 卡号与银行冲突"
                result["detail"] = f"卡号前缀识别为【{inferred_bank}】，但填写的为【{found_main_bank}】"

    # 9. 内容长度异常检测（小于 4 个字通常不合规）
    if len(curr) < 4:
        result["status"] = "❌ 内容过短"
        result["detail"] = "开户行名称过短，无法正常汇款"

    result["cleaned_bank"] = curr
    if result["status"] == "✅ 正常" and result["is_repaired"]:
        result["detail"] = "已自动修复并格式达标"

    return result


def clean_dataframe_banks(
    df: pd.DataFrame,
    bank_col: str,
    card_col: Optional[str] = None,
    name_col: Optional[str] = None
) -> Tuple[pd.DataFrame, Dict[str, int]]:
    """
    批量清洗与质检 DataFrame 中的开户行数据
    在 DataFrame 末尾新增 3 列：
    - 【清洗后】开户行规范全称
    - 【开户行质检状态】
    - 【质检核验说明】
    返回: (处理后的 DataFrame, 统计指标字典)
    """
    df_out = df.copy()

    cleaned_banks = []
    statuses = []
    details = []
    
    stats = {
        "total": len(df_out),
        "normal": 0,
        "repaired": 0,
        "warning": 0,
        "error": 0
    }

    for idx, row in df_out.iterrows():
        raw_b = row.get(bank_col, "")
        card_v = row.get(card_col, "") if card_col else None
        name_v = row.get(name_col, "") if name_col else None

        res = clean_single_bank_record(raw_b, card_v, name_v)

        cleaned_banks.append(res["cleaned_bank"])
        statuses.append(res["status"])
        
        # 拼接详情与修复说明
        full_detail = res["detail"]
        if res["repair_note"]:
            full_detail += f" ({res['repair_note'].rstrip('; ')})"
        details.append(full_detail)

        # 统计计数
        if res["is_repaired"]:
            stats["repaired"] += 1

        if res["status"].startswith("✅"):
            stats["normal"] += 1
        elif res["status"].startswith("⚠️"):
            stats["warning"] += 1
        else:
            stats["error"] += 1

    df_out["【清洗后】开户行规范全称"] = cleaned_banks
    df_out["【开户行质检状态】"] = statuses
    df_out["【质检核验说明】"] = details

    return df_out, stats


def export_styled_excel(df: pd.DataFrame, output_path: str):
    """
    导出高保真带样式 Excel：
    1. 冻结首行、深色专业表头
    2. 自动标记异常行（浅黄/浅红高亮）
    3. 卡号、身份证、手机号强制文本格式，杜绝科学计数法
    4. 自适应列宽
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "开户行清洗质检明细"

    # 样式定义
    font_header = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # 深海蓝

    font_data = Font(name="微软雅黑", size=10)
    font_status_ok = Font(name="微软雅黑", size=10, bold=True, color="15803D")
    font_status_warn = Font(name="微软雅黑", size=10, bold=True, color="B45309")
    font_status_err = Font(name="微软雅黑", size=10, bold=True, color="B91C1C")

    fill_warn_row = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid") # 浅黄色
    fill_err_row = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # 浅红色

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    # 写入表头
    headers = list(df.columns)
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    status_col_idx = headers.index("【开户行质检状态】") + 1 if "【开户行质检状态】" in headers else None
    clean_col_idx = headers.index("【清洗后】开户行规范全称") + 1 if "【清洗后】开户行规范全称" in headers else None

    # 写入数据行
    for r_idx, row in df.iterrows():
        excel_row_num = r_idx + 2
        status_val = str(row.get("【开户行质检状态】", ""))

        for c_idx, col_name in enumerate(headers, start=1):
            val = row[col_name]
            if pd.isna(val):
                val = ""
            
            # 卡号、手机号、身份证号防科学计数法
            col_str = str(col_name)
            is_sensitive_num = any(k in col_str for k in ["卡号", "账号", "手机", "电话", "身份证", "编码", "编号"])
            if is_sensitive_num and val != "":
                val = str(val).strip()
                if val.endswith(".0") and len(val) > 2 and val[:-2].isdigit():
                    val = val[:-2]

            cell = ws.cell(row=excel_row_num, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", horizontal="left")

            if is_sensitive_num:
                cell.number_format = "@"

            # 异常高亮
            if "⚠️" in status_val:
                cell.fill = fill_warn_row
            elif "❌" in status_val:
                cell.fill = fill_err_row

        # 单独针对质检状态列着色
        if status_col_idx:
            scell = ws.cell(row=excel_row_num, column=status_col_idx)
            scell.alignment = Alignment(horizontal="center", vertical="center")
            if "✅" in status_val:
                scell.font = font_status_ok
            elif "⚠️" in status_val:
                scell.font = font_status_warn
            else:
                scell.font = font_status_err

        # 清洗后开户行加粗突出显示
        if clean_col_idx and "✅" in status_val:
            ccell = ws.cell(row=excel_row_num, column=clean_col_idx)
            ccell.font = Font(name="微软雅黑", size=10, bold=True, color="0F172A")

        ws.row_dimensions[excel_row_num].height = 22

    # 冻结第一行
    ws.freeze_panes = "A2"

    # 自适应列宽
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            length = sum(2 if ord(c) > 127 else 1 for c in val_str)
            if length > max_len:
                max_len = length
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    wb.save(output_path)
    return output_path


# ==============================================================================
# 3. 命令行直接执行入口 (CLI)
# ==============================================================================
if __name__ == "__main__":
    # 快速自测
    test_cases = [
        ("中国建设银行建设银行太原住房支行", "6227001234567890", "李医生"),
        ("中国民生银行民生银行太原市体育南路支行", None, "王医生"),
        ("中国农业银行农行太原迎泽支", "6228481234567890", "张医生"),
        ("太原住房支行", "6227009876543210", "赵医生"), # 只有支行，靠卡号补全建行
        ("体育南路支行", None, "孙医生"), # 只有支行且无卡号
        ("中国工商银行", "6222021234567890", "周医生"), # 缺少支行
        ("中国工商银行太原分行营业部", "6227001234567890", "吴医生"), # 卡号是建行，填的是工行（冲突）
        ("太原市南郊区农村信用社", None, "郑医生"), # 信用社正常
        ("郑医生", "6227001234567890", "郑医生"), # 误填医生姓名
    ]

    print("=" * 60)
    print("HCP 开户行清洗与质检引擎 - 单元测试")
    print("=" * 60)
    for raw, card, name in test_cases:
        res = clean_single_bank_record(raw, card, name)
        print(f"原始填写: {raw}")
        print(f" -> 清洗后: {res['cleaned_bank']}")
        print(f" -> 状态: {res['status']} | 详情: {res['detail']}")
        if res['repair_note']:
            print(f" -> 修复点: {res['repair_note']}")
        print("-" * 50)
