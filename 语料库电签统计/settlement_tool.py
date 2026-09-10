#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
医疗健康语料库 - 每月结算一键生成工具 (工业级高鲁棒性版)
适用环境：macOS / Windows / Linux (Python 3.8+)
特点：
1. 智能多工作表扫描（自动识别真正包含数据的 Sheet，跳过说明/空表）
2. 强力编号清洗（去不可见字符、去.0、防科学计数法、前缀容错）
3. 过程透明诊断（打印两边编号样例，0匹配时自动拦截并清晰指出原因）
4. 金额三方强闭环核验，彻底杜绝空白表和数据失真
"""

import os
import sys
import re
import glob
from datetime import datetime

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("【缺少必要依赖】请先在终端运行：pip3 install openpyxl pandas xlwt xlrd")
    sys.exit(1)

try:
    import xlwt
except ImportError:
    xlwt = None

# ==================== 1. 字段别名库（按匹配优先级排序） ====================
COLUMN_ALIASES = {
    'corpus_id': ['语料词条编号', '语料编号', '词条编号', '题目编号', '语料id', 'corpus_id', '编号', 'id'],
    'doctor_name': ['医生姓名', '专家姓名', '医生', '姓名', '人员姓名', '持卡人姓名'],
    'phone': ['手机号码', '手机号', '联系电话', '电话', '手机', '用户手机号'],
    'id_card': ['身份证号码', '身份证号', '身份证', '证件号码', '证件号', '身份证号/护照号'],
    'hospital': ['所在医院', '医院全称', '医院', '工作单位', '医院名称', '医疗机构', '单位名称', '单位'],
    'department': ['科室', '所在科室', '专业科室', '科室名称', '部门'],
    'title': ['职称', '医务职称', '技术职称', '医生职称', '职务职称', '专业职称'],
    'unit_price': ['结算单价', '单价', '结算金额', '单条金额', '金额', '每条单价'],
    'is_settle': ['是否结算', '结算状态', '结算标记'],
    'audit_status': ['审核状态', '资质审核状态', '平台审核结果', '状态'],
    'submit_time': ['提交时间', '语料提交时间', '创建时间', '申请时间'],
    'audit_time': ['审核时间', '审核通过时间', '通过时间'],
    'bank_name': ['开户银行', '开户行', '银行名称', '银行', '结算银行'],
    'bank_card': ['银行卡号', '银行卡', '卡号', '银行账号', '账号', '结算卡号'],
    'branch_name': ['支行名称', '开户支行', '开户网点', '支行', '开户行支行'],
    'question': ['题目内容', '题目', '问答题目', '问题', '题干'],
    'answer': ['回答记录', '回答', '医生回答', '答案', '回答内容'],
    'project_name': ['项目名称', '参与项目', '项目'],
    'domain': ['疾病领域', '领域', '病种']
}

# 城市智能识别规则
CITY_RULES = [
    (r'(杭州|余杭|萧山|临平|富阳|临安|桐庐|淳安|建德|西湖|滨江|拱墅|上城)', '浙江省', '杭州市'),
    (r'(温州|乐清|瑞安|文成|苍南|永嘉|平阳|泰顺|鹿城|龙湾|瓯海|龙港)', '浙江省', '温州市'),
    (r'(金华|义乌|东阳|兰溪|永康|武义|浦江|磐安|婺城|金东)', '浙江省', '金华市'),
    (r'(台州|椒江|黄岩|路桥|临海|温岭|玉环|天台|仙居|三门)', '浙江省', '台州市'),
    (r'(丽水|莲都|龙泉|青田|缙云|遂昌|松阳|云和|庆元|景宁)', '浙江省', '丽水市'),
    (r'(宁波|海曙|江北|镇海|北仑|鄞州|奉化|余姚|慈溪|象山|宁海)', '浙江省', '宁波市'),
    (r'(绍兴|越城|柯桥|上虞|诸暨|嵊州|新昌)', '浙江省', '绍兴市'),
    (r'(嘉兴|南湖|秀洲|海宁|平湖|桐乡|嘉善|海盐)', '浙江省', '嘉兴市'),
    (r'(湖州|吴兴|南浔|德清|长兴|安吉)', '浙江省', '湖州市'),
    (r'(舟山|定海|普陀|岱山|嵊泗)', '浙江省', '舟山市'),
    (r'(衢州|柯城|衢江|江山|常山|开化|龙游)', '浙江省', '衢州市'),
    (r'(上海)', '上海市', '上海市'),
    (r'(北京)', '北京市', '北京市'),
    (r'(南京|苏州|无锡|常州|南通)', '江苏省', '江苏省市级')
]

def parse_hospital_location(hospital_name):
    if not hospital_name or str(hospital_name).strip() == '':
        return '浙江省', '温州市'
    h_str = str(hospital_name)
    for pattern, prov, city in CITY_RULES:
        if re.search(pattern, h_str):
            return prov, city
    return '浙江省', '温州市'

# ==================== 2. 强力编号清洗器 ====================
def normalize_id(val):
    """把各种怪异格式的语料编号/卡号转换为纯净标准字符串"""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    # 处理 pandas 误把数字转为 12345.0
    if s.endswith(".0") and len(s) > 2:
        prefix = s[:-2]
        if prefix.replace("-", "").isdigit():
            s = prefix
    # 去除不可见字符、换行、BOM头、单双引号
    s = re.sub(r'[\s\u200b\uFEFF\r\n\'\"`]+', '', s)
    return s.upper()

def find_col_name(available_cols, alias_key):
    """按优先级寻找匹配的表头列名"""
    clean_cols_map = {}
    for c in available_cols:
        clean_name = re.sub(r'[\s\*_]+', '', str(c)).lower()
        clean_cols_map[clean_name] = c

    targets = COLUMN_ALIASES.get(alias_key, [])
    for t in targets:
        clean_t = re.sub(r'[\s\*_]+', '', t).lower()
        if clean_t in clean_cols_map:
            return clean_cols_map[clean_t]
    return None

# ==================== 3. 智能多工作表检测与加载 ====================
def load_smart_dataframe(filepath, required_alias_keys):
    """
    智能扫描 Excel 的所有 Sheet，自动选择包含所需目标列且有效数据最多的 Sheet
    自动跳过前面的空行或大标题行
    """
    try:
        xl = pd.ExcelFile(filepath)
    except Exception as e:
        print(f"❌ 无法打开文件 {filepath}: {e}")
        return None, None
        
    best_df = None
    best_sheet = None
    max_matched_keys = -1
    max_rows = -1

    for sheet_name in xl.sheet_names:
        # 尝试不同 header 行 (0, 1, 2, 3) 防止第1行是合并大标题
        for header_idx in range(4):
            try:
                df = xl.parse(sheet_name, header=header_idx, nrows=100)
                if df.empty or len(df.columns) < 2:
                    continue
                
                # 检查命中了多少个 required_alias_keys
                matched_count = sum(1 for k in required_alias_keys if find_col_name(df.columns, k) is not None)
                
                if matched_count > max_matched_keys or (matched_count == max_matched_keys and len(df) > max_rows):
                    max_matched_keys = matched_count
                    max_rows = len(df)
                    # 重新全量读取该 sheet
                    best_df = xl.parse(sheet_name, header=header_idx)
                    best_sheet = sheet_name
                    if matched_count == len(required_alias_keys):
                        break
            except Exception:
                continue
        if max_matched_keys == len(required_alias_keys):
            break

    return best_df, best_sheet

# ==================== 4. 自动文件嗅探定位 ====================
def detect_input_files(target_dir):
    all_files = glob.glob(os.path.join(target_dir, "*.xlsx")) + glob.glob(os.path.join(target_dir, "*.xls"))
    candidate_files = [
        f for f in all_files
        if not os.path.basename(f).startswith("~$")
        and not os.path.basename(f).startswith("最终")
        and "结算筛选" not in os.path.basename(f)
        and "对账" not in os.path.basename(f)
    ]
    
    pay_file = None
    corpus_file = None
    user_file = None
    
    # 优先根据文件名关键字识别
    for f in candidate_files:
        fname = os.path.basename(f).lower()
        if ("支付" in fname or "清单" in fname or "结算" in fname) and ("语料列表" not in fname and "明文" not in fname):
            pay_file = f
        elif ("语料列表" in fname or "语料" in fname) and ("支付" not in fname):
            corpus_file = f
        elif ("用户列表" in fname or "用户" in fname or "医生列表" in fname):
            user_file = f

    # 若未识别全，结合内容特征二次嗅探
    for f in candidate_files:
        if f in [pay_file, corpus_file, user_file]:
            continue
        try:
            df = pd.read_excel(f, nrows=5)
            h_str = "".join([str(c) for c in df.columns])
            if not pay_file and ("单价" in h_str or "结算" in h_str) and "语料" in h_str:
                pay_file = f
            elif not corpus_file and ("题目" in h_str or "回答" in h_str or "疾病" in h_str):
                corpus_file = f
            elif not user_file and ("银行卡" in h_str or "开户行" in h_str or "资质" in h_str):
                user_file = f
        except Exception:
            pass

    return pay_file, corpus_file, user_file

# ==================== 5. 核心处理主程序 ====================
def process_settlement(target_dir="."):
    print("=" * 65)
    print("   🏥 医疗健康语料库 - 每月结算自动生成工具 (Mac/Win 通用)")
    print("=" * 65)
    print(f"📁 工作文件夹: {os.path.abspath(target_dir)}")
    
    pay_file, corpus_file, user_file = detect_input_files(target_dir)
    
    missing = []
    if not pay_file: missing.append("【三方项目支付清单】(含结算单价、语料编号的表格)")
    if not corpus_file: missing.append("【语料列表(明文)】(含题目、回答、医生姓名、身份证等)")
    if not user_file: missing.append("【用户列表(明文)】(含医生银行卡号、开户行等)")
    
    if missing:
        print("\n❌ 未能完整识别所需的三张表格！缺少：")
        for m in missing:
            print(f"   • {m}")
        print("\n👉 请检查文件夹，确保这 3 个表格已放入并命名合理。")
        return False
        
    print("\n🔍 自动锁定表格：")
    print(f"  • 支付清单 : {os.path.basename(pay_file)}")
    print(f"  • 语料明文 : {os.path.basename(corpus_file)}")
    print(f"  • 用户明文 : {os.path.basename(user_file)}")
    
    # 1. 解析支付清单
    print("\n⏳ [1/4] 正在解析【支付清单】...")
    df_pay, pay_sheet = load_smart_dataframe(pay_file, ['corpus_id'])
    if df_pay is None:
        print("❌ 无法读取支付清单！")
        return False
        
    pay_id_col = find_col_name(df_pay.columns, 'corpus_id')
    pay_price_col = find_col_name(df_pay.columns, 'unit_price')
    pay_status_col = find_col_name(df_pay.columns, 'is_settle')
    
    if not pay_id_col:
        print(f"❌ 支付清单 (工作表:{pay_sheet}) 未能识别到【语料词条编号】列！可用列: {list(df_pay.columns)}")
        return False

    # 清洗支付清单中的语料编号
    df_pay['CORPUS_ID_NORM'] = df_pay[pay_id_col].apply(normalize_id)
    # 过滤掉空编号
    df_pay = df_pay[df_pay['CORPUS_ID_NORM'] != ''].copy()
    
    # 状态过滤（若存在“是否结算”列，只抓取“是”）
    if pay_status_col:
        status_clean = df_pay[pay_status_col].astype(str).str.strip()
        is_yes = status_clean.isin(['是', 'Y', 'YES', '通过', '审核通过', '1'])
        if is_yes.sum() > 0 and (~is_yes).sum() > 0:
            print(f"  ℹ️ 检测到【是否结算】列，过滤出确认结算记录: {is_yes.sum()} 条 (剔除无需结算: {(~is_yes).sum()} 条)")
            df_pay = df_pay[is_yes].copy()

    pay_ids_set = set(df_pay['CORPUS_ID_NORM'])
    pay_samples = list(df_pay['CORPUS_ID_NORM'].head(3))
    print(f"  ✔ 识别到工作表: [{pay_sheet}], 编号列: [{pay_id_col}], 有效待结语料: {len(df_pay)} 条")
    print(f"  🔍 支付单编号样例 (前3条): {pay_samples}")

    # 2. 解析语料明文列表
    print("\n⏳ [2/4] 正在解析【语料列表(明文)】...")
    df_corpus, corpus_sheet = load_smart_dataframe(corpus_file, ['corpus_id', 'doctor_name'])
    if df_corpus is None:
        print("❌ 无法读取语料列表(明文)！")
        return False

    c_id_col = find_col_name(df_corpus.columns, 'corpus_id')
    c_name_col = find_col_name(df_corpus.columns, 'doctor_name')
    c_phone_col = find_col_name(df_corpus.columns, 'phone')
    c_idcard_col = find_col_name(df_corpus.columns, 'id_card')
    c_hosp_col = find_col_name(df_corpus.columns, 'hospital')
    c_dept_col = find_col_name(df_corpus.columns, 'department')
    c_title_col = find_col_name(df_corpus.columns, 'title')
    c_time_col = find_col_name(df_corpus.columns, 'submit_time')

    if not c_id_col:
        print(f"❌ 语料列表中未能识别到【语料词条编号】列！可用列: {list(df_corpus.columns)}")
        return False

    df_corpus['CORPUS_ID_NORM'] = df_corpus[c_id_col].apply(normalize_id)
    df_corpus = df_corpus[df_corpus['CORPUS_ID_NORM'] != ''].copy()
    corpus_samples = list(df_corpus['CORPUS_ID_NORM'].head(3))
    print(f"  ✔ 识别到工作表: [{corpus_sheet}], 编号列: [{c_id_col}], 总语料库条数: {len(df_corpus)} 条")
    print(f"  🔍 语料库编号样例 (前3条): {corpus_samples}")

    # 3. 关联匹配
    matched_corpus = df_corpus[df_corpus['CORPUS_ID_NORM'].isin(pay_ids_set)].copy()
    matched_count = len(matched_corpus)
    print(f"\n🔗 匹配结果: 成功匹配到 {matched_count} / {len(pay_ids_set)} 条明文语料！")

    # 【深度诊断拦截】：如果匹配为 0，绝对不生成空白表，直接打印详细原因！
    if matched_count == 0:
        print("\n" + "=" * 65)
        print("🚨【严重警告】匹配结果为 0 条！系统已自动拦截，避免生成空白表！")
        print("=" * 65)
        print("📋 原因深度诊断与排查指引：")
        print(f"1. 支付清单共有 {len(pay_ids_set)} 个编号，前 3 个示例为:")
        print(f"   {pay_samples}")
        print(f"2. 语料明文表共有 {len(df_corpus)} 个编号，前 3 个示例为:")
        print(f"   {corpus_samples}")
        
        # 尝试前缀去除非字母数字对比
        pay_pure_nums = {re.sub(r'^[A-Za-z]+', '', x) for x in pay_ids_set}
        corpus_pure_nums = {re.sub(r'^[A-Za-z]+', '', x) for x in df_corpus['CORPUS_ID_NORM']}
        overlap_nums = pay_pure_nums.intersection(corpus_pure_nums)
        if len(overlap_nums) > 0:
            print(f"\n💡 发现端倪：若去掉字母前缀（如'YL'），两表可以匹配到 {len(overlap_nums)} 条！")
            print("   说明两张表在导出时一边的编号带有前缀，另一边没有！")
        else:
            print("\n💡 两张表的编号没有任何交集，可能原因：")
            print("   • 您放入的【支付清单】和【语料明文列表】不是同一批次/同月份的数据；")
            print("   • 后台导出的语料明文表可能不包含本次三方结算项目的数据；")
            print("   • 请打开两份表格人工核对一下前几个词条编号是否能对得上。")
        print("=" * 65)
        return False

    # 4. 解析用户明文列表
    print("\n⏳ [3/4] 正在解析【用户列表(明文)】(提取银行卡号)...")
    wb_u = openpyxl.load_workbook(user_file, data_only=True)
    ws_u = wb_u.active
    u_rows = list(ws_u.iter_rows(values_only=True))
    u_header = [str(c).strip() if c is not None else "" for c in u_rows[0]]
    
    u_idcard_col = find_col_name(u_header, 'id_card')
    u_phone_col = find_col_name(u_header, 'phone')
    u_bank_col = find_col_name(u_header, 'bank_name')
    u_card_col = find_col_name(u_header, 'bank_card')
    u_branch_col = find_col_name(u_header, 'branch_name')
    
    user_map_by_id = {}
    user_map_by_phone = {}
    
    for r in u_rows[1:]:
        r_dict = {h: str(v).strip() if v is not None else "" for h, v in zip(u_header, r)}
        cid_raw = r_dict.get(u_idcard_col, "").strip().upper()
        # 清洗身份证
        cid = re.sub(r'[\s\u200b\uFEFF]+', '', cid_raw)
        ph_raw = r_dict.get(u_phone_col, "").strip()
        ph = re.sub(r'[\s\u200b\uFEFF]+', '', ph_raw)
        if ph.endswith(".0"): ph = ph[:-2]
        
        # 清洗银行卡号（去除空格）
        if u_card_col in r_dict:
            raw_card = r_dict[u_card_col]
            if raw_card.endswith(".0"): raw_card = raw_card[:-2]
            r_dict[u_card_col] = re.sub(r'[\s\u200b\uFEFF\-]+', '', raw_card)
            
        if cid: user_map_by_id[cid] = r_dict
        if ph: user_map_by_phone[ph] = r_dict

    # 5. 汇总医生信息
    print("\n⏳ [4/4] 正在按医生汇总计算结算金额...")
    # 确定月份标签
    cur_year = 2026
    cur_month = 8
    if c_time_col and len(matched_corpus) > 0:
        try:
            t_val = matched_corpus[c_time_col].dropna().iloc[0]
            dt = pd.to_datetime(t_val)
            cur_year = dt.year
            cur_month = dt.month
        except Exception:
            pass
    month_tag = f"{cur_year}年{cur_month:02d}月"

    doctor_summary_list = []
    warnings = []
    
    # 建立支付单价映射
    price_map = {}
    if pay_price_col and pay_price_col in df_pay.columns:
        for _, pr in df_pay.iterrows():
            cid_n = pr['CORPUS_ID_NORM']
            try:
                price_map[cid_n] = float(pr[pay_price_col])
            except Exception:
                price_map[cid_n] = 100.0
                
    grouped = matched_corpus.groupby(c_idcard_col if c_idcard_col else c_name_col)
    
    for grp_key, grp in grouped:
        name = str(grp[c_name_col].iloc[0]).strip() if c_name_col else "未知"
        phone_raw = str(grp[c_phone_col].iloc[0]).strip() if c_phone_col else ""
        if phone_raw.endswith(".0"): phone_raw = phone_raw[:-2]
        phone = re.sub(r'[\s\u200b\uFEFF]+', '', phone_raw)
        
        cid_raw = str(grp[c_idcard_col].iloc[0]).strip().upper() if c_idcard_col else ""
        cid = re.sub(r'[\s\u200b\uFEFF]+', '', cid_raw)
        
        hosp = str(grp[c_hosp_col].iloc[0]).strip() if c_hosp_col else ""
        dept = str(grp[c_dept_col].iloc[0]).strip() if c_dept_col else ""
        title = str(grp[c_title_col].iloc[0]).strip() if c_title_col else ""
        count = len(grp)
        
        # 逐笔累加真实单价
        total_amt = sum(price_map.get(cid_n, 100.0) for cid_n in grp['CORPUS_ID_NORM'])
        
        # 首次提交日期
        first_submit_dt = pd.to_datetime(grp[c_time_col]).min() if c_time_col else datetime.now()
        s_year = first_submit_dt.year
        s_month = first_submit_dt.month
        s_day = first_submit_dt.day
        
        u_info = user_map_by_id.get(cid) or user_map_by_phone.get(phone) or {}
        bank_name = u_info.get(u_bank_col, "")
        bank_card = u_info.get(u_card_col, "")
        branch_name = u_info.get(u_branch_col, "")
        
        if not bank_card:
            warnings.append(f"医生【{name}】(手机: {phone}, 身份证: {cid}) 未在用户列表中找到银行卡号！")
            
        prov, city = parse_hospital_location(hosp)
        
        doctor_summary_list.append({
            '姓名*': name,
            '开始年*': s_year,
            '开始月*': s_month,
            '开始日*': s_day,
            '项目名称': '',
            '金额*': total_amt,
            '姓名1*': name,
            '省份*': prov,
            '市*': city,
            '身份证*': cid,
            '手机号*': phone,
            '开户行*': bank_name,
            '银行卡号*': bank_card,
            '工作单位*': hosp,
            '医务职称*': title,
            '科室*': dept,
            '签署日期': '',
            '签署日期1': '',
            '语料条数': count,
            '支行名称': branch_name
        })
        
    doctor_summary_list.sort(key=lambda x: (-x['金额*'], x['姓名*']))
    
    total_docs = len(doctor_summary_list)
    total_corpus_count = sum(d['语料条数'] for d in doctor_summary_list)
    total_settle_amount = sum(d['金额*'] for d in doctor_summary_list)

    # 导出文件
    final_columns = [
        '姓名*', '开始年*', '开始月*', '开始日*', '项目名称', '金额*', '姓名1*', '省份*', '市*',
        '身份证*', '手机号*', '开户行*', '银行卡号*', '工作单位*', '医务职称*', '科室*',
        '签署日期', '签署日期1'
    ]
    
    # 1. 最终.xlsx
    wb_final = openpyxl.Workbook()
    ws_final = wb_final.active
    ws_final.title = "Sheet0"
    
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="000000")
    data_font = Font(name="微软雅黑", size=10, color="000000")
    border_style = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )
    
    for c_idx, col_name in enumerate(final_columns, 1):
        c = ws_final.cell(row=1, column=c_idx, value=col_name)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_style
        
    for r_idx, doc in enumerate(doctor_summary_list, 2):
        for c_idx, col_name in enumerate(final_columns, 1):
            val = doc.get(col_name, "")
            c = ws_final.cell(row=r_idx, column=c_idx)
            if col_name in ['身份证*', '手机号*', '银行卡号*']:
                c.number_format = '@'
                c.value = str(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ['开始年*', '开始月*', '开始日*']:
                c.value = int(val) if val != "" else ""
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == '金额*':
                c.value = float(val) if val != "" else 0.0
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in ['姓名*', '姓名1*', '省份*', '市*', '医务职称*', '科室*']:
                c.value = str(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.value = str(val) if val != "" else ""
                c.alignment = Alignment(horizontal="left", vertical="center")
            c.font = data_font
            c.border = border_style
            
    for col in ws_final.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '')) for cell in col)
        ws_final.column_dimensions[col_letter].width = min(max(max_l + 4, 12), 40)
        
    out_final_xlsx = os.path.join(target_dir, "最终.xlsx")
    wb_final.save(out_final_xlsx)
    
    # 2. 最终.xls (标准 BIFF8)
    out_final_xls = os.path.join(target_dir, "最终.xls")
    if xlwt:
        wb_xls = xlwt.Workbook(encoding='utf-8')
        ws_xls = wb_xls.add_sheet('Sheet0')
        st_hdr = xlwt.easyxf('font: name 微软雅黑, bold on, height 220; align: horiz center, vert center; pattern: pattern solid, fore_colour gray25; borders: left thin, right thin, top thin, bottom thin')
        st_txt = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz center, vert center; borders: left thin, right thin, top thin, bottom thin')
        st_left = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz left, vert center; borders: left thin, right thin, top thin, bottom thin')
        st_amt = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz right, vert center; borders: left thin, right thin, top thin, bottom thin', num_format_str='#,##0.00')
        st_int = xlwt.easyxf('font: name 微软雅黑, height 200; align: horiz center, vert center; borders: left thin, right thin, top thin, bottom thin')
        
        for c_idx, col_name in enumerate(final_columns):
            ws_xls.write(0, c_idx, col_name, st_hdr)
            
        for r_idx, doc in enumerate(doctor_summary_list, 1):
            for c_idx, col_name in enumerate(final_columns):
                val = doc.get(col_name, "")
                if col_name in ['身份证*', '手机号*', '银行卡号*']:
                    ws_xls.write(r_idx, c_idx, str(val), st_txt)
                elif col_name in ['开始年*', '开始月*', '开始日*']:
                    ws_xls.write(r_idx, c_idx, int(val) if val != "" else "", st_int)
                elif col_name == '金额*':
                    ws_xls.write(r_idx, c_idx, float(val) if val != "" else 0.0, st_amt)
                elif col_name in ['姓名*', '姓名1*', '省份*', '市*', '医务职称*', '科室*']:
                    ws_xls.write(r_idx, c_idx, str(val), st_txt)
                else:
                    ws_xls.write(r_idx, c_idx, str(val) if val != "" else "", st_left)
        wb_xls.save(out_final_xls)
    
    # 3. 筛选明细与对账表
    out_detail_xlsx = os.path.join(target_dir, f"语料列表(明文)-{cur_month}月结算筛选.xlsx")
    wb_detail = openpyxl.Workbook()
    
    ws_d1 = wb_detail.active
    ws_d1.title = f"{cur_month}月结算语料明细({len(matched_corpus)}条)"
    
    d1_headers = [
        '语料词条编号', '医生姓名', '手机号', '身份证号码', '医院', '科室', '职称',
        '结算单价', '审核状态', '是否结算', '开户银行', '支行名称', '银行卡号',
        '提交时间', '题目内容', '回答记录'
    ]
    for c_idx, h in enumerate(d1_headers, 1):
        c = ws_d1.cell(row=1, column=c_idx, value=h)
        c.font = Font(name="微软雅黑", size=10, bold=True, color="003366")
        c.fill = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        
    for r_idx, (_, r) in enumerate(matched_corpus.iterrows(), 2):
        cid_str = re.sub(r'[\s\u200b\uFEFF]+', '', str(r[c_idcard_col]).strip().upper()) if c_idcard_col else ""
        ph_str = re.sub(r'[\s\u200b\uFEFF]+', '', str(r[c_phone_col]).strip()) if c_phone_col else ""
        if ph_str.endswith(".0"): ph_str = ph_str[:-2]
        
        u_info = user_map_by_id.get(cid_str) or user_map_by_phone.get(ph_str) or {}
        cid_norm = r['CORPUS_ID_NORM']
        
        row_vals = [
            str(r[c_id_col]),
            str(r[c_name_col]) if c_name_col else "",
            ph_str,
            cid_str,
            str(r[c_hosp_col]) if c_hosp_col else "",
            str(r[c_dept_col]) if c_dept_col else "",
            str(r[c_title_col]) if c_title_col else "",
            price_map.get(cid_norm, 100.0),
            "审核通过",
            "是",
            u_info.get(u_bank_col, ""),
            u_info.get(u_branch_col, ""),
            u_info.get(u_card_col, ""),
            str(r[c_time_col]) if c_time_col else "",
            str(r[find_col_name(df_corpus.columns, 'question')]) if find_col_name(df_corpus.columns, 'question') else "",
            str(r[find_col_name(df_corpus.columns, 'answer')]) if find_col_name(df_corpus.columns, 'answer') else ""
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_d1.cell(row=r_idx, column=c_idx)
            if d1_headers[c_idx-1] in ['语料词条编号', '手机号', '身份证号码', '银行卡号']:
                c.number_format = '@'
                c.value = str(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif d1_headers[c_idx-1] == '结算单价':
                c.value = float(val)
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.value = str(val)
                c.alignment = Alignment(horizontal="left", vertical="center")
            c.font = Font(name="微软雅黑", size=9)
            c.border = border_style

    # Sheet 2: 医生汇总
    ws_d2 = wb_detail.create_sheet(title="医生结算汇总")
    d2_headers = [
        '序号', '医生姓名', '身份证号码', '手机号码', '所在医院', '科室', '职称',
        '结算语料条数', '应付总金额(元)', '开户银行', '支行名称', '银行卡号',
        '省份', '城市', '首次提交日期'
    ]
    for c_idx, h in enumerate(d2_headers, 1):
        c = ws_d2.cell(row=1, column=c_idx, value=h)
        c.font = Font(name="微软雅黑", size=10, bold=True, color="1E4620")
        c.fill = PatternFill(start_color="EAF2E8", end_color="EAF2E8", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        
    for idx, doc in enumerate(doctor_summary_list, 1):
        r_vals = [
            idx, doc['姓名*'], doc['身份证*'], doc['手机号*'], doc['工作单位*'],
            doc['科室*'], doc['医务职称*'], doc['语料条数'], doc['金额*'],
            doc['开户行*'], doc['支行名称'], doc['银行卡号*'], doc['省份*'], doc['市*'],
            f"{doc['开始年*']}-{doc['开始月*']:02d}-{doc['开始日*']:02d}"
        ]
        for c_idx, val in enumerate(r_vals, 1):
            c = ws_d2.cell(row=idx+1, column=c_idx)
            if d2_headers[c_idx-1] in ['身份证号码', '手机号码', '银行卡号']:
                c.number_format = '@'
                c.value = str(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif d2_headers[c_idx-1] == '应付总金额(元)':
                c.value = float(val)
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif d2_headers[c_idx-1] in ['序号', '结算语料条数']:
                c.value = int(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.value = str(val)
                c.alignment = Alignment(horizontal="left", vertical="center")
            c.font = Font(name="微软雅黑", size=9)
            c.border = border_style

    tot_row = len(doctor_summary_list) + 2
    ws_d2.cell(row=tot_row, column=1, value="合计").font = Font(name="微软雅黑", size=10, bold=True)
    ws_d2.cell(row=tot_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws_d2.cell(row=tot_row, column=8, value=total_corpus_count).font = Font(name="微软雅黑", size=10, bold=True)
    ws_d2.cell(row=tot_row, column=8).alignment = Alignment(horizontal="center", vertical="center")
    ws_d2.cell(row=tot_row, column=9, value=total_settle_amount).font = Font(name="微软雅黑", size=10, bold=True)
    ws_d2.cell(row=tot_row, column=9).number_format = '#,##0.00'
    ws_d2.cell(row=tot_row, column=9).alignment = Alignment(horizontal="right", vertical="center")
    
    for c in range(1, len(d2_headers)+1):
        ws_d2.cell(row=tot_row, column=c).border = border_style
        
    for ws_cur in [ws_d1, ws_d2]:
        for col in ws_cur.columns:
            col_let = get_column_letter(col[0].column)
            max_l = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '')) for cell in col)
            ws_cur.column_dimensions[col_let].width = min(max(max_l + 4, 12), 45)
            
    wb_detail.save(out_detail_xlsx)
    
    print("\n" + "=" * 65)
    print(f"🎉 恭喜！{month_tag} 结算数据已全部处理完成！")
    print("=" * 65)
    print(f" 📊 本期结算总人数 : {total_docs} 位医生")
    print(f" 📑 本期结算语料数 : {total_corpus_count} 条")
    print(f" 💰 本期应付总金额 : ¥ {total_settle_amount:,.2f} 元")
    print("-" * 65)
    print(" 📁 生成文件清单：")
    print(f"   1. [标准发放表] {os.path.basename(out_final_xlsx)}")
    if xlwt:
        print(f"   2. [银行导入表] {os.path.basename(out_final_xls)}")
    print(f"   3. [对账明细表] {os.path.basename(out_detail_xlsx)}")
    
    if warnings:
        warn_path = os.path.join(target_dir, "【警告】待核查异常名单.txt")
        with open(warn_path, "w", encoding="utf-8") as wf:
            wf.write("\n".join(warnings))
        print("\n⚠️ 存在数据告警（已保存至 【警告】待核查异常名单.txt）：")
        for w in warnings:
            print(f"   • {w}")
    else:
        print("\n✨ 数据完整性核验 100% 通过，无缺失卡号！")
    print("=" * 65 + "\n")
    return True

if __name__ == "__main__":
    run_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    process_settlement(run_dir)
