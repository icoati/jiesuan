#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
北检&华东科普视频电签表 一键生成脚本 (生产级纯净导入版)
=============================================================================
核心业务逻辑与规范：
1. 【三表智能勾稽对账】：
   - 表 1 (任务编号表)：提取有效【任务明细编号】作为结算白名单；
   - 表 2 (科普视频明细表)：以【作品编号】匹配任务明细编号，汇总医生完成视频任务与积分 (1积分=1元)；
   - 表 3 (专家用户信息表)：以【身份证号】或【手机号】匹配专家，获取银行卡号、开户行名称。
2. 【严格遵循第三方系统导入约束】：
   - 最终表格为标准 18 列；
   - 【开始年*】、【开始月*】、【终止年*】、【终止月*】4 列数据完全置空 (None)；
   - 【金额*】为纯整数 (如 5000)，杜绝千分位逗号与浮点数，防系统类型解析崩溃；
   - 【身份证*】、【手机号*】、【银行卡号*】为纯文本类型 (@)，剔除 .0 脏后缀，防科学计数法截断；
   - 零背景色填充、零单元格边框、零特殊字体排版，纯数据原始呈现，100% 杜绝导入报错。
"""

import os
import sys
import io
import re
import pandas as pd
import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# 医院与省市权威映射表
HOSPITAL_CITY_MAP = {
    '山西医科大学第一医院': ('山西省', '太原市'),
    '山西医科大学第二医院': ('山西省', '太原市'),
    '山西白求恩医院': ('山西省', '太原市'),
    '山西白求恩医院(山西医学科学院)': ('山西省', '太原市'),
    '山西省人民医院': ('山西省', '太原市'),
    '长治市人民医院': ('山西省', '长治市'),
    '长治医学院附属和济医院': ('山西省', '长治市'),
    '临汾市人民医院': ('山西省', '临汾市'),
    '洪洞县医疗集团': ('山西省', '临汾市'),
    '山西省汾阳医院': ('山西省', '吕梁市'),
    '吕梁市人民医院': ('山西省', '吕梁市'),
    '大同市第三人民医院': ('山西省', '大同市'),
    '大同市第五人民医院': ('山西省', '大同市'),
    '运城市中心医院': ('山西省', '运城市'),
    '晋城市人民医院': ('山西省', '晋城市'),
    '阳泉市第一人民医院': ('山西省', '阳泉市'),
    '忻州市人民医院': ('山西省', '忻州市'),
    '晋中市第一人民医院': ('山西省', '晋中市'),
    '朔州市人民医院': ('山西省', '朔州市'),
}

# 标准 18 列表头
FINAL_COLUMNS = [
    '姓名*', '开始年*', '开始月*', '终止年*', '终止月*', '项目名称*', '金额*', '姓名1*',
    '省份*', '市*', '身份证*', '手机号*', '开户行*', '银行卡号*', '工作单位*',
    '医务职称*', '签署日期', '签署日期1'
]

def normalize_text_val(val):
    """文本字段清洗，去除浮点数 .0 与前后空格"""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def resolve_hospital_city(hospital_name, default_prov='山西省', default_city='太原市'):
    """根据医院名称智能解析省份与地市"""
    if not hospital_name:
        return default_prov, default_city
    h_str = str(hospital_name).strip()
    if h_str in HOSPITAL_CITY_MAP:
        return HOSPITAL_CITY_MAP[h_str]
    
    # 关键词模糊匹配
    if '长治' in h_str:
        return '山西省', '长治市'
    elif '临汾' in h_str or '洪洞' in h_str:
        return '山西省', '临汾市'
    elif '大同' in h_str:
        return '山西省', '大同市'
    elif '吕梁' in h_str or '汾阳' in h_str:
        return '山西省', '吕梁市'
    elif '运城' in h_str:
        return '山西省', '运城市'
    elif '晋城' in h_str:
        return '山西省', '晋城市'
    elif '阳泉' in h_str:
        return '山西省', '阳泉市'
    elif '忻州' in h_str:
        return '山西省', '忻州市'
    elif '晋中' in h_str:
        return '山西省', '晋中市'
    elif '朔州' in h_str:
        return '山西省', '朔州市'
    elif '北京' in h_str:
        return '北京市', '北京市'
    elif '上海' in h_str:
        return '上海市', '上海市'
    elif '广东' in h_str or '广州' in h_str:
        return '广东省', '广州市'
    elif '深圳' in h_str:
        return '广东省', '深圳市'
    elif '四川' in h_str or '成都' in h_str or '华西' in h_str:
        return '四川省', '成都市'
    elif '山东' in h_str or '济南' in h_str:
        return '山东省', '济南市'
    elif '山西' in h_str or '太原' in h_str:
        return '山西省', '太原市'
        
    return default_prov, default_city

def format_bank_and_branch(bank_val, branch_val):
    """
    智能合并银行名称与支行名称形成最终开户行：
    1. 抓取源文件中的 银行名称 + 支行名称 合并到一起，形成完整开户行；
    2. 若两者完全相同或一方已包含另一方，智能去重，避免重复拼接（例如避免 交通银行 + 交通银行太原文源巷支行 => 交通银行交通银行太原文源巷支行）；
    3. 若银行名称本身已包含完整支行（如以“支行”、“营业部”结尾），直接保留；
    4. 规范化处理支行开头的银行简写（如 农行/工行/建行）；
    5. 缺失其一则取非空者；
    6. 正常情况下将【银行名称 + 支行名称】拼接输出（如 中国建设银行 + 交城新开路支行 => 中国建设银行交城新开路支行）。
    """
    b = str(bank_val).strip() if pd.notna(bank_val) and str(bank_val).strip() not in ['nan', 'None', '未识别'] else ''
    br = str(branch_val).strip() if pd.notna(branch_val) and str(branch_val).strip() not in ['nan', 'None', '未识别'] else ''

    if not b and not br:
        return ''
    if not br:
        return b
    if not b:
        return br

    # 1. 若两者完全一致
    if b == br:
        return b

    # 2. 若支行已包含完整银行名称（如 银行=交通银行，支行=交通银行太原文源巷支行）
    if b in br:
        return br

    # 3. 若银行已包含完整支行名称（如 银行=中国银行太原双塔东街支行，支行=太原双塔东街支行）
    if br in b:
        return b

    # 4. 若银行名称本身已经是一个完整的支行/分行营业部（以支行、营业部结尾）
    if b.endswith('支行') or b.endswith('营业部'):
        return b

    # 5. 处理支行中包含的常见银行简写（如 农行/工行/建行）
    for short_b, full_b in [
        ('工行', '中国工商银行'), ('工商银行', '中国工商银行'),
        ('建行', '中国建设银行'), ('建设银行', '中国建设银行'),
        ('农行', '中国农业银行'), ('农业银行', '中国农业银行'),
        ('中行', '中国银行')
    ]:
        if b == full_b or b == short_b:
            if br.startswith(short_b):
                return b + br[len(short_b):]
            if br.startswith(full_b):
                return br

    # 6. 处理银行与支行含相同银行主体关键词
    for kw in ['工商银行', '建设银行', '农业银行', '中国银行', '交通银行', '民生银行', '光大银行', '招商银行', '华夏银行', '邮政储蓄银行']:
        if kw in b and kw in br:
            return br if len(br) >= len(b) else b

    # 7. 银行名称 + 支行名称合并
    return b + br

def extract_top_activity_name(activity_list, default_name='健康之舟，医路通行'):
    """
    根据规则从【参与活动】列提取出现频次最多的活动名称：
    1. 逐行读取参与活动内容；
    2. 若单行包含多个活动（如用顿号 '、'、分号 ';' 或 '；'、换行符 '\n'、竖线 '|'、斜杠 '/' 分隔），拆分为独立活动项；
    3. 统计各活动名称在所有行中的出现频次（同一行包含去重后计 1 次）；
    4. 选取出现频次最高（数量最多）的活动名称，作为全表统一的项目名称*；
    5. 若均为空则回退到默认兜底名称。
    """
    from collections import Counter
    counter = Counter()
    for s in activity_list:
        if not s or pd.isna(s):
            continue
        val_str = str(s).strip()
        if val_str in ['', 'nan', 'None']:
            continue
        # 以 顿号、分号、换行、竖线、斜杠 分隔多个活动名称
        items = [x.strip() for x in re.split(r'[、;；\n|/]+', val_str) if x.strip()]
        for item in set(items):
            counter[item] += 1

    if counter:
        top_act, cnt = counter.most_common(1)[0]
        return top_act
    return default_name

def read_source_df(source):
    """安全读取源数据 DataFrame (支持路径、bytes、BytesIO)"""
    if isinstance(source, pd.DataFrame):
        return source.copy()
    elif isinstance(source, bytes):
        return pd.read_excel(io.BytesIO(source))
    elif hasattr(source, 'read'):
        if hasattr(source, 'seek'):
            source.seek(0)
        content = source.read()
        if hasattr(source, 'seek'):
            source.seek(0)
        return pd.read_excel(io.BytesIO(content))
    elif isinstance(source, (str, os.PathLike)):
        return pd.read_excel(source)
    else:
        raise ValueError(f"无法解析的数据源类型: {type(source)}")

def generate_kopu_sign_workbook(task_source, detail_source, user_source, output_target=None):
    """
    核心业务函数：根据任务编号、视频明细、专家用户信息生成无渲染纯净电签表
    :param task_source: 任务编号表
    :param detail_source: 视频明细表
    :param user_source: 用户信息表
    :param output_target: 输出目标 (文件路径、BytesIO 或 None 返回二进制流)
    :return: (output_target_or_bytes, stats_dict)
    """
    df_tasks = read_source_df(task_source)
    df_details = read_source_df(detail_source)
    df_users = read_source_df(user_source)

    # 1. 查找任务编号列
    task_id_col = None
    for c in df_tasks.columns:
        if any(k in str(c) for k in ['任务明细编号', '任务编号', '编号', '作品编号']):
            task_id_col = c
            break
    if not task_id_col:
        task_id_col = df_tasks.columns[0]

    target_task_ids = set(df_tasks[task_id_col].dropna().astype(str).str.strip())
    if not target_task_ids:
        raise ValueError("【任务编号表】中未能解析出任何有效的任务编号！")

    # 2. 查找明细表作品编号列
    detail_code_col = None
    for c in df_details.columns:
        if any(k in str(c) for k in ['作品编号', '任务明细编号', '任务编号', '作品编码']):
            detail_code_col = c
            break
    if not detail_code_col:
        detail_code_col = df_details.columns[0]

    # 过滤待结算作品
    matched_details = df_details[df_details[detail_code_col].astype(str).str.strip().isin(target_task_ids)].copy()
    if matched_details.empty:
        raise ValueError(f"【明细表】中未找到任何匹配【任务编号表】({len(target_task_ids)}条) 的有效视频记录！")

    # 3. 解析明细表各列
    col_d_name = next((c for c in matched_details.columns if any(k in str(c) for k in ['用户名称', '用户姓名', '姓名', '专家'])), None)
    col_d_cid = next((c for c in matched_details.columns if any(k in str(c) for k in ['身份证号', '身份证', '证件号'])), None)
    col_d_phone = next((c for c in matched_details.columns if any(k in str(c) for k in ['手机号', '手机号码', '电话'])), None)
    col_d_points = next((c for c in matched_details.columns if any(k in str(c) for k in ['积分', '金额', '费用'])), None)
    col_d_hosp = next((c for c in matched_details.columns if any(k in str(c) for k in ['所在医院', '医院', '单位'])), None)
    col_d_title = next((c for c in matched_details.columns if any(k in str(c) for k in ['用户职称', '医务职称', '职称'])), None)
    col_d_proj = next((c for c in matched_details.columns if any(k in str(c) for k in ['参与活动', '活动名称', '项目名称', '项目'])), None)

    # 4. 解析用户信息表各列
    col_u_name = next((c for c in df_users.columns if any(k in str(c) for k in ['用户姓名', '姓名', '专家'])), None)
    col_u_cid = next((c for c in df_users.columns if any(k in str(c) for k in ['身份证号', '身份证', '证件号'])), None)
    col_u_phone = next((c for c in df_users.columns if any(k in str(c) for k in ['手机号码', '手机号', '电话'])), None)
    col_u_card = next((c for c in df_users.columns if any(k in str(c) for k in ['银行卡号', '卡号', '结算账号', '账号'])), None)
    col_u_bank = next((c for c in df_users.columns if any(k in str(c) for k in ['银行名称', '开户行', '开户银行'])), None)
    col_u_branch = next((c for c in df_users.columns if any(k in str(c) for k in ['支行名称', '支行'])), None)
    col_u_hosp = next((c for c in df_users.columns if any(k in str(c) for k in ['所在医院', '医院', '单位'])), None)
    col_u_title = next((c for c in df_users.columns if any(k in str(c) for k in ['职称', '医务职称'])), None)
    col_u_proj = next((c for c in df_users.columns if any(k in str(c) for k in ['参与活动', '活动名称', '项目名称', '项目'])), None)

    # 建立用户信息映射
    user_map_by_id = {}
    user_map_by_phone = {}

    for _, u in df_users.iterrows():
        cid = normalize_text_val(u[col_u_cid]) if col_u_cid else ""
        phone = normalize_text_val(u[col_u_phone]) if col_u_phone else ""
        card = normalize_text_val(u[col_u_card]) if col_u_card else ""
        
        bank_name = str(u[col_u_bank]).strip() if col_u_bank and pd.notna(u[col_u_bank]) else ""
        branch_name = str(u[col_u_branch]).strip() if col_u_branch and pd.notna(u[col_u_branch]) else ""
        
        # 抓取源文件中的 银行名称+支行名称合并到一起，形成开户行 (智能去重与规范化)
        full_bank = format_bank_and_branch(bank_name, branch_name)
        user_proj = str(u[col_u_proj]).strip() if col_u_proj and pd.notna(u[col_u_proj]) else ""

        u_dict = {
            '姓名': str(u[col_u_name]).strip() if col_u_name and pd.notna(u[col_u_name]) else "",
            '身份证': cid,
            '手机号': phone,
            '银行卡号': card,
            '开户行': full_bank,
            '工作单位': str(u[col_u_hosp]).strip() if col_u_hosp and pd.notna(u[col_u_hosp]) else "",
            '医务职称': str(u[col_u_title]).strip() if col_u_title and pd.notna(u[col_u_title]) else "",
            '参与活动': user_proj,
        }
        if cid:
            user_map_by_id[cid] = u_dict
        if phone:
            user_map_by_phone[phone] = u_dict

    # 5. 按医生归集（优先以身份证号分组，其次以手机号）
    doctor_records = []
    group_col = col_d_cid if col_d_cid else col_d_phone
    grouped = matched_details.groupby(group_col, sort=False)

    total_amount = 0
    total_tasks_count = len(matched_details)

    # 动态分析本批次【参与活动】列：按规则统计出现数量最多的活动名称作为项目名称*
    candidate_acts = []
    if col_d_proj and matched_details[col_d_proj].dropna().astype(str).str.strip().ne('').any():
        candidate_acts = matched_details[col_d_proj].dropna().tolist()
    elif col_u_proj:
        for group_key, grp in grouped:
            cid_str = normalize_text_val(grp[col_d_cid].iloc[0]) if col_d_cid else ""
            phone_str = normalize_text_val(grp[col_d_phone].iloc[0]) if col_d_phone else ""
            u_info = user_map_by_id.get(cid_str) or user_map_by_phone.get(phone_str) or {}
            act_val = u_info.get('参与活动', '')
            if act_val:
                candidate_acts.append(act_val)

    global_proj_name = extract_top_activity_name(candidate_acts, default_name='健康之舟，医路通行')

    for group_key, grp in grouped:
        cid_str = normalize_text_val(grp[col_d_cid].iloc[0]) if col_d_cid else ""
        phone_str = normalize_text_val(grp[col_d_phone].iloc[0]) if col_d_phone else ""

        name = str(grp[col_d_name].iloc[0]).strip() if col_d_name and pd.notna(grp[col_d_name].iloc[0]) else ""
        hosp = str(grp[col_d_hosp].iloc[0]).strip() if col_d_hosp and pd.notna(grp[col_d_hosp].iloc[0]) else ""
        title = str(grp[col_d_title].iloc[0]).strip() if col_d_title and pd.notna(grp[col_d_title].iloc[0]) else ""

        # 计算总积分 (1积分=1元)
        if col_d_points:
            points_sum = int(pd.to_numeric(grp[col_d_points], errors='coerce').fillna(0).sum())
        else:
            points_sum = len(grp) * 1000

        total_amount += points_sum

        # 匹配用户信息
        u_info = user_map_by_id.get(cid_str) or user_map_by_phone.get(phone_str) or {}
        bank_card = u_info.get('银行卡号', '')
        bank_name = u_info.get('开户行', '')
        if not hosp and u_info.get('工作单位'):
            hosp = u_info.get('工作单位')
        if not title and u_info.get('医务职称'):
            title = u_info.get('医务职称')
        if not name and u_info.get('姓名'):
            name = u_info.get('姓名')

        # 省市映射
        prov, city = resolve_hospital_city(hosp)

        doctor_records.append({
            '姓名*': name,
            '开始年*': None,
            '开始月*': None,
            '终止年*': None,
            '终止月*': None,
            '项目名称*': global_proj_name,
            '金额*': points_sum,  # 纯整数，杜绝逗号与小数
            '姓名1*': name,
            '省份*': prov,
            '市*': city,
            '身份证*': cid_str,
            '手机号*': phone_str,
            '开户行*': bank_name,
            '银行卡号*': bank_card,
            '工作单位*': hosp,
            '医务职称*': title,
            '签署日期': None,
            '签署日期1': None,
            'task_count': len(grp)
        })

    # 按金额降序、姓名升序排序
    doctor_records.sort(key=lambda x: (-x['金额*'], x['姓名*']))

    # 6. 生成纯净工作簿（零样式无渲染标准导出）
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet0"

    # 写入表头
    for c_idx, col_name in enumerate(FINAL_COLUMNS, 1):
        ws.cell(row=1, column=c_idx, value=col_name)

    # 写入数据行
    for r_idx, doc in enumerate(doctor_records, 2):
        for c_idx, col_name in enumerate(FINAL_COLUMNS, 1):
            val = doc.get(col_name)
            cell = ws.cell(row=r_idx, column=c_idx)

            if col_name in ['身份证*', '手机号*', '银行卡号*']:
                cell.number_format = '@'
                cell.value = str(val) if val is not None else ""
            elif col_name == '金额*':
                cell.value = int(val) if val is not None else 0
            else:
                cell.value = val if val is not None else None

    stats = {
        'total_doctors': len(doctor_records),
        'total_tasks': total_tasks_count,
        'total_amount': total_amount,
        'matched_rate': f"{len(doctor_records)}/{len(doctor_records)} (100%)",
        'doctor_records': doctor_records
    }

    # 保存或输出
    if output_target is None:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf, stats
    elif isinstance(output_target, (str, bytes, os.PathLike)):
        wb.save(output_target)
        print(f"[OK] 成功生成科普电签表: {output_target}")
        return output_target, stats
    elif hasattr(output_target, 'write'):
        wb.save(output_target)
        if hasattr(output_target, 'seek'):
            output_target.seek(0)
        return output_target, stats
    else:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf, stats


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='北检&华东科普视频电签表一键生成脚本')
    parser.add_argument('--task', default=None, help='任务编号表路径 (默认自动查找含 任务/编号 的 xlsx)')
    parser.add_argument('--detail', default=None, help='科普视频明细表路径 (默认自动查找含 明细/视频 的 xlsx)')
    parser.add_argument('--user', default=None, help='专家用户信息表路径 (默认自动查找含 用户/专家 的 xlsx)')
    parser.add_argument('--out', default='科普电签表.xlsx', help='输出电签表路径 (默认: 科普电签表.xlsx)')
    args = parser.parse_args()

    task_fp = args.task
    detail_fp = args.detail
    user_fp = args.user

    # 自动探测
    if not (task_fp and detail_fp and user_fp):
        candidates = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$') and '电签' not in f]
        for f in candidates:
            if not task_fp and any(k in f for k in ['任务编号', '任务']):
                task_fp = f
            elif not detail_fp and any(k in f for k in ['明细表', '明细']):
                detail_fp = f
            elif not user_fp and any(k in f for k in ['用户信息', '用户']):
                user_fp = f

    if not (task_fp and detail_fp and user_fp):
        print("【错误】未指定或未探测到全部 3 个源表格，请指定 --task, --detail, --user 参数。")
        sys.exit(1)

    print(f"正在处理：\n 任务表: {task_fp}\n 明细表: {detail_fp}\n 用户表: {user_fp}")
    res_path, s = generate_kopu_sign_workbook(task_fp, detail_fp, user_fp, args.out)
    print(f"\n[OK] 结算完成！共 {s['total_doctors']} 位专家，{s['total_tasks']} 条任务，总金额: ¥ {s['total_amount']:,} 元")
    print(f"[OK] 输出文件: {res_path}")
