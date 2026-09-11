# -*- coding: utf-8 -*-
"""
=============================================================================
医院等级智能识别与检索工具 (Hospital Grade Matcher & Search Tool) - 超级双引擎版
=============================================================================
核心特点:
1. 双模式高效检索引擎:
   - 模式 A (首选 SQLite 数据库): 若存在 hospital_master.db，毫秒级连接，0 内存开销，单机每秒 3.6 万次索引检索
   - 模式 B (轻量 GZ/JSON 字典): 若部署在云端/Git，仅需 6.7 MB 的 hospital_base_cache.json.gz，0.5 秒载入半秒内完成 50 万条缓存检索
2. 全国 49.4 万条超全医疗机构知识库覆盖:
   - 包含全国各省市三级甲等、三级乙等、二级甲等、二级乙等、一级医院及基层卫生院
3. 三层智能流水线:
   - 第 1 层: 规则秒判 (村卫生室/诊所/门诊部/药房 -> 无等级; 卫生院/服务中心 -> 一级医院)
   - 第 2 层: 本地 49.4 万超级库三级索引查找 (全称 -> 简称去括号 -> 去省市归一化)
   - 第 3 层: 自动化在线检索保底 (针对历史库缺失等级的重点公立医院/新设医院，自动联网请求、正则解析并持久化写入自定义增量库)
4. 严格规范化为系统限定 10 个标准下拉菜单项:
   [无等级, 一级医院, 一级乙等, 一级甲等, 二级医院, 二级乙等, 二级甲等, 三级医院, 三级乙等, 三级甲等]
5. 支持多种使用模式:
   - 命令行单条即时查询: python hospital_grade_tool.py -q "成都市第三人民医院"
   - 交互式连续查询: python hospital_grade_tool.py -i
   - Excel 批量补齐与 300 条自动分批: python hospital_grade_tool.py -e "数据.xlsx" -s 300
=============================================================================
"""

import os
import sys
import re
import json
import gzip
import time
import sqlite3
import urllib.parse
import argparse
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

# 确保控制台 UTF-8 输出
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# 允许的 10 个标准下拉菜单项
VALID_GRADES = [
    '无等级', '一级医院', '一级乙等', '一级甲等',
    '二级医院', '二级乙等', '二级甲等',
    '三级医院', '三级乙等', '三级甲等'
]

# 非医院/基层服务机构关键词（直接定级为“无等级”）
NON_HOSPITAL_KEYWORDS = [
    '卫生室', '村卫生', '服务站', '卫生服务站', '门诊部', '门诊所', '门诊',
    '诊所', '卫生所', '药房', '药店', '大药房', '医药', '工会', '大厦',
    '物流', '急救站', '保健所', '检验所', '体检中心', '研究所'
]

# 基层卫生院/社区卫生服务中心关键词
PRIMARY_CARE_KEYWORDS = ['卫生院', '社区卫生服务中心', '防保所', '保健站']

# 重点公立医院特征关键词（若库中标为无等级，则需触发联网二次校验）
MAJOR_HOSPITAL_KEYWORDS = [
    '人民医院', '大学附属', '中医院', '中医医院', '中心医院', '第一医院',
    '第二医院', '第三医院', '第四医院', '妇幼保健院', '妇幼保健', '骨伤医院',
    '肿瘤医院', '儿童医院', '脑科医院', '胸科医院', '传染病医院'
]

class HospitalGradeMatcher:
    def __init__(self, db_file="hospital_master.db", base_cache_file="hospital_base_cache.json", custom_cache_file="hospital_custom_cache.json"):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.db_path = os.path.join(self.script_dir, db_file)
        self.base_cache_path = os.path.join(self.script_dir, base_cache_file)
        self.base_gz_path = self.base_cache_path + ".gz"
        self.custom_cache_path = os.path.join(self.script_dir, custom_cache_file)
        
        self.use_db = False
        self.conn = None
        self.cache = {}
        self.custom_cache = {}
        
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        
        # 1. 加载增量自定义缓存 (新在线搜索到的内容，优先级最高)
        if os.path.exists(self.custom_cache_path):
            try:
                with open(self.custom_cache_path, 'r', encoding='utf-8') as f:
                    self.custom_cache = json.load(f)
            except Exception:
                self.custom_cache = {}
                
        # 2. 初始化存储引擎 (优先 SQLite，其次 GZ/JSON)
        if os.path.exists(self.db_path):
            try:
                self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
                self.cursor = self.conn.cursor()
                self.use_db = True
            except Exception:
                self.use_db = False
                
        if not self.use_db:
            self.load_json_cache()

    def load_json_cache(self):
        """若无 SQLite 数据库，则从 Gzip 压缩包或 JSON 文件秒级加载字典"""
        if os.path.exists(self.base_gz_path):
            try:
                with gzip.open(self.base_gz_path, 'rt', encoding='utf-8') as f:
                    self.cache = json.load(f)
            except Exception:
                pass
        elif os.path.exists(self.base_cache_path):
            try:
                with open(self.base_cache_path, 'r', encoding='utf-8') as f:
                    self.cache = json.load(f)
            except Exception:
                pass

    def normalize_grade_text(self, text):
        """将任意文本中提取到的等级词归一化为系统的 10 个标准下拉项"""
        if not text:
            return '无等级'
        text = text.strip()
        
        if '三级甲等' in text or '三甲' in text:
            return '三级甲等'
        elif '三级乙等' in text or '三乙' in text:
            return '三级乙等'
        elif '二级甲等' in text or '二甲' in text:
            return '二级甲等'
        elif '二级乙等' in text or '二乙' in text:
            return '二级乙等'
        elif '一级甲等' in text or '一甲' in text:
            return '一级甲等'
        elif '一级乙等' in text or '一乙' in text:
            return '一级乙等'
        elif '三级' in text or '三丙' in text or '三级综合' in text:
            return '三级医院'
        elif '二级' in text or '二丙' in text or '二级综合' in text:
            return '二级医院'
        elif '一级' in text or '一丙' in text or '一级综合' in text:
            return '一级医院'
        elif '未定级' in text or '未评级' in text or '无等级' in text:
            return '无等级'
        return '无等级'

    def clean_name(self, name, prov='', city=''):
        """去除括号别名、公司后缀、省市前缀以用于高命中率归一化检索"""
        name = re.sub(r'[\(（\[【].*?[\)）\]】]', '', name)
        name = re.sub(r'有限公司|有限责任公司|股份有限公司|集团', '', name)
        
        if prov and name.startswith(prov):
            name = name[len(prov):]
        prov_short = prov.rstrip('省').rstrip('市').rstrip('自治区')
        if prov_short and len(prov_short) >= 2 and name.startswith(prov_short):
            name = name[len(prov_short):]
            
        if city and name.startswith(city):
            name = name[len(city):]
        city_short = city.rstrip('市').rstrip('地区').rstrip('州').rstrip('盟')
        if city_short and len(city_short) >= 2 and name.startswith(city_short):
            name = name[len(city_short):]
            
        return name.strip()

    def query_db(self, raw_name, clean, norm, city=''):
        """利用 SQLite 索引高速多级检索 (耗时约 0.02 毫秒)"""
        if not self.use_db or not self.conn:
            return None
        # 1. 精确匹配全称
        self.cursor.execute("SELECT grade FROM hospitals WHERE name = ? LIMIT 1", (raw_name,))
        r = self.cursor.fetchone()
        if r:
            return r[0], '超级知识库精确匹配'
            
        # 2. 简称匹配 (去括号)
        if clean and clean != raw_name:
            self.cursor.execute("SELECT grade FROM hospitals WHERE clean_name = ? LIMIT 1", (clean,))
            r = self.cursor.fetchone()
            if r:
                return r[0], '超级知识库简称匹配'
                
        # 3. 归一化匹配 (去省市前缀)
        if norm:
            if city:
                self.cursor.execute("SELECT grade FROM hospitals WHERE norm_name = ? AND (city = ? OR city = '') LIMIT 1", (norm, city))
                r = self.cursor.fetchone()
                if r:
                    return r[0], '超级知识库同城归一化匹配'
            self.cursor.execute("SELECT grade FROM hospitals WHERE norm_name = ? LIMIT 1", (norm,))
            r = self.cursor.fetchone()
            if r:
                return r[0], '超级知识库归一化匹配'
                
        return None

    def query_memory(self, raw_name, clean, norm):
        """内存字典查询"""
        if raw_name in self.cache:
            return self.cache[raw_name], '本地字典全称匹配'
        if clean in self.cache:
            return self.cache[clean], '本地字典简称匹配'
        if norm in self.cache:
            return self.cache[norm], '本地字典归一化匹配'
        return None

    def save_custom_cache(self, new_records):
        """保存新发现/新处理的医院等级到增量缓存，并同步写入 SQLite (若就绪)"""
        if not new_records:
            return
        self.custom_cache.update(new_records)
        try:
            with open(self.custom_cache_path, 'w', encoding='utf-8') as f:
                json.dump(self.custom_cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[警告] 保存自定义缓存失败: {e}")

        # 如果 SQLite 引擎在线，同步写入 SQLite 数据库以持久化
        if self.use_db and self.conn:
            try:
                for name, grade in new_records.items():
                    clean = re.sub(r'[\(（\[【].*?[\)）\]】]', '', name).strip()
                    norm = self.clean_name(name)
                    self.cursor.execute(
                        "INSERT OR REPLACE INTO hospitals (name, clean_name, norm_name, grade, province, city) VALUES (?, ?, ?, ?, '', '')",
                        (name, clean, norm, grade)
                    )
                self.conn.commit()
            except Exception as e:
                print(f"[警告] 同步写入 SQLite 失败: {e}")

    def search_online(self, hosp_name, prov='', city=''):
        """联网搜索医院等级（使用免封禁高速搜索引擎与公开医疗百科）"""
        clean = self.clean_name(hosp_name, prov, city)
        query = f"{prov}{city}{clean} 医院等级" if (prov or city) else f"{clean} 医院等级"
        
        url = f"https://www.so.com/s?q={urllib.parse.quote(query)}"
        try:
            r = requests.get(url, headers=self.headers, timeout=5)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'html.parser')
                for res in soup.find_all('li', class_='res-list')[:6]:
                    txt = res.get_text(separator=' ', strip=True)
                    m = re.search(r'(三级甲等|三级乙等|三级丙等|二级甲等|二级乙等|二级丙等|一级甲等|一级乙等|三甲|二甲|二乙|三乙|一甲|一级综合|二级综合|三级综合|未定级)', txt)
                    if m:
                        grade_token = m.group(1)
                        res_grade = self.normalize_grade_text(grade_token)
                        return res_grade, f"互联网搜索命中: '{m.group(1)}'"
        except Exception:
            pass

        return '无等级', '互联网搜索未发现明确等级评定'

    def get_grade(self, raw_name, prov='', city='', enable_online_search=True):
        """
        核心匹配函数：给定医院名称（及可选省市），返回 (grade, reason)
        """
        raw_name = str(raw_name).strip()
        if not raw_name:
            return '无等级', '名称为空'

        # 0. 增量自定义缓存查询 (优先级最高)
        if raw_name in self.custom_cache:
            return self.custom_cache[raw_name], '历史自学习缓存命中'

        clean = re.sub(r'[\(（\[【].*?[\)）\]】]', '', raw_name).strip()
        norm = self.clean_name(raw_name, prov, city)

        # -------------------------------------------------------------
        # 第 1 步: 本地 49.4 万超级知识库匹配 (0.02毫秒)
        # -------------------------------------------------------------
        db_res = None
        if self.use_db:
            db_res = self.query_db(raw_name, clean, norm, city)
        else:
            db_res = self.query_memory(raw_name, clean, norm)

        # 如果库中命中且是有等级的（如三甲、二甲、一级等），直接返回
        if db_res and db_res[0] != '无等级':
            return db_res

        # -------------------------------------------------------------
        # 第 2 步: 规则分流 (非医院机构 / 基层医疗机构) (0秒)
        # -------------------------------------------------------------
        # 2.1 非医院基层实体 (村卫生室、门诊、诊所、药房等) -> 坚决定为无等级
        for kw in NON_HOSPITAL_KEYWORDS:
            if kw in raw_name:
                return '无等级', f'非医院实体(含关键词: {kw})'

        # 2.2 基层卫生院/社区卫生服务中心
        for kw in PRIMARY_CARE_KEYWORDS:
            if kw in raw_name:
                for g_kw in ['二甲', '二级甲等', '二乙', '二级乙等', '一甲', '一级甲等']:
                    if g_kw in raw_name:
                        return self.normalize_grade_text(g_kw), f'基层医疗机构自标等级({g_kw})'
                return '一级医院', '基层卫生院/服务中心法定一级'

        # -------------------------------------------------------------
        # 第 3 步: 重点公立医院在线检索校验
        # -------------------------------------------------------------
        # 判断是否属于公立/重点专科医院（如果库中标为无等级，往往是历史数据缺漏，需联网复核）
        is_major_hosp = any(k in raw_name for k in MAJOR_HOSPITAL_KEYWORDS) or any(k in raw_name for k in ['医院', '保健院', '疗养院'])
        
        if enable_online_search and is_major_hosp:
            grade, reason = self.search_online(raw_name, prov, city)
            if grade != '无等级':
                self.save_custom_cache({raw_name: grade})
                return grade, reason
            # 如果网上搜索结果也确实无等级，但本地库有记录，返回本地库原因
            if db_res:
                return db_res
            return '无等级', '未评级或普通民营医院'

        # 如果本地库命中了无等级且不属于重点公立医院
        if db_res:
            return db_res

        return '无等级', '未评级机构'

    def exists_in_knowledge_base(self, raw_name, prov='', city=''):
        """
        判断某医院是否已存在于 49.4 万超级知识库中
        返回: (exists: bool, reason: str)
        """
        raw_name = str(raw_name).strip()
        if not raw_name:
            return False, '机构名称为空'

        # 1. 增量自定义缓存
        if raw_name in self.custom_cache:
            return True, '增量自学习知识库已收录'

        clean = re.sub(r'[\(（\[【].*?[\)）\]】]', '', raw_name).strip()
        norm = self.clean_name(raw_name, prov, city)

        # 2. SQLite 数据库检索
        if self.use_db and self.conn:
            self.cursor.execute("SELECT 1 FROM hospitals WHERE name = ? LIMIT 1", (raw_name,))
            if self.cursor.fetchone():
                return True, '49.4万知识库全称精确收录'

            if clean and clean != raw_name:
                self.cursor.execute("SELECT 1 FROM hospitals WHERE clean_name = ? LIMIT 1", (clean,))
                if self.cursor.fetchone():
                    return True, '49.4万知识库简称收录'

            if norm:
                if city:
                    self.cursor.execute("SELECT 1 FROM hospitals WHERE norm_name = ? AND (city = ? OR city = '') LIMIT 1", (norm, city))
                    if self.cursor.fetchone():
                        return True, '49.4万知识库同城归一化收录'
                self.cursor.execute("SELECT 1 FROM hospitals WHERE norm_name = ? LIMIT 1", (norm,))
                if self.cursor.fetchone():
                    return True, '49.4万知识库归一化收录'

        # 3. 内存字典检索
        else:
            if raw_name in self.cache:
                return True, '49.4万知识库全称精确收录'
            if clean in self.cache:
                return True, '49.4万知识库简称收录'
            if norm in self.cache:
                return True, '49.4万知识库归一化收录'

        return False, '知识库未收录'

    def get_saas_template_hospitals(self):
        """加载老 SaaS 模板已导入的 2.6 万机构名单缓存"""
        if hasattr(self, '_saas_hosp_cache') and self._saas_hosp_cache is not None:
            return self._saas_hosp_cache

        tpl_path = os.path.join(self.script_dir, "医院管理-批量导入模板.xlsx")
        names = set()
        clean_names = set()
        if os.path.exists(tpl_path):
            try:
                wb = openpyxl.load_workbook(tpl_path, read_only=True)
                ws = wb.active
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i == 0 or not row or not row[0]:
                        continue
                    raw = str(row[0]).strip()
                    names.add(raw)
                    clean = re.sub(r'[\(（\[【].*?[\)）\]】]', '', raw).strip()
                    if clean:
                        clean_names.add(clean)
                wb.close()
            except Exception as e:
                print(f"[警告] 加载老SaaS模板失败: {e}")
        self._saas_hosp_cache = (names, clean_names)
        return self._saas_hosp_cache

    def exists_in_saas_template(self, raw_name):
        """判断某医院是否已在老 SaaS 已导入历史模板中"""
        raw_name = str(raw_name).strip()
        if not raw_name:
            return False, '名称为空'
        names, clean_names = self.get_saas_template_hospitals()
        if raw_name in names:
            return True, '老SaaS历史模板全称已存在'
        clean = re.sub(r'[\(（\[【].*?[\)）\]】]', '', raw_name).strip()
        if clean and clean in clean_names:
            return True, '老SaaS历史模板简称已存在'
        return False, '老SaaS历史模板未收录'

    def clean_and_deduplicate(self, input_excel, output_cleaned_excel, output_excluded_excel, dedup_mode='49.4w', dedup_internal=True):
        """
        洗稿与增量去重：
        根据指定知识库基准剔除已有医院，并可选剔除表格内部同名重复行。
        返回清洗统计字典。
        """
        wb = openpyxl.load_workbook(input_excel)
        ws = wb.active

        header_row = [str(cell.value).strip() if cell.value is not None else '' for cell in ws[1]]

        name_col_idx = None
        prov_col_idx = None
        city_col_idx = None

        for idx, col_name in enumerate(header_row, start=1):
            if any(k in col_name for k in ['机构名称', '医院名称', '机构全称', '医院名']):
                name_col_idx = idx
            elif '省' in col_name:
                prov_col_idx = idx
            elif '市' in col_name:
                city_col_idx = idx

        if not name_col_idx:
            wb.close()
            raise ValueError("未在表头第1行找到'机构名称'或'医院名称'列")

        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        header = rows[0]
        data_rows = rows[1:]
        total_input = len(data_rows)

        seen_internal = set()
        cleaned_rows = []
        excluded_rows = []

        count_excluded_internal = 0
        count_excluded_kb = 0

        for row in data_rows:
            raw_name = str(row[name_col_idx - 1]).strip() if (name_col_idx - 1 < len(row) and row[name_col_idx - 1] is not None) else ''
            if not raw_name:
                continue

            prov = str(row[prov_col_idx - 1]).strip() if (prov_col_idx and prov_col_idx - 1 < len(row) and row[prov_col_idx - 1] is not None) else ''
            city = str(row[city_col_idx - 1]).strip() if (city_col_idx and city_col_idx - 1 < len(row) and row[city_col_idx - 1] is not None) else ''

            # 1. 表格内部自身去重检查
            if dedup_internal:
                clean_self = re.sub(r'[\(（\[【].*?[\)）\]】]', '', raw_name).strip()
                if raw_name in seen_internal or (clean_self and clean_self in seen_internal):
                    count_excluded_internal += 1
                    excluded_rows.append(list(row) + ['表格内部重复', f'与表格前序记录 [{raw_name}] 重复'])
                    continue
                seen_internal.add(raw_name)
                if clean_self:
                    seen_internal.add(clean_self)

            # 2. 知识库比对检查 (49.4万系统已有全量库洗稿剔除)
            is_dup = False
            dup_reason = ''

            exists, r = self.exists_in_knowledge_base(raw_name, prov, city)
            if exists:
                is_dup = True
                dup_reason = f'49.4万系统已有库已收录 ({r})'
            elif dedup_mode == 'saas_2.6w':
                exists_saas, r_saas = self.exists_in_saas_template(raw_name)
                if exists_saas:
                    is_dup = True
                    dup_reason = f'老SaaS历史模板已收录 ({r_saas})'

            if is_dup:
                count_excluded_kb += 1
                excluded_rows.append(list(row) + ['系统已有机构', dup_reason])
            else:
                cleaned_rows.append(row)

        # 保存洗稿后全新机构 Excel
        clean_wb = openpyxl.Workbook()
        clean_ws = clean_wb.active
        clean_ws.title = "洗稿后全新机构"
        clean_ws.append(header)
        for r in cleaned_rows:
            clean_ws.append(r)
        clean_wb.save(output_cleaned_excel)
        clean_wb.close()

        # 保存已剔除机构名单 Excel
        ex_wb = openpyxl.Workbook()
        ex_ws = ex_wb.active
        ex_ws.title = "已剔除系统已有机构名单"
        ex_header = list(header) + ['*剔除类型', '*判定依据与匹配说明']
        ex_ws.append(ex_header)
        for r in excluded_rows:
            ex_ws.append(r)
        ex_wb.save(output_excluded_excel)
        ex_wb.close()

        return {
            "total_input": total_input,
            "cleaned_count": len(cleaned_rows),
            "excluded_kb_count": count_excluded_kb,
            "excluded_internal_count": count_excluded_internal,
            "total_excluded": count_excluded_kb + count_excluded_internal,
            "output_cleaned_excel": output_cleaned_excel,
            "output_excluded_excel": output_excluded_excel
        }

    def process_excel(self, input_excel, output_excel=None, split_size=None, enable_online=True, progress_callback=None,
                      enable_dedup=False, dedup_mode='49.4w', dedup_internal=True):
        """批量处理 Excel 文件，支持洗稿剔除已有医院、自动补充等级并切分批次"""
        if not os.path.exists(input_excel):
            print(f"[错误] 文件不存在: {input_excel}")
            return None

        if not output_excel:
            base, ext = os.path.splitext(input_excel)
            output_excel = f"{base}_已补充等级{ext}"

        dedup_stats = None
        target_process_excel = input_excel

        # =============================================================
        # 步骤 0: 洗稿去重（直接剔除知识库已有机构与内部重复机构）
        # =============================================================
        if enable_dedup:
            print(f"正在启动洗稿去重流水线 (基准模式: {dedup_mode}, 表格内去重: {dedup_internal})...")
            dir_name = os.path.dirname(output_excel) or '.'
            cleaned_file = os.path.join(dir_name, "temp_洗稿过滤后全新机构.xlsx")
            excluded_file = os.path.join(dir_name, "已剔除已有机构名单.xlsx")
            
            dedup_stats = self.clean_and_deduplicate(
                input_excel=input_excel,
                output_cleaned_excel=cleaned_file,
                output_excluded_excel=excluded_file,
                dedup_mode=dedup_mode,
                dedup_internal=dedup_internal
            )
            print(f"洗稿完成: 原始 {dedup_stats['total_input']} 行，剔除已有 {dedup_stats['excluded_kb_count']} 行，内部重复 {dedup_stats['excluded_internal_count']} 行，保留新机构 {dedup_stats['cleaned_count']} 行")
            
            if dedup_stats['cleaned_count'] == 0:
                print("[提示] 洗稿过滤后无全新机构需要导入（全部均已在知识库中存在）！")
                return {
                    "total_rows": 0,
                    "grade_counts": {g: 0 for g in VALID_GRADES},
                    "output_excel": None,
                    "elapsed": 0.0,
                    "batch_files": [],
                    "split_dir": None,
                    "dedup_stats": dedup_stats
                }
            target_process_excel = cleaned_file

        print(f"正在读取待定级表格: {target_process_excel} ...")
        wb = openpyxl.load_workbook(target_process_excel)
        ws = wb.active

        header_row = [str(cell.value).strip() if cell.value is not None else '' for cell in ws[1]]

        name_col_idx = None
        prov_col_idx = None
        city_col_idx = None
        grade_col_idx = None

        for idx, col_name in enumerate(header_row, start=1):
            if any(k in col_name for k in ['机构名称', '医院名称', '机构全称', '医院名']):
                name_col_idx = idx
            elif '省' in col_name:
                prov_col_idx = idx
            elif '市' in col_name:
                city_col_idx = idx
            elif '等级' in col_name:
                grade_col_idx = idx

        if not name_col_idx:
            print("[错误] 未在第1行表头中找到'机构名称'或'医院名称'列！")
            return None
        if not grade_col_idx:
            grade_col_idx = len(header_row) + 1
            ws.cell(row=1, column=grade_col_idx, value="*医院等级")
            print(f"[提示] 未找到医院等级列，已自动在第 {grade_col_idx} 列追加 '*医院等级'")

        total_rows = ws.max_row - 1
        print(f"找到数据行数: {total_rows} 条，开始匹配等级 (数据库就绪: {'SQLite 引擎' if self.use_db else 'GZ/JSON 字典'})...")

        new_searched = {}
        grade_counts = {g: 0 for g in VALID_GRADES}
        t0 = time.time()

        for r in range(2, ws.max_row + 1):
            raw_name = ws.cell(row=r, column=name_col_idx).value or ''
            prov = ws.cell(row=r, column=prov_col_idx).value or '' if prov_col_idx else ''
            city = ws.cell(row=r, column=city_col_idx).value or '' if city_col_idx else ''

            grade, reason = self.get_grade(raw_name, prov, city, enable_online_search=enable_online)
            ws.cell(row=r, column=grade_col_idx, value=grade)
            grade_counts[grade] = grade_counts.get(grade, 0) + 1

            clean_str = str(raw_name).strip()
            if clean_str:
                if enable_dedup or '搜索' in reason:
                    new_searched[clean_str] = grade

            if progress_callback and (r % 50 == 0 or r == ws.max_row):
                progress_callback(r - 1, total_rows)

            if (r - 1) % 1000 == 0 or r == ws.max_row:
                speed = (r - 1) / max(time.time() - t0, 0.001)
                print(f"   已处理 {r - 1} / {total_rows} 条 ({(r-1)/total_rows*100:.1f}%) [速度: {speed:.0f} 条/秒] ...")

        if new_searched:
            self.save_custom_cache(new_searched)
            print(f"[提示] 本次已将 {len(new_searched)} 家新机构成功持久化沉淀入系统增量自学习知识库。")

        wb.save(output_excel)
        wb.close()
        t1 = time.time()
        print(f"\n[成功] 处理完成！耗时: {t1-t0:.2f} 秒。已保存至: {output_excel}")

        batch_files = []
        split_dir = None
        if split_size and split_size > 0:
            batch_files, split_dir = self.split_batches(output_excel, split_size)

        return {
            "total_rows": total_rows,
            "grade_counts": grade_counts,
            "output_excel": output_excel,
            "elapsed": round(t1 - t0, 2),
            "batch_files": batch_files,
            "split_dir": split_dir,
            "dedup_stats": dedup_stats
        }

    def split_batches(self, excel_path, batch_size=300):
        """按指定大小切分 Excel 文件，保留第 1 行表头和原模板所有下拉约束"""
        print(f"\n开始按每批 {batch_size} 条切分文件...")
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        
        split_dir = os.path.join(os.path.dirname(excel_path), f"分批导入_每批{batch_size}条")
        os.makedirs(split_dir, exist_ok=True)
        batch_files = []

        if not rows:
            return batch_files, split_dir
            
        header = rows[0]
        data_rows = rows[1:]
        total_count = len(data_rows)
        
        import math
        total_batches = math.ceil(total_count / batch_size)
        
        for b in range(total_batches):
            start_i = b * batch_size
            end_i = min((b + 1) * batch_size, total_count)
            batch_data = data_rows[start_i:end_i]
            
            sub_wb = openpyxl.Workbook()
            sub_ws = sub_wb.active
            sub_ws.title = ws.title
            
            sub_ws.append(header)
            for row in batch_data:
                sub_ws.append(row)
                
            dv_grade = DataValidation(
                type="list",
                formula1='"无等级,一级医院,一级乙等,一级甲等,二级医院,二级乙等,二级甲等,三级医院,三级乙等,三级甲等"',
                allow_blank=True
            )
            sub_ws.add_data_validation(dv_grade)
            dv_grade.add(f"G2:G{len(batch_data)+100}")

            dv_nature = DataValidation(
                type="list",
                formula1='"公立,私立"',
                allow_blank=True
            )
            sub_ws.add_data_validation(dv_nature)
            dv_nature.add(f"H2:H{len(batch_data)+100}")
            
            filename = f"医院管理-批量导入_第{b+1:02d}批({start_i+1:05d}-{end_i:05d}).xlsx"
            out_file = os.path.join(split_dir, filename)
            sub_wb.save(out_file)
            sub_wb.close()
            batch_files.append(out_file)
            
        print(f"[成功] 已成功切分为 {total_batches} 个文件，输出目录: {split_dir}")
        return batch_files, split_dir

def main():
    parser = argparse.ArgumentParser(description="医院等级智能查询与批量补全工具 (全国 49.4 万超级库版)")
    parser.add_argument("--query", "-q", type=str, help="单个医院名称查询")
    parser.add_argument("--prov", "-p", type=str, default="", help="所在省份 (可选)")
    parser.add_argument("--city", "-c", type=str, default="", help="所在城市 (可选)")
    parser.add_argument("--excel", "-e", type=str, help="待处理的 Excel 文件路径")
    parser.add_argument("--output", "-o", type=str, help="输出 Excel 文件路径 (可选)")
    parser.add_argument("--split", "-s", type=int, default=0, help="切分批次大小 (例如: 300)")
    parser.add_argument("--no-online", action="store_true", help="禁用在线联网搜索，仅使用本地库和规则")
    parser.add_argument("--interactive", "-i", action="store_true", help="进入交互式查询模式")

    args = parser.parse_args()
    matcher = HospitalGradeMatcher()

    if args.interactive:
        engine_str = "SQLite 索引引擎 (0 内存)" if matcher.use_db else "GZ 高速内存字典"
        print(f"\n=== 进入医院等级查询交互模式 [{engine_str} / 49.4 万条全国库] (输入 exit 退出) ===")
        while True:
            try:
                name = input("\n请输入医院名称: ").strip()
                if not name or name.lower() in ['exit', 'quit', 'q']:
                    break
                city = input("请输入所在城市 (可直接回车跳过): ").strip()
                t0 = time.time()
                grade, reason = matcher.get_grade(name, city=city, enable_online_search=not args.no_online)
                ms = (time.time() - t0) * 1000
                print(f"-> 匹配等级: 【{grade}】 (依据: {reason}, 耗时: {ms:.2f} ms)")
            except (KeyboardInterrupt, EOFError):
                break
        print("已退出。")
        return

    if args.query:
        t0 = time.time()
        grade, reason = matcher.get_grade(args.query, args.prov, args.city, enable_online_search=not args.no_online)
        ms = (time.time() - t0) * 1000
        print(f"\n医院名称: {args.query}")
        print(f"匹配等级: 【{grade}】")
        print(f"匹配依据: {reason}")
        print(f"查询耗时: {ms:.2f} ms\n")
        return

    if args.excel:
        matcher.process_excel(
            input_excel=args.excel,
            output_excel=args.output,
            split_size=args.split,
            enable_online=not args.no_online
        )
        return

    parser.print_help()

if __name__ == "__main__":
    main()
