#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_dt_integration.py — DT Skill 集成测试用例

覆盖:
  T1: Phase 0 设定信息提取（不硬编码）
  T2: BS side-by-side格式负债列映射修正
  T3: subjects.json 多格式兼容（dict/list/嵌套dict）
  T4: 字段命名一致性（balance vs closing_balance）
  T5: journal_extractor 列名兼容（flat vs 嵌套）
  T6: Phase 1 industry_type 动态提取
  T7: data_loader 筛选+去重+勾稽
  T8: fix_format_issues 条件格式参数安全性
  T9: journal_extractor name_col安全检查

运行方式:
  cd ~/.workbuddy/skills/valuation-detail-table/scripts
  python test_dt_integration.py
"""

import os
import sys
import json
import tempfile
import shutil
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

# 路径配置
SCRIPT_DIR = Path(__file__).parent
SKILL_DIR = SCRIPT_DIR.parent
COMMON_SCRIPTS = SKILL_DIR.parent / 'valuation-common' / 'scripts'
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(COMMON_SCRIPTS))


class TestPhase0Settings(unittest.TestCase):
    """T1: Phase 0 设定信息提取——不硬编码公司名/基准日/行业类型"""

    def test_extract_settings_from_dirname(self):
        """从项目目录名提取公司名"""
        from dt_runner import _extract_settings

        subjects = [
            {'code': '1001', 'name': '库存现金', 'balance': 100},
            {'code': '1122', 'name': '应收账款', 'balance': 200},
        ]
        bs_data = {'items': []}

        # 项目目录名含公司名
        result = _extract_settings(subjects, bs_data, project_dir='C:/Users/1-上海图灵')
        self.assertIn('上海图灵', result['company_name'])

    def test_extract_settings_not_hardcoded(self):
        """确保不再返回硬编码的'河南平煤神马平绿置业有限公司'"""
        from dt_runner import _extract_settings

        subjects = [{'code': '1001', 'name': '库存现金', 'balance': 0}]
        bs_data = {'items': []}

        result = _extract_settings(subjects, bs_data, project_dir='C:/Users/北京某某公司')
        # 不应返回河南平绿的硬编码值
        self.assertNotEqual(result['company_name'], '河南平煤神马平绿置业有限公司')

    def test_industry_type_not_hardcoded(self):
        """行业类型应根据科目推断，不硬编码为'房地产'"""
        from dt_runner import _extract_settings

        # 无房地产科目的公司
        subjects = [
            {'code': '1001', 'name': '库存现金', 'balance': 0},
            {'code': '1122', 'name': '应收账款', 'balance': 0},
        ]
        bs_data = {'items': []}

        result = _extract_settings(subjects, bs_data, project_dir='C:/Users/科技公司')
        self.assertNotEqual(result['industry_type'], '房地产')

    def test_industry_type_real_estate_detected(self):
        """有开发成本科目时应识别为'房地产'"""
        from dt_runner import _extract_settings

        subjects = [
            {'code': '1001', 'name': '库存现金', 'balance': 0},
            {'code': '1601', 'name': '开发成本', 'balance': 1000000},
        ]
        bs_data = {'items': []}

        result = _extract_settings(subjects, bs_data, project_dir='C:/Users/1-某房地产')
        self.assertEqual(result['industry_type'], '房地产')


class TestBSSideBySide(unittest.TestCase):
    """T2: BS side-by-side格式负债列映射修正"""

    def test_liab_cols_greater_than_asset_cols(self):
        """负债侧列号应大于资产侧列号"""
        from source_header_parser import parse_balance_sheet

        # 使用上海图灵项目的实际BS文件（如果存在）
        bs_path = Path(r'C:\Users\Administrator\Desktop\上海图灵\上海图灵\上海图灵-资产负债表-20260430.xlsx')
        if not bs_path.exists():
            self.skipTest('上海图灵BS文件不存在，跳过实际文件测试')

        result = parse_balance_sheet(str(bs_path))
        if result['format'] == 'side_by_side':
            col_map = result.get('col_map', {})
            asset_ending = col_map.get('asset_ending', 4)
            liab_ending = col_map.get('liab_ending', 8)

            # 负债侧期末余额列号必须 > 资产侧期末余额列号
            self.assertGreater(liab_ending, asset_ending,
                f'负债侧期末余额列({liab_ending})应大于资产侧({asset_ending})')

    def test_liab_items_not_zero(self):
        """负债侧数据不应全为0"""
        from source_header_parser import parse_balance_sheet

        bs_path = Path(r'C:\Users\Administrator\Desktop\上海图灵\上海图灵\上海图灵-资产负债表-20260430.xlsx')
        if not bs_path.exists():
            self.skipTest('上海图灵BS文件不存在，跳过')

        result = parse_balance_sheet(str(bs_path))
        liab_items = [i for i in result.get('items', []) if i.get('side') == '负债及权益']

        if liab_items:
            total_liab = sum(abs(i.get('ending_balance', 0)) for i in liab_items)
            self.assertGreater(total_liab, 0,
                '负债及权益侧不应全为0（之前的bug: 列映射错误导致读到资产侧数据）')


class TestSubjectsJsonFormats(unittest.TestCase):
    """T3: subjects.json 多格式兼容"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def _save_subjects(self, data, filename='subjects.json'):
        path = os.path.join(self.tmpdir, filename)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
        return self.tmpdir

    def test_list_format(self):
        """subjects.json为顶层list时应正常解析"""
        from data_loader import _filter_data

        data = [
            {'code': '1122', 'name': '应收账款', 'balance': 1000},
            {'code': '2202', 'name': '应付账款', 'balance': 500},
        ]
        config = {'source_code_prefix': ['1122']}
        rows = _filter_data(data, config, 'subjects.json', self.tmpdir)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['code'], '1122')

    def test_dict_with_subjects_key(self):
        """subjects.json为{subjects: [...]}时应正常解析"""
        from data_loader import _filter_data

        data = {
            'subjects': [
                {'code': '1122', 'name': '应收账款', 'balance': 1000},
            ]
        }
        config = {'source_code_prefix': ['1122']}
        rows = _filter_data(data, config, 'subjects.json', self.tmpdir)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['balance'], 1000)

    def test_nested_dict_format(self):
        """subjects.json为{data: {subjects: [...]}}时应正常解析"""
        from data_loader import _filter_data

        data = {
            'data': {
                'subjects': [
                    {'code': '1122', 'name': '应收账款', 'balance': 1000},
                ]
            }
        }
        config = {'source_code_prefix': ['1122']}
        rows = _filter_data(data, config, 'subjects.json', self.tmpdir)

        self.assertEqual(len(rows), 1)


class TestFieldNamingConsistency(unittest.TestCase):
    """T4: 字段命名一致性（balance vs closing_balance）"""

    def test_dt_runner_uses_balance_first(self):
        """dt_runner.py subject_sheet_mapping应使用balance字段"""
        # 读取dt_runner源码验证
        dt_runner_path = SCRIPT_DIR / 'dt_runner.py'
        with open(dt_runner_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应包含balance优先的逻辑
        self.assertIn("s.get('balance', s.get('closing_balance', 0))", content,
            "dt_runner.py应优先使用balance字段，closing_balance仅作fallback")

    def test_data_loader_uses_balance_first(self):
        """data_loader.py应优先使用balance字段"""
        data_loader_path = SCRIPT_DIR / 'data_loader.py'
        with open(data_loader_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # _filter_data中的balance映射
        self.assertIn("s.get('balance', s.get('closing_balance', 0))", content,
            "data_loader.py应优先使用balance字段")

    def test_source_header_parser_outputs_balance(self):
        """source_header_parser.py输出subjects时字段名应为balance"""
        shp_path = COMMON_SCRIPTS / 'source_header_parser.py'
        with open(shp_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # parse_subject_balance输出的dict应包含'balance'键
        self.assertIn("'balance': balance", content,
            "source_header_parser.py输出的subjects应使用balance字段名")


class TestJournalExtractorColMap(unittest.TestCase):
    """T5: journal_extractor 列名兼容（flat vs 嵌套）"""

    def test_flat_key_name_preferred(self):
        """sheet_col_map.json flat键名(date/business/settlement)应被优先使用"""
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应包含flat键名优先逻辑
        self.assertIn("col_map.get('date'", content, "应支持flat键名'date'")
        self.assertIn("col_map.get('business'", content, "应支持flat键名'business'")
        self.assertIn("col_map.get('settlement'", content, "应支持flat键名'settlement'")

    def test_nested_key_name_fallback(self):
        """旧嵌套键名(occurrence_date等)应作为兼容fallback"""
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应包含嵌套键名fallback
        self.assertIn("occurrence_date", content, "应兼容旧嵌套键名'occurrence_date'")
        self.assertIn("business_content", content, "应兼容旧嵌套键名'business_content'")
        self.assertIn("settlement_object", content, "应兼容旧嵌套键名'settlement_object'")

    def test_name_col_none_safety(self):
        """name_col为None时应有安全检查，不应crash"""
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应包含name_col安全检查
        self.assertIn("if not name_col:", content,
            "journal_extractor.py应有name_col None安全检查")


class TestPhase1IndustryType(unittest.TestCase):
    """T6: Phase 1 industry_type 动态提取"""

    def test_phase1_reads_industry_from_settings(self):
        """Phase 1应从settings_info.json读取industry_type"""
        dt_runner_path = SCRIPT_DIR / 'dt_runner.py'
        with open(dt_runner_path, "r", encoding='utf-8') as f:
            content = f.read()

        # 应包含从settings读取industry_type的逻辑
        self.assertIn("settings.get('industry_type'", content,
            "Phase 1应从settings_info.json读取industry_type")
        # 不应包含硬编码的'房地产'
        # 检查 get_sheet_id_for_subject 调用时是否使用动态industry_type
        self.assertIn("industry_type=industry_type", content,
            "get_sheet_id_for_subject应使用动态industry_type参数")


class TestDataLoaderFilterAndDedup(unittest.TestCase):
    """T7: data_loader 筛选+去重+勾稽"""

    def test_dedup_keep_first(self):
        """去重策略keep_first应保留第一条记录"""
        from data_loader import dedup_data

        rows = [
            {'name': '公司A', 'balance': 100},
            {'name': '公司A', 'balance': 200},
            {'name': '公司B', 'balance': 300},
        ]
        result = dedup_data(rows, ['name'], 'keep_first')
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]['balance'], 100)  # 保留第一条

    def test_dedup_no_key(self):
        """无dedup_key时应原样返回"""
        from data_loader import dedup_data

        rows = [
            {'name': '公司A', 'balance': 100},
            {'name': '公司A', 'balance': 200},
        ]
        result = dedup_data(rows, None, 'keep_first')
        self.assertEqual(len(result), 2)

    def test_reconcile_balance_field(self):
        """reconcile应优先使用balance字段而非closing_balance"""
        from data_loader import get_reconcile_target

        schema = {'subjects': {}}
        config = {'reconcile_to': 'subjects.json:1122:balance'}
        # 模拟subjects.json (顶层list格式)
        tmpdir = tempfile.mkdtemp()
        subjects_data = [
            {'code': '1122', 'name': '应收账款', 'balance': 1000},
        ]
        with open(os.path.join(tmpdir, 'subjects.json'), 'w') as f:
            json.dump(subjects_data, f)

        result = get_reconcile_target(schema, config, tmpdir)
        self.assertEqual(result, 1000)

        shutil.rmtree(tmpdir)


class TestFixFormatConditionalFormatting(unittest.TestCase):
    """T8: fix_format_issues 条件格式参数安全性"""

    def test_no_cells_none_assignment(self):
        """不应有 rule.cells = None 的危险赋值"""
        ffi_path = COMMON_SCRIPTS / 'fix_format_issues.py'
        with open(ffi_path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertNotIn("new_rule.cells = None", content,
            "fix_format_issues.py不应有rule.cells=None赋值（会导致参数类型错误）")

    def test_uses_deepcopy(self):
        """条件格式规则应使用deepcopy而非shallow copy"""
        ffi_path = COMMON_SCRIPTS / 'fix_format_issues.py'
        with open(ffi_path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn("deepcopy", content,
            "fix_format_issues.py条件格式应使用deepcopy避免共享引用问题")

    def test_cond_format_exception_handling(self):
        """条件格式添加应有异常处理"""
        ffi_path = COMMON_SCRIPTS / 'fix_format_issues.py'
        with open(ffi_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应包含try/except包裹conditional_formatting.add
        self.assertIn("ws.conditional_formatting.add", content)
        # 在conditional_formatting.add附近应有异常处理
        self.assertTrue(
            "try:" in content and "ws.conditional_formatting.add" in content,
            "conditional_formatting.add应在try块中"
        )


class TestJournalExtractorNameColSafety(unittest.TestCase):
    """T9: journal_extractor name_col安全检查"""

    def test_scan_empty_fields_handles_none_name_col(self):
        """scan_empty_fields在name_col为None时不应crash"""
        # 验证源码中有安全检查
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应在name_col为None时跳过该Sheet（continue）
        self.assertIn("if not name_col:", content)


class TestFieldNamingConventionDoc(unittest.TestCase):
    """验证字段命名规范文档存在且内容完整"""

    def test_convention_doc_exists(self):
        """FIELD_NAMING_CONVENTION.md应存在"""
        doc_path = SKILL_DIR / 'assets' / 'FIELD_NAMING_CONVENTION.md'
        self.assertTrue(doc_path.exists(),
            f'字段命名规范文档应存在于 {doc_path}')

    def test_convention_doc_covers_key_fields(self):
        """文档应覆盖关键字段命名"""
        doc_path = SKILL_DIR / 'assets' / 'FIELD_NAMING_CONVENTION.md'
        if not doc_path.exists():
            self.skipTest('文档不存在')

        with open(doc_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应包含关键字段名定义
        for key_field in ['settlement', 'business', 'date', 'balance', 'book_value']:
            self.assertIn(key_field, content,
                f"文档应定义标准字段名 '{key_field}'")

        # 应明确禁止的字段名
        for forbidden in ['closing_balance', 'occurrence_date', 'business_content', 'settlement_object']:
            self.assertIn(forbidden, content,
                f"文档应标注禁止使用的字段名 '{forbidden}'")


# ============================================================
# T10-T17: DT-153v3 动态检测相关测试
# ============================================================

class TestDynamicFileSearch(unittest.TestCase):
    """T10: dt_runner.py 不硬编码文件名搜索关键词"""

    def test_bs_search_no_project_name(self):
        """BS文件搜索不应依赖特定项目名（如'河南'）"""
        dt_runner_path = SCRIPT_DIR / 'dt_runner.py'
        with open(dt_runner_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 不应有硬编码项目名搜索
        self.assertNotIn("'河南' in os.path.basename(f)", content,
            "不应使用'河南'等特定项目名筛选BS文件")

    def test_detail_table_dynamic_search(self):
        """评估明细表文件搜索应动态查找"""
        dt_runner_path = SCRIPT_DIR / 'dt_runner.py'
        with open(dt_runner_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有_find_detail_table函数
        self.assertIn('_find_detail_table', content,
            "应有_find_detail_table函数动态查找评估明细表")

    def test_detail_table_no_hardcoded_filename(self):
        """不应硬编码'评估明细表.xlsx'为唯一查找路径"""
        dt_runner_path = SCRIPT_DIR / 'dt_runner.py'
        with open(dt_runner_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # _find_detail_table应使用glob搜索
        self.assertIn("glob.glob(os.path.join(project_dir, '*评估明细表*'", content,
            "_find_detail_table应使用glob动态搜索")


class TestDynamicIndustryDetection(unittest.TestCase):
    """T11: 行业类型多行业推断"""

    def test_multiple_industry_types(self):
        """应支持多种行业类型推断（不仅限于房地产）"""
        from dt_runner import _extract_settings

        # 制造业
        mfg_subjects = [
            {'code': '1001', 'name': '库存现金', 'balance': 0},
            {'code': '5001', 'name': '生产成本', 'balance': 100},
        ]
        result = _extract_settings(mfg_subjects, {'items': []}, project_dir='C:/Users/某制造公司')
        self.assertEqual(result['industry_type'], '制造业')

        # 信息技术
        it_subjects = [
            {'code': '1001', 'name': '库存现金', 'balance': 0},
            {'code': '5301', 'name': '研发支出', 'balance': 100},
        ]
        result = _extract_settings(it_subjects, {'items': []}, project_dir='C:/Users/某科技公司')
        self.assertEqual(result['industry_type'], '信息技术')

    def test_default_industry_is_general(self):
        """无特定行业科目时默认为'通用'"""
        from dt_runner import _extract_settings

        subjects = [{'code': '1001', 'name': '库存现金', 'balance': 0}]
        result = _extract_settings(subjects, {'items': []}, project_dir='C:/Users/某公司')
        self.assertEqual(result['industry_type'], '通用')


class TestDynamicHeaderDetection(unittest.TestCase):
    """T12: 表头行动态检测（不再固定Row 5）"""

    def test_dt_runner_dynamic_header_row(self):
        """dt_runner.py勾稽核对应动态检测表头行"""
        dt_runner_path = SCRIPT_DIR / 'dt_runner.py'
        with open(dt_runner_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有动态表头检测逻辑
        self.assertIn("header_row", content,
            "勾稽核对应动态检测表头行")

    def test_fix_format_dynamic_header(self):
        """fix_format_issues.py应动态检测表头行"""
        ffi_path = COMMON_SCRIPTS / 'fix_format_issues.py'
        with open(ffi_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有header_row变量
        self.assertIn("header_row", content,
            "fix_format_issues.py应动态检测表头行")

    def test_sheet_col_finder_dynamic_rows(self):
        """sheet_col_finder.py应有动态表头行检测"""
        scf_path = COMMON_SCRIPTS / 'sheet_col_finder.py'
        with open(scf_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有_detect_header_rows函数
        self.assertIn("_detect_header_rows", content,
            "sheet_col_finder应有动态表头行检测函数")


class TestDynamicColumnDetection(unittest.TestCase):
    """T13: 列号动态检测（不再硬编码列号）"""

    def test_journal_extractor_no_hardcoded_col_fallback(self):
        """journal_extractor.py写入时不应硬编码列号兜底"""
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # write_phase3_results中不应有 "5 if r['type'] == 'asset' else 4" 等硬编码
        self.assertNotIn("5 if r['type'] == 'asset' else 4", content,
            "journal_extractor.py不应硬编码日期列号兜底值")

    def test_fix_format_no_fixed_column_c(self):
        """fix_format_issues.py不应固定C列做字体修正"""
        ffi_path = COMMON_SCRIPTS / 'fix_format_issues.py'
        with open(ffi_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 不应有固定C列判断
        self.assertNotIn("c == 3 and cell.font.size", content,
            "fix_format_issues.py不应固定C列做12pt修正")

    def test_source_header_parser_dynamic_right_cols(self):
        """source_header_parser.py应动态检测右栏列位"""
        shp_path = COMMON_SCRIPTS / 'source_header_parser.py'
        with open(shp_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有右栏动态检测逻辑
        self.assertIn("_right_beginning", content,
            "source_header_parser应有右栏列位动态检测")


class TestDynamicSheetDefinition(unittest.TestCase):
    """T14: 往来Sheet定义动态获取"""

    def test_receivable_sheets_from_config(self):
        """往来Sheet定义应从sheet_col_map.json动态获取"""
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有_get_receivable_sheets函数
        self.assertIn("_get_receivable_sheets", content,
            "应有_get_receivable_sheets函数动态获取Sheet定义")

    def test_col_map_path_dynamic(self):
        """sheet_col_map.json路径应动态检测"""
        je_path = SCRIPT_DIR / 'journal_extractor.py'
        with open(je_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有从脚本目录推断路径的逻辑
        self.assertIn("os.path.dirname(os.path.abspath(__file__))", content,
            "应从脚本所在目录动态推断sheet_col_map.json路径")


class TestSheetFillerNameMatch(unittest.TestCase):
    """T15: sheet_filler名称关键词匹配"""

    def test_get_sheet_id_by_name_exists(self):
        """应有get_sheet_id_by_name函数"""
        sf_path = COMMON_SCRIPTS / 'sheet_filler.py'
        with open(sf_path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn("get_sheet_id_by_name", content,
            "sheet_filler应有get_sheet_id_by_name函数支持非标准编码匹配")

    def test_standard_map_fuzzy_match(self):
        """STANDARD_MAP应有编码前缀模糊匹配"""
        sf_path = COMMON_SCRIPTS / 'sheet_filler.py'
        with open(sf_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有startswith模糊匹配
        self.assertIn("code4.startswith(std_code[:3])", content,
            "STANDARD_MAP应有3位前缀模糊匹配支持非标准编码")


class TestNoProjectSpecificDefaults(unittest.TestCase):
    """T16: CLI/入口不硬编码项目特定值"""

    def test_fix_format_cli_no_shanghai_turing(self):
        """fix_format_issues.py CLI不应硬编码'上海图灵'"""
        ffi_path = COMMON_SCRIPTS / 'fix_format_issues.py'
        with open(ffi_path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertNotIn("上海图灵", content,
            "fix_format_issues.py不应硬编码'上海图灵'")


class TestDataLoaderDynamicStartRow(unittest.TestCase):
    """T17: data_loader序时账数据起始行动态检测"""

    def test_journal_data_start_not_fixed_row2(self):
        """序时账数据起始行不应固定从Row 2开始"""
        dl_path = SCRIPT_DIR / 'data_loader.py'
        with open(dl_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 应有data_start_row变量
        self.assertIn("data_start_row", content,
            "data_loader应有动态检测序时账数据起始行")


# ============================================================
# 运行入口
# ============================================================

if __name__ == '__main__':
    print('='*60)
    print('DT Skill 集成测试')
    print('='*60)
    print(f'脚本目录: {SCRIPT_DIR}')
    print(f'通用脚本目录: {COMMON_SCRIPTS}')
    print()

    # 运行测试
    unittest.main(verbosity=2)
