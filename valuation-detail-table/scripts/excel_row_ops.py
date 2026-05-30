# excel_row_ops.py - 桥接模块
# DT Skill本地scripts/指向valuation-common/scripts/excel_row_ops.py

import sys
import os
import importlib.util

_common_scripts = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    '..', '..', 'valuation-common', 'scripts'
))

_target = os.path.join(_common_scripts, 'excel_row_ops.py')
_spec = importlib.util.spec_from_file_location('_excel_row_ops_impl', _target)
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)

smart_insert_row = _mod.smart_insert_row
smart_delete_rows = _mod.smart_delete_rows
smart_insert_rows_for_data = _mod.smart_insert_rows_for_data
_apply_direct_format = _mod._apply_direct_format
_find_header_structure = _mod._find_header_structure
_find_total_row = _mod._find_total_row

__all__ = ['smart_insert_row', 'smart_delete_rows', 'smart_insert_rows_for_data',
           '_apply_direct_format', '_find_header_structure', '_find_total_row']
