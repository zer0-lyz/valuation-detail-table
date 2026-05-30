# scripts/__init__.py - DT Skill脚本包
# 路径桥接：将valuation-common/scripts/加入模块搜索路径
# 使Agent可以直接 from scripts.sheet_filler import fill_sheet

import sys
import os

# 将valuation-common/scripts/加入sys.path（优先级最高）
_common_scripts = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    '..', '..', 'valuation-common', 'scripts'
))
if _common_scripts not in sys.path:
    sys.path.insert(0, _common_scripts)
