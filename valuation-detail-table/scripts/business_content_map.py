# business_content_map.py - 桥接模块
# DT Skill本地scripts/指向valuation-common/scripts/business_content_map.py

import sys
import os
import importlib.util

_common_scripts = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    '..', '..', 'valuation-common', 'scripts'
))

_target = os.path.join(_common_scripts, 'business_content_map.py')
_spec = importlib.util.spec_from_file_location('_business_content_map_impl', _target)
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)

infer_business_content = _mod.infer_business_content

__all__ = ['infer_business_content']
