# phase_gate.py - 桥接模块
# DT Skill本地scripts/指向valuation-common/scripts/phase_gate.py

import sys
import os
import importlib.util

_common_scripts = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    '..', '..', 'valuation-common', 'scripts'
))

_target = os.path.join(_common_scripts, 'phase_gate.py')
_spec = importlib.util.spec_from_file_location('_phase_gate_impl', _target)
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)

auto_gate_after_fill = _mod.auto_gate_after_fill
run_gate_check = _mod.run_gate_check

__all__ = ['auto_gate_after_fill', 'run_gate_check']
