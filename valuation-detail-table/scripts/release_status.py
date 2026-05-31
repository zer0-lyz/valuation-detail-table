"""release_status.py - 唯一发布状态判定。

判定规则:
  DELIVERED =
    source_validation.pass
    AND pending_confirmations == 0
    AND reconciliation.external_pass
    AND gate_summary.pass
    AND qa.pass
    AND formula_cache_status.pass

  DRAFT_REVIEW_REQUIRED =
    NOT DELIVERED
    AND (source_validation.pass OR has_project_exception)
"""

import json
import os
from datetime import datetime
from pathlib import Path


REQUIRED_GATES = {'G0', 'G1', 'G1F', 'G3', 'G-DT182'}


def _read_json(path: str, default=None):
    if not os.path.exists(path):
        return default
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _write_json(path: str, payload: dict):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def ensure_formula_cache_status(cache_dir: str) -> dict:
    """创建或读取公式缓存状态。未显式确认时必须阻断正式交付。"""
    status_path = os.path.join(cache_dir, 'formula_cache_status.json')
    status = _read_json(status_path)
    if status is None:
        template_path = Path(__file__).resolve().parent.parent / 'assets' / 'formula_cache_template.json'
        status = _read_json(str(template_path), {}) or {}
        status['generated_at'] = datetime.now().isoformat()
        _write_json(status_path, status)
    return status


def _summarize_gates(gate_results: list) -> tuple:
    """按每个 Gate 的最后一次结果汇总，缺少必需 Gate 时失败。"""
    gate_detail = {}
    for record in gate_results or []:
        result = record.get('result', record)
        gate = record.get('gate', result.get('gate', record.get('id', 'unknown')))
        gate_detail[gate] = result.get('status', 'unknown')
    missing = sorted(REQUIRED_GATES - set(gate_detail))
    gate_pass = not missing and all(status == 'passed' for status in gate_detail.values())
    return gate_pass, gate_detail, missing


def _count_pending(cache_dir: str) -> int:
    pending = _read_json(os.path.join(cache_dir, 'pending_confirmations.json'), {}) or {}
    return len([item for item in pending.get('items', []) if not item.get('resolved')])


def build_release_status(cache_dir: str, gate_results: list = None,
                         qa_result: dict = None, recon_result: dict = None,
                         pending_count: int = None,
                         formula_cache_status: dict = None,
                         dt139_status: dict = None) -> dict:
    """构建唯一发布状态。任何证据缺失均按失败处理。"""
    gate_pass, gate_detail, missing_gates = _summarize_gates(gate_results)

    qa_pass = bool((qa_result or {}).get('passed', False))

    recon_result = recon_result or {}
    recon_pass = bool(recon_result.get('external_pass', False))

    formula_cache_status = formula_cache_status or ensure_formula_cache_status(cache_dir)
    fc_pass = (
        formula_cache_status.get('status') == 'OK'
        and formula_cache_status.get('formula_chain_valid') is True
        and formula_cache_status.get('cache_values_valid') is True
        and formula_cache_status.get('blocking') is False
    )

    if pending_count is None:
        pending_count = _count_pending(cache_dir)
    confirm_ok = pending_count == 0

    dt139_status = dt139_status or _read_json(
        os.path.join(cache_dir, 'dt139_validation_status.json'), {}
    ) or {}
    source_ok = dt139_status.get('status') == 'PASS'
    has_dt139_exception = dt139_status.get('status') == 'EXCEPTION_DRAFT'

    is_deliverable = (source_ok and confirm_ok and recon_pass
                      and gate_pass and qa_pass and fc_pass
                      and not has_dt139_exception)

    if is_deliverable:
        status = 'DELIVERED'
    elif source_ok or has_dt139_exception:
        status = 'DRAFT_REVIEW_REQUIRED'
    else:
        status = 'FAILED'

    status_detail = {
        'source_validation': 'pass' if source_ok else 'fail',
        'pending_confirmations': 'clear' if confirm_ok else f'{pending_count} pending',
        'gate_summary': 'pass' if gate_pass else 'fail',
        'reconciliation': 'pass' if recon_pass else 'fail',
        'qa': 'pass' if qa_pass else 'fail',
        'formula_cache': 'pass' if fc_pass else 'fail',
        'dt139_exception': 'active' if has_dt139_exception else 'none',
        'missing_gates': missing_gates,
    }

    result = {
        'status': status,
        'is_deliverable': is_deliverable,
        'detail': status_detail,
        'gate_detail': gate_detail,
        'formula_cache_status': formula_cache_status,
        'dt139_validation_status': dt139_status,
        'generated_at': datetime.now().isoformat(),
    }

    _write_json(os.path.join(cache_dir, 'release_status.json'), result)
    return result


def load_release_status(cache_dir: str) -> dict:
    """从缓存加载发布状态"""
    path = os.path.join(cache_dir, 'release_status.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'status': 'UNKNOWN', 'is_deliverable': False}
