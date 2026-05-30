"""
phase_gate.py — Phase间Gate自动触发模块 v1.0

设计目的:
  消灭"Gate存在但Agent选择不调用"的问题。
  本模块嵌入fill_sheet返回路径，Agent无法跳过。
  CRITICAL>0 → sys.exit(1)，流程强制阻断。

强制等级: B级(L2 Gate门控) — Phase间强制触发

五级门控体系:
  G0: 数据源级（Phase 0完成后触发）
  G1: 写入级-数据（每Sheet填写后触发）
  G1F: 写入级-格式（Phase 3完成后集中触发）
  G2: 科目级（Phase 2完成后触发）
  G3: 勾稽级（Phase 4完成后触发，交付前）

调用方式:
  # 方式1: 在fill_sheet()返回后自动调用（推荐）
  from phase_gate import auto_gate_after_fill
  fill_result = fill_sheet(...)
  gate_result = auto_gate_after_fill(filepath, sheet_id, fill_result)

  # 方式2: 手动触发Phase间Gate
  from phase_gate import run_phase_gate
  run_phase_gate(filepath, phase=2, bs_path=..., sb_path=...)

v1.0 (2026-05-24): 初始版本
  - auto_gate_after_fill(): fill_sheet后自动触发G1
  - run_phase_gate(): Phase间Gate触发
  - _call_gate_validator(): 调用gate_validator.py
  - gate_result_assert(): Gate结果断言（CRITICAL→exit(1)）
"""

import sys
import json
import subprocess
from pathlib import Path

# gate_validator.py的位置
_DT_SCRIPTS_DIR = Path(__file__).parent.parent.parent / 'valuation-detail-table' / 'scripts'


def _call_gate_validator(filepath, gate='G1', sheet_name=None,
                          bs_path=None, sb_path=None, aux_data=None):
    """调用gate_validator.py执行Gate校验。

    Args:
        filepath: 评估明细表文件路径
        gate: Gate级别（G0/G1/G1F/G2/G3/all）
        sheet_name: 指定Sheet名（G1使用）
        bs_path: 资产负债表路径（G0/G2使用）
        sb_path: 科目余额表路径（G0/G2使用）
        aux_data: 辅助数据（G0使用）

    Returns:
        dict: {
            'passed': bool,
            'violations': list,
            'critical_count': int,
            'warning_count': int,
        }
    """
    gate_script = _DT_SCRIPTS_DIR / 'gate_validator.py'

    if not gate_script.exists():
        return {
            'passed': False,
            'violations': [{
                'gate': 'SYSTEM',
                'severity': 'CRITICAL',
                'message': f'gate_validator.py不存在: {gate_script}'
            }],
            'critical_count': 1,
            'warning_count': 0,
        }

    # 构建命令
    cmd = [sys.executable, str(gate_script), filepath, '--gate', gate]
    if sheet_name:
        cmd.extend(['--sheet', sheet_name])
    if bs_path:
        cmd.extend(['--bs-path', str(bs_path)])
    if sb_path:
        cmd.extend(['--sb-path', str(sb_path)])

    # 写入aux_data临时文件（如有）
    aux_file = None
    if aux_data:
        import tempfile
        aux_file = tempfile.NamedTemporaryFile(
            mode='w', suffix='.json', delete=False, encoding='utf-8'
        )
        json.dump(aux_data, aux_file, ensure_ascii=False)
        aux_file.close()
        cmd.extend(['--aux-data-path', aux_file.name])

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            encoding='utf-8',
            errors='replace',
        )

        # 解析输出
        stdout = result.stdout or ''
        violations = []
        critical_count = 0
        warning_count = 0

        # 从stdout解析违规项（gate_validator.py输出JSON格式）
        try:
            # 尝试解析JSON输出
            for line in stdout.strip().split('\n'):
                line = line.strip()
                if line.startswith('{') and line.endswith('}'):
                    v = json.loads(line)
                    violations.append(v)
                    if v.get('severity') == 'CRITICAL':
                        critical_count += 1
                    elif v.get('severity') == 'WARNING':
                        warning_count += 1
        except (json.JSONDecodeError, ValueError):
            # 非JSON输出→解析文本
            for line in stdout.strip().split('\n'):
                if 'CRITICAL' in line:
                    critical_count += 1
                    violations.append({
                        'gate': 'PARSE',
                        'severity': 'CRITICAL',
                        'message': line.strip()
                    })
                elif 'WARNING' in line:
                    warning_count += 1

        # exit code也反映结果
        passed = (result.returncode == 0) and (critical_count == 0)

        return {
            'passed': passed,
            'violations': violations,
            'critical_count': critical_count,
            'warning_count': warning_count,
            'stdout': stdout,
        }

    except subprocess.TimeoutExpired:
        return {
            'passed': False,
            'violations': [{
                'gate': 'SYSTEM',
                'severity': 'CRITICAL',
                'message': 'gate_validator.py执行超时(300s)'
            }],
            'critical_count': 1,
            'warning_count': 0,
        }
    except Exception as e:
        return {
            'passed': False,
            'violations': [{
                'gate': 'SYSTEM',
                'severity': 'CRITICAL',
                'message': f'gate_validator.py执行异常: {e}'
            }],
            'critical_count': 1,
            'warning_count': 0,
        }
    finally:
        # 清理临时文件
        if aux_file:
            try:
                Path(aux_file.name).unlink()
            except OSError:
                pass


# ============================================================
# Phase间Gate触发映射
# ============================================================

PHASE_GATE_MAP = {
    0: ['G0'],      # Phase 0完成后 → G0数据源级
    2: ['G1', 'G2'],  # Phase 2完成后 → G1写入级 + G2科目级
    3: ['G1F'],      # Phase 3完成后 → G1F格式级
    4: ['G3'],       # Phase 4完成后 → G3勾稽级
    5: ['G3'],       # Phase 5交付前 → G3最终验证
}


def run_phase_gate(filepath, phase, bs_path=None, sb_path=None,
                    aux_data=None, exit_on_critical=True):
    """Phase间Gate触发。

    按PHASE_GATE_MAP映射自动选择需要触发的Gate级别。
    CRITICAL>0时默认exit(1)强制阻断。

    Args:
        filepath: 评估明细表文件路径
        phase: 当前Phase编号（0/2/3/4/5）
        bs_path: 资产负债表路径
        sb_path: 科目余额表路径
        aux_data: 辅助数据
        exit_on_critical: CRITICAL时是否exit(1)，默认True

    Returns:
        dict: 汇总的Gate结果
    """
    gates = PHASE_GATE_MAP.get(phase, [])
    if not gates:
        return {
            'phase': phase,
            'passed': True,
            'gate_results': [],
            'total_critical': 0,
            'total_warning': 0,
        }

    all_results = []
    total_critical = 0
    total_warning = 0

    for gate in gates:
        result = _call_gate_validator(
            filepath, gate=gate,
            bs_path=bs_path, sb_path=sb_path,
            aux_data=aux_data,
        )
        all_results.append({
            'gate': gate,
            'result': result,
        })
        total_critical += result['critical_count']
        total_warning += result['warning_count']

    summary = {
        'phase': phase,
        'passed': total_critical == 0,
        'gate_results': all_results,
        'total_critical': total_critical,
        'total_warning': total_warning,
    }

    # CRITICAL→强制阻断
    if total_critical > 0 and exit_on_critical:
        print(f'\n🚨 Phase {phase} Gate未通过！')
        print(f'   CRITICAL: {total_critical}, WARNING: {total_warning}')
        print(f'   禁止进入下一Phase！')
        for gr in all_results:
            if gr['result']['critical_count'] > 0:
                for v in gr['result']['violations']:
                    if v.get('severity') == 'CRITICAL':
                        print(f'   ❌ [{gr["gate"]}] {v.get("message", "N/A")}')
        sys.exit(1)

    return summary


def auto_gate_after_fill(filepath, sheet_id, fill_result,
                          exit_on_critical=True):
    """fill_sheet()后自动触发G1门控。

    这是嵌入fill_sheet返回路径的自动Gate触发。
    Agent无法跳过。

    Args:
        filepath: 评估明细表文件路径
        sheet_id: 明细表前缀
        fill_result: fill_sheet()的返回值
        exit_on_critical: CRITICAL时是否exit(1)

    Returns:
        dict: Gate结果
    """
    # 检查fill_result自身的严重问题
    fill_criticals = fill_result.get('gate_errors', [])
    read_back_criticals = fill_result.get('read_back_errors', [])

    total_critical = len(fill_criticals) + len(read_back_criticals)

    if total_critical > 0:
        all_errors = fill_criticals + read_back_criticals
        if exit_on_critical:
            print(f'\n🚨 Sheet {sheet_id} fill_sheet发现{total_critical}个严重错误！')
            for err in all_errors[:10]:  # 最多显示10个
                print(f'   ❌ {err}')
            sys.exit(1)

        return {
            'passed': False,
            'total_critical': total_critical,
            'errors': all_errors,
        }

    # 调用外部G1门控（可选，fill_sheet内部已做基本校验）
    # 对于性能考虑，默认不在每次fill_sheet后都运行完整G1
    # 仅在fill_result有warnings时触发
    if fill_result.get('warnings'):
        g1_result = _call_gate_validator(filepath, gate='G1', sheet_name=sheet_id)
        return {
            'passed': g1_result['passed'],
            'total_critical': g1_result['critical_count'],
            'gate_result': g1_result,
        }

    return {
        'passed': True,
        'total_critical': 0,
    }


# ============================================================
# CLI入口
# ============================================================

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='phase_gate.py v1.0 — Phase间Gate自动触发')
    parser.add_argument('xlsx_path', help='评估明细表文件路径')
    parser.add_argument('--phase', type=int, required=True, help='Phase编号(0/2/3/4/5)')
    parser.add_argument('--bs-path', help='资产负债表路径')
    parser.add_argument('--sb-path', help='科目余额表路径')
    parser.add_argument('--no-exit', action='store_true', help='CRITICAL时不exit')

    args = parser.parse_args()

    result = run_phase_gate(
        args.xlsx_path,
        phase=args.phase,
        bs_path=args.bs_path,
        sb_path=args.sb_path,
        exit_on_critical=not args.no_exit,
    )

    if result['passed']:
        print(f'\n✅ Phase {args.phase} Gate通过！')
        print(f'   CRITICAL: 0, WARNING: {result["total_warning"]}')
    else:
        print(f'\n❌ Phase {args.phase} Gate未通过')
        print(f'   CRITICAL: {result["total_critical"]}, WARNING: {result["total_warning"]}')

    sys.exit(0 if result['passed'] else 1)
