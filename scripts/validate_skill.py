#!/usr/bin/env python3
"""Run lightweight release checks for the valuation-detail-table skill."""

from __future__ import annotations

import compileall
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
INNER = ROOT / 'valuation-detail-table'


def run(command: list[str], *, allow_warning: bool = False) -> bool:
    print(f"\n$ {' '.join(command)}")
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode == 0:
        return True
    if allow_warning and result.returncode == 2:
        print('  [WARN] 文档编号存在历史告警；未发现阻断错误。')
        return True
    print(f'  [FAIL] exit={result.returncode}')
    return False


def check_layout() -> bool:
    required = [
        ROOT / 'SKILL.md',
        ROOT / 'agents' / 'openai.yaml',
        INNER / 'RULES.md',
        INNER / 'FLOW.md',
        INNER / 'scripts' / 'dt_runner.py',
        INNER / 'scripts' / 'validate_numbering.py',
        INNER / 'steps' / 'S-1_5_normalize.md',
        INNER / 'steps' / 'S4_bs_verify.md',
    ]
    missing = [str(path.relative_to(ROOT)) for path in required if not path.exists()]
    if missing:
        print(f'[FAIL] 缺少必要文件: {missing}')
        return False
    print('[PASS] Skill目录结构完整')
    return True


def check_summary_rule_consistency() -> bool:
    targets = [
        ROOT / 'SKILL.md',
        INNER / 'RULES.md',
        INNER / 'FLOW.md',
        INNER / 'steps' / 'S4_bs_verify.md',
    ]
    stale_phrases = [
        'I列(硬编码)',
        'I 列 AI 直接录入',
        '逐行写入I列（硬编码数值）',
    ]
    failures = []
    for path in targets:
        content = path.read_text(encoding='utf-8')
        for phrase in stale_phrases:
            if phrase in content:
                failures.append(f'{path.relative_to(ROOT)}: {phrase}')
    if failures:
        print(f'[FAIL] 汇总表规则仍有历史硬编码路径: {failures}')
        return False
    print('[PASS] DT-182汇总表公式链说明一致')
    return True


def check_python_compile() -> bool:
    ok = True
    for path in [
        ROOT / 'valuation-common' / 'scripts',
        INNER / 'scripts',
        ROOT / 'financial-normalizer',
        ROOT / 'scripts',
    ]:
        ok = compileall.compile_dir(str(path), quiet=1) and ok
    print('[PASS] Python静态编译通过' if ok else '[FAIL] Python静态编译失败')
    return ok


def main() -> int:
    checks = [
        check_layout(),
        check_summary_rule_consistency(),
        check_python_compile(),
        run([
            sys.executable,
            str(INNER / 'scripts' / 'validate_numbering.py'),
        ], allow_warning=True),
        run([
            sys.executable,
            str(INNER / 'scripts' / 'test_dt_integration.py'),
            '-q',
        ]),
    ]
    if all(checks):
        print('\n[PASS] valuation-detail-table skill自检完成')
        return 0
    print('\n[FAIL] valuation-detail-table skill自检未通过')
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
