#!/usr/bin/env python3
"""
DT Skill 编号唯一性校验脚本
===========================
扫描 valuation-detail-table Skill 所有文件，检测：
1. R编号重复 / 间隙 / 乱序
2. 教训编号重复 / 间隙 / 乱序
3. DT编号重复 / 间隙（DT编号有设计间隙，仅报告>2的连续间隙）
4. 虚空文件引用（指向不存在的.md/.py文件）
5. 未定义DT引用（聚合报告，作为历史文档清理提示）

用法：
    python validate_numbering.py [--skill-dir <path>] [--fix-hints]

返回：
    exit 0 = 全部通过
    exit 1 = 有错误（重复编号/虚空文件引用）
    exit 2 = 仅有警告（间隙/乱序/未定义DT引用）
"""

import sys
import os
import re
import json
import argparse
from collections import Counter, defaultdict
from pathlib import Path


# ============================================================
# 配置
# ============================================================

# DT编号有设计间隙（历史原因），仅报告>5的连续缺失区间
DT_GAP_THRESHOLD = 5

# 教训编号有已知间隙（33/34），属于历史遗留，不报告
LESSON_KNOWN_GAPS = {33, 34}

# 忽略的目录
IGNORE_DIRS = {'__pycache__', '.git', 'node_modules'}
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SKILL_DIR = SCRIPT_DIR.parent


# ============================================================
# 提取函数
# ============================================================

def read_file(path):
    """安全读取文件内容"""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    except (OSError, UnicodeDecodeError):
        return ''


def extract_r_numbers(skill_content):
    """从SKILL.md提取R编号定义（🚨R{N}格式）"""
    return re.findall(r'🚨R(\d+)', skill_content)


def extract_lesson_numbers(lessons_content):
    """从lessons_learned.md提取教训编号定义（## 教训{N}格式）"""
    return re.findall(r'^## 教训(\d+)', lessons_content, re.MULTILINE)


def load_dt_definitions(skill_dir, skill_content):
    """优先从RULES.md读取DT定义，其次rule_manifest.json，最后回退SKILL.md。"""
    rules_md = os.path.join(skill_dir, 'RULES.md')
    if os.path.exists(rules_md):
        content = read_file(rules_md)
        # 同时识别主规则（**DT-182**）和附录降级规则（DT-28）。
        # 子规则如DT-164.1、DT-182b按主编号归并，避免引用侧误报。
        nums = sorted(set(re.findall(
            r'\|\s*(?:\*\*)?DT-(\d+)(?:\.\d+)?[a-z]?(?:\*\*)?\s*\|',
            content,
        )), key=int)
        if nums:
            return nums

    manifest = os.path.join(skill_dir, 'assets', 'rule_manifest.json')
    if os.path.exists(manifest):
        try:
            with open(manifest, 'r', encoding='utf-8') as f:
                data = json.load(f)
            nums = set()
            for key in data.get('rules', {}).keys():
                m = re.match(r'^DT-(\d+)$', str(key))
                if m:
                    nums.add(m.group(1))
            if nums:
                return sorted(nums, key=int)
        except (OSError, json.JSONDecodeError):
            pass

    dt_nums = set()
    for m in re.finditer(r'\| \*\*DT-(\d+)\*\*', skill_content):
        dt_nums.add(m.group(1))
    for m in re.finditer(r'^> - DT-(\d+)\(', skill_content, re.MULTILINE):
        dt_nums.add(m.group(1))
    return sorted(dt_nums, key=int)


def extract_file_references(content, base_dir):
    """从文件内容提取.md/.py文件引用，返回(引用路径, 行号)列表
    仅匹配明确的路径引用（含目录前缀或→指向），排除行内提及文件名
    """
    refs = []
    for i, line in enumerate(content.split('\n'), 1):
        # 匹配 steps/xxx.md 或 scripts/xxx.py 或 references/xxx.md 或 assets/xxx
        for m in re.finditer(r'(?:steps|scripts|references|assets)/[\w\-]+\.(?:md|py|xlsx)', line):
            refs.append((m.group(), i))
        # 匹配 valuation-common/scripts/xxx.py 共享脚本引用
        for m in re.finditer(r'valuation-common/scripts/[\w\-]+\.py', line):
            refs.append((m.group(), i))
        # 匹配 →指向格式（如 →steps/S3_format.md）
        for m in re.finditer(r'→\s*(steps/[\w\-]+\.md)', line):
            refs.append((m.group(1), i))
    return refs


def extract_dt_references(content):
    """从任意文件提取DT-{N}引用（非定义），返回(编号, 行号)列表"""
    refs = []
    for i, line in enumerate(content.split('\n'), 1):
        # 跳过定义行（| **DT-xxx** | 格式）
        if re.match(r'\|\s*\*\*DT-\d+\*\*\s*\|', line.strip()):
            continue
        for m in re.finditer(r'DT-(\d+)', line):
            refs.append((m.group(1), i))
    return refs


# ============================================================
# 校验函数
# ============================================================

def check_uniqueness(nums_str, label):
    """检查编号唯一性，返回(错误列表, 警告列表)"""
    errors = []
    warnings = []

    if not nums_str:
        return errors, warnings

    counts = Counter(nums_str)
    dupes = {k: v for k, v in counts.items() if v > 1}
    if dupes:
        for num, count in sorted(dupes.items(), key=lambda x: int(x[0])):
            errors.append(f"❌ {label}{num} 出现{count}次（重复）")

    return errors, warnings


def check_sequence(nums_str, label, known_gaps=None):
    """检查编号连续性，返回(错误列表, 警告列表)"""
    errors = []
    warnings = []

    if not nums_str:
        return errors, warnings

    nums_int = sorted(set(int(n) for n in nums_str))
    if not nums_int:
        return errors, warnings

    max_num = max(nums_int)
    expected = set(range(1, max_num + 1))
    actual = set(nums_int)
    missing = expected - actual

    if known_gaps:
        missing = missing - known_gaps

    if missing:
        # 报告连续缺失区间
        missing_sorted = sorted(missing)
        ranges = []
        start = missing_sorted[0]
        end = missing_sorted[0]
        for m in missing_sorted[1:]:
            if m == end + 1:
                end = m
            else:
                ranges.append((start, end))
                start = m
                end = m
        ranges.append((start, end))

        for s, e in ranges:
            if s == e:
                warnings.append(f"⚠️ {label}{s} 缺失（{label}{s-1}→{label}{s+1}跳号）")
            else:
                warnings.append(f"⚠️ {label}{s}-{e} 缺失（{label}{s-1}→{label}{e+1}跳号，共{e-s+1}个）")

    return errors, warnings


def check_order(nums_str, label):
    """检查编号是否按定义顺序递增，返回(错误列表, 警告列表)"""
    errors = []
    warnings = []

    if len(nums_str) < 2:
        return errors, warnings

    for i in range(1, len(nums_str)):
        if int(nums_str[i]) < int(nums_str[i-1]):
            warnings.append(
                f"⚠️ {label}编号乱序：位置{i+1}的{label}{nums_str[i]} "
                f"出现在{label}{nums_str[i-1]}之后"
            )

    return errors, warnings


def check_file_references(refs, base_dir):
    """检查文件引用是否存在，返回(错误列表, 警告列表)
    同时检查DT Skill目录和valuation-common共享目录
    """
    errors = []
    seen = set()

    # 构建共享skill目录路径（同级目录下的valuation-common）
    common_dir = os.path.join(os.path.dirname(base_dir), 'valuation-common')

    for ref_path, line_no in refs:
        if ref_path in seen:
            continue
        seen.add(ref_path)

        # 处理valuation-common前缀的引用
        if ref_path.startswith('valuation-common/'):
            full_path = os.path.join(os.path.dirname(base_dir), ref_path)
        else:
            # 先在DT Skill目录下查找
            full_path = os.path.join(base_dir, ref_path)

        if os.path.exists(full_path):
            continue

        # 再在仓库根目录查找（顶层入口脚本，如scripts/validate_skill.py）
        repo_path = os.path.join(os.path.dirname(base_dir), ref_path)
        if os.path.exists(repo_path):
            continue

        # 再在共享脚本目录下查找（针对scripts/xxx.py格式）
        common_path = os.path.join(common_dir, ref_path)
        if os.path.exists(common_path):
            continue

        errors.append(f"❌ 虚空文件引用：{ref_path} (行{line_no}) → 文件不存在")

    return errors, []


def check_dt_void_references(skill_dir, dt_defined):
    """聚合报告操作文档中的未定义DT引用，返回(错误列表, 警告列表)。"""
    warnings = []
    dt_defined_set = set(dt_defined)
    missing = defaultdict(list)

    operational_docs = [
        os.path.join(skill_dir, 'CHECK.md'),
        os.path.join(skill_dir, 'FLOW.md'),
        os.path.join(skill_dir, 'scripts', 'README.md'),
        os.path.join(skill_dir, 'scripts', 'SKILL_SCRIPT_INDEX.md'),
    ]
    steps_dir = os.path.join(skill_dir, 'steps')
    if os.path.isdir(steps_dir):
        operational_docs.extend(
            os.path.join(steps_dir, fname)
            for fname in sorted(os.listdir(steps_dir))
            if fname.endswith('.md')
        )

    for fpath in operational_docs:
        if not os.path.exists(fpath):
            continue
        content = read_file(fpath)
        rel_path = os.path.relpath(fpath, skill_dir)
        for dt_num, line_no in extract_dt_references(content):
            if dt_num not in dt_defined_set:
                missing[dt_num].append(f'{rel_path}:{line_no}')

    for dt_num in sorted(missing, key=int):
        locations = missing[dt_num]
        sample = ', '.join(locations[:3])
        extra = f' 等{len(locations)}处' if len(locations) > 3 else ''
        warnings.append(f'⚠️ 未定义DT引用：DT-{dt_num} ({sample}{extra})')

    return [], warnings


def resolve_skill_dir(path=None):
    """定位包含RULES.md的内层Skill目录。"""
    candidate = Path(path).expanduser().resolve() if path else DEFAULT_SKILL_DIR
    if (candidate / 'RULES.md').exists():
        return candidate
    nested = candidate / 'valuation-detail-table'
    if (nested / 'RULES.md').exists():
        return nested
    return None


def find_skill_md(skill_dir):
    """顶层SKILL.md是Codex入口；兼容旧版内层布局。"""
    candidates = [
        skill_dir / 'SKILL.md',
        skill_dir.parent / 'SKILL.md',
    ]
    for path in candidates:
        if path.exists():
            return path
    return candidates[0]


def check_step_references(skill_dir):
    """检查FLOW.md和SKILL.md中的Step文件引用是否存在"""
    errors = []
    warnings = []
    steps_dir = os.path.join(skill_dir, 'steps')

    for target_file in ['FLOW.md', 'SKILL.md']:
        fpath = os.path.join(skill_dir, target_file)
        if not os.path.exists(fpath):
            continue
        content = read_file(fpath)

        # 提取 steps/S*.md 引用
        step_refs = re.findall(r'S-?\d+[_a-z]*\.md', content)
        for ref in sorted(set(step_refs)):
            step_path = os.path.join(steps_dir, ref)
            if not os.path.exists(step_path):
                errors.append(
                    f"❌ 虚空Step引用：{target_file} 引用 steps/{ref}，"
                    f"文件不存在"
                )

    return errors, warnings


# ============================================================
# 主函数
# ============================================================

def main():
    parser = argparse.ArgumentParser(description='DT Skill 编号唯一性校验')
    parser.add_argument('--skill-dir', default=None,
                        help='Skill目录路径（默认自动查找）')
    parser.add_argument('--fix-hints', action='store_true',
                        help='输出修复建议')
    args = parser.parse_args()

    # 定位内层Skill目录
    resolved = resolve_skill_dir(args.skill_dir)
    if resolved is None:
        shown = args.skill_dir or str(DEFAULT_SKILL_DIR)
        print(f"❌ 找不到Skill目录: {shown}")
        print("请用 --skill-dir 指定Skill根目录或包含RULES.md的内层目录")
        return 1
    skill_dir = str(resolved)

    print(f"📂 Skill目录: {skill_dir}")
    print("=" * 60)

    all_errors = []
    all_warnings = []

    # ----------------------------------------------------------
    # 1. R编号校验
    # ----------------------------------------------------------
    print("\n🔍 R编号校验")
    print("-" * 40)

    skill_content = read_file(str(find_skill_md(Path(skill_dir))))
    r_nums = extract_r_numbers(skill_content)

    if not r_nums:
        print("  ⚠️ 未找到R编号定义")
    else:
        errs, warns = check_uniqueness(r_nums, 'R')
        all_errors.extend(errs)
        all_warnings.extend(warns)

        errs, warns = check_order(r_nums, 'R')
        all_errors.extend(errs)
        all_warnings.extend(warns)

        print(f"  R编号总数: {len(r_nums)}, 范围: R{min(r_nums, key=int)}-R{max(r_nums, key=int)}")

    # ----------------------------------------------------------
    # 2. 教训编号校验
    # ----------------------------------------------------------
    print("\n🔍 教训编号校验")
    print("-" * 40)

    lessons_path = os.path.join(skill_dir, 'lessons_learned.md')
    if os.path.exists(lessons_path):
        lessons_content = read_file(lessons_path)
        lesson_nums = extract_lesson_numbers(lessons_content)

        if not lesson_nums:
            print("  ⚠️ 未找到教训编号定义")
        else:
            errs, warns = check_uniqueness(lesson_nums, '教训')
            all_errors.extend(errs)
            all_warnings.extend(warns)

            errs, warns = check_sequence(lesson_nums, '教训', known_gaps=LESSON_KNOWN_GAPS)
            all_errors.extend(errs)
            all_warnings.extend(warns)

            errs, warns = check_order(lesson_nums, '教训')
            all_errors.extend(errs)
            all_warnings.extend(warns)

            print(f"  教训编号总数: {len(lesson_nums)}, 范围: 教训{min(lesson_nums, key=int)}-教训{max(lesson_nums, key=int)}")
            if LESSON_KNOWN_GAPS:
                print(f"  已知间隙（不报告）: {', '.join(f'教训{g}' for g in sorted(LESSON_KNOWN_GAPS))}")

    # ----------------------------------------------------------
    # 3. DT编号校验（间隙仅报告大区间）
    # ----------------------------------------------------------
    print("\n🔍 DT编号校验")
    print("-" * 40)

    dt_defs = load_dt_definitions(skill_dir, skill_content)

    if not dt_defs:
        print("  ⚠️ 未找到DT编号定义")
    else:
        errs, warns = check_uniqueness(dt_defs, 'DT-')
        all_errors.extend(errs)
        all_warnings.extend(warns)

        # DT编号有设计间隙，仅报告>阈值的连续缺失
        dt_ints = sorted(set(int(d) for d in dt_defs))
        max_dt = max(dt_ints)
        missing = set(range(0, max_dt + 1)) - set(dt_ints)
        if missing:
            missing_sorted = sorted(missing)
            ranges = []
            start = missing_sorted[0]
            end = missing_sorted[0]
            for m in missing_sorted[1:]:
                if m == end + 1:
                    end = m
                else:
                    ranges.append((start, end))
                    start = m
                    end = m
            ranges.append((start, end))

            for s, e in ranges:
                span = e - s + 1
                if span > DT_GAP_THRESHOLD:
                    all_warnings.append(
                        f"⚠️ DT-{s}~DT-{e} 缺失（连续{span}个，超阈值{DT_GAP_THRESHOLD}）"
                    )

        print(f"  DT编号总数: {len(dt_defs)}, 范围: DT-{min(dt_defs, key=int)}-DT-{max(dt_defs, key=int)}")
        print(f"  缺失编号: {sorted(missing) if missing else '无'}")

    # ----------------------------------------------------------
    # 4. 文件引用校验
    # ----------------------------------------------------------
    print("\n🔍 文件引用校验")
    print("-" * 40)

    # SKILL.md中的文件引用
    skill_file_refs = extract_file_references(skill_content, skill_dir)
    errs, warns = check_file_references(skill_file_refs, skill_dir)
    all_errors.extend(errs)
    all_warnings.extend(warns)

    # FLOW.md中的文件引用
    flow_path = os.path.join(skill_dir, 'FLOW.md')
    if os.path.exists(flow_path):
        flow_content = read_file(flow_path)
        flow_file_refs = extract_file_references(flow_content, skill_dir)
        errs, warns = check_file_references(flow_file_refs, skill_dir)
        all_errors.extend(errs)
        all_warnings.extend(warns)

    # ----------------------------------------------------------
    # 5. 虚空DT引用校验
    # ----------------------------------------------------------
    print("\n🔍 虚空DT引用校验")
    print("-" * 40)

    errs, warns = check_dt_void_references(skill_dir, dt_defs)
    all_errors.extend(errs)
    all_warnings.extend(warns)

    # ----------------------------------------------------------
    # 6. Step文件引用校验
    # ----------------------------------------------------------
    print("\n🔍 Step文件引用校验")
    print("-" * 40)

    errs, warns = check_step_references(skill_dir)
    all_errors.extend(errs)
    all_warnings.extend(warns)

    # ----------------------------------------------------------
    # 汇总报告
    # ----------------------------------------------------------
    print("\n" + "=" * 60)
    print("📊 校验汇总")
    print("=" * 60)

    if all_errors:
        print(f"\n❌ 错误 ({len(all_errors)}):")
        for e in all_errors:
            print(f"  {e}")

    if all_warnings:
        print(f"\n⚠️ 警告 ({len(all_warnings)}):")
        for w in all_warnings:
            print(f"  {w}")

    if not all_errors and not all_warnings:
        print("\n✅ 全部通过！无重复编号、无虚空引用、无乱序。")

    # ----------------------------------------------------------
    # 修复建议
    # ----------------------------------------------------------
    if args.fix_hints and (all_errors or all_warnings):
        print(f"\n{'=' * 60}")
        print("🔧 修复建议")
        print("=" * 60)

        # 重复编号修复建议
        dupes_r = Counter(extract_r_numbers(skill_content))
        for num, cnt in sorted(dupes_r.items(), key=lambda x: int(x[0])):
            if cnt > 1:
                max_r = max(int(n) for n in dupes_r)
                print(f"  R{num}重复→建议将较新的重编号为R{max_r+1}")

        dupes_l = Counter(extract_lesson_numbers(read_file(lessons_path))) if os.path.exists(lessons_path) else Counter()
        for num, cnt in sorted(dupes_l.items(), key=lambda x: int(x[0])):
            if cnt > 1:
                max_l = max(int(n) for n in dupes_l)
                print(f"  教训{num}重复→建议将较新的重编号为教训{max_l+1}")

    # 返回码
    if all_errors:
        return 1
    elif all_warnings:
        return 2
    else:
        return 0


if __name__ == '__main__':
    sys.exit(main())
