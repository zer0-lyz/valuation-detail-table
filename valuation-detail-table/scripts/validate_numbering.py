#!/usr/bin/env python3
"""
DT Skill 编号唯一性校验脚本
===========================
扫描 valuation-detail-table Skill 所有文件，检测：
1. R编号重复 / 间隙 / 乱序
2. 教训编号重复 / 间隙 / 乱序
3. DT编号重复 / 间隙（DT编号有设计间隙，仅报告>2的连续间隙）
4. 虚空文件引用（指向不存在的.md/.py文件）
5. 虚空DT引用（Step文件/CHECK.md中引用了SKILL.md中不存在的DT编号）

用法：
    python validate_numbering.py [--skill-dir <path>] [--fix-hints]

返回：
    exit 0 = 全部通过
    exit 1 = 有错误（重复/虚空引用）
    exit 2 = 仅有警告（间隙/乱序）
"""

import sys
import os
import re
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


def extract_dt_definitions(skill_content):
    """从SKILL.md提取DT编号定义，包含两种格式：
    1. 完整定义行：| **DT-{N}** |
    2. 已迁移规则引用行：> - DT-{N}(...)
    """
    dt_nums = set()
    # 格式1：完整定义行
    for m in re.finditer(r'\| \*\*DT-(\d+)\*\*', skill_content):
        dt_nums.add(m.group(1))
    # 格式2：已迁移规则引用行
    for m in re.finditer(r'^> - DT-(\d+)\(', skill_content, re.MULTILINE):
        dt_nums.add(m.group(1))
    return list(dt_nums)


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

        # 再在共享脚本目录下查找（针对scripts/xxx.py格式）
        common_path = os.path.join(common_dir, ref_path)
        if os.path.exists(common_path):
            continue

        errors.append(f"❌ 虚空文件引用：{ref_path} (行{line_no}) → 文件不存在")

    return errors, []


def check_dt_void_references(skill_dir, dt_defined):
    """检查Step文件/CHECK.md中的DT引用是否都有定义，返回(错误列表, 警告列表)"""
    errors = []
    warnings = []
    dt_defined_set = set(dt_defined)

    # 扫描除SKILL.md外的所有.md文件
    for root, dirs, files in os.walk(skill_dir):
        dirs[:] = [d for d in dirs if d not in IGNORE_DIRS]
        for fname in files:
            if not fname.endswith('.md'):
                continue
            fpath = os.path.join(root, fname)
            # SKILL.md是定义文件，跳过
            if fname == 'SKILL.md':
                continue
            # CHANGELOG.md是历史记录，跳过（历史DT编号可能已被剥离/合并）
            if fname == 'CHANGELOG.md':
                continue

            content = read_file(fpath)
            dt_refs = extract_dt_references(content)
            rel_path = os.path.relpath(fpath, skill_dir)

            for dt_num, line_no in dt_refs:
                if dt_num not in dt_defined_set:
                    errors.append(
                        f"❌ 虚空DT引用：{rel_path}:{line_no} 引用 DT-{dt_num}，"
                        f"但SKILL.md中无此定义"
                    )

    return errors, warnings


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

    # 定位Skill目录
    if args.skill_dir:
        skill_dir = args.skill_dir
    else:
        # 默认路径
        default = os.path.join(
            os.path.expanduser('~'), '.workbuddy', 'skills',
            'valuation-detail-table'
        )
        if os.path.exists(default):
            skill_dir = default
        else:
            print(f"❌ 找不到Skill目录: {default}")
            print("请用 --skill-dir 指定路径")
            return 1

    print(f"📂 Skill目录: {skill_dir}")
    print("=" * 60)

    all_errors = []
    all_warnings = []

    # ----------------------------------------------------------
    # 1. R编号校验
    # ----------------------------------------------------------
    print("\n🔍 R编号校验")
    print("-" * 40)

    skill_content = read_file(os.path.join(skill_dir, 'SKILL.md'))
    r_nums = extract_r_numbers(skill_content)

    if not r_nums:
        print("  ⚠️ 未找到R编号定义")
    else:
        errs, warns = check_uniqueness(r_nums, 'R')
        all_errors.extend(errs)
        all_warnings.extend(warns)

        errs, warns = check_sequence(r_nums, 'R')
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

    dt_defs = extract_dt_definitions(skill_content)

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
