#!/usr/bin/env python3
"""
apply.py — 第二轮：按映射配置执行标准化

用法：
    python apply.py <映射配置JSON路径>
    
流程：
    1. 加载映射配置
    2. 读取源文件
    3. 按映射关系转换
    4. 数据质量校验
    5. 输出标准化 Excel + JSON
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import config as cfg, mapper, validator


def main():
    if len(sys.argv) < 2:
        print("用法: python apply.py <映射配置JSON路径>")
        print("       python apply.py --validate <映射配置JSON路径>  (含校验)")
        print("示例: python apply.py mapping_configs/客户A_科目余额表_20260529_mapping.json")
        sys.exit(1)

    # 解析参数
    run_validate = False
    if '--validate' in sys.argv:
        run_validate = True
        sys.argv.remove('--validate')
    
    config_path = sys.argv[1]

    # 加载映射配置
    try:
        mapping = cfg.load_config(config_path)
    except Exception as e:
        print(f"❌ 加载映射配置失败: {e}")
        sys.exit(1)

    # 检查是否已确认
    if not mapping.get("confirmed"):
        print("❌ 映射配置未确认！请检查并修改后设 'confirmed': true")
        print(f"   配置文件: {config_path}")
        sys.exit(1)

    print(f"📋 加载映射配置: {config_path}")
    cfg.print_config_summary(mapping)
    print()

    # 执行映射
    print("🔄 开始标准化...")
    try:
        df = mapper.apply(mapping)
    except Exception as e:
        print(f"❌ 标准化失败: {e}")
        sys.exit(1)

    # 校验
    msgs = validator.validate(mapping, df)
    print()
    for msg in msgs:
        print(f"   {msg}")
    print()

    # 保存
    try:
        output = mapper.apply_and_save(mapping)
    except Exception as e:
        print(f"❌ 保存失败: {e}")
        sys.exit(1)

    # 可选：完整校验
    if run_validate:
        print()
        print("=" * 60)
        print("🔍 执行完整校验...")
        print("=" * 60)
        try:
            from validate import validate as full_validate
            src_file = mapping.get("_source_file")
            doc_type = mapping.get("_detected_type")
            if src_file and os.path.exists(src_file):
                full_validate(output, src_file, doc_type)
            else:
                full_validate(output, doc_type=doc_type)
        except ImportError:
            print("  ⚠️ validate.py 未找到，跳过完整校验")
        except Exception as e:
            print(f"  ⚠️ 校验过程出错: {e}")

    print()
    print(f"🎉 完成！可以打开输出文件检查结果。")


if __name__ == "__main__":
    main()
