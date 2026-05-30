#!/usr/bin/env python3
"""
guess.py — 第一轮：自动检测 + 生成映射配置

用法：
    python guess.py <Excel文件路径>

输出：
    mapping_configs/<文件名>_<时间戳>_mapping.json

流程：
    1. 读取 Excel
    2. 自动检测文档类型、Sheet、表头行
    3. 匹配列名到标准字段
    4. 生成映射配置 JSON
    5. 展示配置摘要
    6. 提示用户编辑确认后运行 apply.py
"""

import sys
import os

# 确保核心模块可导入
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import detector, config as cfg


def main():
    if len(sys.argv) < 2:
        print("用法: python guess.py <Excel文件路径>")
        print("示例: python guess.py 数据/客户A_科目余额表.xlsx")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    print(f"🔍 开始检测: {excel_path}")
    print()
    
    try:
        mapping = detector.guess(excel_path)
    except Exception as e:
        print(f"❌ 检测失败: {e}")
        sys.exit(1)
    
    # 保存映射配置
    config_path = cfg.save_config(mapping)
    
    print()
    cfg.print_config_summary(mapping)
    print()
    print(f"💾 映射配置已保存到: {config_path}")
    print()
    print("📝 下一步:")
    print(f"   1. 打开 {config_path}")
    print(f"   2. 检查列映射是否正确，必要时修改")
    print(f"   3. 将 'confirmed' 改为 true")
    print(f"   4. 运行: python apply.py {config_path}")
    print()


if __name__ == "__main__":
    main()
