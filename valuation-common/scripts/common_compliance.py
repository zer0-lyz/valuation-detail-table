#!/usr/bin/env python3
"""
common_compliance.py — MR-18 CCEP合规验证脚本

功能：验证checkpoint.json中的common_compliance字段是否完整
逻辑：检查C0合规项、CCEP关卡项、零幻觉原则

用法：
  python common_compliance.py <checkpoint_dir> [phase_number]
  python common_compliance.py <checkpoint_dir> all

退出码：
  0 = 全部通过
  1 = 有合规项未通过
"""

import json
import sys
from pathlib import Path


def validate_common_compliance(checkpoint: dict, skill_type: str = "audit") -> dict:
    """
    验证单个checkpoint的common_compliance字段
    
    Args:
        checkpoint: checkpoint.json内容
        skill_type: "audit" 或 "preparation"
    
    Returns:
        验证结果dict
    """
    result = {
        "phase": checkpoint.get("phase", "?"),
        "phase_name": checkpoint.get("phase_name", ""),
        "compliance_items": [],
        "verdict": "pass",
        "blocking_reasons": []
    }
    
    # 1. 检查common_compliance字段存在性
    cc = checkpoint.get("common_compliance")
    if cc is None:
        result["compliance_items"].append({
            "id": "CCEP-FIELD",
            "desc": "common_compliance字段存在性",
            "status": "FAIL",
            "note": "checkpoint中缺少common_compliance字段（MR-18要求）"
        })
        result["verdict"] = "BLOCKED"
        result["blocking_reasons"].append("common_compliance字段缺失")
        return result
    
    # 2. 检查CCEP项
    ccep_items = cc.get("ccep_items", [])
    if not ccep_items:
        result["compliance_items"].append({
            "id": "CCEP-ITEMS",
            "desc": "CCEP关卡项",
            "status": "FAIL",
            "note": "common_compliance中无ccep_items（至少需要CCEP-1和CCEP-2）"
        })
        result["verdict"] = "BLOCKED"
        result["blocking_reasons"].append("CCEP关卡项缺失")
    
    for item in ccep_items:
        status = item.get("status", "unknown")
        result["compliance_items"].append({
            "id": item.get("id", "?"),
            "desc": item.get("desc", ""),
            "status": status,
            "note": item.get("reason", "")
        })
        if status != "pass":
            result["blocking_reasons"].append(
                f"{item.get('id', '?')}: {item.get('desc', '')} (status={status})"
            )
    
    # 3. 零幻觉原则检查：findings中每条问题应有单元格引用
    findings = checkpoint.get("findings", {})
    details = findings.get("details", [])
    zero_hallucination_pass = True
    for detail in details:
        # 检查是否包含单元格引用模式（如 A1, B2, SY4-F32 等）
        has_cell_ref = any(c in detail for c in ["行", "列", "Cell", "cell"])
        if not has_cell_ref:
            # 不阻断，仅警告
            zero_hallucination_pass = False
    
    result["compliance_items"].append({
        "id": "C0-5-ZH",
        "desc": "零幻觉原则：findings有单元格引用",
        "status": "pass" if zero_hallucination_pass else "warn",
        "note": "" if zero_hallucination_pass else "部分finding缺少单元格引用，请核实"
    })
    
    # 4. self_check_completed
    if not checkpoint.get("self_check_completed", False):
        result["compliance_items"].append({
            "id": "SELF-CHECK",
            "desc": "Phase自检完成",
            "status": "FAIL",
            "note": "self_check_completed=False"
        })
        result["blocking_reasons"].append("self_check_completed=False")
    
    # 5. 脚本执行记录
    scripts = checkpoint.get("scripts_executed", [])
    result["compliance_items"].append({
        "id": "SCRIPTS-EXEC",
        "desc": "已执行脚本记录",
        "status": "pass" if scripts else "warn",
        "note": f"已记录{len(scripts)}个脚本" if scripts else "无脚本执行记录"
    })
    
    # 汇判定
    if result["blocking_reasons"]:
        result["verdict"] = "BLOCKED"
    
    return result


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    checkpoint_dir = Path(sys.argv[1])
    phase_arg = sys.argv[2] if len(sys.argv) > 2 else "all"
    
    if not checkpoint_dir.exists():
        print(f"ERROR: 目录不存在: {checkpoint_dir}")
        sys.exit(1)
    
    # 收集checkpoint文件
    checkpoints = []
    if phase_arg == "all":
        for cp_file in sorted(checkpoint_dir.glob("phase*_checkpoint.json")):
            try:
                with open(cp_file, "r", encoding="utf-8") as f:
                    checkpoints.append(json.load(f))
            except (json.JSONDecodeError, IOError) as e:
                print(f"WARN: 无法读取 {cp_file.name}: {e}")
    else:
        try:
            phase_num = int(phase_arg)
            cp_file = checkpoint_dir / f"phase{phase_num}_checkpoint.json"
            if cp_file.exists():
                with open(cp_file, "r", encoding="utf-8") as f:
                    checkpoints.append(json.load(f))
            else:
                print(f"ERROR: {cp_file.name} 不存在")
                sys.exit(1)
        except ValueError:
            print(f"ERROR: 无效的Phase参数: {phase_arg}")
            sys.exit(1)
    
    # 验证
    all_pass = True
    for cp in checkpoints:
        result = validate_common_compliance(cp)
        
        phase = result["phase"]
        phase_name = result["phase_name"]
        print(f"\n{'='*50}")
        print(f"Phase {phase}: {phase_name}")
        print(f"{'='*50}")
        
        for item in result["compliance_items"]:
            icon = {"pass": "✅", "FAIL": "❌", "warn": "⚠️", "BLOCKED": "🚫"}.get(item["status"], "❓")
            print(f"  {icon} {item['id']}: {item['desc']}")
            if item.get("note"):
                print(f"     {item['note']}")
        
        print(f"\n  判定: {result['verdict']}")
        if result["blocking_reasons"]:
            print(f"  阻断原因:")
            for br in result["blocking_reasons"]:
                print(f"    - {br}")
        
        if result["verdict"] != "pass":
            all_pass = False
    
    print(f"\n{'='*50}")
    if all_pass:
        print("✅ Common合规检查全部通过")
        sys.exit(0)
    else:
        print("❌ Common合规检查未通过 — 需补完后再继续")
        sys.exit(1)


if __name__ == "__main__":
    main()
