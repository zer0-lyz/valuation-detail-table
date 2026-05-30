#!/usr/bin/env python3
"""
资产负债表校验工具
==================
用法:
    python validate_bs.py <标准化Excel路径> [源文件Excel路径]

校验内容:
    1. 项目分类正确性
    2. 借贷平衡 (资产 = 负债 + 权益)
    3. 摘要行交叉校对
    4. 源数据交叉校对 (如有)
    5. 漏项/错项检测
"""

import sys
import os
import pandas as pd
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def validate_balance_sheet(standardized_path, source_path=None):
    """执行资产负债表全量校验"""
    df = pd.read_excel(standardized_path)
    
    # 识别列名（兼容中英文）
    def col(name_cn, name_en):
        if name_cn in df.columns:
            return name_cn
        if name_en in df.columns:
            return name_en
        return None
    
    nc = col("项目名称", "item_name")
    cc = col("项目类别", "item_category")
    dc = col("方向", "item_direction")
    bc = col("期末余额", "closing_balance")
    oc = col("期初余额", "opening_balance")
    
    if not all([nc, cc, bc, oc]):
        print("❌ 无法识别表头，需要列: 项目名称/项目类别/期末余额/期初余额")
        return False
    
    lines = []
    def log(msg=""):
        print(msg)
        lines.append(msg)
    
    log("=" * 60)
    log(f"📋 资产负债表校验 — {Path(standardized_path).stem}")
    log("=" * 60)
    
    n_errors = 0
    
    # ---- 1. 分类检查 ----
    log("\n[1/5] 项目分类检查")
    for _, r in df.iterrows():
        name = str(r[nc]).strip()
        cat = str(r[cc]).strip() if pd.notna(r[cc]) else ""
        if '库存股' in name and '所有者权益' not in cat:
            log(f"  ❌ '{name}' 应为所有者权益类(当前:{cat})")
            n_errors += 1
    if n_errors == 0:
        log("  ✅ 无分类异常")
    
    # ---- 2. 平衡校验 ----
    log("\n[2/5] 借贷平衡校验")
    detail = df[df[nc].notna()].copy()
    detail = detail[~detail[nc].astype(str).str.contains('[:：]', na=False)]
    detail = detail[~detail[nc].astype(str).str.contains('合计|总计|小计', na=False)]
    
    for col, label in [(bc, "期末"), (oc, "期初")]:
        rows = detail.dropna(subset=[col])
        if rows.empty:
            continue
        assets = rows[rows[cc] == '资产类'][col].sum()
        liab = rows[rows[cc] == '负债类'][col].sum()
        equity = rows[rows[cc] == '所有者权益类'][col].sum()
        diff = abs(assets - (liab + equity))
        mark = "✅" if diff < 1 else "❌"
        log(f"  {mark} {label}: 资产={assets:>14,.2f}  负债={liab:>14,.2f}  权益={equity:>14,.2f}  差={diff:,.2f}")
        if diff >= 1:
            n_errors += 1
    
    # ---- 3. 摘要行校对 ----
    log("\n[3/5] 摘要行交叉校对")
    ta = df[df[nc].astype(str).str.strip() == '资产总计']
    tle = df[df[nc].astype(str).str.contains('负债和所有者权益|负债及所有者权益', na=False)]
    
    for col, label in [(bc, "期末"), (oc, "期初")]:
        if not ta.empty and not tle.empty:
            v1, v2 = ta.iloc[0][col], tle.iloc[0][col]
            if pd.notna(v1) and pd.notna(v2):
                d = abs(v1 - v2)
                log(f"  {'✅' if d<1 else '❌'} {label}: 资产总计={v1:>14,.2f} = 负债+权益总计={v2:>14,.2f}  差={d:,.2f}")
    
    # ---- 4. 源数据对比 ----
    if source_path and Path(source_path).exists():
        log(f"\n[4/5] 源数据交叉校对 — {Path(source_path).name}")
        try:
            src = pd.read_excel(source_path, sheet_name='资产负债表', header=None)
            for item_name in ['资产总计', '负债合计', '所有者权益（或股东权益)合计']:
                # 找左列
                for i in range(len(src)):
                    name = str(src.iloc[i, 0]).strip()
                    if name == item_name:
                        src_close = pd.to_numeric(src.iloc[i, 2], errors='coerce')
                        src_open = pd.to_numeric(src.iloc[i, 3], errors='coerce')
                        std_row = df[df[nc].astype(str).str.strip() == item_name]
                        if not std_row.empty:
                            r = std_row.iloc[0]
                            d_c = abs(src_close - r[bc]) if pd.notna(src_close) and pd.notna(r[bc]) else None
                            d_o = abs(src_open - r[oc]) if pd.notna(src_open) and pd.notna(r[oc]) else None
                            log(f"  {'✅' if d_c is not None and d_c<1 else '❌'} {item_name}(期末): 源={src_close:>14,.2f}  标准={r[bc]:>14,.2f}")
                            log(f"  {'✅' if d_o is not None and d_o<1 else '❌'} {item_name}(期初): 源={src_open:>14,.2f}  标准={r[oc]:>14,.2f}")
                    
                    # 找右列
                    if src.shape[1] > 4:
                        name_r = str(src.iloc[i, 4]).strip()
                        if name_r == item_name or (item_name == '负债合计' and name_r == '负债合计'):
                            src_close = pd.to_numeric(src.iloc[i, 6], errors='coerce')
                            src_open = pd.to_numeric(src.iloc[i, 7], errors='coerce')
                            std_row = df[df[nc].astype(str).str.strip() == name_r]
                            if not std_row.empty:
                                r = std_row.iloc[0]
                                d_c = abs(src_close - r[bc]) if pd.notna(src_close) and pd.notna(r[bc]) else None
                                d_o = abs(src_open - r[oc]) if pd.notna(src_open) and pd.notna(r[oc]) else None
                                log(f"  {'✅' if d_c is not None and d_c<1 else '❌'} {name_r}(期末): 源={src_close:>14,.2f}  标准={r[bc]:>14,.2f}")
                                log(f"  {'✅' if d_o is not None and d_o<1 else '❌'} {name_r}(期初): 源={src_open:>14,.2f}  标准={r[oc]:>14,.2f}")
        except Exception as e:
            log(f"  ⚠️ 源数据读取失败: {e}")
    
    # ---- 5. 漏项/错项 ----
    log("\n[5/5] 漏项/错项检测")
    if source_path and Path(source_path).exists():
        try:
            src = pd.read_excel(source_path, sheet_name='资产负债表', header=None)
            src_set = set()
            for i in range(5, len(src)):
                for ci in [0, 4]:
                    if ci < src.shape[1]:
                        n = str(src.iloc[i, ci]).strip()
                        if n and n not in ('nan', ''):
                            src_set.add(n)
            std_set = set(df[nc].dropna().astype(str).str.strip())
            missing = src_set - std_set
            extra = std_set - src_set
            log(f"  {'✅ 无漏项' if not missing else '⚠️ 缺少 '+str(len(missing))+'项'}")
            log(f"  {'✅ 无多余项' if not extra else '⚠️ 多出 '+str(len(extra))+'项'}")
            if missing:
                for m in sorted(missing)[:5]:
                    log(f"    - {m}")
        except Exception as e:
            log(f"  ⚠️  漏项检测失败: {e}")
    else:
        log("  ⏭️  跳过（未提供源文件）")
    
    status = "❌ 存在问题" if n_errors > 0 else "✅ 全部通过"
    log(f"\n{'='*60}\n[结果] {status} | {len(df)}行 {n_errors}个问题\n{'='*60}")
    
    # 保存报告
    report_path = Path(standardized_path).with_suffix(".校验报告.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    log(f"📄 校验报告已保存: {report_path}")
    
    return n_errors == 0


def main():
    if len(sys.argv) < 2:
        print("用法: python validate_bs.py <标准化Excel路径> [源文件路径]")
        print("示例: python validate_bs.py output/资产负债表_standardized.xlsx src.xlsx")
        sys.exit(1)
    validate_balance_sheet(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)


if __name__ == "__main__":
    main()
