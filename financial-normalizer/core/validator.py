"""
数据校验器
==========
对标准化后的数据进行质量检查。
"""

import pandas as pd


def check_trial_balance(df: pd.DataFrame) -> list:
    """检查科目余额表的借贷平衡关系"""
    msgs = []
    amount_cols = ["opening_debit", "opening_credit", "current_debit",
                   "current_credit", "closing_debit", "closing_credit"]
    has_amount = any(col in df.columns and df[col].notna().any() for col in amount_cols)
    if not has_amount:
        msgs.append("⚠️  未检测到金额数据列")
        return msgs

    if "opening_debit" in df.columns and "opening_credit" in df.columns:
        od_sum = df["opening_debit"].sum()
        oc_sum = df["opening_credit"].sum()
        if abs(od_sum - oc_sum) > 1 and (od_sum > 0 or oc_sum > 0):
            msgs.append(f"⚠️  期初不平衡: 借={od_sum:,.2f} 贷={oc_sum:,.2f} 差={abs(od_sum-oc_sum):,.2f}")
        else:
            msgs.append(f"✅ 期初平衡: 借={od_sum:,.2f} 贷={oc_sum:,.2f}")

    if "closing_debit" in df.columns and "closing_credit" in df.columns:
        cd_sum = df["closing_debit"].sum()
        cc_sum = df["closing_credit"].sum()
        if abs(cd_sum - cc_sum) > 1 and (cd_sum > 0 or cc_sum > 0):
            msgs.append(f"⚠️  期末不平衡: 借={cd_sum:,.2f} 贷={cc_sum:,.2f} 差={abs(cd_sum-cc_sum):,.2f}")
        else:
            msgs.append(f"✅ 期末平衡: 借={cd_sum:,.2f} 贷={cc_sum:,.2f}")

    return msgs


def check_generic(df: pd.DataFrame) -> list:
    """通用检查"""
    msgs = [f"📊 数据行数: {len(df)}"]
    empty_rows = df.isna().all(axis=1).sum()
    if empty_rows > 0:
        msgs.append(f"⚠️  全空行: {empty_rows}")
    return msgs


def validate(config: dict, df: pd.DataFrame) -> list:
    """综合校验"""
    doc_type = config.get("_detected_type", "trial_balance")
    msgs = ["--- 数据质量报告 ---"]
    msgs.extend(check_generic(df))
    if doc_type == "trial_balance":
        msgs.extend(check_trial_balance(df))
    return msgs
