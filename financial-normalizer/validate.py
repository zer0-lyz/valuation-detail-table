#!/usr/bin/env python3
"""
标准化数据通用校验工具
======================
对所有类型的标准化输出执行质量检查。

用法:
    python validate.py <标准化Excel路径> [源文件路径]

支持类型（自动识别）:
    科目余额表, 序时账, 资产负债表, 利润表, 固定资产台账
"""

import sys
import os
import re
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def detect_type(df, doc_type_hint=None):
    """根据表头自动识别文档类型，也可通过 doc_type_hint 指定"""
    if doc_type_hint in ('trial_balance', 'journal', 'balance_sheet', 'income_statement', 'fixed_asset'):
        return doc_type_hint
    
    cols = set(str(c).strip() for c in df.columns)
    # 英文列名兼容
    en_cols = set()
    for c in cols:
        en_cols.add(c)
        # 中文→英文映射
        cn_en = {
            '科目编码': 'account_code', '科目名称': 'account_name',
            '期初借方余额': 'opening_debit', '期初贷方余额': 'opening_credit',
            '本期借方发生': 'current_debit', '本期贷方发生': 'current_credit',
            '期末借方余额': 'closing_debit', '期末贷方余额': 'closing_credit',
            '借方金额': 'debit_amount', '贷方金额': 'credit_amount',
            '日期': 'date', '凭证号': 'voucher_no', '摘要': 'summary',
            '项目名称': 'item_name', '项目类别': 'item_category',
            '期末余额': 'closing_balance', '期初余额': 'opening_balance',
            '本期金额': 'current_period', '本年累计': 'cumulative',
            '原值': 'original_value', '累计折旧': 'accumulated_depreciation',
            '净值': 'net_value', '资产名称': 'asset_name', '资产编码': 'asset_code',
        }
        en_cols.add(cn_en.get(c, c))
    
    # 用中英文混合集合判断
    
    # 强信号特征：只有特定类型才有的字段
    has_tb_signals = {'期初借方余额', '本期借方发生', '期末借方余额', '期初贷方余额', '期末贷方余额'}
    has_journal_signals = {'借方金额', '贷方金额', '凭证号'}
    has_bs_signals = {'项目类别', '期末余额', '期初余额'}
    has_is_signals = {'本期金额', '本年累计'}
    has_fa_signals = {'原值', '累计折旧', '净值'}
    has_common_subject = {'科目编码', '科目名称'}
    
    # 1. 科目余额表：特有借贷分列字段
    if has_tb_signals & cols:
        return "trial_balance"
    
    # 2. 序时账：借贷金额 + 凭证号
    if has_journal_signals & cols and has_common_subject & cols:
        return "journal"
    
    # 3. 资产负债表：项目类别 + 期末余额 + 期初余额（无本期金额）
    if has_bs_signals & cols and '本期金额' not in cols:
        return "balance_sheet"
    
    # 4. 利润表：本期金额 + 本年累计
    if has_is_signals & cols and '项目名称' in cols:
        return "income_statement"
    
    # 5. 固定资产台账
    if has_fa_signals & cols and ('资产名称' in cols or '资产编码' in cols):
        return "fixed_asset"
    
    # ---- fallback ----
    if '期初借方余额' in cols:
        return "trial_balance"
    if '借方金额' in cols and '贷方金额' in cols:
        return "journal"
    if '项目名称' in cols and '期末余额' in cols and '期初余额' in cols:
        if '本年累计' not in cols:
            return "balance_sheet"
    if '项目名称' in cols and '本期金额' in cols and '本年累计' in cols:
        return "income_statement"
    if '资产名称' in cols and '原值' in cols:
        return "fixed_asset"
    
    return None


def col(df, names):
    """找列名（兼容中英文）"""
    for n in names:
        if n in df.columns:
            return n
    return None


def validate_trial_balance(df, source_path):
    """科目余额表校验——只与源数据交叉校对，不做借贷平衡判定"""
    nc = col(df, ["科目名称", "account_name"])
    cc = col(df, ["科目编码", "account_code"])
    od = col(df, ["期初借方余额", "opening_debit"])
    oc = col(df, ["期初贷方余额", "opening_credit"])
    cd = col(df, ["本期借方发生", "current_debit"])
    cc2 = col(df, ["本期贷方发生", "current_credit"])
    ed = col(df, ["期末借方余额", "closing_debit"])
    ec = col(df, ["期末贷方余额", "closing_credit"])
    
    lines = []
    n_err = 0
    
    lines.append("[科目余额表校验]")
    
    if not nc or not cc:
        lines.append("  ❌ 缺少科目名称/科目编码列")
        return lines, 1
    
    # 结构检查
    n_codenull = df[cc].isna().sum()
    if n_codenull > 0:
        lines.append(f"  ⚠️ 科目编码空值: {n_codenull} 行")
    
    n_namenull = df[nc].isna().sum()
    if n_namenull > 0:
        lines.append(f"  ⚠️ 科目名称空值: {n_namenull} 行")
    
    # 重复编码检查
    dupes = df[cc].value_counts()
    dupes = dupes[dupes > 1]
    if len(dupes) > 0:
        lines.append(f"  ⚠️ 重复科目编码: {len(dupes)} 个")
        for code, cnt in dupes.head(5).items():
            names = df[df[cc] == code][nc].unique()
            lines.append(f"      {code} 出现{cnt}次 ({', '.join(str(n) for n in names)})")
    
    # 金额汇总（仅信息展示，不做平衡判定）
    lines.append(f"  📊 总行数: {len(df)}")
    if od and oc:
        lines.append(f"  💰 期初: 借={pd.to_numeric(df[od],errors='coerce').sum():>14,.2f}  "
                     f"贷={pd.to_numeric(df[oc],errors='coerce').sum():>14,.2f}")
    if cd and cc2:
        lines.append(f"  💰 本期: 借={pd.to_numeric(df[cd],errors='coerce').sum():>14,.2f}  "
                     f"贷={pd.to_numeric(df[cc2],errors='coerce').sum():>14,.2f}")
    if ed and ec:
        lines.append(f"  💰 期末: 借={pd.to_numeric(df[ed],errors='coerce').sum():>14,.2f}  "
                     f"贷={pd.to_numeric(df[ec],errors='coerce').sum():>14,.2f}")
    
    # 与源数据交叉校对
    if source_path and Path(source_path).exists():
        lines.append(f"\n  📁 源数据交叉校对 ({Path(source_path).name})")
        try:
            # 读源文件
            # 自动选择第一个有数据的 sheet
            xls = pd.ExcelFile(source_path)
            sheet_name = xls.sheet_names[0]
            src = pd.read_excel(source_path, sheet_name=sheet_name, header=None)
            src_data = src.iloc[6:].copy()
            src_data.columns = ['period', 'src_code', 'src_name', 'src_currency',
                                'src_od', 'src_oc', 'src_cd', 'src_cc', 'src_ed', 'src_ec']
            src_data = src_data[src_data['src_code'].notna() & (src_data['src_code'] != 'nan')]
            
            for c in ['src_od', 'src_oc', 'src_cd', 'src_cc', 'src_ed', 'src_ec']:
                src_data[c] = pd.to_numeric(src_data[c], errors='coerce')
            
            # 汇总金额对比
            pairs = [
                ("期初借方", "src_od", od),
                ("期初贷方", "src_oc", oc),
                ("本期借方", "src_cd", cd),
                ("本期贷方", "src_cc", cc2),
                ("期末借方", "src_ed", ed),
                ("期末贷方", "src_ec", ec),
            ]
            for label, src_col, std_col in pairs:
                if std_col:
                    src_sum = src_data[src_col].sum()
                    std_sum = pd.to_numeric(df[std_col], errors='coerce').sum()
                    diff = abs(src_sum - std_sum)
                    if diff < 1:
                        lines.append(f"    ✅ {label}: 源={src_sum:>14,.2f} = 标准={std_sum:>14,.2f}")
                    else:
                        lines.append(f"    ❌ {label}: 源={src_sum:>14,.2f} ≠ 标准={std_sum:>14,.2f}  差={diff:,.2f}")
                        n_err += 1
            
            # 科目数量对比
            src_codes = set(src_data['src_code'].dropna().astype(str).str.strip())
            std_codes = set(df[cc].dropna().astype(str).str.strip())
            missing = src_codes - std_codes
            extra = std_codes - src_codes
            if missing:
                lines.append(f"    ⚠️ 标准输出缺少科目({len(missing)}个): {sorted(missing)[:5]}...")
            else:
                lines.append(f"    ✅ 源数据科目全部覆盖 ({len(src_codes)}个)")
            if extra:
                lines.append(f"    ⚠️ 标准输出多余科目({len(extra)}个): {sorted(extra)[:5]}...")
            
            # 抽样核对前5个科目
            lines.append(f"    📝 抽样核对(前5个科目):")
            sample_codes = sorted(src_codes)[:5]
            for code in sample_codes:
                src_row = src_data[src_data['src_code'].astype(str).str.strip() == code].iloc[0]
                std_row = df[df[cc].astype(str).str.strip() == code]
                if not std_row.empty:
                    r = std_row.iloc[0]
                    src_name = str(src_row['src_name']).strip()
                    std_name = str(r[nc]).strip() if pd.notna(r[nc]) else ""
                    name_ok = "✅" if src_name == std_name else f"❌({src_name}→{std_name})"
                    lines.append(f"      {code:8s} {name_ok}")
            
        except Exception as e:
            lines.append(f"    ⚠️ 源数据读取失败: {e}")
    
    return lines, n_err


def validate_journal(df, source_path):
    """序时账校验"""
    dc = col(df, ["借方金额", "debit_amount"])
    cc = col(df, ["贷方金额", "credit_amount"])
    nc = col(df, ["科目名称", "account_name"])
    dt = col(df, ["日期", "date"])
    
    lines = []
    n_err = 0
    
    lines.append("[序时账校验]")
    
    # 借贷平衡
    if dc and cc:
        dr_sum = pd.to_numeric(df[dc], errors='coerce').sum()
        cr_sum = pd.to_numeric(df[cc], errors='coerce').sum()
        diff = abs(dr_sum - cr_sum)
        if diff < 1:
            lines.append(f"  ✅ 借贷平衡: 借={dr_sum:>14,.2f}  贷={cr_sum:>14,.2f}")
        else:
            lines.append(f"  ❌ 借贷不平衡: 借={dr_sum:>14,.2f}  贷={cr_sum:>14,.2f}  差={diff:,.2f}")
            n_err += 1
    
    # 日期空值
    if dt:
        n_null = df[dt].isna().sum()
        if n_null > 0:
            lines.append(f"  ⚠️ 日期空值: {n_null} 行")
        
        # 日期范围
        dates = pd.to_datetime(df[dt], errors='coerce')
        valid_dates = dates.dropna()
        if len(valid_dates) > 0:
            lines.append(f"  📅 日期范围: {valid_dates.min().date()} ~ {valid_dates.max().date()}")
    
    lines.append(f"  📊 总行数: {len(df)}")
    return lines, n_err


def validate_balance_sheet(df, source_path):
    """资产负债表校验"""
    nc = col(df, ["项目名称", "item_name"])
    cc = col(df, ["项目类别", "item_category"])
    bc = col(df, ["期末余额", "closing_balance"])
    oc = col(df, ["期初余额", "opening_balance"])
    
    lines = []
    n_err = 0
    
    lines.append("[资产负债表校验]")
    
    if not nc:
        lines.append("  ❌ 缺少项目名称列")
        return lines, 1
    
    # 分类检查
    for _, r in df.iterrows():
        name = str(r[nc]).strip()
        cat = str(r[cc]).strip() if cc and pd.notna(r[cc]) else ""
        if '库存股' in name and cat != '所有者权益类':
            lines.append(f"  ❌ '{name}' 分类应为所有者权益类(当前:{cat})")
            n_err += 1
    else:
        lines.append("  ✅ 项目分类无异常")
    
    # 平衡校验（明细行）
    if bc and cc:
        detail = df[df[nc].notna()].copy()
        detail = detail[~detail[nc].astype(str).str.contains('[:：]', na=False)]
        detail = detail[~detail[nc].astype(str).str.contains('合计|总计|小计', na=False)]
        
        for col_name, label in [(bc, "期末"), (oc, "期初")]:
            if col_name not in df.columns:
                continue
            rows = detail.dropna(subset=[col_name])
            if rows.empty:
                continue
            assets = rows[rows[cc] == '资产类'][col_name].sum()
            liab = rows[rows[cc] == '负债类'][col_name].sum()
            equity = rows[rows[cc] == '所有者权益类'][col_name].sum()
            diff = abs(assets - (liab + equity))
            mark = "✅" if diff < 1 else "❌"
            lines.append(f"  {mark} {label}: 资产={assets:>14,.2f}  负债={liab:>14,.2f}  权益={equity:>14,.2f}  差={diff:,.2f}")
            if diff >= 1:
                n_err += 1
    
    # 摘要行
    ta = df[df[nc].astype(str).str.strip() == '资产总计']
    tle = df[df[nc].astype(str).str.contains('负债和所有者权益|负债及所有者权益', na=False)]
    for col_name, label in [(bc, "期末"), (oc, "期初")]:
        if col_name and not ta.empty and not tle.empty:
            v1, v2 = ta.iloc[0][col_name], tle.iloc[0][col_name]
            if pd.notna(v1) and pd.notna(v2):
                d = abs(v1 - v2)
                lines.append(f"  {'✅' if d<1 else '❌'} {label}: 资产总计={v1:>14,.2f} = 负债+权益={v2:>14,.2f}  差={d:,.2f}")
    
    # 源数据对比
    if source_path and Path(source_path).exists():
        lines.append(f"\n  源数据交叉校对 ({Path(source_path).name})")
        try:
            src = pd.read_excel(source_path, sheet_name='资产负债表', header=None)
            for item in ['资产总计', '负债合计', '所有者权益（或股东权益)合计']:
                for ci in [0, 4]:
                    if ci >= src.shape[1]:
                        continue
                    for i in range(len(src)):
                        name = str(src.iloc[i, ci]).strip()
                        if name == item or (item == '所有者权益（或股东权益)合计' and '所有者权益' in name and '合计' in name):
                            src_close = pd.to_numeric(src.iloc[i, ci+2], errors='coerce') if ci+2 < src.shape[1] else None
                            src_open = pd.to_numeric(src.iloc[i, ci+3], errors='coerce') if ci+3 < src.shape[1] else None
                            std_row = df[df[nc].astype(str).str.strip() == name]
                            if not std_row.empty and bc and oc:
                                r = std_row.iloc[0]
                                d_c = abs(src_close - r[bc]) if pd.notna(src_close) and pd.notna(r[bc]) else None
                                d_o = abs(src_open - r[oc]) if pd.notna(src_open) and pd.notna(r[oc]) else None
                                if d_c is not None:
                                    lines.append(f"    {'✅' if d_c<1 else '❌'} {name}(期末): 源={src_close:>14,.2f}  标准={r[bc]:>14,.2f}")
                                if d_o is not None:
                                    lines.append(f"    {'✅' if d_o<1 else '❌'} {name}(期初): 源={src_open:>14,.2f}  标准={r[oc]:>14,.2f}")
        except Exception as e:
            lines.append(f"  ⚠️ 源数据读取失败: {e}")
    
    lines.append(f"  📊 总行数: {len(df)}")
    return lines, n_err


def validate_income_statement(df, source_path):
    """利润表校验"""
    nc = col(df, ["项目名称", "item_name"])
    pc = col(df, ["本期金额", "current_period"])
    cc = col(df, ["本年累计", "cumulative"])
    
    lines = []
    n_err = 0
    
    lines.append("[利润表校验]")
    
    if not nc:
        lines.append("  ❌ 缺少项目名称列")
        return lines, 1
    
    # 关键行提取
    def find_item(pattern):
        for _, r in df.iterrows():
            if re.search(pattern, str(r[nc]).strip()):
                return r
        return None
    
    total_profit = find_item(r'利润总额')
    income_tax = find_item(r'所得税费用')
    net_profit = find_item(r'^五、|^四、净利润|净利润')
    
    # 勾稽关系：利润总额 - 所得税 = 净利润
    for col_name, label in [(pc, "本期"), (cc, "本年累计")]:
        if col_name and total_profit is not None and income_tax is not None and net_profit is not None:
            tp_val = pd.to_numeric(total_profit[col_name], errors='coerce') if col_name in total_profit.index else None
            tax_val = pd.to_numeric(income_tax[col_name], errors='coerce') if col_name in income_tax.index else None
            np_val = pd.to_numeric(net_profit[col_name], errors='coerce') if col_name in net_profit.index else None
            
            if pd.notna(tp_val) and pd.notna(tax_val) and pd.notna(np_val):
                calc = tp_val - tax_val
                diff = abs(calc - np_val)
                lines.append(f"  {'✅' if diff<1 else '❌'} {label}: 利润总额({tp_val:,.2f}) - 所得税({tax_val:,.2f}) = {calc:,.2f}  = 净利润({np_val:,.2f})  差={diff:,.2f}")
                if diff >= 1:
                    n_err += 1
            elif pd.notna(tp_val) and pd.notna(np_val) and pd.isna(tax_val):
                # No tax line - profit = net profit
                diff = abs(tp_val - np_val)
                lines.append(f"  {'✅' if diff<1 else '❌'} {label}: 利润总额({tp_val:,.2f}) = 净利润({np_val:,.2f}) (无所得税)  差={diff:,.2f}")
                if diff >= 1:
                    n_err += 1
    
    # 营业收入 → 营业利润勾稽
    revenue = find_item(r'营业收入')
    op_profit = find_item(r'营业利润')
    if revenue is not None and op_profit is not None:
        for col_name, label in [(pc, "本期"), (cc, "本年累计")]:
            if col_name and col_name in revenue.index and col_name in op_profit.index:
                rev_val = pd.to_numeric(revenue[col_name], errors='coerce')
                op_val = pd.to_numeric(op_profit[col_name], errors='coerce')
                if pd.notna(rev_val) and pd.notna(op_val):
                    lines.append(f"  ℹ️  {label}: 营业收入={rev_val:>14,.2f}  营业利润={op_val:>14,.2f}  利润率={op_val/rev_val*100 if rev_val!=0 else 0:.2f}%")
    
    # 源数据对比
    if source_path and Path(source_path).exists():
        lines.append(f"\n  源数据交叉校对 ({Path(source_path).name})")
        try:
            src = pd.read_excel(source_path, sheet_name='利润表', header=None)
            for i in range(4, len(src)):
                src_name = str(src.iloc[i, 0]).strip()
                if src_name and src_name not in ('nan', '') and ('利润' in src_name or '营业' in src_name or '所得' in src_name):
                    src_period = pd.to_numeric(src.iloc[i, 2], errors='coerce')
                    src_cumul = pd.to_numeric(src.iloc[i, 3], errors='coerce')
                    std_row = df[df[nc].astype(str).str.strip() == src_name]
                    if not std_row.empty and pc and cc:
                        r = std_row.iloc[0]
                        d_p = abs(src_period - r[pc]) if pd.notna(src_period) and pd.notna(r[pc]) else None
                        d_c = abs(src_cumul - r[cc]) if pd.notna(src_cumul) and pd.notna(r[cc]) else None
                        if d_p is not None and d_c is not None:
                            lines.append(f"    {'✅' if d_p<1 and d_c<1 else '❌'} {src_name}: 本期={src_period:>12,.2f}→{r[pc]:>12,.2f}  累计={src_cumul:>12,.2f}→{r[cc]:>12,.2f}")
        except Exception as e:
            lines.append(f"  ⚠️ 源数据读取失败: {e}")
    
    # 非空行
    n_data = df.dropna(subset=[pc] if pc else []).shape[0] if pc else 0
    lines.append(f"  📊 总行数: {len(df)}  (有数据: {n_data})")
    return lines, n_err


def validate_fixed_asset(df, source_path):
    """固定资产台账校验"""
    ac = col(df, ["资产编码", "asset_code"])
    an = col(df, ["资产名称", "asset_name"])
    ov = col(df, ["原值", "original_value"])
    ad = col(df, ["累计折旧", "accumulated_depreciation"])
    ia = col(df, ["减值准备", "impairment_amount"])
    nv = col(df, ["净值", "net_value"])
    dp = col(df, ["折旧方法", "depreciation_method"])
    st = col(df, ["资产状态", "status"])
    
    lines = []
    n_err = 0
    
    lines.append("[固定资产台账校验]")
    
    if not an:
        lines.append("  ❌ 缺少资产名称列")
        return lines, 1
    
    # 勾稽：原值 - 累计折旧 - 减值 = 净值
    if ov and ad and nv:
        df_check = df.copy()
        df_check['_ov'] = pd.to_numeric(df_check[ov], errors='coerce')
        df_check['_ad'] = pd.to_numeric(df_check[ad], errors='coerce')
        df_check['_ia'] = pd.to_numeric(df_check[ia], errors='coerce').fillna(0) if ia else 0
        df_check['_nv'] = pd.to_numeric(df_check[nv], errors='coerce')
        
        df_check['_calc_nv'] = df_check['_ov'] - df_check['_ad'] - df_check['_ia']
        df_check['_diff'] = abs(df_check['_calc_nv'] - df_check['_nv'])
        
        bad = df_check[df_check['_diff'] > 1]
        n_bad = len(bad)
        
        if n_bad == 0:
            lines.append(f"  ✅ 原值-累计折旧-减值=净值: 全部{len(df_check)}行一致")
        else:
            lines.append(f"  ❌ 原值-累计折旧-减值≠净值: {n_bad}行不一致")
            for _, r in bad.head(5).iterrows():
                name = str(r[an]) if an else "?"
                lines.append(f"      {name}: 原值={r['_ov']:,.2f} - 折旧={r['_ad']:,.2f} - 减值={r['_ia']:,.2f} = {r['_calc_nv']:,.2f} ≠ 净值={r['_nv']:,.2f}")
            n_err += n_bad
    
    # 空值检查
    for field, label in [(ov, "原值"), (ad, "累计折旧")]:
        if field:
            n_null = df[field].isna().sum()
            if n_null > 0:
                lines.append(f"  ⚠️ {label}空值: {n_null}行")
    
    # 资产状态分布
    if st:
        status_counts = df[st].value_counts()
        lines.append(f"  📊 状态分布: {' | '.join(f'{k}={v}' for k,v in status_counts.items())}")
    
    # 汇总
    if ov:
        total_val = pd.to_numeric(df[ov], errors='coerce').sum()
        lines.append(f"  💰 原值合计: {total_val:,.2f}")
    if nv:
        total_nv = pd.to_numeric(df[nv], errors='coerce').sum()
        lines.append(f"  💰 净值合计: {total_nv:,.2f}")
    
    lines.append(f"  📊 总行数: {len(df)}")
    return lines, n_err


# ====== 调度入口 ======
VALIDATORS = {
    "trial_balance": validate_trial_balance,
    "journal": validate_journal,
    "balance_sheet": validate_balance_sheet,
    "income_statement": validate_income_statement,
    "fixed_asset": validate_fixed_asset,
}

TYPE_NAMES = {
    "trial_balance": "科目余额表",
    "journal": "序时账",
    "balance_sheet": "资产负债表",
    "income_statement": "利润表",
    "fixed_asset": "固定资产台账",
}


def validate(standardized_path, source_path=None, doc_type=None):
    """统一校验入口
    doc_type: 指定文档类型（可选），如不指定则自动识别
    """
    std_path = Path(standardized_path)
    if not std_path.exists():
        print(f"❌ 文件不存在: {standardized_path}")
        return False
    
    df = pd.read_excel(std_path)
    doc_type = detect_type(df, doc_type)
    
    if not doc_type:
        print(f"❌ 无法自动识别文档类型。支持的列名组合:")
        print(f"   科目余额表: 科目编码, 科目名称, 期初借方余额...")
        print(f"   序时账:     日期, 凭证号, 借方金额, 贷方金额...")
        print(f"   资产负债表: 项目名称, 项目类别, 期末余额, 期初余额")
        print(f"   利润表:     项目名称, 本期金额, 本年累计")
        print(f"   固定资产台账: 资产名称, 原值, 累计折旧, 净值")
        return False
    
    type_name = TYPE_NAMES.get(doc_type, doc_type)
    validator_fn = VALIDATORS.get(doc_type)
    
    print("=" * 60)
    print(f"📋 {type_name}校验 — {std_path.stem}")
    print("=" * 60)
    
    lines, n_errors = validator_fn(df, source_path)
    
    for l in lines:
        print(l)
    
    status = "❌ 存在问题" if n_errors > 0 else "✅ 全部通过"
    print(f"\n{'='*60}")
    print(f"[结果] {status} | {n_errors} 个问题")
    print(f"{'='*60}")
    
    # 保存报告
    report_path = std_path.with_name(f"{std_path.stem}.校验报告.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"校验类型: {type_name}\n")
        f.write(f"文件: {std_path}\n")
        f.write(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"结果: {status}\n\n")
        f.write("\n".join(lines))
    print(f"📄 校验报告: {report_path}")
    
    return n_errors == 0


def main():
    if len(sys.argv) < 2:
        print("用法: python validate.py <标准化Excel路径> [源文件路径]")
        print("")
        print("示例:")
        print("  python validate.py output/科目余额表_standardized.xlsx")
        print("  python validate.py output/资产负债表_standardized.xlsx 源文件.xlsx")
        print("  python validate.py output/*.xlsx   # 批量校验")
        print("  python validate.py output/*.xlsx 源文件.xlsx  # 批量+源文件")
        sys.exit(1)
    
    args = sys.argv[1:]
    
    # 解析源文件路径（最后一个参数，如果不含通配符且不是标准化文件）
    source_path = None
    std_args = list(args)
    
    if len(args) > 1:
        last = args[-1]
        # 如果最后一个参数不是标准输出文件名模式，视为源文件
        if '_standardized' not in last and '*' not in last and '?' not in last:
            if Path(last).exists():
                source_path = last
                std_args = args[:-1]
    
    # 展开通配符
    paths = []
    for arg in std_args:
        if '*' in arg or '?' in arg:
            import glob
            paths.extend(glob.glob(arg))
        else:
            paths.append(arg)
    
    # 过滤：只处理标准化文件（包含 _standardized 关键字）
    std_paths = [p for p in paths if '_standardized' in p and Path(p).exists()]
    
    if not std_paths:
        print("⚠️ 未找到标准化文件（文件名需包含 _standardized）")
        sys.exit(1)
    
    all_ok = True
    for p in std_paths:
        if not validate(p, source_path):
            all_ok = False
        print()
    
    if not all_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
