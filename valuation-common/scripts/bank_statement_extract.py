# -*- coding: utf-8 -*-
"""
bank_statement_extract.py — 银行对账单PDF混合提取工具（v2.0）

设计理念（混合方案）：
  优先级1: pdfplumber文本提取 — 快速、零成本、数字精确
  优先级2: PDF转图片+多模态识别 — 兜底方案，覆盖扫描件/格式混乱PDF

  当pdfplumber提取为空或明显残缺时，自动切换到多模态识别。
  多模态识别由Agent调用（Agent读取图片后用视觉模型分析），
  本脚本负责：PDF→图片转换 + 生成多模态prompt + 后处理校验。

覆盖银行（6家实测）：
  - 建设银行(CCB): 账号/余额跨行拆分，需合并后提取或从文件名提取
  - 中国银行(BOC): 多种格式(BOCCA汇总/BOCVC明细/BOCCC贷款)
  - 工商银行(ICBC): 常见扫描件，pdfplumber完全失效
  - 交通银行(BOCOM): 标准格式，提取较简单
  - 中原银行: 标准格式，末行余额提取
  - 财务公司: "本月合计"行余额提取

踩坑经验详见: references/PDF_EXTRACTION_PITFALLS.md
多模态prompt详见: references/bank_statement_multimodal_prompt.md

依赖：
  - pdfplumber >= 0.11 — 文本PDF表格提取（首选）
  - pdf2image >= 1.17 — PDF转图片（多模态前置，需poppler）
  - Pillow — 图片处理

v2.0 (2026-05-23):
  - 重构为混合方案：pdfplumber → 多模态识别自动降级
  - 新增 pdf_to_images(): PDF转图片，供Agent多模态识别
  - 新增 generate_multimodal_prompt(): 生成银行对账单识别prompt
  - 新增 validate_balance(): 余额合理性校验
  - 新增 deduplicate_accounts(): 同账号多文件去重
  - 保留6家银行格式适配器，从v1项目实战经验提炼
  - 修复CCB余额提取bug（原extract_ccb_balance_from_text中未定义balance变量）
"""

import os
import sys
import re
import json
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple

# 延迟导入
_pdfplumber = None
_pdf2image = None


def _import_pdfplumber():
    global _pdfplumber
    if _pdfplumber is None:
        try:
            import pdfplumber
            _pdfplumber = pdfplumber
        except ImportError:
            pass
    return _pdfplumber


def _import_pdf2image():
    global _pdf2image
    if _pdf2image is None:
        try:
            from pdf2image import convert_from_path
            _pdf2image = convert_from_path
        except ImportError:
            pass
    return _pdf2image


# ============================================================
# 通用工具函数
# ============================================================

def clean_number(s: Any) -> Optional[float]:
    """将字符串解析为浮点数。支持千分位逗号、换行拆分、货币符号。
    
    关键修复: 建行PDF余额跨行拆分 "298,335.8\\n4" → 298335.84
    """
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip()
    # 先去除换行（处理建行余额拆行: "298,335.8\n4" → "298,335.84"）
    s = s.replace('\n', '').replace('\r', '')
    s = s.replace(',', '').replace('，', '').replace('￥', '').replace('¥', '')
    s = s.replace('元', '').replace('人民币', '').strip()
    if not s or s == '-':
        return None
    try:
        return float(s)
    except ValueError:
        return None


def identify_bank(text: str, filename: str = '') -> str:
    """从PDF文本或文件名识别银行名称。
    
    DT-213: 去除硬编码账号前缀(41050172)，改为通用银行名匹配。
    增加更多银行识别规则。
    
    Returns:
        银行名称: 'CCB'|'BOC'|'ICBC'|'BOCOM'|'ZY'|'FINCO'|'ABC'|'CMB'|'SPDB'|'UNKNOWN'
    """
    fname = filename.lower()
    source = text + ' ' + fname
    
    if '建设' in source or 'ccb' in fname or '建行' in source:
        return 'CCB'
    if '工商' in source or 'icbc' in fname or '工行' in source:
        return 'ICBC'
    if '中国银行' in source or 'boc' in fname or '中行' in source:
        return 'BOC'
    if '交通' in source or 'bocom' in fname or '交行' in source:
        return 'BOCOM'
    if '农业' in source or 'abc' in fname or '农行' in source:
        return 'ABC'
    if '招商' in source or 'cmb' in fname:
        return 'CMB'
    if '浦发' in source or 'spdb' in fname:
        return 'SPDB'
    if '中原' in source:
        return 'ZY'
    if '财务公司' in source:
        return 'FINCO'
    return 'UNKNOWN'


def identify_account_type(text: str) -> str:
    """判断账户类型：保证金/资金监管/活期/贷款"""
    if '保证金' in text:
        return '保证金'
    if '监管' in text:
        return '资金监管'
    if '贷款' in text or '同业' in text or 'LPR' in text:
        return '贷款'
    return '活期'


def extract_account_no_from_filename(filename: str) -> Optional[str]:
    """从文件名中提取账号（建行PDF最可靠的方式）。
    
    建行单账户PDF文件名即账号: "41050172560800001913.pdf"
    """
    basename = os.path.basename(filename)
    m = re.search(r'(\d{16,})', basename)
    if m:
        return m.group(1)
    return None


def validate_balance(balance: float, bank_name: str = '',
                     min_reasonable: float = -100_000_000,
                     max_reasonable: float = 10_000_000_000) -> Dict[str, Any]:
    """余额合理性校验。
    
    Args:
        balance: 提取到的余额
        bank_name: 银行名称（用于日志）
        min_reasonable: 最低合理余额（默认-1亿，允许透支账户）
        max_reasonable: 最高合理余额（默认100亿）
    
    Returns:
        {'valid': bool, 'warnings': list}
    """
    warnings = []
    if balance is None:
        return {'valid': False, 'warnings': ['余额为空']}
    
    if balance < min_reasonable:
        warnings.append(f'余额{balance:,.2f}异常偏低（低于{min_reasonable:,.0f}）')
    if balance > max_reasonable:
        warnings.append(f'余额{balance:,.2f}异常偏高（高于{max_reasonable:,.0f}）')
    if abs(balance) < 0.01 and balance != 0:
        warnings.append(f'余额接近零: {balance}')
    
    return {'valid': len(warnings) == 0, 'warnings': warnings}


def deduplicate_accounts(records: List[Dict]) -> List[Dict]:
    """同账号多文件去重。
    
    优先级:
    1. 来自单账户PDF（文件名含账号）> 综合对账单PDF
    2. 同优先级时，保留余额非None的
    3. 都有余额时，保留余额更大的（更完整）
    """
    seen = {}
    for r in records:
        key = r.get('account_no', '')
        if not key or key == 'UNKNOWN':
            # 无账号记录直接保留
            if '_no_acct' not in seen:
                seen['_no_acct_' + str(len(seen))] = r
            continue
            
        if key in seen:
            existing = seen[key]
            is_single = key in existing.get('source_file', '')
            new_is_single = key in r.get('source_file', '')
            
            if new_is_single and not is_single:
                seen[key] = r
            elif new_is_single and is_single:
                if r.get('balance') is not None and existing.get('balance') is not None:
                    if r['balance'] > existing['balance']:
                        seen[key] = r
                elif r.get('balance') is not None:
                    seen[key] = r
            elif not new_is_single and not is_single:
                if existing.get('type') == '贷款' and r.get('type') != '贷款':
                    seen[key] = r
        else:
            seen[key] = r
    
    return list(seen.values())


# ============================================================
# pdfplumber 文本提取 — 6家银行适配器
# ============================================================

def _extract_ccb(pdf_path: str) -> List[Dict]:
    """建设银行对账单提取。
    
    核心难点: 账号和余额跨行拆分
    - 账号 "41050172\\n56080000\\n1913" → 41050172560800001913
    - 余额 "298,335.8\\n4 人民币" → 298335.84
    
    对策:
    - 账号: 优先从文件名提取（最可靠）
    - 余额: 多策略逐级降级
    """
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            
            if not all_text.strip():
                return []
            
            is_margin = '保证金' in all_text
            acct_type = '保证金' if is_margin else '活期'
            
            # DT-213: 账号提取——通用方式，不依赖硬编码前缀41050172/56080
            # 策略1: 从文件名提取（最可靠）
            acct_no = extract_account_no_from_filename(pdf_path)
            # 策略2: 从合并文本中匹配16位以上连续数字串（排除日期/年份）
            if not acct_no:
                merged = all_text.replace('\n', ' ')
                for m in re.finditer(r'(\d{16,})', merged):
                    candidate = m.group(1)
                    # 排除纯年份/日期
                    if not re.match(r'^(19|20)\d{2}$', candidate[:4]):
                        acct_no = candidate
                        break
            # 策略3: 逐行拼接——查找连续短数字段组成的长账号
            # 建行PDF有时将账号拆成多行: "4105" → "0172" → "5608" → "00001913"
            if not acct_no:
                lines = all_text.split('\n')
                _digit_parts = []
                for line in lines:
                    s = line.strip()
                    m3 = re.match(r'^(\d{4})$', s)
                    if m3 and m3.group(1) not in ('2024', '2025', '2026'):
                        _digit_parts.append(m3.group(1))
                if len(_digit_parts) >= 3:
                    acct_no = ''.join(_digit_parts)
            
            # DT-213: 余额提取——通用方式，不依赖56080前缀
            balance = None
            
            # 策略1: 在含账号行中查找最后的金额（账号行常为: 账号 发生额 发生额 余额）
            if acct_no:
                _acct_short = acct_no[-8:] if len(acct_no) >= 8 else acct_no
                for line in all_text.split('\n'):
                    if _acct_short in line or acct_no in line:
                        # 提取该行所有金额
                        _amounts = re.findall(r'[\d,]+\.\d{2}', line)
                        if _amounts:
                            balance = clean_number(_amounts[-1])
                            break
            
            # 策略2: "X.XX 人民币" 完整模式
            if balance is None:
                matches = re.findall(r'([\d,]+\.\d+)\s*人民币', all_text)
                if matches:
                    balance = clean_number(matches[-1])
            
            # 策略3: "X,XXX 人民币" + ".XX 元" 拆分模式
            if balance is None:
                int_matches = list(re.finditer(r'([\d,]+)\s*人民币', all_text))
                if int_matches:
                    last = int_matches[-1]
                    after = all_text[last.end():last.end()+200]
                    dm = re.search(r'\.(\d{2})\s*\n?\s*元', after)
                    if dm:
                        balance = clean_number(last.group(1) + '.' + dm.group(1))
            
            # 策略4: "X.XX 元" 模式
            if balance is None:
                matches = re.findall(r'([\d,]+\.\d{2})\s*\n?\s*元', all_text)
                if matches:
                    balance = clean_number(matches[-1])
            
            if acct_no and balance is not None:
                results.append({
                    'bank_name': '建行',
                    'bank_full_name': '中国建设银行',
                    'account_no': acct_no,
                    'account_name': '',
                    'balance': balance,
                    'type': acct_type,
                    'source_file': os.path.basename(pdf_path),
                    'strategy': 'pdfplumber',
                })
    except Exception as e:
        results.append({
            'bank_name': '建行',
            'account_no': extract_account_no_from_filename(pdf_path) or 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'error': str(e),
        })
    return results


def _extract_boc(pdf_path: str) -> List[Dict]:
    """中国银行对账单提取。支持3种格式:
    
    1. BOCCA: 综合汇总表（第1页表格含多个账户）
    2. BOCVC: 明细对账单（含账号/账户类型/期末余额）
    3. BOCCC: 贷款明细（无期末余额，跳过）
    """
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    basename = os.path.basename(pdf_path).lower()
    results = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # BOCCC贷款明细跳过
            if basename.startswith('boccc'):
                return []
            
            # BOCCA综合汇总表
            if basename.startswith('bocca'):
                page = pdf.pages[0]
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        m = re.match(r'\|\s*(\d+)\|(.+?)\|(.+?)\|(.+?)\|(.+?)\|(.+?)\|(.+?)\|', line)
                        if m:
                            acct_info = m.group(2).strip()
                            product_type = m.group(4).strip()
                            balance_str = m.group(6).strip()
                            acct_name = m.group(7).strip()
                            balance = clean_number(balance_str)
                            if balance is not None:
                                results.append({
                                    'bank_name': '中行',
                                    'account_no': acct_info,
                                    'account_name': acct_name,
                                    'balance': balance,
                                    'type': identify_account_type(product_type),
                                    'source_file': os.path.basename(pdf_path),
                                    'strategy': 'pdfplumber',
                                })
                return results
            
            # BOCVC/其他中行明细
            for page in pdf.pages:
                text = page.extract_text()
                if not text or '账号' not in text:
                    continue
                
                acct_no = None
                acct_type_str = None
                acct_name = ''
                
                m = re.search(r'账号\s*(\d+)', text)
                if m:
                    acct_no = m.group(1).strip()
                
                m = re.search(r'账户类型\s*(.+?)\s+承前页', text)
                if m:
                    acct_type_str = m.group(1).strip()
                
                m = re.search(r'账户名称\s*(.+?)\s+开户行', text)
                if m:
                    acct_name = m.group(1).strip()

                # DT-216: 提取开户行全称
                bank_full_name = '中国银行'  # 默认全称
                m = re.search(r'开户行\s*(.+?)\s+起始日期', text)
                if m:
                    bank_full_name = m.group(1).strip()

                balance = None
                m = re.search(r'本对账期末余额\s*([\d,]+\.\d+)', text)
                if m:
                    balance = clean_number(m.group(1))
                if balance is None:
                    m = re.search(r'本页余额\s*([\d,]+\.\d+)', text)
                    if m:
                        balance = clean_number(m.group(1))
                
                if acct_no and balance is not None:
                    results.append({
                        'bank_name': '中行',
                        'bank_full_name': bank_full_name,
                        'account_no': acct_no,
                        'account_name': acct_name,
                        'balance': balance,
                        'type': identify_account_type(acct_type_str or ''),
                        'source_file': os.path.basename(pdf_path),
                        'strategy': 'pdfplumber',
                    })
    except Exception as e:
        results.append({
            'bank_name': '中行',
            'account_no': 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'error': str(e),
        })
    return results


def _extract_icbc(pdf_path: str) -> List[Dict]:
    """工商银行对账单提取。
    
    注意: 工行对账单常见扫描件，pdfplumber提取文本为空。
    此函数标记为需多模态兜底，但尝试提取。
    """
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        all_text += ' '.join([str(c) for c in row if c]) + "\n"
            
            if not all_text.strip():
                # 扫描件：[DT-132] MUST使用Agent多模态Read兜底，禁止跳过
                results.append({
                    'bank_name': '工行',
                    'account_no': 'UNKNOWN',
                    'account_name': '',
                    'balance': None,
                    'type': '活期',
                    'source_file': os.path.basename(pdf_path),
                    'strategy': 'pdfplumber',
                    'needs_multimodal': True,
                    'needs_multimodal_read': True,  # [DT-132] 标记
                    'error': '扫描件PDF，pdfplumber无法提取文本，MUST使用Read工具多模态识别',
                })
                return results
            
            m = re.search(r'账号[：:]\s*(\d+)', all_text)
            acct_no = m.group(1) if m else 'UNKNOWN'
            
            balance = None
            m = re.search(r'余额[：:]\s*([\d,]+\.\d{2})', all_text)
            if m:
                balance = clean_number(m.group(1))
            
            results.append({
                'bank_name': '工行',
                'account_no': acct_no,
                'account_name': '',
                'balance': balance,
                'type': '活期',
                'source_file': os.path.basename(pdf_path),
                'strategy': 'pdfplumber',
            })
    except Exception as e:
        results.append({
            'bank_name': '工行',
            'account_no': 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'needs_multimodal': True,
            'error': str(e),
        })
    return results


def _extract_bocom(pdf_path: str) -> List[Dict]:
    """交通银行对账单提取。"""
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                acct_no = None
                acct_name = ''
                
                m = re.search(r'账号：(\d+)', text)
                if m:
                    acct_no = m.group(1).strip()
                
                m = re.search(r'户名：(.+?)\n', text)
                if m:
                    acct_name = m.group(1).strip()
                
                balance = None
                for line in reversed(text.split('\n')):
                    nums = re.findall(r'([\d,]+\.\d{2})', line)
                    if nums:
                        balance = clean_number(nums[-1])
                        break
                
                if acct_no and balance is not None:
                    results.append({
                        'bank_name': '交行',
                        'account_no': acct_no,
                        'account_name': acct_name,
                        'balance': balance,
                        'type': '活期',
                        'source_file': os.path.basename(pdf_path),
                        'strategy': 'pdfplumber',
                    })
    except Exception as e:
        results.append({
            'bank_name': '交行',
            'account_no': 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'error': str(e),
        })
    return results


def _extract_zy(pdf_path: str) -> List[Dict]:
    """中原银行对账单提取。"""
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                acct_no = None
                acct_name = ''
                
                m = re.search(r'账号:\s*(\d+)', text)
                if m:
                    acct_no = m.group(1).strip()
                
                m = re.search(r'账户名称:\s*(.+?)\s+账号', text)
                if m:
                    acct_name = m.group(1).strip()
                
                balance = None
                for line in reversed(text.strip().split('\n')):
                    nums = re.findall(r'([\d,]+\.\d{2})', line)
                    if len(nums) >= 1:
                        balance = clean_number(nums[-1])
                        break
                
                if acct_no and balance is not None:
                    results.append({
                        'bank_name': '中原银行',
                        'account_no': acct_no,
                        'account_name': acct_name,
                        'balance': balance,
                        'type': '活期',
                        'source_file': os.path.basename(pdf_path),
                        'strategy': 'pdfplumber',
                    })
    except Exception as e:
        results.append({
            'bank_name': '中原银行',
            'account_no': 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'error': str(e),
        })
    return results


def _extract_finco(pdf_path: str) -> List[Dict]:
    """财务公司对账单提取。"""
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            acct_no = None
            acct_name = ''
            last_balance = None
            
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                if acct_no is None:
                    m = re.search(r'账号[：:]\s*(\d+)', text)
                    if m:
                        acct_no = m.group(1).strip()
                
                if not acct_name:
                    m = re.search(r'账户名称[：:]\s*(.+?)\s+币种', text)
                    if m:
                        acct_name = m.group(1).strip()
                
                for line in text.split('\n'):
                    if '本月合计' in line or '本年累计' in line:
                        nums = re.findall(r'([\d,]+\.\d{2})', line)
                        if nums:
                            last_balance = clean_number(nums[-1])
            
            if acct_no and last_balance is not None:
                results.append({
                    'bank_name': '财务公司',
                    'account_no': acct_no,
                    'account_name': acct_name,
                    'balance': last_balance,
                    'type': '活期',
                    'source_file': os.path.basename(pdf_path),
                    'strategy': 'pdfplumber',
                })
    except Exception as e:
        results.append({
            'bank_name': '财务公司',
            'account_no': 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'error': str(e),
        })
    return results


# ============================================================
# 银行适配器路由
# ============================================================

BANK_EXTRACTORS = {
    'CCB': _extract_ccb,
    'BOC': _extract_boc,
    'ICBC': _extract_icbc,
    'BOCOM': _extract_bocom,
    'ZY': _extract_zy,
    'FINCO': _extract_finco,
}


def _extract_generic_bank(pdf_path: str, bank_name: str = 'UNKNOWN') -> List[Dict]:
    """DT-213: 通用银行适配器——对未专门适配的银行，尝试结构化表格提取。
    
    策略：
    1. 用pdfplumber.extract_tables()提取所有表格
    2. 在表头行匹配关键词（"日期"/"摘要"/"收入"/"支出"/"余额"等）定位列
    3. 从最后一页的最后一行取余额
    4. 从文件名或文本中提取账号
    
    如果结构化提取也失败，返回needs_multimodal标记。
    """
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return _unknown_bank_result(pdf_path, bank_name, 'pdfplumber不可用')
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ''
            last_balance = None
            acct_no = None
            
            for page in pdf.pages:
                text = page.extract_text() or ''
                all_text += text + '\n'
                
                # 尝试提取结构化表格
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    header = [str(c).strip() if c else '' for c in table[0]]
                    
                    # 定位余额列
                    balance_col = _find_col_by_keywords(header, ['余额', '结余', '当前余额', 'balance'])
                    date_col = _find_col_by_keywords(header, ['日期', 'date'])
                    
                    if balance_col is not None:
                        # 从后往前找有效余额
                        for row in reversed(table[1:]):
                            if row and len(row) > balance_col:
                                val = row[balance_col]
                                if val and str(val).strip():
                                    _bal = clean_number(val)
                                    if _bal is not None:
                                        last_balance = _bal
                                        break
            
            # 提取账号
            acct_no = extract_account_no_from_filename(pdf_path)
            if not acct_no:
                # 从文本中匹配16位以上数字串
                for m in re.finditer(r'(\d{16,})', all_text):
                    candidate = m.group(1)
                    if not re.match(r'^(19|20)\d{2}$', candidate[:4]):
                        acct_no = candidate
                        break
            
            # 提取账户名称
            acct_name = ''
            name_match = re.search(r'(?:账户名[称:]|户名|账\s*户\s*名)[：:\s]*([^\n,，\s]+)', all_text)
            if name_match:
                acct_name = name_match.group(1).strip()
            
            # 判断账户类型
            acct_type = identify_account_type(all_text)
            
            if last_balance is not None or acct_no:
                results.append({
                    'bank_name': bank_name,
                    'account_no': acct_no or 'UNKNOWN',
                    'account_name': acct_name,
                    'balance': last_balance,
                    'type': acct_type,
                    'source_file': os.path.basename(pdf_path),
                    'strategy': 'generic_pdfplumber',
                })
            else:
                results.append(_unknown_bank_result(pdf_path, bank_name, '结构化提取未找到余额和账号'))
    
    except Exception as e:
        results.append(_unknown_bank_result(pdf_path, bank_name, str(e)))
    
    return results


def _find_col_by_keywords(header: List[str], keywords: List[str]) -> Optional[int]:
    """DT-213: 在表头列表中查找包含指定关键词的列索引。"""
    for i, h in enumerate(header):
        h_compact = h.replace(' ', '').replace('\u3000', '')
        for kw in keywords:
            if kw in h_compact:
                return i
    return None


def _unknown_bank_result(pdf_path: str, bank_name: str, error: str) -> Dict:
    """生成未知银行的标准化返回结果。"""
    return {
        'bank_name': bank_name,
        'account_no': 'UNKNOWN',
        'account_name': '',
        'balance': None,
        'type': '活期',
        'source_file': os.path.basename(pdf_path),
        'strategy': 'pdfplumber',
        'needs_multimodal': True,
        'error': error,
    }


def extract_bank_statement_pdfplumber(pdf_path: str) -> List[Dict]:
    """使用pdfplumber提取银行对账单（自动路由到对应银行适配器）。
    
    Args:
        pdf_path: 银行对账单PDF文件路径
    
    Returns:
        提取结果列表，每项含 bank_name/account_no/account_name/balance/type/source_file/strategy
    """
    # 先尝试获取文本以识别银行
    pdfplumber = _import_pdfplumber()
    sample_text = ''
    if pdfplumber:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if pdf.pages:
                    sample_text = pdf.pages[0].extract_text() or ''
        except:
            pass
    
    bank = identify_bank(sample_text, os.path.basename(pdf_path))
    
    # DT-213: 建行路由——不再依赖硬编码账号前缀判断综合/单账户
    # 改用内容特征判断: 含"账户明细信息"为综合对账单，否则为单账户
    basename = os.path.basename(pdf_path)
    if bank == 'CCB':
        if '账户明细信息' in sample_text or '综合' in sample_text:
            return _extract_ccb_combined(pdf_path)
    
    extractor = BANK_EXTRACTORS.get(bank)
    if extractor:
        return extractor(pdf_path)
    
    # DT-213: 未知银行——使用通用适配器尝试结构化提取
    return _extract_generic_bank(pdf_path, bank)


def _extract_ccb_combined(pdf_path: str) -> List[Dict]:
    """建行综合对账单提取（多账户，10页+）。
    
    每个账户以"中国建设银行账户明细信息"开头，
    需要按账户分割页面后分别提取。
    """
    pdfplumber = _import_pdfplumber()
    if not pdfplumber:
        return []
    
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 按账户分割页面
            accounts = []
            current_pages = []
            current_is_margin = False
            
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text:
                    continue
                
                if '中国建设银行账户明细信息' in text:
                    if current_pages:
                        accounts.append({
                            'pages': current_pages[:],
                            'is_margin': current_is_margin,
                        })
                    current_pages = [i]
                    current_is_margin = '保证金' in text
                else:
                    current_pages.append(i)
            
            if current_pages:
                accounts.append({
                    'pages': current_pages[:],
                    'is_margin': current_is_margin,
                })
            
            # 对每个账户提取数据
            for acct_info in accounts:
                all_text = ""
                for page_idx in acct_info['pages']:
                    text = pdf.pages[page_idx].extract_text()
                    if text:
                        all_text += text + "\n"
                
                if not all_text.strip():
                    continue
                
                # DT-213: 账号提取——通用方式
                acct_no = None
                merged = all_text.replace('\n', ' ')
                # 策略1: 匹配16位以上连续数字串
                for m in re.finditer(r'(\d{16,})', merged):
                    candidate = m.group(1)
                    if not re.match(r'^(19|20)\d{2}$', candidate[:4]):
                        acct_no = candidate
                        break
                # 策略2: 逐行拼接
                if not acct_no:
                    lines = all_text.split('\n')
                    _digit_parts = []
                    for line in lines:
                        s = line.strip()
                        m3 = re.match(r'^(\d{4})$', s)
                        if m3 and m3.group(1) not in ('2024', '2025', '2026'):
                            _digit_parts.append(m3.group(1))
                    if len(_digit_parts) >= 3:
                        acct_no = ''.join(_digit_parts)
                
                # DT-213: 余额提取——通用方式
                balance = None
                if acct_no:
                    _acct_short = acct_no[-8:] if len(acct_no) >= 8 else acct_no
                    for line in all_text.split('\n'):
                        if _acct_short in line or acct_no in line:
                            _amounts = re.findall(r'[\d,]+\.\d{2}', line)
                            if _amounts:
                                balance = clean_number(_amounts[-1])
                                break
                
                if balance is None:
                    matches = re.findall(r'([\d,]+\.\d+)\s*人民币', all_text)
                    if matches:
                        balance = clean_number(matches[-1])
                
                is_margin = acct_info['is_margin']
                if acct_no and balance is not None:
                    results.append({
                        'bank_name': '建行',
                        'account_no': acct_no,
                        'account_name': '',
                        'balance': balance,
                        'type': '保证金' if is_margin else '活期',
                        'source_file': os.path.basename(pdf_path),
                        'strategy': 'pdfplumber',
                        'page_range': f"Page {acct_info['pages'][0]+1}-{acct_info['pages'][-1]+1}",
                    })
    except Exception as e:
        results.append({
            'bank_name': '建行',
            'account_no': 'UNKNOWN',
            'account_name': '',
            'balance': None,
            'type': '活期',
            'source_file': os.path.basename(pdf_path),
            'strategy': 'pdfplumber',
            'error': str(e),
        })
    return results


# ============================================================
# 混合提取主入口
# ============================================================

def extract_bank_statement(pdf_path: str,
                           auto_multimodal: bool = True,
                           output_image_dir: str = None,
                           base_date: str = None) -> Dict[str, Any]:
    """银行对账单PDF混合提取（主入口）。
    
    策略:
    1. 先用pdfplumber提取 → 成功则直接返回
    2. 提取失败/残缺 → 生成图片+prompt供Agent多模态识别
    
    [DT-133] 基准日倒序查找规则:
    - 提供base_date时，优先从最后一页向前查找基准日余额
    - 找到基准日最后余额即停止，不继续提取前面页面的交易明细
    - 评估明细表只需基准日余额，不需完整交易流水
    
    Args:
        pdf_path: 银行对账单PDF路径
        auto_multimodal: pdfplumber失败时是否自动准备多模态素材
        output_image_dir: 图片输出目录（None=临时目录）
        base_date: 评估基准日（如'2025-12-31'），启用DT-133倒序查找
    
    Returns:
        {
            'filepath': str,
            'filename': str,
            'strategy': 'pdfplumber' | 'multimodal_needed' | 'failed',
            'records': list,          # pdfplumber提取结果
            'multimodal_images': list, # 图片路径（供Agent读取）
            'multimodal_prompt': str,  # 多模态识别prompt
            'needs_multimodal': bool,
            'base_date': str,          # [DT-133] 基准日
            'warnings': list,
        }
    """
    result = {
        'filepath': pdf_path,
        'filename': os.path.basename(pdf_path),
        'strategy': 'pdfplumber',
        'records': [],
        'multimodal_images': [],
        'multimodal_prompt': '',
        'needs_multimodal': False,
        'base_date': base_date,  # [DT-133]
        'warnings': [],
    }
    
    # [DT-133] 基准日倒序查找：从最后一页向前提取
    # Step 1: pdfplumber提取（倒序页）
    records = extract_bank_statement_pdfplumber(pdf_path)
    
    if base_date and records:
        # DT-133: 找到基准日最后余额即停，过滤掉前面的交易明细
        # 保留：银行名称、账号、户名、币种、基准日余额
        # 不保留：逐笔交易明细（除非余额行本身）
        balance_records = []
        for r in reversed(records):
            if r.get('balance') is not None:
                balance_records.append(r)
                break  # 找到最后一个余额即停 [DT-133 Early Stop]
        if balance_records:
            # 保留银行信息 + 最后余额
            result['records'] = balance_records
            result['warnings'].append(
                f'[DT-133] 基准日{base_date}倒序查找：找到余额即停，'
                f'跳过{len(records)-1}条交易明细'
            )
        else:
            result['records'] = records
            result['warnings'].append(
                f'[DT-133] 基准日{base_date}倒序查找：未找到余额行，保留全部记录'
            )
    else:
        result['records'] = records
    
    # 判断是否需要多模态兜底
    needs_multimodal = False
    if not records:
        needs_multimodal = True
        result['warnings'].append('pdfplumber提取为空，需要多模态识别')
    elif any(r.get('needs_multimodal') for r in records):
        needs_multimodal = True
        result['warnings'].append('pdfplumber提取失败（可能是扫描件），需要多模态识别')
    elif all(r.get('balance') is None for r in records):
        needs_multimodal = True
        result['warnings'].append('pdfplumber提取到文本但余额全部为空，建议多模态核验')
    
    # Step 2: 准备多模态素材
    if needs_multimodal and auto_multimodal:
        result['strategy'] = 'multimodal_needed'
        result['needs_multimodal'] = True
        
        images = pdf_to_images(pdf_path, output_dir=output_image_dir)
        result['multimodal_images'] = images
        
        if images:
            # [DT-133] 多模态识别时也传入base_date，引导倒序查找
            result['multimodal_prompt'] = generate_multimodal_prompt(
                os.path.basename(pdf_path), len(images),
                base_date=base_date or ''
            )
        else:
            result['warnings'].append('PDF转图片失败，无法进行多模态识别')
            result['strategy'] = 'failed'
    
    return result


# ============================================================
# 多模态识别支持
# ============================================================

def pdf_to_images(pdf_path: str, output_dir: str = None,
                  dpi: int = 300, fmt: str = 'png') -> List[str]:
    """将PDF转为图片，供Agent多模态识别。
    
    Args:
        pdf_path: PDF文件路径
        output_dir: 图片输出目录（None=PDF同目录下创建_images子目录）
        dpi: 渲染DPI（300适合OCR/视觉识别）
        fmt: 图片格式 png/jpg
    
    Returns:
        图片文件路径列表
    """
    convert_from_path = _import_pdf2image()
    if not convert_from_path:
        print('[WARNING] pdf2image未安装，无法将PDF转为图片。'
              '请安装: pip install pdf2image 并安装poppler')
        return []
    
    if output_dir is None:
        output_dir = os.path.splitext(pdf_path)[0] + '_images'
    
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        from PIL import Image
        images = convert_from_path(pdf_path, dpi=dpi, fmt=fmt)
        image_paths = []
        
        for i, img in enumerate(images):
            img_path = os.path.join(output_dir, f'page_{i+1:03d}.{fmt}')
            img.save(img_path, fmt.upper() if fmt != 'jpg' else 'JPEG')
            image_paths.append(img_path)
        
        return image_paths
    except Exception as e:
        print(f'[ERROR] PDF转图片失败: {e}')
        # 提示安装poppler
        if 'poppler' in str(e).lower() or 'PDFInfo' in str(e):
            print('[HINT] 需要配置poppler PATH:')
            print('  预编译包(Windows): ~/.workbuddy/skills/valuation-detail-table/scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin/')
            print('  配置PATH: export PATH="$HOME/.workbuddy/skills/valuation-detail-table/scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin:$PATH"')
            print('  源码包(Linux/macOS需编译): ~/.workbuddy/skills/valuation-detail-table/scripts/poppler-26.05.0/')
            print('  详见: scripts/README.md → "依赖工具 → Poppler" 节')
        return []


def generate_multimodal_prompt(filename: str, page_count: int,
                                bank_hint: str = '',
                                base_date: str = '') -> str:
    """生成银行对账单多模态识别prompt。
    
    供Agent读取PDF图片时使用，引导视觉模型提取结构化数据。
    [DT-133] 支持基准日倒序查找，优先识别期末余额而非逐笔交易。
    
    Args:
        filename: PDF文件名
        page_count: 页数
        bank_hint: 银行名称提示（可选）
        base_date: 评估基准日（如'2025-12-31'），启用DT-133倒序查找
    
    Returns:
        多模态识别prompt文本
    """
    base_date_instruction = ''
    if base_date:
        base_date_instruction = f"""
**[DT-133] 基准日倒序查找规则（强制执行）：**
- 评估基准日为 {base_date}
- 优先从对账单末尾查找"账户余额"/"期末余额"行
- 找到基准日当天最后的余额即停止，**不需要提取前面的逐笔交易明细**
- 只需提取：银行名称、账号、户名、币种、基准日余额
- 如对账单末尾有"可用余额"+"冻结余额"，则账户余额 = 可用余额 + 冻结余额
- 如基准日当天无交易，取基准日前最后一个交易日的余额
"""
    
    prompt = f"""请识别这张银行对账单图片，提取以下关键信息：
{base_date_instruction}
文件名: {filename}
页数: {page_count}
{"银行提示: " + bank_hint if bank_hint else ""}

请按以下JSON格式输出提取结果（每页一个对象）:

```json
[{{
  "page": 1,
  "bank_name": "银行全称",
  "account_no": "完整账号（连续数字，不要拆分）",
  "account_name": "账户名称/户名",
  "account_type": "活期/保证金/资金监管/贷款",
  "ending_balance": "期末余额（纯数字，含小数，不含逗号）",
  "frozen_amount": "冻结余额（如有）",
  "available_balance": "可用余额（如有）",
  "currency": "CNY/USD/HKD",
  "period": "对账期间，如2025年12月",
  "balance_date": "余额对应日期，如2025-12-31",
  "notes": "任何识别疑问或特殊情况说明"
}}]
```

**特别注意:**
1. 账号必须完整连续，不要因为换行而拆分。如果账号跨行显示，请拼接为完整号码。
2. 期末余额是最终余额（通常是最后一行的余额），不是中间发生额。
3. 如果同一页有多个账户，请分别提取每个账户的信息。
4. 保证金账户请在account_type中标注"保证金"。
5. 如果是扫描件/图片模糊，请在notes中说明。
6. 数字中的逗号是千分位分隔符，提取时去除。
7. **[DT-133]** 如果是最后一页，重点查找"账户余额"/"余额"行，找到后无需关注前面的逐笔交易。

如果无法识别某些字段，请填null并说明原因。"""
    
    return prompt


def parse_multimodal_result(multimodal_text: str,
                            source_file: str = '') -> List[Dict]:
    """解析多模态识别结果文本，转为标准记录格式。
    
    Args:
        multimodal_text: 视觉模型返回的JSON格式文本
        source_file: 源PDF文件名
    
    Returns:
        标准化的银行对账单记录列表
    """
    records = []
    
    # 尝试提取JSON
    try:
        # 去除markdown代码块标记
        text = multimodal_text.strip()
        if text.startswith('```'):
            text = re.sub(r'^```\w*\n?', '', text)
            text = re.sub(r'\n?```$', '', text)
            text = text.strip()
        
        data = json.loads(text)
        if isinstance(data, dict):
            data = [data]
        
        for item in data:
            balance = item.get('ending_balance')
            if isinstance(balance, str):
                balance = clean_number(balance)
            
            records.append({
                'bank_name': item.get('bank_name', ''),
                'account_no': str(item.get('account_no', '') or ''),
                'account_name': item.get('account_name', ''),
                'balance': balance,
                'type': item.get('account_type', '活期'),
                'source_file': source_file,
                'strategy': 'multimodal',
                'notes': item.get('notes', ''),
            })
    except json.JSONDecodeError:
        # JSON解析失败，尝试正则提取
        records = _parse_multimodal_fallback(multimodal_text, source_file)
    
    return records


def _parse_multimodal_fallback(text: str, source_file: str) -> List[Dict]:
    """多模态结果JSON解析失败时的正则兜底提取。"""
    records = []
    
    # 尝试提取关键信息
    bank_name = ''
    for name in ['建设银行', '工商银行', '中国银行', '交通银行', '中原银行', '财务公司']:
        if name in text:
            bank_name = name
            break
    
    acct_no = ''
    m = re.search(r'account_no["\s:]+["\s]*(\d{10,})', text)
    if m:
        acct_no = m.group(1)
    
    balance = None
    m = re.search(r'ending_balance["\s:]+["\s]*([\d,]+\.\d+)', text)
    if m:
        balance = clean_number(m.group(1))
    
    if bank_name or acct_no:
        records.append({
            'bank_name': bank_name,
            'account_no': acct_no or 'UNKNOWN',
            'account_name': '',
            'balance': balance,
            'type': '活期',
            'source_file': source_file,
            'strategy': 'multimodal',
            'notes': 'JSON解析失败，正则兜底提取',
        })
    
    return records


# ============================================================
# 批量提取
# ============================================================

def batch_extract_bank_statements(filepaths: List[str],
                                   output_dir: str = None,
                                   auto_multimodal: bool = True) -> Dict[str, Any]:
    """批量提取银行对账单PDF（混合方案）。
    
    Args:
        filepaths: PDF文件路径列表
        output_dir: 输出目录
        auto_multimodal: 是否自动准备多模态素材
    
    Returns:
        {
            'total': int,
            'extracted': int,
            'needs_multimodal': int,
            'failed': int,
            'records': list,           # 所有提取到的记录
            'multimodal_tasks': list,  # 需要多模态识别的任务
            'summary': dict,
        }
    """
    all_records = []
    multimodal_tasks = []
    stats = {'extracted': 0, 'needs_multimodal': 0, 'failed': 0}
    
    for fp in filepaths:
        if not os.path.exists(fp):
            continue
        
        result = extract_bank_statement(fp, auto_multimodal=auto_multimodal)
        
        if result['records'] and not result['needs_multimodal']:
            # pdfplumber提取成功
            valid_records = [r for r in result['records'] if r.get('balance') is not None]
            if valid_records:
                all_records.extend(valid_records)
                stats['extracted'] += 1
            else:
                stats['needs_multimodal'] += 1
                multimodal_tasks.append(result)
        elif result['needs_multimodal']:
            stats['needs_multimodal'] += 1
            multimodal_tasks.append(result)
        else:
            stats['failed'] += 1
    
    # 去重
    unique_records = deduplicate_accounts(all_records)
    
    # 汇总
    total_balance = sum(r['balance'] for r in unique_records if r.get('balance') is not None)
    by_bank = {}
    for r in unique_records:
        key = r.get('bank_name', 'UNKNOWN')
        if key not in by_bank:
            by_bank[key] = {'count': 0, 'total': 0}
        by_bank[key]['count'] += 1
        if r.get('balance') is not None:
            by_bank[key]['total'] += r['balance']
    
    summary = {
        'total_accounts': len(unique_records),
        'total_balance': round(total_balance, 2),
        'by_bank': {k: {'count': v['count'], 'total': round(v['total'], 2)} for k, v in by_bank.items()},
    }
    
    output = {
        'total': len(filepaths),
        'extracted': stats['extracted'],
        'needs_multimodal': stats['needs_multimodal'],
        'failed': stats['failed'],
        'records': unique_records,
        'multimodal_tasks': multimodal_tasks,
        'summary': summary,
    }
    
    # 保存结果
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        json_path = os.path.join(output_dir, 'bank_extraction_result.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2, default=str)
        
        # 生成多模态任务摘要
        if multimodal_tasks:
            task_path = os.path.join(output_dir, 'multimodal_tasks.json')
            tasks_summary = []
            for t in multimodal_tasks:
                tasks_summary.append({
                    'filename': t['filename'],
                    'images': t['multimodal_images'],
                    'prompt': t['multimodal_prompt'],
                })
            with open(task_path, 'w', encoding='utf-8') as f:
                json.dump(tasks_summary, f, ensure_ascii=False, indent=2)
    
    return output


# ============================================================
# CLI入口
# ============================================================

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='银行对账单PDF混合提取工具 v2.0')
    parser.add_argument('path', help='PDF文件路径或文件夹路径')
    parser.add_argument('--mode', choices=['auto', 'pdfplumber', 'multimodal'],
                       default='auto', help='提取模式')
    parser.add_argument('--output-dir', help='输出目录')
    parser.add_argument('--dpi', type=int, default=300, help='PDF转图片DPI (默认300)')
    parser.add_argument('--no-multimodal', action='store_true', help='禁用多模态兜底')
    parser.add_argument('--images-only', action='store_true', help='仅将PDF转为图片')
    
    args = parser.parse_args()
    
    if args.images_only:
        # 仅转图片模式
        if os.path.isfile(args.path):
            images = pdf_to_images(args.path, args.output_dir, dpi=args.dpi)
            print(f"生成 {len(images)} 张图片:")
            for img in images:
                print(f"  {img}")
        elif os.path.isdir(args.path):
            for root, dirs, files in os.walk(args.path):
                for f in files:
                    if f.lower().endswith('.pdf'):
                        fp = os.path.join(root, f)
                        images = pdf_to_images(fp, args.output_dir, dpi=args.dpi)
                        print(f"{f}: {len(images)} 页")
        sys.exit(0)
    
    if os.path.isfile(args.path):
        # 单文件模式
        result = extract_bank_statement(
            args.path,
            auto_multimodal=not args.no_multimodal,
            output_image_dir=args.output_dir,
        )
        
        print(f"\n文件: {result['filename']}")
        print(f"策略: {result['strategy']}")
        
        if result['records']:
            print(f"\npdfplumber提取结果 ({len(result['records'])}条):")
            for r in result['records']:
                bal = f"{r['balance']:,.2f}" if r.get('balance') is not None else 'None'
                print(f"  {r.get('bank_name', '?'):8s} {r.get('account_no', '?'):25s} {bal}")
        
        if result['needs_multimodal']:
            print(f"\n⚠️ 需要多模态识别:")
            print(f"  图片: {len(result['multimodal_images'])} 张")
            if result['multimodal_images']:
                for img in result['multimodal_images']:
                    print(f"    {img}")
            if result['multimodal_prompt']:
                print(f"\n  多模态Prompt:")
                print(f"  {result['multimodal_prompt'][:200]}...")
        
        if result['warnings']:
            print(f"\n警告:")
            for w in result['warnings']:
                print(f"  ⚠️ {w}")
    
    elif os.path.isdir(args.path):
        # 批量模式
        filepaths = []
        for root, dirs, files in os.walk(args.path):
            for f in files:
                if f.lower().endswith('.pdf'):
                    filepaths.append(os.path.join(root, f))
        
        print(f"扫描到 {len(filepaths)} 个PDF文件")
        result = batch_extract_bank_statements(
            filepaths,
            output_dir=args.output_dir,
            auto_multimodal=not args.no_multimodal,
        )
        
        print(f"\n{'='*60}")
        print(f"批量提取结果:")
        print(f"  成功: {result['extracted']}")
        print(f"  需多模态: {result['needs_multimodal']}")
        print(f"  失败: {result['failed']}")
        print(f"\n去重后账户数: {result['summary']['total_accounts']}")
        print(f"余额合计: {result['summary']['total_balance']:,.2f}")
        print(f"\n按银行汇总:")
        for bank, info in result['summary']['by_bank'].items():
            print(f"  {bank}: {info['count']}户 合计{info['total']:,.2f}")
        
        if result['multimodal_tasks']:
            print(f"\n⚠️ 需要多模态识别的文件 ({len(result['multimodal_tasks'])}个):")
            for t in result['multimodal_tasks']:
                print(f"  {t['filename']}")
    else:
        print(f"错误: 路径不存在 - {args.path}")
        sys.exit(1)
