# -*- coding: utf-8 -*-
"""
pdf_extract.py — PDF/图片内容提取共享工具

核心思想：
  估值Skill体系中，PDF和图片是重要数据源（银行对账单、固定资产卡片台账、
  辅助余额表扫描件等）。每次读取PDF/图片时，Agent需要重复编写提取逻辑，
  且不同Skill间提取结果格式不一致。

  本工具将常用的PDF/图片提取逻辑封装为标准化函数，Agent只需一次调用即可
  完成提取+结构化+验证，结果格式统一，可供Phase 0~5直接引用。

设计原则：
  - 三级提取策略：pdfplumber文本提取 → PyMuPDF备选 → 图片OCR兜底
  - 结构化输出：所有提取结果统一为dict格式，含提取状态/数据/元信息
  - 验证闭环：提取后自动验证非空（DT-108），空结果标注原因
  - 场景化提取：银行对账单/卡片台账/辅助余额表等预设模板
  - 纯Python实现，不依赖外部API

覆盖场景：
  - 银行对账单PDF：提取银行名/账号/余额
  - 固定资产卡片台账PDF：提取资产名/原值/净值
  - 辅助余额表PDF：提取科目/结算对象/余额
  - 通用PDF/图片：文本+表格提取
  - 图片型PDF（扫描件）：OCR识别

依赖（已安装）：
  - pdfplumber 0.11.9 — 文本PDF表格提取（首选）
  - PyMuPDF 1.27.2.3 — 文本PDF备选+图片提取
  - pytesseract 0.3.13 — 图片OCR（Tesseract引擎）
  - pdf2image 1.17.0 — PDF转图片（OCR前置）

v1.0 (2026-05-22): 初始版本
  - extract_pdf(): 通用PDF提取（文本+表格+OCR三级策略）
  - extract_bank_statement(): 银行对账单结构化提取
  - extract_asset_register(): 固定资产卡片台账结构化提取
  - extract_auxiliary_balance(): 辅助余额表结构化提取
  - extract_image(): 图片OCR提取
  - validate_extraction(): DT-108提取完整性验证
  - batch_extract(): 批量提取+完整性报告
  - save_extraction_json(): 提取结果持久化
"""

import os
import sys
import json
import re
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple

# 延迟导入，避免未安装时崩溃
_pdfplumber = None
_fitz = None
_pytesseract = None
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


def _import_fitz():
    global _fitz
    if _fitz is None:
        try:
            import fitz
            _fitz = fitz
        except ImportError:
            pass
    return _fitz


def _import_pytesseract():
    global _pytesseract
    if _pytesseract is None:
        try:
            import pytesseract
            _pytesseract = pytesseract
        except ImportError:
            pass
    return _pytesseract


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
# 通用PDF提取
# ============================================================

def extract_pdf(filepath: str, use_ocr_fallback: bool = True,
                ocr_lang: str = 'chi_sim+eng') -> Dict[str, Any]:
    """通用PDF内容提取（三级策略）。

    策略优先级：
    1. pdfplumber — 文本PDF首选，支持表格提取
    2. PyMuPDF — 文本PDF备选，速度快
    3. OCR — 图片型PDF兜底（需use_ocr_fallback=True）

    Args:
        filepath: PDF文件路径
        use_ocr_fallback: 文本提取为空时是否尝试OCR
        ocr_lang: OCR语言，默认中文+英文

    Returns:
        dict: {
            'filepath': str,
            'filename': str,
            'status': 'extracted' | 'empty' | 'ocr_fallback' | 'failed',
            'strategy': 'pdfplumber' | 'pymupdf' | 'ocr',
            'pages': int,
            'text': str,           # 全文文本
            'tables': list,        # 表格数据 [{headers, rows}]
            'page_texts': list,    # 逐页文本
            'warnings': list,
        }
    """
    result = {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'status': 'empty',
        'strategy': None,
        'pages': 0,
        'text': '',
        'tables': [],
        'page_texts': [],
        'warnings': [],
    }

    # --- 策略1: pdfplumber ---
    pdfplumber = _import_pdfplumber()
    if pdfplumber:
        try:
            with pdfplumber.open(filepath) as pdf:
                result['pages'] = len(pdf.pages)
                all_text = []
                for i, page in enumerate(pdf.pages):
                    page_text = page.extract_text() or ''
                    all_text.append(page_text)

                    # 提取表格
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            headers = table[0] if table[0] else []
                            rows = table[1:] if len(table) > 1 else []
                            # 清理None值
                            headers = [str(h).strip() if h else '' for h in headers]
                            rows = [
                                [str(c).strip() if c is not None else '' for c in row]
                                for row in rows
                            ]
                            result['tables'].append({
                                'page': i + 1,
                                'headers': headers,
                                'rows': rows,
                            })

                result['text'] = '\n'.join(all_text)
                result['page_texts'] = all_text
                result['strategy'] = 'pdfplumber'

                if result['text'].strip():
                    result['status'] = 'extracted'
                else:
                    result['warnings'].append(
                        'pdfplumber提取文本为空，可能是图片型PDF（扫描件）'
                    )
        except Exception as e:
            result['warnings'].append(f'pdfplumber提取失败: {str(e)}')

    # --- 策略2: PyMuPDF（仅pdfplumber失败时） ---
    if result['status'] != 'extracted':
        fitz = _import_fitz()
        if fitz:
            try:
                doc = fitz.open(filepath)
                result['pages'] = len(doc)
                all_text = []
                for page in doc:
                    all_text.append(page.get_text())
                doc.close()

                result['text'] = '\n'.join(all_text)
                result['page_texts'] = all_text
                result['strategy'] = 'pymupdf'

                if result['text'].strip():
                    result['status'] = 'extracted'
                    # PyMuPDF不直接提取表格结构，但文本可用
                else:
                    result['warnings'].append(
                        'PyMuPDF提取文本也为空，确认是图片型PDF'
                    )
            except Exception as e:
                result['warnings'].append(f'PyMuPDF提取失败: {str(e)}')

    # --- 策略3: OCR兜底 ---
    if result['status'] != 'extracted' and use_ocr_fallback:
        ocr_result = extract_image(filepath, ocr_lang=ocr_lang)
        if ocr_result['status'] == 'extracted':
            result['text'] = ocr_result['text']
            result['strategy'] = 'ocr'
            result['status'] = 'ocr_fallback'
            result['warnings'].append('使用了OCR兜底提取，建议人工核验关键数字')
        else:
            result['status'] = 'failed'
            result['warnings'].extend(ocr_result.get('warnings', []))

    # --- 策略4: [DT-132] 标记需要Agent多模态Read兜底 ---
    # Python脚本全部失败后，MUST由Agent使用Read工具多模态识别
    # 禁止以"扫描件""需人工处理"为由跳过
    if result['status'] == 'failed':
        result['needs_multimodal_read'] = True
        result['warnings'].append(
            'Python脚本提取全部失败，MUST使用Read工具多模态识别（DT-132），'
            '禁止以扫描件/需人工处理为由跳过'
        )

    return result


# ============================================================
# 图片OCR提取
# ============================================================

def extract_image(filepath: str, ocr_lang: str = 'chi_sim+eng') -> Dict[str, Any]:
    """图片OCR内容提取。

    支持：PNG/JPG/JPEG/BMP/TIFF/PDF（PDF先转图片再OCR）

    Args:
        filepath: 图片/PDF文件路径
        ocr_lang: OCR语言

    Returns:
        dict: {
            'filepath': str,
            'status': 'extracted' | 'failed',
            'text': str,
            'warnings': list,
        }
    """
    result = {
        'filepath': filepath,
        'status': 'empty',
        'text': '',
        'warnings': [],
    }

    pytesseract = _import_pytesseract()
    if not pytesseract:
        result['status'] = 'failed'
        result['warnings'].append('pytesseract未安装，无法进行OCR提取')
        return result

    from PIL import Image

    try:
        ext = os.path.splitext(filepath)[1].lower()

        if ext == '.pdf':
            # PDF → 图片 → OCR
            convert_from_path = _import_pdf2image()
            if not convert_from_path:
                result['status'] = 'failed'
                result['warnings'].append('pdf2image未安装或poppler未配置PATH，无法将PDF转为图片进行OCR。'
                                          '预编译包(Windows): ~/.workbuddy/skills/valuation-detail-table/scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin/')
                return result

            images = convert_from_path(filepath, dpi=300)
            all_text = []
            for img in images:
                text = pytesseract.image_to_string(img, lang=ocr_lang)
                all_text.append(text)

            result['text'] = '\n'.join(all_text)
        else:
            # 直接OCR图片
            img = Image.open(filepath)
            result['text'] = pytesseract.image_to_string(img, lang=ocr_lang)

        if result['text'].strip():
            result['status'] = 'extracted'
        else:
            result['status'] = 'failed'
            result['warnings'].append('OCR提取结果为空')

    except Exception as e:
        result['status'] = 'failed'
        result['warnings'].append(f'OCR提取失败: {str(e)}')

    return result


# ============================================================
# 银行对账单结构化提取
# ============================================================

# 银行名称关键词映射
BANK_KEYWORDS = {
    '工商银行': ['工商', 'ICBC'],
    '建设银行': ['建设', 'CCB'],
    '农业银行': ['农业', 'ABC'],
    '中国银行': ['中国银行', 'BOC'],
    '交通银行': ['交通', 'BOCOM'],
    '招商银行': ['招商', 'CMB'],
    '浦发银行': ['浦发', 'SPDB'],
    '民生银行': ['民生', 'CMBC'],
    '兴业银行': ['兴业', 'CIB'],
    '中信银行': ['中信', 'CITIC'],
    '光大银行': ['光大', 'CEB'],
    '华夏银行': ['华夏', 'HXB'],
    '平安银行': ['平安', 'PAB'],
    '邮储银行': ['邮储', 'PSBC'],
    '北京银行': ['北京银行', 'BOB'],
    '宁波银行': ['宁波银行', 'NBCB'],
}


def extract_bank_statement(filepath: str) -> Dict[str, Any]:
    """银行对账单PDF结构化提取。

    提取字段：银行名称、账户名称、账号、期末余额、币种

    Args:
        filepath: 银行对账单PDF文件路径

    Returns:
        dict: {
            'filepath': str,
            'filename': str,
            'status': 'extracted' | 'partial' | 'failed',
            'bank_name': str,
            'account_name': str,
            'account_no': str,
            'ending_balance': float or None,
            'currency': str,
            'raw_text': str,
            'tables': list,
            'warnings': list,
        }
    """
    # 先用通用提取获取全文
    raw = extract_pdf(filepath)

    result = {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'status': raw['status'],
        'bank_name': '',
        'account_name': '',
        'account_no': '',
        'ending_balance': None,
        'currency': 'CNY',
        'raw_text': raw['text'],
        'tables': raw['tables'],
        'warnings': raw['warnings'],
    }

    text = raw['text']
    if not text.strip():
        result['status'] = 'failed'
        return result

    # --- 银行名称 ---
    for bank_name, keywords in BANK_KEYWORDS.items():
        for kw in keywords:
            if kw in text:
                result['bank_name'] = bank_name
                break
        if result['bank_name']:
            break

    # --- 账户名称 ---
    patterns_name = [
        r'(?:账户名称|户名|账户户名)[：:\s]*([^\n\r,，]{2,30})',
        r'(?:单位名称|公司名称)[：:\s]*([^\n\r,，]{2,30})',
    ]
    for pat in patterns_name:
        m = re.search(pat, text)
        if m:
            result['account_name'] = m.group(1).strip()
            break

    # --- 账号 ---
    patterns_no = [
        r'(?:账号|活期账号|帐号|账户号码)[：:\s]*(\d{10,30})',
        r'(?:Account\s*No)[：:\s]*(\d{10,30})',
    ]
    for pat in patterns_no:
        m = re.search(pat, text)
        if m:
            result['account_no'] = m.group(1).strip()
            break

    # --- 期末余额 ---
    patterns_balance = [
        r'(?:期末余额|本对账期末余额|账户余额|结余|余额)[：:\s]*[¥￥]?\s*([\d,]+\.\d{2})',
        r'(?:期末余额|本对账期末余额|账户余额|结余|余额)[：:\s]*[¥￥]?\s*([\d,]+)',
    ]
    for pat in patterns_balance:
        m = re.search(pat, text)
        if m:
            try:
                result['ending_balance'] = float(m.group(1).replace(',', ''))
            except ValueError:
                pass
            break

    # --- 币种 ---
    if '美元' in text or 'USD' in text.upper():
        result['currency'] = 'USD'
    elif '欧元' in text or 'EUR' in text.upper():
        result['currency'] = 'EUR'
    elif '港币' in text or 'HKD' in text.upper():
        result['currency'] = 'HKD'

    # 判断提取完整性
    filled_fields = sum([
        bool(result['bank_name']),
        bool(result['account_name']),
        bool(result['account_no']),
        result['ending_balance'] is not None,
    ])
    if filled_fields >= 3:
        result['status'] = 'extracted'
    elif filled_fields >= 1:
        result['status'] = 'partial'
    else:
        result['status'] = 'failed'
        result['warnings'].append('银行对账单关键字段均未提取到')

    return result


# ============================================================
# 固定资产卡片台账结构化提取
# ============================================================

def extract_asset_register(filepath: str) -> Dict[str, Any]:
    """固定资产卡片台账PDF结构化提取。

    提取每项资产的：资产名称、资产类别、账面原值、账面净值、数量

    Args:
        filepath: 卡片台账PDF文件路径

    Returns:
        dict: {
            'filepath': str,
            'filename': str,
            'status': 'extracted' | 'partial' | 'failed',
            'total_items': int,
            'items': [{name, category, original_value, net_value, quantity}],
            'raw_text': str,
            'tables': list,
            'warnings': list,
        }
    """
    raw = extract_pdf(filepath)

    result = {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'status': raw['status'],
        'total_items': 0,
        'items': [],
        'raw_text': raw['text'],
        'tables': raw['tables'],
        'warnings': raw['warnings'],
    }

    # 优先从表格提取（卡片台账通常是表格形式）
    if raw['tables']:
        for table in raw['tables']:
            headers = [h.replace(' ', '').strip() for h in table['headers']]
            # 查找关键列
            name_col = _find_col(headers, ['资产名称', '设备名称', '项目名称', '名称'])
            cat_col = _find_col(headers, ['类别', '分类', '资产类别', '设备类别'])
            orig_col = _find_col(headers, ['原值', '账面原值', '入账价值', '购置价值'])
            net_col = _find_col(headers, ['净值', '账面净值', '净额'])
            qty_col = _find_col(headers, ['数量', '台数', '台/套'])

            if name_col is not None or orig_col is not None:
                for row in table['rows']:
                    item = {
                        'name': row[name_col] if name_col is not None and name_col < len(row) else '',
                        'category': row[cat_col] if cat_col is not None and cat_col < len(row) else '',
                        'original_value': _parse_number(row[orig_col]) if orig_col is not None and orig_col < len(row) else None,
                        'net_value': _parse_number(row[net_col]) if net_col is not None and net_col < len(row) else None,
                        'quantity': _parse_int(row[qty_col]) if qty_col is not None and qty_col < len(row) else 1,
                    }
                    if item['name'] or item['original_value'] is not None:
                        result['items'].append(item)

    # 表格提取不足时，尝试从文本正则提取
    if len(result['items']) < 3 and raw['text']:
        # 常见模式：每行一个资产，含名称+金额
        # 匹配模式：资产名称 + 数字（原值）+ 数字（净值）
        pattern = r'([^\d\s,，.]{3,30})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})'
        for m in re.finditer(pattern, raw['text']):
            name = m.group(1).strip()
            # 排除表头/合计等非资产行
            if any(kw in name for kw in ['合计', '小计', '名称', '原值', '净值', '项目']):
                continue
            result['items'].append({
                'name': name,
                'category': '',
                'original_value': _parse_number(m.group(2)),
                'net_value': _parse_number(m.group(3)),
                'quantity': 1,
            })

    result['total_items'] = len(result['items'])

    if result['total_items'] > 0:
        result['status'] = 'extracted'
    elif raw['text'].strip():
        result['status'] = 'partial'
        result['warnings'].append('卡片台账文本已提取但未解析出资产明细项')
    else:
        result['status'] = 'failed'

    return result


# ============================================================
# 辅助余额表结构化提取
# ============================================================

def extract_auxiliary_balance(filepath: str) -> Dict[str, Any]:
    """辅助余额表PDF结构化提取。

    提取：科目名称、结算对象、期末借方余额、期末贷方余额

    Args:
        filepath: 辅助余额表PDF文件路径

    Returns:
        dict: {
            'filepath': str,
            'filename': str,
            'status': 'extracted' | 'partial' | 'failed',
            'subject_name': str,
            'counterparties': [{name, debit_balance, credit_balance}],
            'raw_text': str,
            'tables': list,
            'warnings': list,
        }
    """
    raw = extract_pdf(filepath)

    result = {
        'filepath': filepath,
        'filename': os.path.basename(filepath),
        'status': raw['status'],
        'subject_name': '',
        'counterparties': [],
        'raw_text': raw['text'],
        'tables': raw['tables'],
        'warnings': raw['warnings'],
    }

    text = raw['text']

    # --- 科目名称 ---
    pat_subject = r'(?:科目|科目名称)[：:\s]*([^\n\r,，]{2,20})'
    m = re.search(pat_subject, text)
    if m:
        result['subject_name'] = m.group(1).strip()

    # --- 结算对象列表 ---
    # 优先从表格提取
    if raw['tables']:
        for table in raw['tables']:
            headers = [h.replace(' ', '').strip() for h in table['headers']]
            name_col = _find_col(headers, ['结算对象', '对象名称', '客户名称', '供应商名称', '户名', '名称'])
            debit_col = _find_col(headers, ['期末借方', '借方余额', '借方', '借方发生额'])
            credit_col = _find_col(headers, ['期末贷方', '贷方余额', '贷方', '贷方发生额'])

            if name_col is not None:
                for row in table['rows']:
                    name = row[name_col] if name_col < len(row) else ''
                    # 排除合计/小计行
                    if name and not any(kw in name for kw in ['合计', '小计', '合计行']):
                        result['counterparties'].append({
                            'name': name.strip(),
                            'debit_balance': _parse_number(row[debit_col]) if debit_col is not None and debit_col < len(row) else None,
                            'credit_balance': _parse_number(row[credit_col]) if credit_col is not None and credit_col < len(row) else None,
                        })

    # 表格不足时尝试文本正则
    if not result['counterparties'] and text:
        # 常见格式：结算对象名 + 数字
        pattern = r'([^\d\s,，.]{2,25})\s+([\d,]+\.\d{2})'
        for m in re.finditer(pattern, text):
            name = m.group(1).strip()
            if any(kw in name for kw in ['合计', '小计', '科目', '对象', '余额']):
                continue
            result['counterparties'].append({
                'name': name,
                'debit_balance': _parse_number(m.group(2)),
                'credit_balance': None,
            })

    if result['counterparties']:
        result['status'] = 'extracted'
    elif text.strip():
        result['status'] = 'partial'
        result['warnings'].append('辅助余额表文本已提取但未解析出结算对象')
    else:
        result['status'] = 'failed'

    return result


# ============================================================
# DT-108 提取完整性验证
# ============================================================

def validate_extraction(extraction_result: Dict[str, Any],
                       expected_type: str = None) -> Dict[str, Any]:
    """DT-108提取完整性验证。

    Args:
        extraction_result: extract_pdf/bank_statement/...的返回值
        expected_type: 预期类型 'bank' | 'asset' | 'auxiliary' | None(通用)

    Returns:
        dict: {
            'valid': bool,
            'severity': 'PASS' | 'WARNING' | 'CRITICAL',
            'checks': [{name, passed, detail}],
        }
    """
    checks = []
    status = extraction_result.get('status', 'empty')

    # 通用检查：提取状态
    if status in ('extracted', 'ocr_fallback'):
        checks.append({'name': '提取状态', 'passed': True, 'detail': f'状态={status}'})
    elif status == 'partial':
        checks.append({'name': '提取状态', 'passed': False, 'detail': '部分提取，关键字段缺失'})
    else:
        checks.append({'name': '提取状态', 'passed': False, 'detail': f'提取失败或为空，状态={status}'})

    # 通用检查：文本非空
    text = extraction_result.get('raw_text', extraction_result.get('text', ''))
    checks.append({
        'name': '文本非空',
        'passed': bool(text.strip()),
        'detail': f'文本长度={len(text.strip())}字符',
    })

    # 场景化检查
    if expected_type == 'bank':
        for field in ['bank_name', 'account_no', 'ending_balance']:
            val = extraction_result.get(field)
            checks.append({
                'name': f'银行对账单.{field}',
                'passed': bool(val) or val == 0,
                'detail': f'值={val}',
            })

    elif expected_type == 'asset':
        items = extraction_result.get('items', [])
        checks.append({
            'name': '资产项数',
            'passed': len(items) > 0,
            'detail': f'提取到{len(items)}项资产',
        })

    elif expected_type == 'auxiliary':
        cp = extraction_result.get('counterparties', [])
        checks.append({
            'name': '结算对象数',
            'passed': len(cp) > 0,
            'detail': f'提取到{len(cp)}个结算对象',
        })

    # 汇总
    all_passed = all(c['passed'] for c in checks)
    critical_fails = [c for c in checks if not c['passed']]

    if all_passed:
        severity = 'PASS'
    elif len(critical_fails) <= 1:
        severity = 'WARNING'
    else:
        severity = 'CRITICAL'

    return {
        'valid': all_passed,
        'severity': severity,
        'checks': checks,
    }


# ============================================================
# 批量提取 + 完整性报告
# ============================================================

def batch_extract(filepaths: List[str], output_dir: str = None) -> Dict[str, Any]:
    """批量提取PDF/图片文件内容，生成DT-108完整性报告。

    Args:
        filepaths: 文件路径列表
        output_dir: 提取结果JSON输出目录（None=不保存）

    Returns:
        dict: {
            'total': int,
            'extracted': int,
            'partial': int,
            'failed': int,
            'results': [{filepath, filename, type, status, summary, warnings}],
            'report_text': str,   # 可直接输出的完整性报告
        }
    """
    results = []
    type_counts = {'bank': 0, 'asset': 0, 'auxiliary': 0, 'other': 0}
    status_counts = {'extracted': 0, 'partial': 0, 'failed': 0}

    for fp in filepaths:
        ext = os.path.splitext(fp)[1].lower()
        if ext not in ('.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.tiff'):
            continue

        # 根据文件名预判类型
        fname = os.path.basename(fp).lower()
        if any(kw in fname for kw in ['对账单', '银行', '存款']):
            ftype = 'bank'
            raw = extract_bank_statement(fp)
            summary = f"{raw.get('bank_name', '?')} 账号{raw.get('account_no', '?')} 余额{raw.get('ending_balance', '?')}"
        elif any(kw in fname for kw in ['卡片', '台账', '固定资产']):
            ftype = 'asset'
            raw = extract_asset_register(fp)
            summary = f"{raw.get('total_items', 0)}项资产"
        elif any(kw in fname for kw in ['辅助余额', '辅助明细', '辅助账']):
            ftype = 'auxiliary'
            raw = extract_auxiliary_balance(fp)
            summary = f"{len(raw.get('counterparties', []))}个结算对象"
        else:
            ftype = 'other'
            raw = extract_pdf(fp)
            summary = f"文本{len(raw.get('text', ''))}字符, 表格{len(raw.get('tables', []))}个"

        type_counts[ftype] += 1

        # 统一状态
        status = raw.get('status', 'failed')
        if status == 'ocr_fallback':
            status = 'extracted'  # OCR成功也算提取成功
        status_counts[status] = status_counts.get(status, 0) + 1

        results.append({
            'filepath': fp,
            'filename': os.path.basename(fp),
            'type': ftype,
            'status': status,
            'summary': summary,
            'warnings': raw.get('warnings', []),
        })

        # 保存单个文件提取结果
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            json_name = os.path.splitext(os.path.basename(fp))[0] + '.json'
            json_path = os.path.join(output_dir, json_name)
            save_extraction_json(raw, json_path)

    # 生成完整性报告
    total = len(results)
    extracted = sum(1 for r in results if r['status'] == 'extracted')
    partial = sum(1 for r in results if r['status'] == 'partial')
    failed = sum(1 for r in results if r['status'] == 'failed')

    report_lines = [
        "=" * 60,
        "PDF/图片提取完整性报告 (DT-108)",
        "=" * 60,
        f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"总文件数: {total}",
        "",
        "| 类别 | 文件数 | 成功 | 部分成功 | 失败 |",
        "|------|--------|------|----------|------|",
    ]

    for ftype, label in [('bank', '银行对账单'), ('asset', '卡片台账'),
                          ('auxiliary', '辅助余额表'), ('other', '其他')]:
        type_results = [r for r in results if r['type'] == ftype]
        if not type_results:
            continue
        t_total = len(type_results)
        t_ext = sum(1 for r in type_results if r['status'] == 'extracted')
        t_par = sum(1 for r in type_results if r['status'] == 'partial')
        t_fail = sum(1 for r in type_results if r['status'] == 'failed')
        report_lines.append(f"| {label} | {t_total} | {t_ext} | {t_par} | {t_fail} |")

    report_lines.extend([
        "",
        f"✅ 成功提取: {extracted}",
        f"⚠️ 部分成功: {partial}",
        f"❌ 提取失败: {failed}",
        "",
    ])

    if failed > 0:
        report_lines.append("🚨 提取失败文件：")
        for r in results:
            if r['status'] == 'failed':
                report_lines.append(f"  - {r['filename']}: {', '.join(r['warnings'])}")
        report_lines.append("")
        report_lines.append("🚨 CRITICAL: 未提取文件数>0，禁止进入Phase 0")
    else:
        report_lines.append("✅ 全部PDF/图片提取完毕，可进入Phase 0")

    report_text = '\n'.join(report_lines)

    return {
        'total': total,
        'extracted': extracted,
        'partial': partial,
        'failed': failed,
        'results': results,
        'report_text': report_text,
    }


# ============================================================
# 提取结果持久化
# ============================================================

def save_extraction_json(extraction_result: Dict[str, Any], output_path: str) -> None:
    """将提取结果保存为JSON文件。

    Args:
        extraction_result: 提取结果dict
        output_path: 输出JSON文件路径
    """
    # 确保可序列化
    serializable = _make_serializable(extraction_result)

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(serializable, f, ensure_ascii=False, indent=2)


def load_extraction_json(json_path: str) -> Dict[str, Any]:
    """加载已保存的提取结果JSON。

    Args:
        json_path: JSON文件路径

    Returns:
        dict: 提取结果
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ============================================================
# 辅助函数
# ============================================================

def _find_col(headers: List[str], keywords: List[str]) -> Optional[int]:
    """在表头列表中查找包含关键词的列索引。

    [DT-154] DEPRECATED: 本函数已委托给source_header_parser.find_col_by_keywords()
    保留接口兼容性，新代码请直接使用source_header_parser
    """
    try:
        from source_header_parser import find_col_by_keywords
        return find_col_by_keywords(headers, keywords)
    except ImportError:
        pass  # fallback到原实现

    for i, h in enumerate(headers):
        h_clean = h.replace(' ', '').strip()
        for kw in keywords:
            if kw in h_clean:
                return i
    return None


def _parse_number(s: Any) -> Optional[float]:
    """将字符串解析为浮点数。支持千分位逗号和中文数字。"""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace(',', '').replace('，', '').replace('￥', '').replace('¥', '')
    if not s or s == '-':
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_int(s: Any) -> Optional[int]:
    """将字符串解析为整数。"""
    val = _parse_number(s)
    return int(val) if val is not None else None


def _make_serializable(obj: Any) -> Any:
    """确保dict可JSON序列化（处理datetime等特殊类型）。"""
    if isinstance(obj, dict):
        return {k: _make_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_make_serializable(v) for v in obj]
    elif isinstance(obj, datetime):
        return obj.isoformat()
    elif isinstance(obj, (int, float, str, bool, type(None))):
        return obj
    else:
        return str(obj)


# ============================================================
# CLI入口
# ============================================================

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='PDF/图片内容提取工具 v1.0')
    parser.add_argument('path', help='PDF/图片文件路径，或包含PDF/图片的文件夹路径')
    parser.add_argument('--mode', choices=['auto', 'bank', 'asset', 'auxiliary', 'general'],
                       default='auto',
                       help='提取模式: auto=自动判断, bank=银行对账单, asset=卡片台账, auxiliary=辅助余额表, general=通用')
    parser.add_argument('--output-dir', help='提取结果JSON输出目录')
    parser.add_argument('--no-ocr', action='store_true', help='禁用OCR兜底')
    parser.add_argument('--ocr-lang', default='chi_sim+eng', help='OCR语言 (默认chi_sim+eng)')
    parser.add_argument('--batch', action='store_true', help='批量模式（path为文件夹）')

    args = parser.parse_args()

    if args.batch or os.path.isdir(args.path):
        # 批量模式
        filepaths = []
        for root, dirs, files in os.walk(args.path):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in ('.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.tiff'):
                    filepaths.append(os.path.join(root, f))

        print(f"扫描到 {len(filepaths)} 个PDF/图片文件")
        result = batch_extract(filepaths, args.output_dir)
        print(result['report_text'])

        if args.output_dir:
            report_path = os.path.join(args.output_dir, 'extraction_report.txt')
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(result['report_text'])
            print(f"\n报告已保存: {report_path}")

    else:
        # 单文件模式
        if not os.path.exists(args.path):
            print(f"错误: 文件不存在 - {args.path}")
            sys.exit(1)

        if args.mode == 'bank':
            result = extract_bank_statement(args.path)
        elif args.mode == 'asset':
            result = extract_asset_register(args.path)
        elif args.mode == 'auxiliary':
            result = extract_auxiliary_balance(args.path)
        else:
            # auto模式：根据文件名判断
            fname = os.path.basename(args.path).lower()
            if any(kw in fname for kw in ['对账单', '银行', '存款']):
                result = extract_bank_statement(args.path)
            elif any(kw in fname for kw in ['卡片', '台账', '固定资产']):
                result = extract_asset_register(args.path)
            elif any(kw in fname for kw in ['辅助余额', '辅助明细']):
                result = extract_auxiliary_balance(args.path)
            else:
                result = extract_pdf(args.path, use_ocr_fallback=not args.no_ocr)

        # 输出结果
        print(f"\n提取状态: {result['status']}")
        if result.get('warnings'):
            print(f"警告: {result['warnings']}")

        # 场景化输出
        if 'bank_name' in result:
            print(f"\n银行对账单提取结果:")
            print(f"  银行名称: {result.get('bank_name', '未识别')}")
            print(f"  账户名称: {result.get('account_name', '未识别')}")
            print(f"  账号: {result.get('account_no', '未识别')}")
            print(f"  期末余额: {result.get('ending_balance', '未识别')}")
            print(f"  币种: {result.get('currency', 'CNY')}")
        elif 'items' in result:
            print(f"\n固定资产卡片台账提取结果:")
            print(f"  资产项数: {result.get('total_items', 0)}")
            for i, item in enumerate(result.get('items', [])[:5]):
                print(f"  [{i+1}] {item.get('name', '?')} "
                      f"原值={item.get('original_value', '?')} "
                      f"净值={item.get('net_value', '?')}")
            if result.get('total_items', 0) > 5:
                print(f"  ... 共{result['total_items']}项")
        elif 'counterparties' in result:
            print(f"\n辅助余额表提取结果:")
            print(f"  科目名称: {result.get('subject_name', '未识别')}")
            print(f"  结算对象数: {len(result.get('counterparties', []))}")
            for i, cp in enumerate(result.get('counterparties', [])[:5]):
                print(f"  [{i+1}] {cp.get('name', '?')} "
                      f"借方={cp.get('debit_balance', '?')} "
                      f"贷方={cp.get('credit_balance', '?')}")
        else:
            # 通用模式
            text = result.get('text', '')
            print(f"\n通用提取结果:")
            print(f"  页数: {result.get('pages', 0)}")
            print(f"  文本长度: {len(text)}字符")
            print(f"  表格数: {len(result.get('tables', []))}")
            if text:
                print(f"\n  前200字符预览:")
                print(f"  {text[:200]}")

        # 保存JSON
        if args.output_dir:
            os.makedirs(args.output_dir, exist_ok=True)
            json_name = os.path.splitext(os.path.basename(args.path))[0] + '.json'
            json_path = os.path.join(args.output_dir, json_name)
            save_extraction_json(result, json_path)
            print(f"\n提取结果已保存: {json_path}")

    # DT-108验证
    print(f"\n--- DT-108 验证 ---")
    if isinstance(result, dict) and 'results' in result:
        # 批量模式
        for r in result['results']:
            print(f"  {r['filename']}: {r['status']} - {r['summary']}")
    else:
        validation = validate_extraction(result)
        for check in validation['checks']:
            icon = '✅' if check['passed'] else '❌'
            print(f"  {icon} {check['name']}: {check['detail']}")
        print(f"  最终判定: {validation['severity']}")
