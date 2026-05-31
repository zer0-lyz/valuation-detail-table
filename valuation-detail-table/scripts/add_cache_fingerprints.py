#!/usr/bin/env python3
"""add_cache_fingerprints.py — r10: 给关键缓存文件补充幂等指纹"""

import json
import os
import hashlib
from datetime import datetime

VERSION = 'r9'
CACHE_DIR = os.path.dirname(os.path.abspath(__file__))


def fingerprint_file(path: str) -> str:
    """计算文件SHA256"""
    if not os.path.exists(path):
        return 'file_missing'
    sha = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            sha.update(chunk)
    return sha.hexdigest()[:16]


def add_fingerprints(cache_path: str, source_files: list = None):
    """给缓存JSON文件添加幂等指纹"""
    if not os.path.exists(cache_path):
        return
    with open(cache_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 包装为带指纹的格式
    if isinstance(data, dict) and '_meta' not in data:
        data = {
            '_meta': {
                'schema_version': '1.0',
                'producer_version': VERSION,
                'generated_at': datetime.now().isoformat(),
                'source_fingerprints': {os.path.basename(s): fingerprint_file(s) for s in (source_files or []) if os.path.exists(s)},
            },
            'data': data
        }

    with open(cache_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'  ✅ {os.path.basename(cache_path)}: fingerprints added')


if __name__ == '__main__':
    import sys
    cache_dir = sys.argv[1] if len(sys.argv) > 1 else CACHE_DIR
    source_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(cache_dir)

    # 源文件
    source_files = []
    for pat in ('*.xlsx', '*.xls'):
        for p in __import__('glob').glob(os.path.join(source_dir, pat)):
            source_files.append(p)

    # 关键缓存
    for name in ('subjects.json', 'bs_balances.json', 'reclassification.json',
                 'asset_register_by_sheet.json', 'subledger_standardized.json'):
        path = os.path.join(cache_dir, name)
        if os.path.exists(path):
            add_fingerprints(path, source_files)
