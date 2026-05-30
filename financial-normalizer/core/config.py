"""
映射配置管理
============
Guess 阶段生成映射配置（JSON），Apply 阶段加载使用。
用户可编辑这个 JSON 来修正映射关系。
"""

import json
import os
from pathlib import Path
from datetime import datetime


DEFAULT_CONFIG_DIR = "mapping_configs"


def _get_schemas():
    """懒加载 schemas 避免循环引用"""
    from core import schemas
    return schemas


def generate_config_path(source_file: str, doc_type: str = None) -> str:
    """根据源文件名生成映射配置文件名"""
    src = Path(source_file)
    stem = src.stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(DEFAULT_CONFIG_DIR, exist_ok=True)
    return f"{DEFAULT_CONFIG_DIR}/{stem}_{timestamp}_mapping.json"


def save_config(config: dict, path: str = None) -> str:
    """保存映射配置到 JSON 文件"""
    if path is None:
        path = generate_config_path(config.get("_source_file", "unknown"))
    
    config.setdefault("_meta", {})
    config["_meta"]["saved_at"] = datetime.now().isoformat()
    
    os.makedirs(Path(path).parent, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    return path


def load_config(path: str) -> dict:
    """加载映射配置 JSON 文件"""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def create_blank_mapping(source_file: str, doc_type: str) -> dict:
    """创建一个空白映射配置模板"""
    schemas = _get_schemas()
    
    doc_info = schemas.DOCUMENT_TYPES.get(doc_type, {})
    schema = doc_info.get("schema", {})
    
    mapping = {
        "_meta": {
            "version": "0.1.0",
            "description": "自动生成的映射配置，编辑后通过 apply 使用",
            "instructions": "将 source 列改为对应的标准字段名。确认无误后将 confirmed 设为 true",
        },
        "_source_file": source_file,
        "_detected_type": doc_type,
        "_detected_type_name": doc_info.get("name", doc_type),
        "sheet_name": None,
        "header_row": None,
        "data_start_row": None,
        "column_mapping": {},
        "options": {
            "skip_rows": [],
        },
        "confirmed": False,
    }
    
    return mapping


def print_config_summary(config: dict):
    """打印映射配置摘要，方便用户检查"""
    print("=" * 60)
    print(f"📄 源文件  : {config.get('_source_file', '?')}")
    print(f"📋 检测类型 : {config.get('_detected_type_name', config.get('_detected_type', '?'))}")
    print(f"📑 Sheet    : {config.get('sheet_name', '?')}")
    print(f"🔢 表头行   : {config.get('header_row', '?')}")
    print(f"🔢 数据起始 : {config.get('data_start_row', '?')}")
    print("-" * 60)
    print("📊 列映射关系:")
    col_map = config.get("column_mapping", {})
    if col_map:
        for src, tgt in col_map.items():
            print(f"   [{src:20s}]  →  [{tgt}]")
    else:
        print("   (空 - 未检测到列映射)")
    print("-" * 60)
    print(f"✅ 已确认   : {'是' if config.get('confirmed') else '否 - 请检查并修改后设 confirmed=true'}")
    print("=" * 60)
