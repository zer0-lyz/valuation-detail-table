"""
数据映射器
==========
根据映射配置，将源 Excel 数据转换为标准 Schema 格式。

支持双行表头：当 mapping 配置中 _two_row_header=True 时，
直接使用配置中的 column_index_map 定位列（跳过表头名匹配）。
"""

import pandas as pd
import openpyxl
import math
from pathlib import Path
from datetime import datetime


def apply(config: dict) -> pd.DataFrame:
    """
    根据映射配置，读取源文件并输出标准化 DataFrame。
    """
    source_file = config.get("_source_file")
    sheet_name = config.get("sheet_name")
    header_row = config.get("header_row")
    data_start_row = config.get("data_start_row")
    column_mapping = config.get("column_mapping", {})
    options = config.get("options", {})
    is_two_row = config.get("_two_row_header", False)

    if not config.get("confirmed"):
        raise ValueError("映射配置未确认！请检查并修改后设 'confirmed': true")

    # 读取原始数据
    df = pd.read_excel(
        source_file,
        sheet_name=sheet_name,
        header=None,
        skiprows=list(range(data_start_row)) if data_start_row else None,
    )

    skip_rows = options.get("skip_rows", [])
    if skip_rows:
        df = df.drop(index=[r for r in skip_rows if r < len(df)], errors="ignore")

    # 构建 {标准字段名: 列索引} 映射
    field_to_col = {}

    if is_two_row and config.get("_column_index_map"):
        # 双行表头模式：直接用 column_index_map（detector 已算好）
        col_index_map = config["_column_index_map"]
        for std_field, col_idx in col_index_map.items():
            if not std_field.startswith("__unknown"):
                field_to_col[std_field] = col_idx
    else:
        # 单行表头模式：读表头行获取列名→索引映射
        header_df = pd.read_excel(
            source_file,
            sheet_name=sheet_name,
            header=None,
            skiprows=list(range(header_row)),
            nrows=1,
        )
        header_names = header_df.iloc[0].astype(str).tolist()

        for i, h in enumerate(header_names):
            h_clean = h.strip()
            if h_clean in column_mapping:
                target = column_mapping[h_clean]
                if not target.startswith("__unknown"):
                    field_to_col[target] = i

    # 检查必要字段
    from core import schemas
    doc_type = config.get("_detected_type", "trial_balance")
    schema = schemas.DOCUMENT_TYPES.get(doc_type, {}).get("schema", {})

    required_fields = [k for k, v in schema.items() if v.get("required")]
    missing_required = [f for f in required_fields if f not in field_to_col]
    if missing_required:
        raise ValueError(f"缺少必要字段映射: {missing_required}")

    # 构建标准 DataFrame
    result = pd.DataFrame()
    for field_name in schema.keys():
        if field_name in field_to_col:
            col_idx = field_to_col[field_name]
            result[field_name] = df.iloc[:, col_idx]
        else:
            result[field_name] = None

    # 类型转换
    for field_name, field_info in schema.items():
        if field_info["type"] == "number" and field_name in result.columns:
            # 清理千分位逗号等格式再转换
            cleaned = result[field_name].astype(str).str.replace(',', '').str.replace(' ', '')
            result[field_name] = pd.to_numeric(cleaned, errors="coerce")

    # 清理：去掉全空行
    string_cols = [k for k, v in schema.items() if v["type"] in ("str", "date")]
    if string_cols:
        valid_rows = result[string_cols].notna().any(axis=1)
        result = result.loc[valid_rows].reset_index(drop=True)

    return result


def apply_and_save(config: dict, output_path: str = None):
    """Apply 映射配置并保存为标准格式。"""
    df = apply(config)

    if output_path is None:
        src = Path(config.get("_source_file", "unknown"))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"output/{src.stem}_standardized_{timestamp}.xlsx"

    import os
    os.makedirs(Path(output_path).parent, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        from core import schemas
        doc_type = config.get("_detected_type", "unknown")
        doc_name = schemas.DOCUMENT_TYPES.get(doc_type, {}).get("name", doc_type)
        sheet_label = f"{doc_name}(标准化)"
        df.to_excel(writer, sheet_name=sheet_label[:31], index=False)

    json_path = output_path.replace(".xlsx", ".json")
    df.to_json(json_path, orient="records", force_ascii=False, indent=2)

    print(f"✅ 标准化完成!")
    print(f"   Excel: {output_path}")
    print(f"   JSON : {json_path}")
    print(f"   行数  : {len(df)}")

    return output_path
