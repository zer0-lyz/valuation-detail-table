"""
标准 Schema 定义
================
每种财务资料的标准字段结构。
字段命名约定：全小写+下划线，借贷用 debit/credit 区分方向。
"""

# ==================== 科目余额表标准 Schema ====================
TRIAL_BALANCE_SCHEMA = {
    "account_code": {
        "type": "str",
        "required": True,
        "label": "科目编码",
        "desc": "如 1001、1002-01",
    },
    "account_name": {
        "type": "str",
        "required": True,
        "label": "科目名称",
        "desc": "标准科目名称",
    },
    "currency": {
        "type": "str",
        "required": False,
        "label": "币种",
        "desc": "货币种类，如人民币、美元、欧元",
    },
    "level1_name": {
        "type": "str",
        "required": False,
        "label": "一级科目名称",
        "desc": "一级明细科目名称，如'银行存款'",
    },
    "level2_name": {
        "type": "str",
        "required": False,
        "label": "二级科目名称",
        "desc": "二级明细科目名称，如'工商银行'",
    },
    "level3_name": {
        "type": "str",
        "required": False,
        "label": "三级科目名称",
        "desc": "三级明细科目名称，如'北京分行'",
    },
    "opening_debit": {
        "type": "number",
        "required": False,
        "label": "期初借方余额",
        "desc": "年初/月初借方余额",
    },
    "opening_credit": {
        "type": "number",
        "required": False,
        "label": "期初贷方余额",
        "desc": "年初/月初贷方余额",
    },
    "current_debit": {
        "type": "number",
        "required": False,
        "label": "本期借方发生",
        "desc": "本期借方发生额",
    },

    "opening_balance": {
        "type": "number",
        "required": False,
        "label": "期初余额（不分方向）",
        "desc": "年初/月初余额（无需区分借贷）",
    },
    "current_credit": {
        "type": "number",
        "required": False,
        "label": "本期贷方发生",
        "desc": "本期贷方发生额",
    },

    "closing_balance": {
        "type": "number",
        "required": False,
        "label": "期末余额（不分方向）",
        "desc": "年末/月末余额（无需区分借贷）",
    },
    "closing_debit": {
        "type": "number",
        "required": False,
        "label": "期末借方余额",
        "desc": "年末/月末借方余额",
    },
    "closing_credit": {
        "type": "number",
        "required": False,
        "label": "期末贷方余额",
        "desc": "年末/月末贷方余额",
    },
}

# ==================== 序时账标准 Schema ====================
JOURNAL_SCHEMA = {
    "date": {
        "type": "date",
        "required": True,
        "label": "日期",
        "desc": "凭证日期",
    },
    "voucher_no": {
        "type": "str",
        "required": False,
        "label": "凭证号",
        "desc": "如 记-1、银付-001",
    },
    "account_code": {
        "type": "str",
        "required": True,
        "label": "科目编码",
        "desc": "如 1001、1122",
    },
    "account_name": {
        "type": "str",
        "required": True,
        "label": "科目名称",
        "desc": "标准科目名称",
    },
    "currency": {
        "type": "str",
        "required": False,
        "label": "币种",
        "desc": "货币种类，如人民币、美元、欧元",
    },
    "summary": {
        "type": "str",
        "required": False,
        "label": "摘要",
        "desc": "业务摘要说明",
    },
    "debit_amount": {
        "type": "number",
        "required": False,
        "label": "借方金额",
        "desc": "借方发生额",
    },
    "credit_amount": {
        "type": "number",
        "required": False,
        "label": "贷方金额",
        "desc": "贷方发生额",
    },
    "customer_supplier": {
        "type": "str",
        "required": False,
        "label": "往来单位",
        "desc": "客户/供应商/往来单位名称",
    },
    "department": {
        "type": "str",
        "required": False,
        "label": "部门",
        "desc": "所属部门",
    },
    "project_name": {
        "type": "str",
        "required": False,
        "label": "项目",
        "desc": "所属项目/工程",
    },
    "personnel": {
        "type": "str",
        "required": False,
        "label": "经办人",
        "desc": "经手人/报销人/负责人",
    },
}

# ==================== 资产负债表标准 Schema ====================
BALANCE_SHEET_SCHEMA = {
    "item_name": {
        "type": "str",
        "required": True,
        "label": "项目名称",
        "desc": "报表行项目名称",
    },
    "currency": {
        "type": "str",
        "required": False,
        "label": "币种",
        "desc": "货币种类，如人民币、美元、欧元",
    },
    "item_category": {
        "type": "str",
        "required": False,
        "label": "项目类别",
        "desc": "资产类/负债类/所有者权益类/损益类",
    },
    "item_direction": {
        "type": "str",
        "required": False,
        "label": "方向",
        "desc": "借/贷，科目正常余额方向",
    },
    "closing_balance": {
        "type": "number",
        "required": False,
        "label": "期末余额",
        "desc": "期末数",
    },
    "opening_balance": {
        "type": "number",
        "required": False,
        "label": "期初余额",
        "desc": "上年年末数/期初数",
    },
}


# ==================== 固定资产台账标准 Schema ====================
FIXED_ASSET_SCHEMA = {
    "asset_code": {
        "type": "str",
        "required": True,
        "label": "资产编码",
        "desc": "资产卡片编号或资产编码",
    },
    "asset_name": {
        "type": "str",
        "required": True,
        "label": "资产名称",
        "desc": "固定资产名称",
    },
    "currency": {
        "type": "str",
        "required": False,
        "label": "币种",
        "desc": "货币种类，如人民币、美元、欧元",
    },
    "asset_category": {
        "type": "str",
        "required": False,
        "label": "资产类别",
        "desc": "如房屋建筑物、机器设备、运输设备、电子设备等",
    },
    "specification": {
        "type": "str",
        "required": False,
        "label": "规格型号",
        "desc": "规格/型号/技术参数",
    },
    "quantity": {
        "type": "number",
        "required": False,
        "label": "数量",
        "desc": "同规格资产数量",
    },
    "unit": {
        "type": "str",
        "required": False,
        "label": "单位",
        "desc": "计量单位，如台、套、辆、栋、平方米",
    },
    "acquisition_date": {
        "type": "date",
        "required": False,
        "label": "取得日期",
        "desc": "入账日期/购置日期",
    },
    "original_value": {
        "type": "number",
        "required": True,
        "label": "原值",
        "desc": "固定资产原值（入账价值）",
    },
    "accumulated_depreciation": {
        "type": "number",
        "required": True,
        "label": "累计折旧",
        "desc": "截至评估基准日的累计折旧",
    },
    "impairment_amount": {
        "type": "number",
        "required": False,
        "label": "减值准备",
        "desc": "固定资产减值准备余额",
    },
    "net_value": {
        "type": "number",
        "required": False,
        "label": "净值",
        "desc": "原值减去累计折旧后的余额",
    },
    "department": {
        "type": "str",
        "required": False,
        "label": "使用部门",
        "desc": "使用/管理部门",
    },
    "location": {
        "type": "str",
        "required": False,
        "label": "存放地点",
        "desc": "资产存放或使用地点",
    },
    "supplier": {
        "type": "str",
        "required": False,
        "label": "供应商",
        "desc": "供应商/生产厂商名称",
    },
    "depreciation_method": {
        "type": "str",
        "required": False,
        "label": "折旧方法",
        "desc": "如平均年限法、双倍余额递减法、年数总和法等",
    },
    "depreciation_life": {
        "type": "number",
        "required": False,
        "label": "折旧年限",
        "desc": "预计使用年限（年）",
    },
    "residual_rate": {
        "type": "number",
        "required": False,
        "label": "残值率",
        "desc": "预计净残值率（百分比）",
    },
    "monthly_depreciation": {
        "type": "number",
        "required": False,
        "label": "月折旧额",
        "desc": "每月计提折旧金额",
    },
    "status": {
        "type": "str",
        "required": False,
        "label": "资产状态",
        "desc": "在用/停用/报废/出租等",
    },
    "start_date": {
        "type": "date",
        "required": False,
        "label": "启用日期",
        "desc": "开始使用/开始计提折旧日期",
    },
}

# ==================== 利润表标准 Schema ====================
INCOME_STATEMENT_SCHEMA = {
    "item_name": {
        "type": "str",
        "required": True,
        "label": "项目名称",
        "desc": "报表行项目名称",
    },
    "currency": {
        "type": "str",
        "required": False,
        "label": "币种",
        "desc": "货币种类，如人民币、美元、欧元",
    },
    "item_category": {
        "type": "str",
        "required": False,
        "label": "项目类别",
        "desc": "资产类/负债类/所有者权益类/损益类",
    },
    "item_direction": {
        "type": "str",
        "required": False,
        "label": "方向",
        "desc": "借/贷，科目正常余额方向",
    },
    "current_period": {
        "type": "number",
        "required": False,
        "label": "本期金额",
        "desc": "本期发生额",
    },
    "cumulative": {
        "type": "number",
        "required": False,
        "label": "本年累计",
        "desc": "年初至本期累计",
    },
}

# ==================== 文档类型注册表 ====================
DOCUMENT_TYPES = {
    "trial_balance": {
        "name": "科目余额表",
        "schema": TRIAL_BALANCE_SCHEMA,
        "keywords": ["科目余额", "试算平衡", "科目汇总", "trial balance"],
        "sheet_pattern": ["科目余额", "Sheet1", "汇总"],
    },
    "journal": {
        "name": "序时账",
        "schema": JOURNAL_SCHEMA,
        "keywords": ["序时账", "明细账", "日记账", "凭证", "journal"],
        "sheet_pattern": ["序时账", "明细", "Sheet1"],
    },
    "balance_sheet": {
        "name": "资产负债表",
        "schema": BALANCE_SHEET_SCHEMA,
        "keywords": ["资产负债表", "资产", "balance sheet", "负债表"],
        "sheet_pattern": ["资产负债表", "资产", "Sheet1"],
    },
    "fixed_asset": {
        "name": "固定资产台账",
        "schema": FIXED_ASSET_SCHEMA,
        "keywords": ["固定资产", "固定资产台账", "资产台账", "固定资产明细", "资产卡片", "fixed asset", "FA ledger"],
        "sheet_pattern": ["固定资产", "资产台账", "固定", "FA", "Sheet1"],
    },
        "income_statement": {
        "name": "利润表",
        "schema": INCOME_STATEMENT_SCHEMA,
        "keywords": ["利润表", "损益表", "income statement", "利润"],
        "sheet_pattern": ["利润表", "损益", "Sheet1"],
    },
}

# Schema 字段名列表（用于快速索引）
SCHEMA_FIELDS = {
    name: list(doc["schema"].keys())
    for name, doc in DOCUMENT_TYPES.items()
}
