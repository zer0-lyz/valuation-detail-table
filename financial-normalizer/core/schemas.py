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
    # v0.2 (2026-06-01): 中汇查账系统 16 列布局新增字段
    "entity_name": {
        "type": "str",
        "required": False,
        "label": "企业主体名称",
        "desc": "余额表所属企业",
    },
    "query_period": {
        "type": "str",
        "required": False,
        "label": "查询区间",
        "desc": "期间范围，如 2025-01~2025-12",
    },
    "auxiliary_type": {
        "type": "str",
        "required": False,
        "label": "核算类型",
        "desc": "辅助核算维度类别,如 银行账户/客户/供应商",
    },
    "auxiliary_code": {
        "type": "str",
        "required": False,
        "label": "核算编号",
        "desc": "辅助核算对象的编号",
    },
    "auxiliary_name": {
        "type": "str",
        "required": False,
        "label": "核算名称",
        "desc": "辅助核算对象的名称",
    },
    "data_type": {
        "type": "str",
        "required": False,
        "label": "数据类型",
        "desc": "本位币/原币/数量",
    },
    "direction": {
        "type": "str",
        "required": False,
        "label": "余额方向",
        "desc": "借/贷",
    },
    "account_full_path": {
        "type": "str",
        "required": False,
        "label": "科目全路径",
        "desc": "多级科目合并路径,如 库存现金/人民币现金",
    },
    "pnl_carryover": {
        "type": "number",
        "required": False,
        "label": "损益结转金额",
        "desc": "期末损益类结转到本年利润的金额",
    },
    "standard_level1": {
        "type": "str",
        "required": False,
        "label": "标准1级科目",
        "desc": "按标准会计制度归类的 1 级科目",
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
    # v0.2 (2026-06-01): 中汇查账系统 50 列布局新增字段
    "entity_name": {
        "type": "str",
        "required": False,
        "label": "企业主体名称",
        "desc": "凭证所属企业",
    },
    "voucher_type": {
        "type": "str",
        "required": False,
        "label": "凭证类别(字)",
        "desc": "记/收/付/转 等",
    },
    "voucher_number": {
        "type": "str",
        "required": False,
        "label": "凭证号数",
        "desc": "每类凭证自增序号",
    },
    "counter_account": {
        "type": "str",
        "required": False,
        "label": "对方科目",
        "desc": "本笔凭证对方科目全路径",
    },
    "other_account": {
        "type": "str",
        "required": False,
        "label": "本方其他科目",
        "desc": "同一凭证中本方其他科目",
    },
    "direction": {
        "type": "str",
        "required": False,
        "label": "方向",
        "desc": "本笔业务后该科目的余额方向",
    },
    "balance": {
        "type": "number",
        "required": False,
        "label": "本位币余额",
        "desc": "本笔业务后该科目的本位币余额",
    },
    "allocation_type": {
        "type": "str",
        "required": False,
        "label": "结转分配类型",
        "desc": "损益结转/费用分配 类型",
    },
    "subject_code": {
        "type": "str",
        "required": False,
        "label": "末级科目编号",
        "desc": "本行所属末级科目编号",
    },
    "subject_full_path": {
        "type": "str",
        "required": False,
        "label": "末级科目全路径",
        "desc": "本行所属末级科目完整路径",
    },
    "query_subject": {
        "type": "str",
        "required": False,
        "label": "查询科目",
        "desc": "查询入口的科目",
    },
    "query_auxiliary": {
        "type": "str",
        "required": False,
        "label": "查询核算",
        "desc": "查询入口的辅助核算对象",
    },
    "entry_line_no": {
        "type": "str",
        "required": False,
        "label": "分录行号",
        "desc": "同一凭证内的分录行序号",
    },
    "customer_supplier_code": {
        "type": "str",
        "required": False,
        "label": "往来单位编号",
        "desc": "往来单位的本方编号",
    },
    "counter_customer_supplier": {
        "type": "str",
        "required": False,
        "label": "对方往来单位名称",
        "desc": "双抬头场景对方往来单位",
    },
    "bank_account_code": {
        "type": "str",
        "required": False,
        "label": "银行账号编号",
        "desc": "银行账号本方编号",
    },
    "bank_account_name": {
        "type": "str",
        "required": False,
        "label": "银行账号名称",
        "desc": "银行账号本方名称",
    },
    "counter_bank_account": {
        "type": "str",
        "required": False,
        "label": "对方银行账号名称",
        "desc": "双抬头场景对方银行账号",
    },
    "purchase_contract_code": {
        "type": "str",
        "required": False,
        "label": "采购合同编号",
        "desc": "采购合同本方编号",
    },
    "purchase_contract_name": {
        "type": "str",
        "required": False,
        "label": "采购合同名称",
        "desc": "采购合同本方名称",
    },
    "counter_purchase_contract": {
        "type": "str",
        "required": False,
        "label": "对方采购合同名称",
        "desc": "双抬头场景对方采购合同",
    },
    "sales_contract_code": {
        "type": "str",
        "required": False,
        "label": "销售合同编号",
        "desc": "销售合同本方编号",
    },
    "sales_contract_name": {
        "type": "str",
        "required": False,
        "label": "销售合同名称",
        "desc": "销售合同本方名称",
    },
    "counter_sales_contract": {
        "type": "str",
        "required": False,
        "label": "对方销售合同名称",
        "desc": "双抬头场景对方销售合同",
    },
    "project_code": {
        "type": "str",
        "required": False,
        "label": "项目编号",
        "desc": "项目本方编号",
    },
    "counter_project": {
        "type": "str",
        "required": False,
        "label": "对方项目名称",
        "desc": "双抬头场景对方项目",
    },
    "asset_code": {
        "type": "str",
        "required": False,
        "label": "资产项目编号",
        "desc": "资产项目本方编号",
    },
    "asset_name": {
        "type": "str",
        "required": False,
        "label": "资产项目名称",
        "desc": "资产项目本方名称",
    },
    "counter_asset": {
        "type": "str",
        "required": False,
        "label": "对方资产项目名称",
        "desc": "双抬头场景对方资产",
    },
    "material_code": {
        "type": "str",
        "required": False,
        "label": "存货物料编号",
        "desc": "存货物料本方编号",
    },
    "material_name": {
        "type": "str",
        "required": False,
        "label": "存货物料名称",
        "desc": "存货物料本方名称",
    },
    "counter_material": {
        "type": "str",
        "required": False,
        "label": "对方存货物料名称",
        "desc": "双抬头场景对方物料",
    },
    "expense_code": {
        "type": "str",
        "required": False,
        "label": "费用项目编号",
        "desc": "费用项目本方编号",
    },
    "expense_name": {
        "type": "str",
        "required": False,
        "label": "费用项目名称",
        "desc": "费用项目本方名称",
    },
    "counter_expense": {
        "type": "str",
        "required": False,
        "label": "对方费用项目名称",
        "desc": "双抬头场景对方费用",
    },
    "department_code": {
        "type": "str",
        "required": False,
        "label": "部门编号",
        "desc": "部门本方编号",
    },
    "counter_department": {
        "type": "str",
        "required": False,
        "label": "对方部门名称",
        "desc": "双抬头场景对方部门",
    },
    "cash_flow_code": {
        "type": "str",
        "required": False,
        "label": "现金流量编号",
        "desc": "现金流量本方编号",
    },
    "cash_flow_name": {
        "type": "str",
        "required": False,
        "label": "现金流量名称",
        "desc": "现金流量本方名称",
    },
    "counter_cash_flow": {
        "type": "str",
        "required": False,
        "label": "对方现金流量名称",
        "desc": "双抬头场景对方现金流",
    },
    "other_code": {
        "type": "str",
        "required": False,
        "label": "其他编号",
        "desc": "自定义辅助核算本方编号",
    },
    "other_name": {
        "type": "str",
        "required": False,
        "label": "其他名称",
        "desc": "自定义辅助核算本方名称",
    },
    "counter_other": {
        "type": "str",
        "required": False,
        "label": "对方其他名称",
        "desc": "双抬头场景对方自定义辅助",
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
