# DT Skill 字段命名规范 (DT-NAMING)

> 版本: v1.1 | 更新: 2026-06-01 | 规则: DT-153v2
>
> **v1.1 变更**: 新增 2.6/2.7 节说明 financial-normalizer 标准化字段(对齐胡庆余堂集团 16/50 列结构)→ 内部 pipeline 字段 的命名规范。详见 `assets/field_mapping.json` 和 `financial-normalizer/FIELD_STANDARD.md`。

## 1. 核心原则

**单一数据源原则**: `sheet_col_map.json` 是字段命名的唯一权威来源。所有producer/consumer脚本MUST与之一致。

**flat键名原则**: sheet_col_map.json中的col_map使用flat键名（如`date`/`business`/`settlement`），禁止使用嵌套命名（如`occurrence_date`/`business_content`/`settlement_object`）。

**三层字段命名层次**:
1. **源表表头(源)**: 中汇查账系统/胡庆余堂集团 Excel 实际表头,如 "企业主体名称/查询区间/对方科目/银行账号编号…"
2. **标准化字段(桥)**: financial-normalizer 输出的 16/50 字段标准,如 `entity_name/account_code/counter_account/bank_account_code…`
3. **内部 pipeline 字段(终)**: valuation-detail-table 内部消费字段,如 `code/name/date/customer_supplier/balance…`

各层映射见 `assets/field_mapping.json`:
- 源表头 → 标准化字段: `financial-normalizer/core/column_mappings.json` (guess 阶段)
- 标准化字段 → pipeline 字段: `assets/field_mapping.json` (apply 阶段)
- pipeline 字段 → 评估明细表 Sheet 字段: `assets/sheet_col_map.json` (fill 阶段)

## 2. 标准字段名对照表

### 2.1 往来科目通用字段（3-5, 3-7, 3-8-3, 5-5, 5-6, 5-10-3 等）

| 字段语义 | 标准键名 (sheet_col_map.json) | ❌ 禁止使用 | 来源 |
|---------|------|------|------|
| 结算对象/户名 | `settlement` | ~~settlement_object~~, ~~name~~ | sheet_col_map.json |
| 业务内容 | `business` | ~~business_content~~, ~~biz~~ | sheet_col_map.json |
| 发生日期 | `date` | ~~occurrence_date~~, ~~date_occurrence~~ | sheet_col_map.json |
| 账龄 | `age` | ~~aging~~, ~~account_age~~ | sheet_col_map.json |
| 序号 | `seq` | ~~serial~~, ~~no~~ | sheet_col_map.json |

### 2.2 金额类字段

| 字段语义 | 标准键名 (subjects.json) | 标准键名 (sheet_col_map.json) | ❌ 禁止使用 |
|---------|------|------|------|
| 期末余额 | `balance` | `book_value` | ~~closing_balance~~ |
| 余额方向 | `direction` | — | ~~closing_direction~~ |
| 账面价值 | — | `book_value` | ~~bv~~, ~~book_val~~ |
| 评估价值 | — | `assessed_value` | ~~eval_value~~, ~~assessment_value~~ |
| 增值额 | — | `increment` | ~~value_added~~ |
| 增值率 | — | `increment_rate` | ~~value_added_rate~~ |

### 2.3 科目余额表字段 (subjects.json)

source_header_parser.py 输出的subjects数组，每个元素的标准字段：

```json
{
  "code": "112201",
  "name": "应收账款_某某公司",
  "balance": 12345.67,
  "direction": "借",
  "level": 3,
  "beginning_debit": 0.0,
  "beginning_credit": 0.0,
  "current_debit": 12345.67,
  "current_credit": 0.0,
  "ending_debit": 12345.67,
  "ending_credit": 0.0
}
```

**重要**: `balance` 是标准字段名。`closing_balance` 仅作为向后兼容的fallback，不应在新代码中使用。

**v0.2 (2026-06-01) 更新**: subjects.json 现在支持从胡庆余堂集团余额账(16列合一期初/期末)直接构造:
- `balance` ← `ending_balance`(合一列,值带 direction 含义)
- `direction` ← `direction` 列(显式)
- `account_full_path` ← `account_full_path` 列(用于层级还原)
- `standard_level1` ← `standard_level1` 列(用于重分类)

### 2.4 资产负债表字段 (bs_balances.json)

```json
{
  "items": [
    {
      "label": "货币资金",
      "beginning_balance": 100000.00,
      "ending_balance": 120000.00,
      "side": "资产"
    }
  ],
  "total_assets": 1000000.00,
  "total_liab_equity": 1000000.00,
  "total_equity": 0.0
}
```

### 2.5 序时账字段 (journal_data.json)

**v0.1 (旧版 12 字段)**:
```json
{
  "counterparty_data": {
    "某某公司": {
      "date": "2026-03-15",
      "business": "货款",
      "record_count": 5
    }
  }
}
```

**v0.2 (新版 50 字段对齐胡庆余堂明细账)**: 详见 `financial-normalizer/FIELD_STANDARD.md` "序时账"节。完整字段包括 entity_name, voucher_date, voucher_type, voucher_number, summary, counter_account, other_account, debit_amount, credit_amount, direction, balance, allocation_type, subject_code, subject_full_path, query_subject, query_auxiliary, entry_line_no, 以及 9 组辅助核算三件套(往来/银行/合同/项目/资产/物料/费用/部门/现金流/其他)。

### 2.6 标准化后的科目余额表字段(16 列对齐胡庆余堂)

> 完整说明见 `financial-normalizer/FIELD_STANDARD.md` v0.2 "科目余额表"节。
> 源表头 → 标准化字段 → pipeline 字段 完整映射见 `assets/field_mapping.json`。

| 源表头(胡庆余堂) | 标准化字段 | pipeline 字段 | 必填 | 说明 |
|------|------|------|---|------|
| 企业主体名称 | `entity_name` | `entity_name` | ✅ | 多主体合并时区分 |
| 查询区间 | `query_period` | `query_period` | ✅ | 如 "2025-01~2025-12" |
| 科目编号 | `account_code` | `code` | ✅ | 末级科目编号 |
| 科目名称 | `account_name` | `name` | ✅ | 末级科目名称 |
| 核算类型 | `auxiliary_type` | `auxiliary_type` |  | 银行账户/客户/供应商/项目 等 |
| 核算编号 | `auxiliary_code` | `auxiliary_code` |  | 辅助核算对象编号 |
| 核算名称 | `auxiliary_name` | `auxiliary_name` |  | 辅助核算对象名称 |
| 数据类型 | `data_type` | `data_type` | ✅ | 本位币/原币/数量 |
| 方向 | `direction` | `direction` | ✅ | 余额方向(借/贷) |
| 本位币期初 | `opening_balance` | `beginning_balance` |  | 合一列(用 direction 判定正负) |
| 本位币借方 | `current_debit` | `current_debit` |  | 本期借方发生 |
| 本位币贷方 | `current_credit` | `current_credit` |  | 本期贷方发生 |
| 本位币期末 | `closing_balance` | `ending_balance` |  | 合一列(用 direction 判定正负) |
| 科目全路径 | `account_full_path` | `account_full_path` |  | 多级科目合并路径 |
| 损益结转金额 | `pnl_carryover` | `pnl_carryover` |  | 期末损益类结转金额 |
| 标准1级科目 | `standard_level1` | `standard_level1` |  | 跨企业重分类用 |

**v0.1 旧字段(15 字段,胡庆余堂之前的格式)保留兼容**:
- `opening_debit` → `beginning_debit`
- `opening_credit` → `beginning_credit`
- `closing_debit` → `ending_debit`
- `closing_credit` → `ending_credit`
- `level1_name` / `level2_name` / `level3_name` → 改用 `account_full_path` + `standard_level1`

### 2.7 标准化后的序时账字段(50 列对齐胡庆余堂)

> 完整说明见 `financial-normalizer/FIELD_STANDARD.md` v0.2 "序时账"节。

50 列序时账的结构:
- **列 1-12**: 凭证基本信息(`entity_name, voucher_date, voucher_type, voucher_number, summary, counter_account, other_account, debit_amount, credit_amount, direction, balance, allocation_type`)
- **列 13-17**: 末级科目信息(`subject_code, subject_full_path, query_subject, query_auxiliary, entry_line_no`)
- **列 18-50**: 9 组辅助核算三件套(本方编号/本方名称/对方名称,共 33 列):
  - **列 18-20**: 往来单位(`customer_supplier_code, customer_supplier_name, counter_customer_supplier`)
  - **列 21-23**: 银行账号(`bank_account_code, bank_account_name, counter_bank_account`)
  - **列 24-26**: 采购合同(`purchase_contract_code, purchase_contract_name, counter_purchase_contract`)
  - **列 27-29**: 销售合同(`sales_contract_code, sales_contract_name, counter_sales_contract`)
  - **列 30-32**: 项目(`project_code, project_name, counter_project`)
  - **列 33-35**: 资产项目(`asset_code, asset_name, counter_asset`)
  - **列 36-38**: 存货物料(`material_code, material_name, counter_material`)
  - **列 39-41**: 费用项目(`expense_code, expense_name, counter_expense`)
  - **列 42-44**: 部门(`department_code, department_name, counter_department`)
  - **列 45-47**: 现金流量(`cash_flow_code, cash_flow_name, counter_cash_flow`)
  - **列 48-50**: 其他(`other_code, other_name, counter_other`)

## 3. 兼容性规则

### 3.1 读取时兼容（consumer端）

所有consumer脚本读取subjects.json时，MUST优先使用`balance`键，仅当`balance`不存在时才fallback到`closing_balance`：

```python
# ✅ 正确
balance = s.get('balance', s.get('closing_balance', 0))
```

**v0.2 更新**: 读 `direction` 时也兼容:`direction = s.get('direction', '借' if s.get('ending_debit', 0) >= s.get('ending_credit', 0) else '贷')`

### 3.2 写入时统一（producer端）

所有producer脚本（如source_header_parser.py）写入subjects.json时，MUST使用`balance`字段名：

```python
# ✅ 正确
subjects_item = {
    'code': code,
    'name': name,
    'balance': balance,  # 标准字段名
    'direction': direction,
}
```

### 3.3 sheet_col_map.json 读取兼容

journal_extractor.py等脚本读取sheet_col_map.json时，MUST同时支持flat键名和旧嵌套键名：

```python
# ✅ 正确（flat优先 + 旧嵌套兼容）
date_col = col_map.get('date') or col_map.get('occurrence_date', {}).get('col')
biz_col = col_map.get('business') or col_map.get('business_content', {}).get('col')
name_col = col_map.get('settlement') or col_map.get('settlement_object', {}).get('col')
```

### 3.4 v0.1 → v0.2 字段升级兼容(2026-06-01 起)

financial-normalizer/apply.py 加载标准化输出时,MUST:
- 保留 v0.1 14 字段(原 `level1/2/3_name` 等)→ 视为 legacy,自动合成 `account_full_path` 和 `standard_level1`
- 优先采用 v0.2 新字段(`account_full_path, standard_level1, direction, entity_name, query_period` 等)
- 借贷合一列(`opening_balance`/`closing_balance`)直接采用,不再从分列合成

## 4. 新增字段命名规则

新增任何字段时，遵循以下规则：

1. **单数形式**: `balance` 而非 `balances`
2. **无缩写**: `business` 而非 `biz`, `settlement` 而非 `stl`
3. **snake_case**: `book_value` 而非 `bookValue` 或 `BookValue`
4. **语义明确**: `date` 而非 `dt`, `business` 而非 `content`
5. **先查sheet_col_map.json**: 如果已有同义字段，复用而非创造新名
6. **辅助核算三件套**: 本方编号 / 本方名称 / 对方名称 → `{dim}_code, {dim}_name, counter_{dim}`(例: `bank_account_code, bank_account_name, counter_bank_account`)
7. **三件套与中文表头严格对齐**: 中文表头"银行账号编号/银行账号名称/对方银行账号名称" → 英文 `bank_account_code, bank_account_name, counter_bank_account`

## 5. 检查清单

每次修改脚本后，运行以下检查：

- [ ] `grep -n "closing_balance" *.py` → 应无新代码使用（旧代码fallback除外）
- [ ] `grep -n "occurrence_date\|business_content\|settlement_object" *.py` → 应无新代码使用（兼容层除外）
- [ ] 新字段是否与sheet_col_map.json中已有键名冲突
- [ ] consumer端是否同时处理dict/list格式的subjects.json
- [ ] **v0.2 新增**: 新字段名是否符合"辅助核算三件套"规则(本方编号/本方名称/对方名称)
- [ ] **v0.2 新增**: 新字段是否同步登记到 `field_mapping.json` 的 trial_balance / journal 节点
- [ ] **v0.2 新增**: 新增汇总/计算字段是否在 `field_mapping.json` 的 `compute_fields` 写明公式
