# DT Skill 字段命名规范 (DT-NAMING)

> 版本: v1.0 | 创建: 2026-05-25 | 规则: DT-153v2

## 1. 核心原则

**单一数据源原则**: `sheet_col_map.json` 是字段命名的唯一权威来源。所有producer/consumer脚本MUST与之一致。

**flat键名原则**: sheet_col_map.json中的col_map使用flat键名（如`date`/`business`/`settlement`），禁止使用嵌套命名（如`occurrence_date`/`business_content`/`settlement_object`）。

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

## 3. 兼容性规则

### 3.1 读取时兼容（consumer端）

所有consumer脚本读取subjects.json时，MUST优先使用`balance`键，仅当`balance`不存在时才fallback到`closing_balance`：

```python
# ✅ 正确
balance = s.get('balance', s.get('closing_balance', 0))

# ❌ 错误
balance = s.get('closing_balance', 0)  # 会漏掉balance字段
```

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

# ❌ 错误
subjects_item = {
    'code': code,
    'name': name,
    'closing_balance': balance,  # 禁止
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

## 4. 新增字段命名规则

新增任何字段时，遵循以下规则：

1. **单数形式**: `balance` 而非 `balances`
2. **无缩写**: `business` 而非 `biz`, `settlement` 而非 `stl`
3. **snake_case**: `book_value` 而非 `bookValue` 或 `BookValue`
4. **语义明确**: `date` 而非 `dt`, `business` 而非 `content`
5. **先查sheet_col_map.json**: 如果已有同义字段，复用而非创造新名

## 5. 检查清单

每次修改脚本后，运行以下检查：

- [ ] `grep -n "closing_balance" *.py` → 应无新代码使用（旧代码fallback除外）
- [ ] `grep -n "occurrence_date\|business_content\|settlement_object" *.py` → 应无新代码使用（兼容层除外）
- [ ] 新字段是否与sheet_col_map.json中已有键名冲突
- [ ] consumer端是否同时处理dict/list格式的subjects.json
