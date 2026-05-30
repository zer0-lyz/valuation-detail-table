# 财务资料标准化字段标准

> 版本: 0.1.0 | 更新: 2026-05-29

本文档定义了 financial-normalizer 支持的全部 5 种财务资料的标准字段结构、类型、约束和默认值。

---

## 科目余额表

**Schema ID**: `trial_balance` | **字段数**: 14

| 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|
| `account_code` | 科目编码 | str | ✅ | — | — | — |
| `currency` | 币种 | str |  | 人民币 | 人民币/美元/欧元/港元/日元 | 默认人民币 |
| `account_name` | 科目名称 | str | ✅ | — | — | — |
| `level1_name` | 一级科目名称 | str |  | — | — | 如'银行存款' |
| `level2_name` | 二级科目名称 | str |  | — | — | 如'工商银行' |
| `level3_name` | 三级科目名称 | str |  | — | — | 如'北京分行' |
| `opening_debit` | 期初借方余额 | number |  | — | — | — |
| `opening_credit` | 期初贷方余额 | number |  | — | — | — |
| `opening_balance` | 期初余额(不分方向) | number |  | — | — | 借贷合一列时用 |
| `current_debit` | 本期借方发生 | number |  | — | — | — |
| `current_credit` | 本期贷方发生 | number |  | — | — | — |
| `closing_balance` | 期末余额(不分方向) | number |  | — | — | 借贷合一列时用 |
| `closing_debit` | 期末借方余额 | number |  | — | — | — |
| `closing_credit` | 期末贷方余额 | number |  | — | — | — |

---

## 序时账

**Schema ID**: `journal` | **字段数**: 12

| 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|
| `date` | 日期 | date | ✅ | — | — | 凭证日期 |
| `voucher_no` | 凭证号 | str |  | — | — | 如'记-1'、'银付-001' |
| `account_code` | 科目编码 | str | ✅ | — | — | — |
| `currency` | 币种 | str |  | 人民币 | 人民币/美元/欧元/港元/日元 | 默认人民币 |
| `account_name` | 科目名称 | str | ✅ | — | — | — |
| `summary` | 摘要 | str |  | — | — | 业务摘要说明 |
| `debit_amount` | 借方金额 | number |  | — | — | — |
| `credit_amount` | 贷方金额 | number |  | — | — | — |
| `customer_supplier` | 往来单位 | str |  | — | — | 客户/供应商/客商名称 |
| `department` | 部门 | str |  | — | — | 所属部门 |
| `project_name` | 项目 | str |  | — | — | 所属项目/工程 |
| `personnel` | 经办人 | str |  | — | — | 经手人/报销人/制单人 |

---

## 资产负债表

**Schema ID**: `balance_sheet` | **字段数**: 6

| 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|
| `item_name` | 项目名称 | str | ✅ | — | — | 报表行项目名称 |
| `currency` | 币种 | str |  | 人民币 | 人民币/美元/欧元/港元/日元 | 默认人民币 |
| `item_category` | 项目类别 | str |  | — | 资产类/负债类/所有者权益类 | — |
| `item_direction` | 方向 | str |  | — | 借/贷 | 科目正常余额方向 |
| `closing_balance` | 期末余额 | number |  | — | — | 期末数 |
| `opening_balance` | 期初余额 | number |  | — | — | 期初数 |

---

## 利润表

**Schema ID**: `income_statement` | **字段数**: 6

| 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|
| `item_name` | 项目名称 | str | ✅ | — | — | 报表行项目名称 |
| `currency` | 币种 | str |  | 人民币 | 人民币/美元/欧元/港元/日元 | 默认人民币 |
| `item_category` | 项目类别 | str |  | — | 损益类 | — |
| `item_direction` | 方向 | str |  | — | 借/贷 | 科目正常余额方向 |
| `current_period` | 本期金额 | number |  | — | — | 本期发生额 |
| `cumulative` | 本年累计 | number |  | — | — | 年初至本期累计 |

---

## 固定资产台账

**Schema ID**: `fixed_asset` | **字段数**: 20

| 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|
| `asset_code` | 资产编码 | str | ✅ | — | — | 资产卡片编号 |
| `asset_name` | 资产名称 | str | ✅ | — | — | — |
| `currency` | 币种 | str |  | 人民币 | 人民币/美元/欧元/港元/日元 | 默认人民币 |
| `asset_category` | 资产类别 | str |  | — | 房屋建筑物/机器设备/运输设备/电子设备/办公设备/其他 | — |
| `specification` | 规格型号 | str |  | — | — | 规格/型号/技术参数 |
| `quantity` | 数量 | number |  | — | — | 同规格资产数量 |
| `unit` | 单位 | str |  | — | 栋/台/辆/套/条/批/平方米 | 计量单位 |
| `acquisition_date` | 取得日期 | date |  | — | — | 入账日期/购置日期 |
| `original_value` | 原值 | number | ✅ | — | — | 固定资产原值 |
| `accumulated_depreciation` | 累计折旧 | number | ✅ | — | — | 截至基准日累计折旧 |
| `impairment_amount` | 减值准备 | number |  | — | — | 固定资产减值准备 |
| `net_value` | 净值 | number |  | — | — | 原值-累计折旧-减值 |
| `department` | 使用部门 | str |  | — | — | — |
| `location` | 存放地点 | str |  | — | — | — |
| `supplier` | 供应商 | str |  | — | — | 供应商/生产厂商 |
| `depreciation_method` | 折旧方法 | str |  | 平均年限法 | 平均年限法/双倍余额递减法/年数总和法/工作量法 | — |
| `depreciation_life` | 折旧年限 | number |  | — | — | 预计使用年限(年) |
| `residual_rate` | 残值率 | number |  | — | — | 预计净残值率(%) |
| `monthly_depreciation` | 月折旧额 | number |  | — | — | — |
| `status` | 资产状态 | str |  | 在用 | 在用/停用/报废/出租/在建 | — |

---

## 通用约定

- **字段命名**: 全小写 + 下划线，英文
- **中文标签**: 与 `column_mappings.json` 中的 label 一致
- **金额字段**: 数值型，正数记借/贷方，负数用负号
- **未映射字段**: 源文件中有但无法匹配的列标记为 `__unknown_col_N__`，用户可在 mapping config 中手动修正
- **币种**: 默认人民币，多币种场景可指定
- **允许值**: 列出的允许值为建议约束，实际数据可能超出此范围，由下游处理决定是否校验
