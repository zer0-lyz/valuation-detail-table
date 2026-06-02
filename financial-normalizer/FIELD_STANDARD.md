# 财务资料标准化字段标准

> 版本: 0.2.0 | 更新: 2026-06-01
>
> **变更说明（v0.2.0）**：trial_balance 由 14 字段扩展为 16 字段，对齐中汇查账系统余额账列结构；journal 由 12 字段扩展为 50 字段，对齐中汇查账系统明细账列结构，新增 9 组辅助核算（对方科目 / 银行账号 / 采购合同 / 销售合同 / 项目 / 资产项目 / 存货物料 / 费用项目 / 部门 / 现金流量 / 其他）的"本方编号/本方名称/对方名称"三件套。

本文档定义了 financial-normalizer 支持的全部 5 种财务资料的标准字段结构、类型、约束和默认值。

---

## 科目余额表

**Schema ID**: `trial_balance` | **字段数**: 16

> 字段列序严格对齐源表头(企业主体名称 → 查询区间 → 科目编号 → … → 标准1级科目)，共 16 列。

| 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|
| `entity_name` | 企业主体名称 | str | ✅ | — | — | 余额表所属企业(用于多主体合并) |
| `query_period` | 查询区间 | str | ✅ | — | — | 期间范围,如 `2025-01~2025-12` |
| `account_code` | 科目编号 | str | ✅ | — | — | 如 `1001` / `100201` / `410399` |
| `account_name` | 科目名称 | str | ✅ | — | — | 如 `人民币现金` / `中国银行` |
| `auxiliary_type` | 核算类型 | str |  | — | 银行账户/客户/供应商/部门/项目/存货/费用 等 | 辅助核算维度类别(空表示无辅助) |
| `auxiliary_code` | 核算编号 | str |  | — | — | 辅助核算对象的编号 |
| `auxiliary_name` | 核算名称 | str |  | — | — | 辅助核算对象的名称 |
| `data_type` | 数据类型 | str | ✅ | 本位币 | 本位币/原币/数量/外币 | 区分本币/原币/数量行(便于多币种项目) |
| `direction` | 方向 | str | ✅ | — | 借/贷 | 该科目在本期的余额方向 |
| `opening_balance` | 本位币期初 | number |  | — | — | 期初余额(不分借贷方向,正数表示实际借方,负数表示实际贷方) |
| `current_debit` | 本位币借方 | number |  | — | — | 本期借方发生额 |
| `current_credit` | 本位币贷方 | number |  | — | — | 本期贷方发生额 |
| `closing_balance` | 本位币期末 | number |  | — | — | 期末余额(不分借贷方向) |
| `account_full_path` | 科目全路径 | str |  | — | — | 多级科目合并路径,如 `库存现金/人民币现金` |
| `pnl_carryover` | 损益结转金额 | number |  | 0 | — | 期末损益类结转到本年利润的金额(用于重分类) |
| `standard_level1` | 标准1级科目 | str |  | — | — | 按标准会计制度归类的 1 级科目(用于跨企业重分类) |

> **v0.1 → v0.2 字段映射**：v0.1 中的 `opening_debit` / `opening_credit` / `closing_debit` / `closing_credit` 在 v0.2 合并为 `opening_balance` / `closing_balance`(通过 `direction` 区分借贷)；`level1_name` / `level2_name` / `level3_name` 合并为 `account_full_path` + `standard_level1`。旧 mapping 在 `column_mappings.json` 中以变体形式保留,确保存量项目仍可识别。

---

## 序时账

**Schema ID**: `journal` | **字段数**: 50

> 字段列序严格对齐源表头,共 50 列。第 1-12 列为基本凭证信息,第 13-50 列为按 9 组辅助核算展开的"本方编号/本方名称/对方名称"三件套(部分扩展为 3 列)。

| # | 字段名 | 中文标签 | 类型 | 必填 | 默认值 | 允许值 | 说明 |
|---|---|---|---|---|---|---|---|
| 1 | `entity_name` | 企业主体名称 | str | ✅ | — | — | 凭证所属企业(用于多主体) |
| 2 | `voucher_date` | 凭证日期 | date | ✅ | — | — | 凭证录入日期(用于发生日期抽取) |
| 3 | `voucher_type` | 字 | str |  | — | 记/收/付/转 等 | 凭证类别(传统记账凭证分类) |
| 4 | `voucher_number` | 号 | str |  | — | — | 凭证号数(每类凭证自增序号) |
| 5 | `summary` | 摘要 | str |  | — | — | 业务摘要说明(用于业务内容归纳) |
| 6 | `counter_account` | 对方科目 | str |  | — | — | 本笔凭证对方科目全路径,如 `银行存款/中国银行` |
| 7 | `other_account` | 本方其他科目 | str |  | — | — | 同一凭证中本方其他科目(多借多贷时填) |
| 8 | `debit_amount` | 本位币借方 | number |  | — | — | 本笔本位币借方金额 |
| 9 | `credit_amount` | 本位币贷方 | number |  | — | — | 本笔本位币贷方金额 |
| 10 | `direction` | 方向 | str | ✅ | — | 借/贷 | 本笔业务后该科目的余额方向 |
| 11 | `balance` | 本位币余额 | number |  | — | — | 本笔业务后该科目的本位币余额 |
| 12 | `allocation_type` | 结转分配类型 | str |  | — | — | 损益结转/费用分配 类型 |
| 13 | `subject_code` | 末级科目编号 | str | ✅ | — | — | 本行所属末级科目编号 |
| 14 | `subject_full_path` | 末级科目全路径 | str | ✅ | — | — | 本行所属末级科目完整路径 |
| 15 | `query_subject` | 查询科目 | str |  | — | — | 查询入口的科目(用于多维度查询) |
| 16 | `query_auxiliary` | 查询核算 | str |  | — | — | 查询入口的辅助核算对象 |
| 17 | `entry_line_no` | 分录行号 | int |  | — | — | 同一凭证内的分录行序号(1, 2, 3, …) |
| 18 | `customer_supplier_code` | 往来单位编号 | str |  | — | — | 客户/供应商的编号(辅助核算-往来单位-本方) |
| 19 | `customer_supplier_name` | 往来单位名称 | str |  | — | — | 客户/供应商的名称(辅助核算-往来单位-本方) |
| 20 | `counter_customer_supplier` | 对方往来单位名称 | str |  | — | — | 对方单位的往来单位名称(辅助核算-往来单位-对方) |
| 21 | `bank_account_code` | 银行账号编号 | str |  | — | — | 银行账号编号(辅助核算-银行账户-本方) |
| 22 | `bank_account_name` | 银行账号名称 | str |  | — | — | 银行账号名称(辅助核算-银行账户-本方) |
| 23 | `counter_bank_account` | 对方银行账号名称 | str |  | — | — | 对方银行账号名称(辅助核算-银行账户-对方) |
| 24 | `purchase_contract_code` | 采购合同编号 | str |  | — | — | 采购合同编号(辅助核算-采购合同-本方) |
| 25 | `purchase_contract_name` | 采购合同名称 | str |  | — | — | 采购合同名称(辅助核算-采购合同-本方) |
| 26 | `counter_purchase_contract` | 对方采购合同名称 | str |  | — | — | 对方采购合同名称(辅助核算-采购合同-对方) |
| 27 | `sales_contract_code` | 销售合同编号 | str |  | — | — | 销售合同编号(辅助核算-销售合同-本方) |
| 28 | `sales_contract_name` | 销售合同名称 | str |  | — | — | 销售合同名称(辅助核算-销售合同-本方) |
| 29 | `counter_sales_contract` | 对方销售合同名称 | str |  | — | — | 对方销售合同名称(辅助核算-销售合同-对方) |
| 30 | `project_code` | 项目编号 | str |  | — | — | 项目编号(辅助核算-项目-本方) |
| 31 | `project_name` | 项目名称 | str |  | — | — | 项目名称(辅助核算-项目-本方) |
| 32 | `counter_project` | 对方项目名称 | str |  | — | — | 对方项目名称(辅助核算-项目-对方) |
| 33 | `asset_code` | 资产项目编号 | str |  | — | — | 固定资产/无形资产 编号(辅助核算-资产项目-本方) |
| 34 | `asset_name` | 资产项目名称 | str |  | — | — | 资产名称(辅助核算-资产项目-本方) |
| 35 | `counter_asset` | 对方资产项目名称 | str |  | — | — | 对方资产名称(辅助核算-资产项目-对方) |
| 36 | `material_code` | 存货物料编号 | str |  | — | — | 物料编号(辅助核算-存货物料-本方) |
| 37 | `material_name` | 存货物料名称 | str |  | — | — | 物料名称(辅助核算-存货物料-本方) |
| 38 | `counter_material` | 对方存货物料名称 | str |  | — | — | 对方物料名称(辅助核算-存货物料-对方) |
| 39 | `expense_code` | 费用项目编号 | str |  | — | — | 费用项目编号(辅助核算-费用项目-本方) |
| 40 | `expense_name` | 费用项目名称 | str |  | — | — | 费用项目名称(辅助核算-费用项目-本方) |
| 41 | `counter_expense` | 对方费用项目名称 | str |  | — | — | 对方费用项目名称(辅助核算-费用项目-对方) |
| 42 | `department_code` | 部门编号 | str |  | — | — | 部门编号(辅助核算-部门-本方) |
| 43 | `department_name` | 部门名称 | str |  | — | — | 部门名称(辅助核算-部门-本方) |
| 44 | `counter_department` | 对方部门名称 | str |  | — | — | 对方部门名称(辅助核算-部门-对方) |
| 45 | `cash_flow_code` | 现金流量编号 | str |  | — | — | 现金流量编号(辅助核算-现金流量-本方) |
| 46 | `cash_flow_name` | 现金流量名称 | str |  | — | — | 现金流量名称(辅助核算-现金流量-本方) |
| 47 | `counter_cash_flow` | 对方现金流量名称 | str |  | — | — | 对方现金流量名称(辅助核算-现金流量-对方) |
| 48 | `other_code` | 其他编号 | str |  | — | — | 其他自定义辅助核算编号 |
| 49 | `other_name` | 其他名称 | str |  | — | — | 其他自定义辅助核算名称 |
| 50 | `counter_other` | 对方其他名称 | str |  | — | — | 对方其他自定义辅助核算名称 |

> **v0.1 → v0.2 字段映射**：v0.1 中的 12 字段(date, voucher_no, account_code, account_name, summary, debit_amount, credit_amount, customer_supplier, department, project_name, personnel, currency)在 v0.2 中：
> - `date` → `voucher_date`
> - `voucher_no` → 拆分为 `voucher_type`(字) + `voucher_number`(号)
> - `customer_supplier` / `department` / `project_name` / `personnel` 升级为三件套(本方编号/本方名称/对方名称)
> - 新增 6 组三件套:`counter_account`(对方科目)、`bank_account`、`purchase_contract`、`sales_contract`、`asset_item`(资产项目)、`material`(存货物料)、`expense_item`(费用项目)、`cash_flow_item`(现金流量)、`other_item`(其他)
> - 新增 `direction` / `balance` / `allocation_type` / `subject_full_path` / `query_subject` / `query_auxiliary` / `entry_line_no` / `entity_name` 字段

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
- **辅助核算三件套**: 序时账中,任意辅助核算维度(往来/银行/合同/项目/资产/物料/费用/部门/现金流/其他)均以"本方编号/本方名称/对方名称"三列为一组,共同表达一笔分录的核算信息
- **方向/余额关系**: 序时账中每行有 `direction`(借/贷) + `balance`(本笔后余额)。重分类时按 `direction` 判断科目正常方向是否一致
