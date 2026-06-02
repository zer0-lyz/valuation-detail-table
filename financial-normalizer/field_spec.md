# 标准化字段完整说明(中汇查账系统格式)

> 版本: 1.0.0 | 创建: 2026-06-01
>
> **定位**: 这是 financial-normalizer 输出的"标准化字段"体系的完整说明,Agent 在 Phase 0 / Phase 1 / Phase 3 时查阅。
> 配套文档:
> - `FIELD_STANDARD.md` — 字段标准(每个字段 label/类型/必填/默认值)
> - `core/column_mappings.json` — 源表表头变体 → 标准化字段
> - `assets/field_mapping.json` — 标准化字段 → pipeline 内部字段
> - `valuation-detail-table/assets/sheet_col_map.json` — pipeline 字段 → 评估明细表 Sheet 列位
> - `valuation-detail-table/字段定义总表.md` — 评估明细表 Sheet 字段内容规范

---

## 1. 字段流转全景

```
源表(中汇查账系统 Excel)     标准化字段          pipeline 字段         评估明细表
═══════════════════════     ══════════          ═══════════         ═════════
50 列序时账 ──guess()──→ 50 标准化字段 ──apply()──→ 50 内部字段 ──fill_sheet()──→ 67 Sheet
16 列余额账 ──guess()──→ 16 标准化字段 ──apply()──→ 24 内部字段 ──fill_sheet()──→ 67 Sheet
6 列资产负债表 ──guess()──→ 6 标准化字段 ──apply()──→ 6 内部字段 ──fill_sheet()──→ 2-分类汇总
```

每一层映射都有唯一配置文件:
- guess() 用 `core/column_mappings.json`
- apply() 用 `assets/field_mapping.json` (在 valuation-detail-table 内部)
- fill_sheet() 用 `assets/sheet_col_map.json`

---

## 2. 科目余额表 — 16 字段(中汇查账系统格式)

> 源表头(胡庆余堂集团余额账-2.xlsx)与字段对应表。

### 2.1 字段速查表

| # | 标准化字段 | 中文表头 | 英文 alias | 类型 | 必填 | 默认值 |
|---|---|---|---|---|---|---|
| 1 | `entity_name` | 企业主体名称 | entity_name | str | ✅ | — |
| 2 | `query_period` | 查询区间 | query_period | str | ✅ | — |
| 3 | `account_code` | 科目编号 | account_code | str | ✅ | — |
| 4 | `account_name` | 科目名称 | account_name | str | ✅ | — |
| 5 | `auxiliary_type` | 核算类型 | auxiliary_type | str |  | — |
| 6 | `auxiliary_code` | 核算编号 | auxiliary_code | str |  | — |
| 7 | `auxiliary_name` | 核算名称 | auxiliary_name | str |  | — |
| 8 | `data_type` | 数据类型 | data_type | str | ✅ | 本位币 |
| 9 | `direction` | 方向 | direction | str | ✅ | — |
| 10 | `opening_balance` | 本位币期初 | opening_balance | number |  | — |
| 11 | `current_debit` | 本位币借方 | current_debit | number |  | — |
| 12 | `current_credit` | 本位币贷方 | current_credit | number |  | — |
| 13 | `closing_balance` | 本位币期末 | closing_balance | number |  | — |
| 14 | `account_full_path` | 科目全路径 | account_full_path | str |  | — |
| 15 | `pnl_carryover` | 损益结转金额 | pnl_carryover | number |  | 0 |
| 16 | `standard_level1` | 标准1级科目 | standard_level1 | str |  | — |

### 2.2 字段逐项说明

#### 2.2.1 `entity_name` — 企业主体名称
- **作用**: 多主体项目(如母公司+子公司合并)时区分科目所属主体
- **示例**: "杭州胡庆余堂集团有限公司" / "江苏宇狮薄膜科技有限公司" / "样本科技有限公司"
- **使用**: 写入元数据,供后续按主体分别汇总/审计
- **DT 关联**: DT-139(BS 自校验),DT-182(汇总表公式链)

#### 2.2.2 `query_period` — 查询区间
- **作用**: 标识数据所属期间(余额表是时点数,区间标识取数期间)
- **示例**: "2025-01~2025-07" / "2026-01~2026-03" / "2025-01~2025-12"
- **格式**: `YYYY-MM~YYYY-MM`(起始月~结束月)
- **使用**: 元数据,不直接进评估明细表

#### 2.2.3 `account_code` — 科目编号
- **作用**: 末级科目编号(可能 4-6 位,如 1001 / 100201 / 410399)
- **规则**: 唯一标识末级科目;非末级科目不出现在源表(已展开到末级)
- **关联**: 与 `code_to_sheet` 映射(见 `assets/field_mapping.json`)配合,确定数据写入哪个 Sheet

#### 2.2.4 `account_name` — 科目名称
- **作用**: 末级科目中文名称
- **示例**: "库存现金" / "中国工商银行" / "杭州坤川贸易有限公司"
- **DT 关联**: DT-149(业务内容自动映射,不要用科目名当业务内容)

#### 2.2.5-7 `auxiliary_type` / `auxiliary_code` / `auxiliary_name` — 辅助核算三件套
- **作用**: 标识该行的辅助核算维度(银行账户/客户/供应商/部门/项目/存货 等)
- **规则**:
  - 父科目行(无辅助):三件套都为空(如 "1002 银行存款" 父行)
  - 末级辅助行:三件套都有值(如 "100201 工商银行" 配 "银行账户" 类型)
- **示例**:
  - 银行账户: `auxiliary_type="银行账户"`, `auxiliary_code="0011-1202021119900266558"`, `auxiliary_name="中国工商银行浙江省分行人民币活期户"`
  - 客户: `auxiliary_type="客户"`, `auxiliary_code="0004-021101487"`, `auxiliary_name="杭州坤川贸易有限公司"`

#### 2.2.8 `data_type` — 数据类型
- **作用**: 区分本币/原币/数量(多币种场景)
- **取值**:
  - `本位币` — 人民币(默认)
  - `原币` — 外币原币
  - `数量` — 物料数量
- **规则**: 多币种项目需保留原币行;单币种项目通常全部 `本位币`

#### 2.2.9 `direction` — 余额方向
- **作用**: 标识该科目的余额方向(借/贷)
- **取值**: `借` / `贷`
- **重要**: 与 `closing_balance` 配合判定正负号
  - 借方余额: `closing_balance > 0`(资产类、费用类)
  - 贷方余额: `closing_balance < 0`(负债类、权益类、收入类)

#### 2.2.10 `opening_balance` — 本位币期初(合一列)
- **作用**: 期初余额,**合一列**(不再分借贷两列)
- **规则**:
  - 借方余额:正数
  - 贷方余额:负数
- **v0.1 兼容**: v0.1 旧格式 `opening_debit` + `opening_credit` 分列,apply 阶段自动合成合一列
- **计算**: `opening_balance = opening_debit - opening_credit`

#### 2.2.11-12 `current_debit` / `current_credit` — 本位币借/贷方
- **作用**: 本期借/贷方发生额
- **规则**: 与 v0.1 一致(分列)

#### 2.2.13 `closing_balance` — 本位币期末(合一列)
- **作用**: 期末余额,**合一列**
- **规则**:
  - 借方余额:正数
  - 贷方余额:负数
- **计算**: `closing_balance = opening_balance + current_debit - current_credit`
- **DT 关联**: DT-139(BS 自校验,期末余额=资产=负债+权益)

#### 2.2.14 `account_full_path` — 科目全路径
- **作用**: 多级科目合并路径(替代 v0.1 的 `level1/2/3_name` 三列)
- **示例**: "库存现金" / "银行存款/工商银行" / "应收账款/杭州坤川贸易有限公司"
- **使用**: 还原科目层级,`level = len(account_full_path.split('/'))`
- **DT 关联**: sheet_filler 匹配 `subject_schema.json:source_code_prefix`

#### 2.2.15 `pnl_carryover` — 损益结转金额
- **作用**: 期末损益类科目(收入/费用)结转到"本年利润"的金额
- **规则**:
  - 损益类科目期末应有结转(收入类借方结转,费用类贷方结转)
  - 结转后损益类余额为 0
  - 结转金额=本年利润借/贷方对应金额
- **DT 关联**: 重分类计算(DT-19,DT-191)

#### 2.2.16 `standard_level1` — 标准1级科目
- **作用**: 按标准会计制度归类的 1 级科目,**用于跨企业重分类**
- **示例**: 
  - 实际科目 "1601 固定资产" / "1602 累计折旧" / "1610 机器设备" → standard_level1 统一为 "固定资产"
  - 实际科目 "1821 长期待摊费用" / "1811 递延所得税资产" → standard_level1 统一为标准一级
- **DT 关联**: 重分类路由(DT-191),评估明细表 Sheet 分类

### 2.3 v0.1 → v0.2 字段映射

| v0.1 (旧) | v0.2 (新) | 说明 |
|---|---|---|
| `opening_debit` + `opening_credit` | `opening_balance` (合一) | 旧分列 → 新合一列 + `direction` |
| `closing_debit` + `closing_credit` | `closing_balance` (合一) | 同上 |
| `level1_name` / `level2_name` / `level3_name` | `account_full_path` + `standard_level1` | 三独立列 → 路径+标准一级 |
| (无) | `entity_name` / `query_period` | 新增,主体+期间元数据 |
| (无) | `auxiliary_type` / `auxiliary_code` / `auxiliary_name` | 新增,辅助核算三件套 |
| (无) | `data_type` | 新增,本币/原币/数量 |
| (无) | `direction` | 新增,显式方向列 |
| (无) | `pnl_carryover` | 新增,损益结转 |
| (无) | `standard_level1` | 新增,重分类用 |

### 2.4 v0.1 旧字段保留(向后兼容)

v0.2 仍接受以下 v0.1 字段(作为 fallback):
- `opening_debit` / `opening_credit` / `opening_balance`
- `current_debit` / `current_credit`
- `closing_debit` / `closing_credit` / `closing_balance`
- `level1_name` / `level2_name` / `level3_name`
- `currency` (与 `data_type` 部分重叠,优先 `data_type`)

---

## 3. 序时账(明细账) — 50 字段(中汇查账系统格式)

> 源表头(胡庆余堂集团明细账.xlsx)与字段对应表。

### 3.1 字段速查表(按 5 个区域分组)

#### 3.1.1 凭证基本信息(列 1-12)

| # | 标准化字段 | 中文表头 | 类型 | 必填 |
|---|---|---|---|---|
| 1 | `entity_name` | 企业主体名称 | str | ✅ |
| 2 | `voucher_date` | 凭证日期 | date | ✅ |
| 3 | `voucher_type` | 字 | str |  |
| 4 | `voucher_number` | 号 | str |  |
| 5 | `summary` | 摘要 | str |  |
| 6 | `counter_account` | 对方科目 | str |  |
| 7 | `other_account` | 本方其他科目 | str |  |
| 8 | `debit_amount` | 本位币借方 | number |  |
| 9 | `credit_amount` | 本位币贷方 | number |  |
| 10 | `direction` | 方向 | str | ✅ |
| 11 | `balance` | 本位币余额 | number |  |
| 12 | `allocation_type` | 结转分配类型 | str |  |

#### 3.1.2 末级科目信息(列 13-17)

| # | 标准化字段 | 中文表头 | 类型 | 必填 |
|---|---|---|---|---|
| 13 | `subject_code` | 末级科目编号 | str | ✅ |
| 14 | `subject_full_path` | 末级科目全路径 | str | ✅ |
| 15 | `query_subject` | 查询科目 | str |  |
| 16 | `query_auxiliary` | 查询核算 | str |  |
| 17 | `entry_line_no` | 分录行号 | int |  |

#### 3.1.3 9 组辅助核算三件套(列 18-50)

每组 3 列:本方编号 / 本方名称 / 对方名称。

| 列 | 往来单位 | 银行账号 | 采购合同 | 销售合同 | 项目 | 资产项目 | 存货物料 | 费用项目 | 部门 | 现金流量 | 其他 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 本方编号 | `customer_supplier_code` | `bank_account_code` | `purchase_contract_code` | `sales_contract_code` | `project_code` | `asset_code` | `material_code` | `expense_code` | `department_code` | `cash_flow_code` | `other_code` |
| 本方名称 | `customer_supplier_name` | `bank_account_name` | `purchase_contract_name` | `sales_contract_name` | `project_name` | `asset_name` | `material_name` | `expense_name` | `department_name` | `cash_flow_name` | `other_name` |
| 对方名称 | `counter_customer_supplier` | `counter_bank_account` | `counter_purchase_contract` | `counter_sales_contract` | `counter_project` | `counter_asset` | `counter_material` | `counter_expense` | `counter_department` | `counter_cash_flow` | `counter_other` |

### 3.2 字段逐项说明(凭证基本信息)

#### 3.2.1 `entity_name` — 企业主体名称
- 同 2.2.1

#### 3.2.2 `voucher_date` — 凭证日期
- **作用**: 凭证录入日期(业务实际发生日期)
- **格式**: `YYYY-MM-DD`(字符串或 datetime 均可,apply 阶段统一转 datetime)
- **使用**: 评估明细表的"发生日期"列取数来源
- **DT 关联**: DT-30(发生日期格式 yyyy"年"m"月"),DT-51(取距基准日最近的同方向凭证日期),DT-178(同方向原则)

#### 3.2.3 `voucher_type` — 字
- **作用**: 凭证类别(传统记账凭证分类)
- **取值**: `记`(记账凭证) / `收`(收款凭证) / `付`(付款凭证) / `转`(转账凭证) / `银收` / `银付` 等
- **示例**: "记账凭证" / "付款凭证" / "收款凭证"
- **v0.1 兼容**: v0.1 的 `voucher_no` 字段包含"类别+号数",v0.2 拆分为 `voucher_type`(字) + `voucher_number`(号)

#### 3.2.4 `voucher_number` — 号
- **作用**: 凭证号数(每类凭证自增序号)
- **示例**: "1" / "2" / "19" / "53"
- **格式**: 字符串(可能有小数点,如 "19.0000")

#### 3.2.5 `summary` — 摘要
- **作用**: 业务摘要说明(本笔凭证的业务实质)
- **示例**: "收到客户A货款" / "支付供应商B材料款" / "差旅费报销" / ">期初余额"(期初行)
- **DT 关联**: DT-60(摘要归纳业务内容),DT-172(业务内容精简到 2-6 字)
- **特殊值**: `>期初余额`(期初行,胡庆余堂格式特有)

#### 3.2.6 `counter_account` — 对方科目
- **作用**: 本笔凭证的对方科目(全路径)
- **示例**: "银行存款/中国银行" / "应收账款/客户A"
- **使用**: 跨分录匹配,双抬头场景识别

#### 3.2.7 `other_account` — 本方其他科目
- **作用**: 同一凭证中本方其他科目(多借多贷拆行时填)
- **示例**: 一笔付款凭证拆为 2 行,第 1 行 `other_account="应付账款"`, 第 2 行 `other_account="银行存款"`
- **使用**: 多借多贷场景的拆行匹配

#### 3.2.8-9 `debit_amount` / `credit_amount` — 本位币借/贷方
- **作用**: 本笔本币借/贷方金额
- **规则**: 二选一(每行只有借贷一方有值)
- **DT 关联**: DT-91(数值列禁止文本),DT-144(无金额时空)

#### 3.2.10 `direction` — 方向
- **作用**: 本笔业务后该科目的余额方向
- **取值**: `借` / `贷`
- **使用**: 配合 `voucher_date` 取发生日期(DT-178: 取同方向最近日期)

#### 3.2.11 `balance` — 本位币余额
- **作用**: 本笔业务后该科目的本币余额
- **使用**: 账龄/同方向匹配辅助

#### 3.2.12 `allocation_type` — 结转分配类型
- **作用**: 损益结转/费用分配的类型
- **取值**: 
  - `损益结转` — 期末损益类结转
  - `费用分配` — 制造费用/辅助生产成本分配
  - 空 — 普通凭证

### 3.3 字段逐项说明(末级科目信息)

#### 3.3.1 `subject_code` — 末级科目编号
- **作用**: 本行所属末级科目编号
- **示例**: "100201" / "112201" / "6601"
- **使用**: 路由到评估明细表对应 Sheet

#### 3.3.2 `subject_full_path` — 末级科目全路径
- **作用**: 本行所属末级科目完整路径
- **示例**: "银行存款/工商银行" / "应收账款/杭州坤川贸易有限公司"
- **使用**: 辅助 Sheet 匹配(如 3-1-2 银行存款的"开户银行"列)

#### 3.3.3 `query_subject` — 查询科目
- **作用**: 查询入口的科目(用户查账时的查询维度)
- **示例**: "应收账款" / "银行存款"
- **使用**: 辅助识别用户关注维度

#### 3.3.4 `query_auxiliary` — 查询核算
- **作用**: 查询入口的辅助核算对象
- **示例**: "客户A" / "工商银行"
- **使用**: 配合 `query_subject` 形成"科目+核算"查询维度

#### 3.3.5 `entry_line_no` — 分录行号
- **作用**: 同一凭证内的分录行序号(1, 2, 3, …)
- **示例**: 1 / 2 / 3
- **使用**: 多借多贷拆行匹配,`other_account` 配对

### 3.4 字段逐项说明(9 组辅助核算三件套)

> 每组三列规则相同:本方编号 / 本方名称 / 对方名称。
> 9 组覆盖了评估明细表 67 个 Sheet 中 80% 的辅助列需求。

#### 3.4.1 往来单位(列 18-20) — `customer_supplier_code` / `customer_supplier_name` / `counter_customer_supplier`
- **适用 Sheet**: 3-5 应收账款 / 3-7 预付款项 / 3-8-3 其他应收款 / 5-5 应付账款 / 5-6 预收款项 / 5-10-3 其他应付款
- **本方名称** → 评估明细表"户名/结算对象"列(DT-111)
- **对方名称** → 双抬头场景(集团内部交易)

#### 3.4.2 银行账号(列 21-23) — `bank_account_code` / `bank_account_name` / `counter_bank_account`
- **适用 Sheet**: 3-1-2 银行存款 / 3-1-3 其他货币资金
- **本方名称** → 评估明细表"开户银行"列
- **本方编号** → 评估明细表"账号"列
- **DT 关联**: DT-65(银行关键值校对),DT-104(逐行展开),DT-134(银行对账单结构化)

#### 3.4.3 采购合同(列 24-26) — `purchase_contract_code` / `purchase_contract_name` / `counter_purchase_contract`
- **适用**: 合同台账关联,大额采购业务披露
- **本方编号/名称** → 内部合同台账
- **对方名称** → 对方合同台账

#### 3.4.4 销售合同(列 27-29) — `sales_contract_code` / `sales_contract_name` / `counter_sales_contract`
- **适用**: 合同台账关联,大额销售业务披露
- **本方编号/名称** → 内部合同台账
- **对方名称** → 对方合同台账

#### 3.4.5 项目(列 30-32) — `project_code` / `project_name` / `counter_project`
- **适用 Sheet**: 在建工程/工程物资(4-9-x),项目核算的损益类(管理费用/销售费用/研发支出 等)
- **本方名称** → 评估明细表"项目"辅助列
- **DT 关联**: DT-150(递延所得税项目披露)

#### 3.4.6 资产项目(列 33-35) — `asset_code` / `asset_name` / `counter_asset`
- **适用 Sheet**: 4-8-1 房屋建筑物 / 4-8-4 机器设备 / 4-8-7 其他设备 / 4-13 无形资产 等
- **本方编号/名称** → 评估明细表"资产编号/资产名称"列
- **DT 关联**: DT-88(PDF 卡片台账全量提取),DT-101(符号方向)

#### 3.4.7 存货物料(列 36-38) — `material_code` / `material_name` / `counter_material`
- **适用 Sheet**: 3-9 存货 / 3-9-1 原材料 / 3-9-2 库存商品 / 3-9-7 周转材料
- **本方编号/名称** → 评估明细表"物料编号/物料名称"列

#### 3.4.8 费用项目(列 39-41) — `expense_code` / `expense_name` / `counter_expense`
- **适用 Sheet**: 4-19 销售费用 / 4-20 管理费用 / 4-21 财务费用 等损益类 Sheet
- **本方名称** → 评估明细表"费用项目"列
- **DT 关联**: DT-149(业务内容自动映射,expense_name 不能是空的)

#### 3.4.9 部门(列 42-44) — `department_code` / `department_name` / `counter_department`
- **适用 Sheet**: 4-19 销售费用 / 4-20 管理费用 / 4-21 财务费用 等
- **本方名称** → 评估明细表"部门"列
- **DT 关联**: DT-167(非往来科目名称/内容列非空)

#### 3.4.10 现金流量(列 45-47) — `cash_flow_code` / `cash_flow_name` / `counter_cash_flow`
- **适用**: 现金流量表关联,大额业务现金流分类
- **本方名称** → 现金流量表项目
- **对方名称** → 对方现金流

#### 3.4.11 其他(列 48-50) — `other_code` / `other_name` / `counter_other`
- **适用**: 自定义辅助核算(非以上 10 类的特殊维度)
- **规则**: 仅在以上 10 类都不适用时使用,优先归到 10 类

---

## 4. 资产负债表 — 6 字段(无变化)

| # | 标准化字段 | 中文表头 | 类型 | 必填 |
|---|---|---|---|---|
| 1 | `item_name` | 项目名称 | str | ✅ |
| 2 | `currency` | 币种 | str |  |
| 3 | `item_category` | 项目类别 | str |  |
| 4 | `item_direction` | 方向 | str |  |
| 5 | `closing_balance` | 期末余额 | number |  |
| 6 | `opening_balance` | 期初余额 | number |  |

---

## 5. 字段流转示例(以胡庆余堂集团应收A客户为例)

### 5.1 源表(胡庆余堂明细账 R3)

```
A=企业主体名称 = 杭州胡庆余堂集团有限公司
B=凭证日期 = 2025-01-15
C=字 = 记账凭证
D=号 = 1
E=摘要 = 收到客户A货款
F=对方科目 = 应收账款/客户A
G=本方其他科目 = 
H=本位币借方 = 50000
I=本位币贷方 = 
J=方向 = 借
K=本位币余额 = 71825.59
L=结转分配类型 = 
M=末级科目编号 = 100201
N=末级科目全路径 = 银行存款/工商银行
O=查询科目 = 银行存款
P=查询核算 = 工商银行
Q=分录行号 = 1
R=往来单位编号 = 0001-001
S=往来单位名称 = 北京客户A有限公司
T=对方往来单位名称 = 
U=银行账号编号 = 0011-1101021119900266558
V=银行账号名称 = 工商银行北京分行人民币活期户
W=对方银行账号名称 = 
X-AC=...
AD-AF=...
AG-AI=...
AJ-AL=...
AM-AO=...
AP-AR=...
AS-AU=...
AV-AX=...
```

### 5.2 标准化输出(50 字段 dict)

```python
{
    "entity_name": "杭州胡庆余堂集团有限公司",
    "voucher_date": "2025-01-15",
    "voucher_type": "记账凭证",
    "voucher_number": "1",
    "summary": "收到客户A货款",
    "counter_account": "应收账款/客户A",
    "other_account": None,
    "debit_amount": 50000.0,
    "credit_amount": None,
    "direction": "借",
    "balance": 71825.59,
    "allocation_type": None,
    "subject_code": "100201",
    "subject_full_path": "银行存款/工商银行",
    "query_subject": "银行存款",
    "query_auxiliary": "工商银行",
    "entry_line_no": 1,
    "customer_supplier_code": "0001-001",
    "customer_supplier_name": "北京客户A有限公司",
    "counter_customer_supplier": None,
    "bank_account_code": "0011-1101021119900266558",
    "bank_account_name": "工商银行北京分行人民币活期户",
    "counter_bank_account": None,
    # ...其余三件套为 None 或其他值
}
```

### 5.3 评估明细表 Sheet 写入(3-1-2 银行存款)

```python
data_row = {
    "seq": 1,                                # 序号
    "settlement": "工商银行北京分行人民币活期户",  # 开户银行(来自 bank_account_name)
    "account": "0011-1101021119900266558",   # 账号(来自 bank_account_code)
    "currency": "人民币",                      # 币种(本币)
    "book_value": 50000.0,                    # 账面价值(借方发生额)
    "remark": "客户A货款",                    # 备注(来自 summary)
}
```

---

## 6. 常见问题

### Q1: v0.1 旧版"分借贷期初/期末"格式还能用吗?

**能**。`assets/field_mapping.json` 的 `compute_fields.balance` 公式兼容两种格式:
- 优先: `ending_debit` / `ending_credit` 分列(v0.1)
- 其次: `ending_balance` 合一列(v0.2)

### Q2: 如果源表只有部分辅助核算(没有现金流量/资产项目),行怎么填?

空白字段填空字符串 `""` 或 `None`。`column_mappings.json` 的 guess 阶段不会把空白误识别为其他字段。

### Q3: 评估明细表只用到 50 列中的一部分,其余列需要全部消费吗?

不需要。`field_mapping.json` 中标记为 `null` 的字段不消费;其余字段按需消费(sheet_filler 只读取它需要的字段)。

### Q4: 如何区分"客户"和"供应商"?

源表 `auxiliary_type` 字段:
- `客户` / `客户(应收账款)` / `客户(预收)` → 应收账款/预收款项
- `供应商` / `供应商(应付账款)` / `供应商(预付)` → 应付账款/预付款项
- 通用 `往来单位` → 需根据科目方向判定

### Q5: `account_full_path` 多级怎么拆?

```python
parts = account_full_path.split("/")  # ["银行存款", "工商银行"]
level = len(parts)                    # 2
```

### Q6: 同一凭证多借多贷怎么拆行?

胡庆余堂明细账中,每行已经是一个"借/贷分录",`entry_line_no` 标识同一凭证内的不同分录行。无需再拆。

### Q7: 重分类(应收账款贷方余额)如何处理?

- 源表 1122 应收账款期末贷方余额 → `direction="贷"`, `closing_balance < 0`
- `pnl_carryover` 标记 0
- sheet_filler 阶段:按 `direction` 路由到"预收款项" Sheet(DT-191 重分类规则)
- 详见 `assets/field_mapping.json` 的 `standard_subject` 公式

---

## 7. 版本与变更

| 版本 | 日期 | 变更 |
|---|---|---|
| 1.0.0 | 2026-06-01 | 初版,完整说明 16+50 字段体系(中汇查账系统格式) |
