# S2_fill_liability: 负债类科目填写

> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2规则节
> **📋 内置规则覆盖**：以下步骤通过fill_sheet()管线自动执行DT-0/46/66/97/120/125/136/143/144/147/149/150/151/152/153/155/156/157/158/159，Agent无需手动实现

> **🚨 DT-182b 插行后公式完整性修复**：每个 `fill_sheet()` 完成后，MUST 依次执行：
> 1. `fix_summary_sheet_refs(wb, sheet_name)` — 修复汇总表跨Sheet引用行号（仅金额列E-P）
> 2. 全部Phase 2完成后：`fix_intra_sheet_total2_formulas(wb)` — 修复明细表内合计2行公式
> 详见 `scripts/fix_summary_refs.py`。
## 🚨 Step 0: 脚本选择（DT-165强制，编写任何代码前MUST执行）

> **⚠️ 本步骤在编写任何Python代码之前必须执行。禁止跳过。**
> 复盘根因：v2脚本完全绕开了fill_sheet()，从零写了fill_sheet_data()→6/9个Sheet列位错误+合计行被覆盖+Phase 2e被跳过。

**必读索引**：Read `scripts/SKILL_SCRIPT_INDEX.md`，确认关键接口：
- 数据写入→`fill_sheet()`（禁止自写写入函数）
- 插入行→`smart_insert_row()`（禁止裸insert_rows）
- 列位映射→`sheet_col_map.json`（禁止硬编码列号）
- 数据加载→`load_subject_data()`/`load_auxiliary_balance()`
- Gate验证→`gate_G2()`

## 🚨 唯一写入接口（DT-128+DT-160）

**评估明细表数据写入MUST通过`sheet_filler.fill_sheet()`接口执行。**
- ✅ 允许：`fill_sheet(ws, sheet_id, data_rows, ...)` — 12条DT规则内部自动执行
- ❌ 禁止：直接`ws.cell(row=r, column=10).value = xxx` — 绕过全部规则断言
- ❌ 禁止：`import openpyxl`后直接写入ws — DT-160裸openpyxl写入=绕过管线=4类错误必现

**openpyxl仅允许**：`load_workbook()`加载 / `save()`保存 / `wb[sheetname]`获取ws对象传给fill_sheet

## 🚨 DT-46: 负债类列序（最高优先级）

> **负债类科目C/D列与资产类相反！**
> - 资产类：C=业务内容(文字), D=发生日期(datetime)
> - **负债类：C=发生日期(datetime), D=业务内容(文字)**
>
> fill_sheet()内部通过sheet_col_map.json自动处理列序差异，Agent无需手动判断C/D列含义。
> 但Agent必须理解业务逻辑差异，以便在data_rows中正确组织date和business_content字段。

## ⚠️ DT-140 高金额科目操作级指导覆盖声明

> 本文件中BS金额>1000万或占负债总计>5%的科目，MUST提供数据来源+勾稽公式。低于此标准=步骤文件不完整=Agent凭猜测执行=错误概率极高。

## 输入

- 科目余额表负债类科目数据（Phase 0已解析，_dt_cache/subjects.json）
- 辅助余额表结算对象数据（Phase 0 Step 0.5已提取，_dt_cache/auxiliary_balance_*.json）
- 资产负债表负债方数据（Phase 0 Step 0.3已解析，_dt_cache/bs_balances.json）
- D1/D2/D3映射表（Phase 0 Step 0.5a已建立，_dt_cache/d1d2d3_mapping.json）
- 评估明细表模板

## 操作

### 通用填写步骤（所有负债科目适用）

每个负债科目Sheet的填写遵循统一3步管线：

**Step 1：组织数据 — `prepare_data_rows()`**
```python
from sheet_filler import fill_sheet, prepare_data_rows, get_sheet_id_for_subject

data_rows = prepare_data_rows(
    subject_code='2202',            # 科目编码
    aux_data=aux_balance_data,      # 辅助余额表结算对象数据（DT-111优先）
    kmye_data=kmye_items,           # 科目余额表末级科目数据
    subject_name='应付账款',
    has_journal=True,               # DT-143: 无序时账→日期留空
    has_foreign_currency=False,     # DT-144: 无外币→币种/汇率留空
    industry_type='房地产',          # DT-145: 行业映射
)
```

**Step 2：写入 — `fill_sheet()`**
```python
ws = wb['5-5应付账款']              # 仅获取ws对象传给fill_sheet
result = fill_sheet(
    ws=ws,
    sheet_id='5-5',                 # DT-153: 从sheet_col_map.json读取列位
    data_rows=data_rows,
    settings=settings_info,
    wb=wb,
    has_journal=True,
    has_foreign_currency=False,
    bad_debt_amount=None,
    provision_amount=None,
)
```
- fill_sheet()内部自动执行：插行判断(DT-164:数据行数>模板预留行数才插行)、插行(DT-120/152)、合计行保护断言(DT-164.1:插行后合计1/坏账准备/合计2三行A列标记+B列内容+SUM公式完整性)、列位映射(DT-136/153)、列序校验(DT-46/66)、回读验证(DT-97)、幂等clear(DT-155)、即时勾稽(DT-158)、合计行B:C合并校验(DT-163)
- **负债类列序(DT-46)由sheet_col_map.json自动处理**，Agent不需要手动判断C=日期还是D=日期

**Step 3：检查结果 — DT-138**
```python
if not result['success']:
    print(f"🚨 fill_sheet失败: {result['gate_errors'] + result['read_back_errors']}")
    sys.exit(1)
```

---

### Step 2d.1 短期借款（5-1）

1. 从subjects.json读取2001短期借款子科目
2. `prepare_data_rows(subject_code='2001', aux_data=..., subject_name='短期借款')`
3. `fill_sheet(ws, sheet_id='5-1', data_rows=data_rows, ...)` — 列位由sheet_col_map.json自动映射(DT-153)
4. 检查`result['success']`

**数据来源**：D2=科目余额表2001短期借款子科目；D3=辅助余额表（如有）

**勾稽公式**：5-1合计1行账面价值 = BS短期借款（流动负债方）

### Step 2d.2 应付账款（5-5）⚡ 高金额科目

1. 从辅助余额表获取应付账款结算对象数据（DT-111优先，通常有3+类辅助余额表：应付工程款/暂估应付/集团内部等）
2. `prepare_data_rows(subject_code='2202', aux_data=..., subject_name='应付账款')`
3. `fill_sheet(ws, sheet_id='5-5', data_rows=data_rows, ...)` — 负债类列序由sheet_col_map.json自动处理(DT-46)
4. 区分正常应付和暂估应付
5. **DT-137**：MUST校验填写行数≥辅助余额表结算对象总数。行数<总数=WARNING
6. 检查`result['success']`

**数据来源**：D2=科目余额表2202应付账款子科目；D3=辅助余额表

**勾稽公式**：5-5合计1行账面价值 = BS应付账款（流动负债方）

**特殊注意**：
- 5-5为高合并单元格风险Sheet — fill_sheet()内部使用_safe_write_cell()自动处理(DT-125)
- 按供应商逐行填写，区分正常应付和暂估应付

### Step 2d.3 预收款项（5-6）

1. 从辅助余额表获取预收款项结算对象数据（DT-111优先）
2. `prepare_data_rows(subject_code='2203', aux_data=..., subject_name='预收款项')`
3. `fill_sheet(ws, sheet_id='5-6', data_rows=data_rows, ...)`
4. 关注长期挂账的预收
5. **[DT-93]** BS同时有预收款项和合同负债时，以科目余额表2203/2210分类为准
6. 检查`result['success']`

**勾稽公式**：5-6合计1行账面价值 = BS预收款项

### Step 2d.4 应付职工薪酬（5-7/5-8）

1. 从subjects.json读取2211应付职工薪酬子科目
2. `prepare_data_rows(subject_code='2211', aux_data=..., subject_name='应付职工薪酬')`
3. `fill_sheet(ws, sheet_id='5-8', data_rows=data_rows, ...)` — 注意：职工薪酬sheet_id需根据实际模板确认
4. 按薪酬项目逐行填写
5. 检查`result['success']`

**勾稽公式**：5-8合计1行账面价值 = BS应付职工薪酬

### Step 2d.5 应交税费（5-9）⚡ 高金额科目（需拆分）

1. 从subjects.json读取2221应交税费子科目
2. `prepare_data_rows(subject_code='2221', aux_data=..., subject_name='应交税费')` — fill_sheet()内部自动调用`infer_tax_details()`逐税种拆分(DT-147)
3. `fill_sheet(ws, sheet_id='5-9', data_rows=data_rows, ...)`
4. 检查`result['success']`

**⚠️ DT-126关键规则**：2221应交税费子科目MUST按BS口径拆分处理：
1. **贷方余额子项**（222109土地使用税/222111个人所得税等）→ 填写5-9应交税费
2. **待转销项税额**（222123）→ 填写5-13其他流动负债
3. **贷方负数子项**（222107/222113/222115等进项税额相关）→ 重分类至3-13其他流动资产（DT-87/DT-118）
4. 校验：5-9合计=BS应交税费，5-13合计=BS其他流动负债，3-13=2221贷方负数合计

**勾稽公式**：5-9合计1行账面价值 = BS应交税费

### Step 2d.6 其他应付款（5-10-3）⚡⚡⚡ 最高金额科目（DT-140操作级指导）

> **🚨 DT-140 覆盖声明**：本科目在河南平绿项目BS金额3.95亿（占负债总计>70%），原S2_fill_liability.md仅3行描述导致列偏移+数据遗漏+勾稽失败。以下为操作级指导。

**填写流程（5步，MUST逐步执行）**：

**Step 1：确认数据源完整性**
- 读取`_dt_cache/auxiliary_balance_2241*.json`和`_dt_cache/auxiliary_balance_summary.json`
- 确认结算对象总数（DT-137）
- 检查BS其他应付款与2241贷方合计的差异，确认2102/2231归并金额

**Step 2：合并结算对象清单**
- 合并所有辅助余额表的2241结算对象（去重）
- 加入2102结算中心借款（如有，作为独立行项）
- 加入2231应付利息结算对象（如有且BS=0）
- 合计行数MUST≥辅助余额表结算对象总数

**Step 3：调用fill_sheet()写入**
```python
data_rows = prepare_data_rows(
    subject_code='2241',
    aux_data=merged_settlement_objects,   # Step 2合并后的结算对象清单
    subject_name='其他应付款',
    has_journal=True,
    has_foreign_currency=False,
)
ws = wb['5-10-3其他应付款']
result = fill_sheet(ws=ws, sheet_id='5-10-3', data_rows=data_rows, wb=wb)
```
- 列位由sheet_col_map.json自动映射(DT-153) — 负债类C=发生日期/D=业务内容由fill_sheet()自动处理
- 2102结算中心借款行：business_content标注"[推论映射]结算中心借款，BS编制时并入其他应付款"（DT-117）
- 2231应付利息行：business_content标注"[重分类]应付利息BS=0，已并入其他应付款"（DT-100/DT-118）

**Step 4：勾稽验证**
- 5-10-3合计1行账面价值 MUST = BS其他应付款 ±1元
- 不等时检查：辅助余额表是否完全覆盖？2102/2231是否正确归并？
- fill_sheet()内部即时勾稽(DT-158)会自动校验

**Step 5：检查结果**
```python
if not result['success']:
    sys.exit(1)
```

**BS归并处理**（MUST执行，DT-117差额推论+DT-100重分类检测）：

```
BS其他应付款 = 2241贷方合计 + 2102结算中心借款(如有) + 2231应付利息(如BS=0)
                        │
                        ├─ 差额=2102结算中心借款 → 确认映射→归入5-10-3
                        │  （BS"结算中心借款"行year_end=null → 编制时并入其他应付款）
                        │
                        └─ 差额=2231应付利息 → 确认映射→归入5-10-3
                           （BS"应付利息"行year_end=0 → 已并入其他应付款，DT-100）
```

**数据来源**：
- D2=科目余额表2241其他应付款子科目（贷方余额子项）
- D2补充=2102结算中心借款（BS并入其他应付款，DT-117差额推论）
- D2补充=2231应付利息（BS=0时需并入其他应付款，DT-100重分类检测）
- D3=辅助余额表（通常有2+类：集团外部/集团内部/部门等）
- **DT-137**：MUST校验填写行数≥辅助余额表结算对象总数

**勾稽公式**：5-10-3合计1行账面价值 = BS其他应付款

### Step 2d.7 长期借款（6-1）

1. 从subjects.json读取2501长期借款子科目
2. `prepare_data_rows(subject_code='2501', aux_data=..., subject_name='长期借款')`
3. `fill_sheet(ws, sheet_id='6-1', data_rows=data_rows, ...)` — DT-148: 6-1长期借款Sheet MUST填写
4. 按借款合同逐行填写
5. 关注一年内到期的非流动负债
6. 检查`result['success']`

**勾稽公式**：6-1合计1行账面价值 = BS长期借款

### Step 2d.8 应付利息/应付股利

- **应付利息（5-10-1）**：按债权人填写
  - **[DT-100]** 若BS应付利息=0但科目余额表2231有值→已并入其他应付款，不在5-10-1单独填写
- **应付股利（5-10-2）**：按股东填写

### Step 2d.9 租赁负债（6-3）& 长期应付款（6-4）

> **双行表头注意**：6-3/6-4为双行表头sheet，Row5=检索表头1，Row6=检索表头2，data_start_row=7。
> 6-4长期应付款的账面价值有子列（初始额/利息及汇率净损失/合计），book_value映射到"合计"列(col11)。
> sheet_col_map.json v2.0已修复，fill_sheet()通过col_map自动处理。

1. 从subjects.json读取租赁负债/长期应付款科目
2. `prepare_data_rows(subject_code=..., aux_data=..., subject_name='租赁负债')`
3. `fill_sheet(ws, sheet_id='6-3', data_rows=data_rows, ...)`
4. 检查`result['success']`
5. 长期应付款同理：`fill_sheet(ws, sheet_id='6-4', data_rows=data_rows, ...)`

### Step 2d.10 递延所得税负债（6-8）

1. 从subjects.json读取递延所得税负债科目
2. `prepare_data_rows(subject_code=..., aux_data=..., subject_name='递延所得税负债')`
3. `fill_sheet(ws, sheet_id='6-8', data_rows=data_rows, ...)`
4. 检查`result['success']`

## 输出

- 所有负债类科目Sheet数据已填入
- 差异项已标注
- 每个Sheet已通过fill_sheet()内部回读验证(DT-97)+即时勾稽(DT-158)

### 🚨 后续必经步骤：Phase 2e序时账核实（DT-161）

> 负债类科目填写完成后，如项目有序时账数据，**MUST执行Phase 2e（S2_seq_verify.md）核实发生日期和业务内容**。
> 负债类科目的发生日期取辅助余额表/科目余额表末笔日期，业务内容由`infer_business_content()`推断——两者都需从序时账摘要中核实替换。
> 三层保障确保不跳过：L1=prepare_data_rows(has_journal=True)触发校验；L2=gate_G2(G2-18)检查非空；L3=S5 Step 5.6a交付前兜底。
> **进入Phase 2e**：执行 `S2_seq_verify.md` Step 2e.0~2e.9

## 约束

- **DT-46**：负债类C列=发生日期(日期), D列=业务内容(文字)，与资产类相反 → fill_sheet()内部自动处理
- **DT-30**：发生日期必须为datetime类型 → fill_sheet()内部自动校验
- **DT-0**：零幻觉原则
- **DT-5**：差异必须标注
- **DT-126**：应交税费MUST按BS口径拆分
- **DT-137**：填写行数≥辅助余额表结算对象总数
- **DT-140**：BS金额>1000万科目MUST有数据来源+勾稽公式
- **DT-147**：应交税费逐税种 → fill_sheet()内部自动调用infer_tax_details()
- **DT-148**：长期借款必填
- **DT-160**：禁止直接用openpyxl写入ws，MUST通过fill_sheet()接口

## 异常处理

- 科目余额表无对应科目 → 清空该Sheet数据行，备注说明
- 暂估应付 → 标注"暂估"
- 一年内到期的非流动负债 → 标注提醒
- BS与2241贷方合计差异 → DT-117差额推论映射（2102结算中心借款/2231应付利息）
