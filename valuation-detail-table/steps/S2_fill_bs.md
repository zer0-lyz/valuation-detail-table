# S2_fill_bs: 资产负债表科目填写（货币资金+固定资产+无形资产+长期待摊费用等）

> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2规则节
> **📋 内置规则覆盖**：以下步骤通过fill_sheet()管线自动执行DT-0/46/66/97/120/125/136/143/144/149/150/151/152/153/155/156/157/158/159，Agent无需手动实现

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

## 输入

- 科目余额表数据（末级科目期末余额，_dt_cache/subjects.json）
- 辅助余额表数据（_dt_cache/auxiliary_balance_*.json）
- 资产负债表数据（_dt_cache/bs_balances.json）
- D1/D2/D3映射表（_dt_cache/d1d2d3_mapping.json）
- 评估明细表模板

## 操作

### 2a.1 货币资金类

#### 3-1-1现金

1. 从subjects.json读取1001科目期末余额
2. 用`prepare_data_rows()`组织数据
3. 调用`fill_sheet(ws, sheet_id='3-1-1', data_rows=data_rows, ...)`写入 — 列位由sheet_col_map.json自动映射(DT-153)
4. 检查`result['success']`：False则中断(DT-138)
5. 若期末余额为0，备注栏标注"库存现金期末余额为0"

#### 3-1-2银行存款

1. 从辅助余额表/科目余额表获取各银行账户明细
2. 用`prepare_data_rows()`组织数据 — 按账户逐行，B列=开户银行+支行(DT-149自动映射业务内容)
3. 调用`fill_sheet(ws, sheet_id='3-1-2', data_rows=data_rows, ...)`写入
4. 若只有科目余额表汇总数，填一行汇总，备注"仅有汇总数，待银行对账单PDF核实"
5. 检查`result['success']`

#### 3-1-3其他货币资金

1. 从subjects.json读取1012科目
2. 若科目余额表无该科目，不填入3-1-3，备注说明
3. 若有数据，用`prepare_data_rows()`组织后调用`fill_sheet()`写入
4. **[DT-117] 待确认映射处理**：保证金对账单PDF提取的保证金数据但科目余额表/BS无对应编码→标记为"待确认映射"，不在Phase 2填入3-1-3，备注栏标注原因，待Phase 4差额推论确认后回填

### 2a.2 固定资产类（DT-88/DT-94/DT-21）

> **🚨 关键规则**：固定资产MUST逐项展开填写，**禁止只写汇总行**（DT-88/DT-94）。
> 评估明细表中固定资产必须按类别拆分到4-8-1房屋/4-8-4机器/4-8-5车辆/4-8-6电子设备各子表，每个子表逐项列出具体资产。

**填写流程（3步）**：

**Step 1：确认数据源优先级**
1. **PDF卡片台账**（最优先，DT-88）：Phase -1已提取的PDF固定资产卡片台账
2. **固定资产清单Excel**（次优先）：企业提供的固定资产清单/台账Excel
3. **科目余额表1601子科目**（兜底，DT-94）：1601仅有总科目时，按清单/台账拆分

**Step 2：按类别拆分至子表（DT-94）**
- 科目余额表仅有1601总科目时，MUST按PDF卡片台账或固定资产清单拆分为4-8-1/4-8-4/4-8-5/4-8-6
- 禁止将1601全部归入4-8-4机器设备（DT-94）
- 禁止只写汇总行（DT-88）

**Step 3：各子表逐项填写**
- 对每个子表：用`prepare_data_rows()`组织逐项资产数据 → 调用`fill_sheet(ws, sheet_id='4-8-x', data_rows=data_rows, ...)`写入
- 列位由sheet_col_map.json自动映射(DT-153)，无需手动判断G/H/I/J列含义
- **双行表头**（4-8-x固定资产子表）：Row5=检索表头1（大标题如"账面价值"占2列），Row6=检索表头2（子列如"原值"/"净值"），data_start_row=7
  - **关键**：book_value映射到R6"净值"列（如4-8-4的col12），不是R5"账面价值"大标题列（col11=原值）
  - assessed_value映射到R6"净值"列（如4-8-4的col15），不是R5"评估价值"大标题列（col13=原值）
  - sheet_col_map.json v2.0已修复（2026-05-24），fill_sheet()通过col_map自动写入正确列位
- fill_sheet()内部自动处理：双行表头识别(DT-116)、插行判断(DT-164:数据行数>模板预留行数才插行)、插行(DT-120/152)、合计行保护断言(DT-164.1:插行后合计1/坏账准备/合计2三行A列标记+B列内容+SUM公式完整性)、列序校验(DT-46)、回读验证(DT-97)、即时勾稽(DT-158)、合计行B:C合并校验(DT-163)
- 账面原值/净值MUST取PDF第7-9页（非10-12页净额列）（DT-88）
- 填写后MUST与BS勾稽验证：各子表净值合计=BS固定资产净值

### 2a.3 无形资产

1. 从subjects.json读取1701无形资产科目
2. 用`prepare_data_rows()`组织数据 — 按权属逐项
3. 调用`fill_sheet(ws, sheet_id='4-13-3', data_rows=data_rows, ...)`写入

### 2a.4 长期待摊费用（DT-129）

> **🚨 关键规则**：长期待摊费用MUST按科目余额表1801的二级子科目逐行展开填写，**禁止只写1行汇总**（DT-129）。

**填写流程（3步）**：

**Step 1：从科目余额表提取1801子科目明细**
- 读取科目余额表1801下所有二级子科目
- 每个有余额的二级子科目MUST在4-16中占一行
- 余额方向="平"的子项排除（DT-95）

**Step 2：逐行填写4-16**
- 用`prepare_data_rows()`组织数据 → 调用`fill_sheet(ws, sheet_id='4-16', data_rows=data_rows, ...)`写入
- 列位由sheet_col_map.json自动映射(DT-153)
- C列=摊销起始日期（如无明确日期可从序时账提取首笔发生日期，DT-30要求datetime类型）
- 如有辅助余额表/辅助明细账可进一步展开（参照DT-111原则）

**Step 3：勾稽验证**
- 4-16合计MUST与BS长期待摊费用一致
- 合计≠BS时检查是否存在重分类（DT-118）或子项遗漏

### 2a.5 递延所得税资产

1. 从subjects.json读取1811递延所得税资产科目
2. 用`prepare_data_rows()`组织数据 — DT-150要求名称披露具体内容
3. 调用`fill_sheet(ws, sheet_id='4-17', data_rows=data_rows, ...)`写入

### 2a.6 减值准备填写

- 减值准备由fill_sheet()内部自动处理（DT-151/DT-18）
- 传入`bad_debt_amount`/`provision_amount`参数即可
- fill_sheet()自动定位"坏账准备"/"预计风险"行（通过A列标记DT-116），自动填写正确列位
- **禁止手动查找"减："行写入** — 列位由fill_sheet()内部处理

## 输出

- 货币资金三个子表已填写
- 固定资产各子表已填写
- 无形资产已填写
- 长期待摊费用已填写
- 减值准备已填入对应科目"减："行
- 差异项已在备注栏标注

### 🚨 后续必经步骤：Phase 3序时账查阅（DT-161）

> 资产类科目填写完成后，如项目有序时账数据，**MUST执行Phase 3（S3_journal_extract.md）核实发生日期和业务内容**。
> 资产类科目的发生日期取辅助余额表/科目余额表末笔日期，业务内容由`infer_business_content()`推断——两者都需从序时账摘要中核实替换。
> 四层保障确保不跳过：L1=journal_extractor.py封装完整流程(DT-166)；L2=gate_G2(G2-18)检查非空；L3=Phase 5勾稽字段完整性校验；L4=S6 Step 6.6a交付前兜底。
> **进入Phase 3**：执行 `S3_journal_extract.md` Step 3.0~3.8

## 约束

- **DT-0**：零幻觉原则，每个数值必须有科目余额表或资产负债表数据支撑
- **DT-5**：科目余额表与资产负债表差异必须标注，不自行调和
- **DT-8**：openpyxl加载时保留公式（data_only=False）
- **DT-21**：固定资产明细逐行核对源报表
- **DT-160**：禁止直接用openpyxl写入ws，MUST通过fill_sheet()接口

## 异常处理

- 科目余额表无对应科目 → 清空该Sheet数据行，备注说明
- 库存现金期末余额为0 → 备注说明
- 固定资产无台账 → 按科目余额表汇总填写，备注"需资产台账"，但仍必须按类别拆分（DT-88）
- 减值准备科目在科目余额表中不存在 → `bad_debt_amount=None`，fill_sheet()自动跳过减值行
