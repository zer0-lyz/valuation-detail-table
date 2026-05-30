# S2_fill_re: 往来科目填写（应收/应付/预收/预付/其他应收应付）

> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2规则节
> **📋 内置规则覆盖**：以下步骤通过fill_sheet()管线自动执行DT-0/46/66/97/120/125/136/143/144/149/150/151/152/153/155/156/157/158/159/163/164，Agent无需手动实现

> **🚨 DT-182b 插行后公式完整性修复**：每个 `fill_sheet()` 完成后，MUST 依次执行：
> 1. `fix_summary_sheet_refs(wb, sheet_name)` — 修复汇总表跨Sheet引用行号（仅金额列E-P）
> 2. 全部Phase 2完成后：`fix_intra_sheet_total2_formulas(wb)` — 修复明细表内合计2行公式
> 详见 `scripts/fix_summary_refs.py`。
## 🚨 Step 0: 脚本选择（DT-165强制，编写任何代码前MUST执行）

> **⚠️ 本步骤在编写任何Python代码之前必须执行。禁止跳过。**
> 复盘根因：v2脚本完全绕开了fill_sheet()，从零写了fill_sheet_data()→6/9个Sheet列位错误+合计行被覆盖+Phase 2e被跳过。

**执行判定**：

| 条件 | 动作 |
|------|------|
| 需求能被已有脚本覆盖 | **MUST调用已有脚本**（import+调用），禁止从零重写 |
| 已有脚本不满足 | **MUST在已有脚本基础上增量扩展**，禁止另起炉灶 |
| 确需新写脚本 | **MUST注释原因+更新索引** |

**必读索引**：Read `scripts/SKILL_SCRIPT_INDEX.md`，确认以下关键接口：

| 需求 | 必须调用 | 禁止自写 |
|------|---------|---------|
| 数据写入评估明细表 | `fill_sheet()` | ❌ 自写`fill_sheet_data()`/裸openpyxl写入 |
| 插入行 | `smart_insert_row()` | ❌ 裸`ws.insert_rows()` |
| 列位映射 | `sheet_col_map.json`/`_build_col_map()` | ❌ 硬编码`column=9` |
| 发生日期/业务内容 | `load_journal_data()` | ❌ 自写序时账解析 |
| 业务内容推断 | `infer_business_content()` | ❌ 仅填科目名 |
| Gate验证 | `gate_G2()` | ❌ 跳过验证 |
| 数据加载 | `load_subject_data()`/`load_auxiliary_balance()` | ❌ 自写科目余额表解析 |

## 🚨 唯一写入接口（DT-128+DT-160）

**评估明细表数据写入MUST通过`sheet_filler.fill_sheet()`接口执行。**
- ✅ 允许：`fill_sheet(ws, sheet_id, data_rows, ...)` — 12条DT规则内部自动执行
- ❌ 禁止：直接`ws.cell(row=r, column=10).value = xxx` — 绕过全部规则断言
- ❌ 禁止：`import openpyxl`后直接写入ws — DT-160裸openpyxl写入=绕过管线=4类错误必现

**openpyxl仅允许**：`load_workbook()`加载 / `save()`保存 / `wb[sheetname]`获取ws对象传给fill_sheet

## 🚨 DT-46: 业务内容与发生日期严禁混淆（最高优先级）

> fill_sheet()内部通过sheet_col_map.json自动处理列序差异，Agent无需手动判断C/D列含义。
> 但Agent必须理解业务逻辑差异，以便正确组织data_rows数据。

**资产类科目**（应收账款/预付款项/其他应收款/合同资产）：
- data_rows中`business_content`字段 = 文字（如"管理服务费"、"代扣住房公积金"）
- data_rows中`date`字段 = datetime对象（如`datetime(2023,1,5)`）

**负债类科目**（应付账款/其他应付款/预收款项）：
- data_rows中`date`字段 = datetime对象
- data_rows中`business_content`字段 = 文字

**合同资产列序特殊**：3-10合同资产sheet的列定义为C=发生日期, D=结算内容（与负债类列序相同！但属于资产类科目取借方末笔）

## 输入

- 科目余额表明细数据（各往来科目末级科目及结算对象）
- **辅助余额表提取结果**（DT-111 Step 0.5输出的"科目→结算对象清单"）
- 序时账/明细账（**本Phase不处理**，由独立的Phase 2e执行核实，DT-161）
- 评估明细表模板

## 操作

### 🚨 DT-111: 辅助余额表强制引用（填写前必检）

> 往来科目填写前MUST先检查DT-111 Step 0.5输出的"科目→结算对象清单"。如有辅助余额表数据，MUST按结算对象逐行填写，禁止仅填科目余额表汇总数。

**填写前检查流程**：

```
填写往来科目Sheet
  │
  ├── 检查"科目→结算对象清单"中是否有该科目的辅助余额表数据
  │     │
  │     ├── 有 → MUST按辅助余额表结算对象逐行填写
  │     │     └── 合计MUST与科目余额表/BS勾稽一致
  │     │
  │     └── 无 → 使用科目余额表末级子科目逐行填写
  │           └── 两者取并集（DT-111-3）
  │
  └── 填写完成后反向校验：辅助余额表结算对象是否全部覆盖（DT-111-4）
```

### 通用填写步骤（所有往来科目适用）

每个往来科目Sheet的填写遵循统一3步管线：

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
    bad_debt_amount=None,           # 往来科目通常无坏账准备
    provision_amount=None,
)
```
- fill_sheet()内部自动执行以下**7步管线**（Agent不得跳过任何步骤）：
  1. **Step 2a: 结构识别** — `_find_header_structure(ws)`获取data_start_row/total_row/bad_debt_row/total2_row
  2. **Step 2b: 列位映射** — `_build_col_map(ws)`从sheet_col_map.json读取列位(DT-153)
  3. **Step 2c: 插行判断** — 比较`needed_rows`与`template_data_rows`(DT-164)
     - **needed_rows > template_data_rows → 必须调用`smart_insert_row()`插入差额行**
     - needed_rows ≤ template_data_rows → 直接写入，不插行
  4. **Step 2d: 合计行保护断言** — 插行完成后，立即断言合计1/坏账准备/合计2三行完整(DT-164.1)
     - A列标记必须存在("合计1"/"坏账准备"/"合计2")
     - B列含"合"/"减"关键字
     - 账面价值列有SUM公式或非零值
     - **任一断言失败 = RuntimeError阻断写入**
  5. **Step 2e: 数据写入** — 逐行写入，列位由col_map自动映射
  6. **Step 2f: 回读验证** — 每行写入后立即回读校验(DT-97)
  7. **Step 2g: 即时勾稽** — 写入后与科目余额表比对(DT-158)
- Agent不需要知道C/D列具体是什么，sheet_col_map.json + fill_sheet()自动处理

**Step 3：检查结果 — DT-138**
```python
if not result['success']:
    print(f"🚨 fill_sheet失败: {result['gate_errors'] + result['read_back_errors']}")
    sys.exit(1)
```

**Step 4：🚨 Phase 2e序时账核实 — 必经下一步（DT-161）**

> **⚠️ 本步骤不是可选的"后续优化"，而是填写完成后的必经下一步骤。**
> Phase 2填写的数据中，发生日期取辅助余额表/科目余额表的末笔日期，业务内容由`infer_business_content()`推断——两者都可能不准确。
> Phase 2e从序时账摘要中核实真实发生日期和业务内容，是保证底稿可审的必要操作。

**执行判定（有序时账时MUST执行）**：

| 条件 | 动作 |
|------|------|
| 项目有序时账数据 | **MUST执行S2_seq_verify.md全部步骤** — 本Step结束→进入Phase 2e |
| 使用人未提供序时账 | 跳过Phase 2e，发生日期留空(DT-143) |
| 使用人明确要求不填写发生日期 | 跳过Phase 2e，记录使用人要求(DT-161②) |

**三层物理保障确保Phase 2e不被跳过**：
1. **L1 脚本层**：`prepare_data_rows(has_journal=True)`传入fill_sheet()，触发date列非空校验
2. **L2 Gate层**：gate_G2(G2-18)检查发生日期+业务内容列非空，CRITICAL=禁止进入Phase 3
3. **L3 交付层**：S5_deliver Step 5.6a交付前兜底验证，发生日期/业务内容缺失=禁止交付

**进入Phase 2e**：执行 `S2_seq_verify.md` Step 2e.0~2e.9 全部步骤

### D列业务内容精炼规则

| 原始描述 | 精炼为 | 说明 |
|---------|--------|------|
| 收货款/采购款 | 货款 | 统一用"货款" |
| 退回/退货 | 退货款 | |
| 投标保证金 | 投标保证金 | |
| 办公/信息技术押金 | 押金 | |
| 个人借款/备用金 | 备用金 | |
| 报销/费用报销 | 报销款 | |
| 往来款/集团内部 | 往来款 | 关联方往来标注"往来款" |
| 暂估应付 | 暂估货款 | |

> fill_sheet()内部通过`infer_business_content()`自动执行精炼映射（DT-149），Agent无需手动逐条映射

### 坏账准备/减值准备填写

- 传入`bad_debt_amount`参数即可，fill_sheet()内部自动定位"坏账准备"/"预计风险"行（DT-151/DT-18）
- **禁止手动查找"减："行写入** — 行定位和列位由fill_sheet()内部处理
- DT-18规则：坏账准备行仅填账面价值列，预计风险行仅填评估价值列

### 应收账款（3-5）

1. 从辅助余额表获取应收账款结算对象数据（DT-111优先）
2. `prepare_data_rows(subject_code='1122', aux_data=..., subject_name='应收账款')`
3. `fill_sheet(ws, sheet_id='3-5', data_rows=data_rows, bad_debt_amount=坏账金额, ...)`
4. 区分正常应收和合同资产重分类
5. 检查`result['success']`

### 预付款项（3-7）

1. 从辅助余额表获取预付款项结算对象数据（DT-111优先）
2. `prepare_data_rows(subject_code='1123', aux_data=..., subject_name='预付款项')`
3. `fill_sheet(ws, sheet_id='3-7', data_rows=data_rows, ...)` — 列位由sheet_col_map.json自动映射(DT-153)，不再手动判断C/D列含义
4. 金额较大关注期后到货情况
5. 检查`result['success']`

### 其他应收款（3-8-3）

1. 从辅助余额表获取其他应收款结算对象数据（DT-111优先）
2. `prepare_data_rows(subject_code='1221', aux_data=..., subject_name='其他应收款')`
3. `fill_sheet(ws, sheet_id='3-8-3', data_rows=data_rows, bad_debt_amount=坏账金额, ...)`
4. 关注押金/保证金/备用金分类
5. 检查`result['success']`

### 合同资产（3-10）

1. 从辅助余额表获取合同资产结算对象数据
2. `prepare_data_rows(subject_code='..., aux_data=..., subject_name='合同资产')`
3. `fill_sheet(ws, sheet_id='3-10', data_rows=data_rows, ...)`
4. ⚠️ 列序特殊：C=发生日期, D=结算内容（与负债类相同，但sheet_col_map.json已正确映射）
5. 检查`result['success']`

### 应付账款（5-5）

1. 从辅助余额表获取应付账款结算对象数据（DT-111优先，通常有3+类辅助余额表：应付工程款/暂估应付/集团内部等）
2. `prepare_data_rows(subject_code='2202', aux_data=..., subject_name='应付账款')`
3. `fill_sheet(ws, sheet_id='5-5', data_rows=data_rows, ...)` — 负债类列序由sheet_col_map.json自动处理(DT-46)
4. 区分正常应付和暂估应付
5. **DT-137**：MUST校验填写行数≥辅助余额表结算对象总数
6. 检查`result['success']`

### 预收款项（5-6）

1. 从辅助余额表获取预收款项结算对象数据（DT-111优先）
2. `prepare_data_rows(subject_code='2203', aux_data=..., subject_name='预收款项')`
3. `fill_sheet(ws, sheet_id='5-6', data_rows=data_rows, ...)`
4. 关注长期挂账的预收
5. 检查`result['success']`

### 其他应付款（5-10-3）

1. 从辅助余额表获取其他应付款结算对象数据（DT-111优先，通常有3类辅助余额表：集团外部其他/集团内部其他/集团内部部门，需合并填写）
2. `prepare_data_rows(subject_code='2241', aux_data=..., subject_name='其他应付款')`
3. `fill_sheet(ws, sheet_id='5-10-3', data_rows=data_rows, ...)`
4. 关注集团内部往来（往来款）
5. 检查`result['success']`

## 输出

- 所有往来科目Sheet数据已按结算对象逐行填入
- 坏账准备/减值准备已由fill_sheet()自动填入"减："行
- 业务内容已由infer_business_content()自动精炼(DT-149)——⚠️ 此推断仅基于科目编码+结算对象名称，结果可能是通用文字（如"销售商品/提供服务""其他应收款"），**Phase 2e Step 2e.5将从序时账摘要归纳替换为具体业务实质（DT-60）**
- 差异项已在备注栏标注

## 约束

- **DT-46**：业务内容与发生日期严禁混淆（最严重历史事故）→ fill_sheet()内部通过sheet_col_map.json自动处理
- **DT-30**：发生日期必须为datetime类型，禁止字符串 → fill_sheet()内部自动校验
- **DT-18**：坏账准备/预计风险损失的I/J列填写规则 → fill_sheet()内部自动处理(DT-151)
- **DT-0**：零幻觉原则
- **DT-5**：科目余额表与资产负债表差异必须标注
- **DT-8**：openpyxl加载时保留公式
- **DT-111**：🚨 辅助余额表强制引用
- **DT-137**：结算对象总数校验
- **DT-160**：禁止直接用openpyxl写入ws，MUST通过fill_sheet()接口
- **DT-161**：🚨 Phase 2e序时账核实不可跳过（有序时账时MUST执行），**本Step 4为必经下一步骤**——填完往来科目后MUST立即进入Phase 2e(S2_seq_verify.md)。三层保障：L1=prepare_data_rows(has_journal=True)触发date列校验；L2=gate_G2(G2-18)检查非空；L3=S5 Step 5.6a交付前兜底。业务内容核实也由Phase 2e从序时账摘要归纳替换（DT-60）

## 异常处理

- 科目余额表无对应科目 → 清空该Sheet数据行，备注说明
- 结算对象数量超过模板预留行 → fill_sheet()内部自动判断(DT-164)并调用smart_insert_row()插入(DT-120/152)，自动校验合计行B:C合并(DT-163)
- 科目余额表与BS差异 → 备注栏标注差异金额和原因（DT-5）
- 暂估应付账款处理 → 标注"暂估"
