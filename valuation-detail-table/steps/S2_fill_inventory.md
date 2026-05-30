# S2_fill_inventory: 存货科目填写

> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2规则节
> **📋 内置规则覆盖**：以下步骤通过fill_sheet()管线自动执行DT-0/46/66/97/120/125/136/143/144/149/150/151/152/153/155/156/157/158/159，Agent无需手动实现

> **🚨 DT-182b 插行后公式完整性修复**：每个 `fill_sheet()` 完成后，MUST 依次执行：
> 1. `fix_summary_sheet_refs(wb, sheet_name)` — 修复汇总表跨Sheet引用行号（仅金额列E-P）
> 2. 全部Phase 2完成后：`fix_intra_sheet_total2_formulas(wb)` — 修复明细表内合计2行公式
> 详见 `scripts/fix_summary_refs.py`。
## 🚨 唯一写入接口（DT-128+DT-160）

**评估明细表数据写入MUST通过`sheet_filler.fill_sheet()`接口执行。**
- ✅ 允许：`fill_sheet(ws, sheet_id, data_rows, ...)` — 12条DT规则内部自动执行
- ❌ 禁止：直接`ws.cell(row=r, column=6).value = xxx` — 绕过全部规则断言
- ❌ 禁止：`import openpyxl`后直接写入ws — DT-160裸openpyxl写入=绕过管线

**openpyxl仅允许**：`load_workbook()`加载 / `save()`保存 / `wb[sheetname]`获取ws对象传给fill_sheet

## ⚠️ 双行表头特殊处理（3-9-2/3-9-5等存货子表）

> 存货子表（3-9-2原材料/3-9-5产成品等）为**双行表头**结构：
> - Row4=标题行（"被评估单位"等）
> - Row5=检索表头1（合并大标题：如"账面价值"占col6-8，"评估价值"占col9-11）
> - Row6=检索表头2（子标题：如"数量"/"单价"/"金额"）
> - Row7+=数据行
>
> **关键映射逻辑**：账面价值应写入检索表头2的**金额**列（如3-9-2的col8），不是检索表头1的大标题列（col6=数量列）。评估价值应写入检索表头2的**金额**列（如3-9-2的col11），不是检索表头1的大标题列（col9=实际数量列）。
>
> **sheet_col_map.json v2.0已正确映射到检索表头2数据列**（2026-05-24修复后），fill_sheet()通过col_map自动写入正确列位，Agent无需手动判断。
> data_start_row=7（R6是检索表头2行，数据从R7开始）。

## ⚠️ DT-145 行业特殊科目映射（房地产企业必须检查）

> 房地产企业的"5002开发成本"MUST映射到3-9存货(在产品)。
> `get_sheet_id_for_subject('5002', industry_type='房地产')` → 自动返回'3-9'
> fill_sheet()内部自动从industry_mapping.json读取映射，Agent无需手动判断。

## 输入

- 科目余额表存货科目数据（1405原材料、1410周转材料、1406库存商品、1407自制半成品等）
- 收发存明细表（可选，用于细分品种）
- 评估明细表模板

## 操作

### Step 2c.1 读取存货科目

从科目余额表提取以下存货相关科目的期末余额：
- 1403/1405 原材料 → sheet_id='3-9-2'
- 1406 库存商品/产成品 → sheet_id='3-9-5'
- 1407 自制半成品 → sheet_id='3-9'
- 1408 委托加工物资 → sheet_id='3-9'
- 1410 周转材料/包装物 → sheet_id='3-9-3'
- 1401 服务成本/合同履约成本（需确认是否计入存货）

### Step 2c.2 按类别调用fill_sheet()填写

每个存货子表统一3步管线：

**Step 1：组织数据 — `prepare_data_rows()`**
```python
from sheet_filler import fill_sheet, prepare_data_rows, get_sheet_id_for_subject

# 原材料示例
data_rows = prepare_data_rows(
    subject_code='1405',
    kmye_data=raw_material_items,
    subject_name='原材料',
    industry_type='房地产',
)
```

**Step 2：写入 — `fill_sheet()`**
```python
ws = wb['3-9-2原材料']
result = fill_sheet(ws=ws, sheet_id='3-9-2', data_rows=data_rows, wb=wb)
```
- 列位由sheet_col_map.json自动映射(DT-153) — 双行表头自动处理，评估价值写入检索表头2的金额列(K列)，不是检索表头1的大标题列
- fill_sheet()内部自动执行：双行表头识别(DT-116)、插行判断(DT-164:数据行数>模板预留行数才插行)、插行(DT-120/152)、列序校验(DT-46/66)、回读验证(DT-97)、即时勾稽(DT-158)、合计行B:C合并校验(DT-163)

**Step 3：检查结果**
```python
if not result['success']:
    sys.exit(1)
```

**行业映射检查**（DT-145）：
```python
# 检查5002开发成本科目（房地产企业特有）
code5002 = [s for s in subjects if s['code'].startswith('5002')]
if code5002:
    sheet_id = get_sheet_id_for_subject('5002', industry_type='房地产')  # → '3-9'
    data_rows = prepare_data_rows(subject_code='5002', kmye_data=code5002, subject_name='开发成本', industry_type='房地产')
    ws = wb['3-9存货']
    result = fill_sheet(ws=ws, sheet_id='3-9', data_rows=data_rows, wb=wb)
```

### Step 2c.3 存货跌价准备

- 传入`provision_amount`参数即可，fill_sheet()内部自动定位"减："行并填写(DT-151/DT-18)
- 若无法确认分项归属，暂全部填入产成品子表，备注"暂全部填入产成品，需收发存明细确认分项"
- **禁止手动查找"减："行写入** — 行定位和列位由fill_sheet()内部处理

### Step 2c.4 合同履约成本特殊处理

- 1401服务成本/合同履约成本：需确认是否在BS存货范围内
- 不在BS存货范围内的科目不填入存货，标注原因

## 输出

- 存货各子表已填写
- 存货跌价准备已由fill_sheet()自动填入"减："行
- 差异项已标注

## 约束

- **DT-0**：零幻觉原则
- **DT-5**：差异必须标注
- **DT-6**：存货跌价准备位置必须正确 → fill_sheet()内部自动处理(DT-151)
- **DT-145**：行业特殊科目映射 → get_sheet_id_for_subject()自动处理
- **DT-160**：禁止直接用openpyxl写入ws，MUST通过fill_sheet()接口
- 合同履约成本(1401)不在BS存货范围内时不填入存货

## 异常处理

- 科目余额表无对应存货科目 → 清空该Sheet，备注说明
- 存货跌价准备无法确认分项 → 暂全部填入产成品，备注待确认
- 合同履约成本是否计入存货不确定 → 以BS口径为准
