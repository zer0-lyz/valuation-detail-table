# S4: 公式修复与格式修复（格式集中处置区）

> **📋 Common规则适用声明**：本步骤适用 META_RULES MR-1/MR-4/MR-5/MR-6/MR-7/MR-9/MR-10 + preparation_discipline_rules G0-G13 + G5(版本承接)
> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2-3规则节+Phase 3规则节：DT-2(插入行公式修复)、DT-6(减值行填写)、DT-24(删行公式修复)、DT-33(COM resave datetime)、DT-67(公式列覆写禁止)、DT-75(子表头保护)、DT-76(增值率格式)、DT-77(行高统一)、DT-78(结构行保护)、DT-82(空白行格式)、DT-83(合并单元格验证)、DT-84(A列居中)、DT-85(合计行公式)、DT-112(格式集中处置)、DT-113(禁止裸insert)、DT-114(验证-修复闭环)、DT-120(smart_insert_row强制)、DT-161(Phase 3（序时账查阅）前置检查)、DT-163(合计行B:C合并强制校验)

> **🎯 Phase定位（DT-112）**：本Phase是**格式集中处置区**。Phase 2仅负责"写数据+最小格式继承"（插入行时copy参考行格式+SUM即时扩展），所有深度格式修复统一在本Phase集中完成。Phase 2不做行高统一、边框扫描、数字格式校验、合并单元格验证等深度格式操作——这些全部在本Phase完成。**分段集中的原因**：填写与格式职责分离，降低认知负荷，减少因"填写→格式→填写"反复切换导致的遗漏。

> **🔄 闭环原则（DT-114）**：本Phase所有步骤遵循"验证→修复→重验"闭环。Phase 2使用`smart_insert_row`已自动完成的项（SUM扩展、格式继承、跨sheet引用）降级为"仅验证"；验证失败时自动触发`auto_fix_formats`修复并重验（最多3次），3次仍FAIL则BLOCKED。

## 🚨 Step 4.pre: Phase 3前置Gate检查（DT-161）

> **本步骤在Step 4.0之前执行，是Phase 4的入口检查。未通过=强制回退Phase 3执行S3_journal_extract.md。**

**执行代码**（直接运行，非声明）：

```python
import openpyxl, os, json

# === Step 4.pre.1: 判定项目是否有序时账数据 ===
has_journal_data = False
cache_dir = os.path.join(PROJECT_DIR, '_dt_cache')
manifest_path = os.path.join(cache_dir, 'file_manifest.json')
if os.path.exists(manifest_path):
    with open(manifest_path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    for item in manifest.get('files', []):
        fname = item.get('name', '').lower()
        if '序时账' in fname or '明细账' in fname or 'journal' in fname:
            has_journal_data = True
            break
if not has_journal_data:
    for f in os.listdir(PROJECT_DIR):
        if '序时账' in f or '明细账' in f:
            has_journal_data = True
            break

# === Step 4.pre.2: 扫描往来科目Sheet的发生日期列 ===
wb = openpyxl.load_workbook(DETAIL_FILE, data_only=True)
receivable_sheets = {
    '3-5应收账款': {'date_col': 5, 'type': 'asset'},       # 资产类D列=发生日期(模板col5)
    '3-7预付款项': {'date_col': 5, 'type': 'asset'},
    '3-8-3其他应收款': {'date_col': 5, 'type': 'asset'},
    '5-5应付账款': {'date_col': 4, 'type': 'liability'},   # 负债类C列=发生日期(模板col4)
    '5-6预收款项': {'date_col': 4, 'type': 'liability'},
    '5-10-3其他应付款': {'date_col': 4, 'type': 'liability'},
}

all_date_empty = True
checked_sheets = 0
for sname, info in receivable_sheets.items():
    if sname not in wb.sheetnames:
        continue
    ws = wb[sname]
    if ws.sheet_state == 'hidden':
        continue
    checked_sheets += 1
    date_col = info['date_col']
    for r in range(7, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        if a_val and ('合' in str(a_val) or '减' in str(a_val)):
            break
        date_val = ws.cell(row=r, column=date_col).value
        if date_val is not None and date_val != '':
            all_date_empty = False
            break
    if not all_date_empty:
        break
wb.close()

# === Step 4.pre.3: 判定 ===
# 注意：使用人明确要求不填写发生日期时，Phase 3（序时账查阅）已在Step 2e.0a中跳过并标注，
# 此处不再重复判定。若Phase 3（序时账查阅）被跳过，往来科目发生日期列必然为空，
# 但has_journal_data=True且标注了"使用人明确要求(DT-161②)"时，Gate应通过。
# 此处通过读取缓存中的Phase 3（序时账查阅）执行标注来判断：
user_skip_phase2e = False
phase2e_status_path = os.path.join(cache_dir, 'phase3_status.json')
if os.path.exists(phase2e_status_path):
    with open(phase2e_status_path, 'r', encoding='utf-8') as f:
        p2e_status = json.load(f)
    if p2e_status.get('skipped_reason') == 'user_explicit':
        user_skip_phase2e = True

if checked_sheets == 0:
    print("✅ Gate通过: 无往来科目Sheet，无需Phase 3（序时账查阅）")
elif not all_date_empty:
    print("✅ Gate通过: 往来科目发生日期已有值，Phase 3（序时账查阅）已执行或部分执行")
elif user_skip_phase2e:
    print("✅ Gate通过: 使用人明确要求不填写发生日期(DT-161②)")
elif has_journal_data:
    print("🚨 GATE阻断(DT-161): 往来科目发生日期全空且项目有序时账数据")
    print("   → MUST回退执行S3_journal_extract.md Step 2e.0a~2e.9")
    print("   → 执行完毕后重新进入Phase 4，再次运行本Gate检查")
    # 回退执行Phase 3（序时账查阅）后，Gate应通过
else:
    print("✅ Gate通过: 未提供序时账(DT-143)，发生日期留空")
```

**Gate未通过时的动作**（DT-161强制）：
1. 输出`🚨 Phase 3（序时账查阅） GATE阻断`
2. **立即回退执行S3_journal_extract.md Step 2e.0a~2e.9**
3. Phase 3（序时账查阅）执行完毕后重新进入Phase 4，再次运行Step 4.pre
4. 第二次仍不通过 → BLOCKED

**Step 4.pre.4: 业务内容质量校验（DT-149 + DT-60）**

> 业务内容填写质量校验：检测"仅填科目名"或"通用模板文字"的不达标情况，WARNING提醒但不阻断。

```python
# DT-149 + DT-60: 业务内容质量校验
wb_biz = openpyxl.load_workbook(DETAIL_FILE, data_only=True)

GENERIC_BIZ_CONTENTS = {
    '其他应收款', '其他应付款', '其他往来', '往来款',
    '销售商品/提供服务', '采购商品/接受服务', '预付货款/服务费', '预收货款/服务费',
}

re_biz_sheets = {
    # 资产类：D列=业务内容
    '3-5应收账款': {'biz_col': 4, 'type': 'asset'},
    '3-7预付款项': {'biz_col': 4, 'type': 'asset'},
    '3-8-3其他应收款': {'biz_col': 4, 'type': 'asset'},
    # 负债类：E列=业务内容
    '5-5应付账款': {'biz_col': 5, 'type': 'liability'},
    '5-6预收款项': {'biz_col': 5, 'type': 'liability'},
    '5-10-3其他应付款': {'biz_col': 5, 'type': 'liability'},
}

biz_warnings = 0
for sname, info in re_biz_sheets.items():
    if sname not in wb_biz.sheetnames:
        continue
    ws = wb_biz[sname]
    if ws.sheet_state == 'hidden':
        continue
    for r in range(7, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        if a_val and ('合' in str(a_val) or '减' in str(a_val) or '预' in str(a_val)):
            break
        biz_val = ws.cell(row=r, column=info['biz_col']).value
        if biz_val and str(biz_val).strip() in GENERIC_BIZ_CONTENTS:
            biz_warnings += 1
            name_val = ws.cell(row=r, column=3).value
            print(f"  ⚠️ 业务内容不达标 | {sname} Row{r} | 结算对象={str(name_val)[:20] if name_val else '?'} | 业务内容='{biz_val}' (DT-149: 通用模板文字，Phase 3（序时账查阅）应从序时账摘要归纳)")

if biz_warnings > 0:
    print(f"\\n⚠️ 业务内容质量WARNING: {biz_warnings}行的业务内容为通用模板文字")
    if has_journal_data:
        print("   → 项目有序时账，Phase 3（序时账查阅） Step 2e.5应从序时账摘要归纳业务内容(DT-60)")
    else:
        print("   → 项目无序时账，业务内容由infer_business_content()推断，质量受限")
else:
    print("\\n✅ 业务内容质量校验通过: 往来科目业务内容均非通用模板文字")
wb_biz.close()
```


## 输入

- 已填写数据的评估明细表（Phase 2已完成数据写入和最小格式继承）

## 操作

> **⚠️ 这是最容易遗漏的环节，也是最关键的环节。历史教训：多次因插入行后未修复公式导致合计错误、勾稽失败。**

### Step 4.0 操作结果断言（⚠️ Phase 2+3结束后强制执行，DT-114）

> **Phase 2中每次调用smart_insert_row/smart_delete_rows后，MUST立即断言返回值。**

```python
from excel_row_ops import smart_insert_row, assert_result

result = smart_insert_row(ws, target_row=10, count=3, total_row=25, wb=wb)
if not assert_result(result):
    # 检测到静默失败（如合计行未识别、SUM未扩展），触发验证-修复闭环
    from excel_row_ops import validate_and_fix
    vf = validate_and_fix(filepath, max_retries=3)
    if not vf['passed']:
        raise RuntimeError(f"Phase 2行操作存在不可自动修复的问题: {vf['remaining_issues']}")
```

**断言检查项**：
- 合计行是否被正确识别（`old_total_row`不为None）
- SUM扩展是否有结果（`sum_extended`非空 或 合计行无SUM公式）
- 新合计行号是否合理（`new_total_row` > `old_total_row`）

### Step 4.1 公式修复（⚠️ 降级为验证，DT-114闭环保障）

> **如果Phase 2已使用 `excel_row_ops.smart_insert_row()`，SUM扩展和跨sheet引用更新已自动完成，本步骤主要做验证。如果Phase 2仍使用了裸 `ws.insert_rows()`，则必须手动修复。**

**插入行后必须检查的三类公式：**

| 公式类型 | 检查方法 | 修复方法 |
|---------|---------|---------|
| **SUM范围** | 检查合计行SUM公式范围是否覆盖插入行 | `SUM(F6:F28)` → `SUM(F6:F38)` |
| **合计行引用** | 检查最终合计行引用的行号是否正确 | `F28` → `F38`（合计行上移） |
| **跨sheet引用** | 检查汇总表引用子表的行号是否正确 | `'3-5应收账款'!I29` → `'3-5应收账款'!I42` |

**DT-114闭环方式——使用validate_and_fix：**

```python
from excel_row_ops import validate_and_fix

# 验证→修复→重验闭环（最多3次）
vf = validate_and_fix(filepath, max_retries=3)
if not vf['passed']:
    print(f"🚨 自动修复{vf['retries']}次后仍存在问题:")
    for issue in vf['remaining_issues']:
        print(f"  ❌ {issue}")
    # BLOCKED: 禁止进入Phase 4
else:
    print(f"✅ 验证通过 (修复{len(vf['fixes_applied'])}项)")
```

**手动修复脚本模板（仅限validate_and_fix无法自动修复时）：**

```python
import openpyxl

filepath = 'path/to/file.xlsx'
wb = openpyxl.load_workbook(filepath)

# 1. 修复SUM范围
ws = wb['Sheet名']
ws['F27'] = '=ROUND(SUM(F6:F26),2)'  # 修改为正确范围

# 2. 修复合计行引用
ws['I42'] = '=I40-I41'  # 修改为正确行号

# 3. 修复跨sheet引用
ws_sum = wb['3-流动资产汇总']
ws_sum['D10'] = "='3-5应收账款'!I42"  # 修改为正确行号

# 4. 修复隐藏汇总表引用（⚠️ DT-61: 隐藏≠不更新，必须同等对待）
HIDDEN_SUMMARY_SHEETS = ['3-辅-流动资产汇总', '3-9-辅-存货汇总', '8-减值准备汇总表', '9-非财务信息汇总表']
for hidden_name in HIDDEN_SUMMARY_SHEETS:
    if hidden_name in wb.sheetnames:
        ws_hidden = wb[hidden_name]
        for row in ws_hidden.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    pass  # 检查是否引用了行号变动的sheet，更新行号

wb.save(filepath)
```

### Step 4.2 深度格式修复（⚠️ DT-114闭环保障，auto_fix_formats可自动执行3.2.1~3.2.4）

> **Phase 2已完成"最小格式继承"（插入行时copy参考行格式），本步骤在此基础上做全量深度扫描和修复。3.2.1~3.2.4已封装在`auto_fix_formats`中，可在validate_and_fix中自动触发。3.2.5~3.2.8由Agent执行校验并记录结果，无需人工确认（DT-151）。**

**3.2.1 行高统一 [DT-77] ✅ auto_fix_formats可自动修复（G1F-2）**

同一Sheet内所有数据行行高MUST一致（与第一行数据行高对齐）。不同Sheet模板默认行高可能不同（银行存款18.75、职工薪酬18.0、应交税费17.25、其他16.5），不强制统一值，但同一Sheet内MUST一致。

**3.2.2 数字格式校验与修复 [DT-76] ✅ auto_fix_formats可自动修复（G1F-1）**

增值额列和增值率列MUST使用与模板一致的格式`#,##0.00_);[Red]\-#,##0.00_);_(* ""_)`。增值率列格式与增值额相同（非0.00%），因为模板中增值率公式乘以100。

**3.2.3 结构行A列居中对齐 [DT-84] ✅ auto_fix_formats可自动修复（G1F-3）**

所有合计行、减值行（"减：xxx准备"）的A列MUST设置为`horizontal='center'`。

**3.2.4 边框完整性扫描 [DT-82] ✅ auto_fix_formats可自动修复（G1F-4）**

数据行区域（序号1所在行至合计行上方）内，所有行MUST保持与模板一致的完整格式（thin边框）。

**3.2.5 多行表头合并单元格验证 [DT-83] ⚠️ Agent校验并记录**

修复后MUST对比模板与当前sheet的merged_cells.ranges，确认合并范围完全一致。差集`template_merges - current_merges`不为空=必须补充。子表头行所有单元格值MUST与模板一致。auto_fix_formats仅能检测子表头全部为空的严重情况，其他由Agent执行校验+记录结果+标注[待核实]（DT-151）。

**3.2.6 合计行下方清理 [DT-78②] ⚠️ Agent校验并记录**

合计行下方的所有行（模板预留空行）MUST清理：清除边框/合并/值/格式，使其为真正的空白行。

**3.2.7 空白数据行格式补齐 [DT-82②] ⚠️ Agent校验并记录**

数据行区域内空白行MUST保留thin边框和公式（如G列=F-E、H列=IF(E=0,"",G/E*100)），合计行前必须有分隔行（A:B合并，无边框）。

**3.2.8 公式列覆写检查 [DT-67] ⚠️ auto_fix_formats可检测但不可自动修复（G1F-6）**

J列（增值额=评估价值-账面价值）和K列（增值率%）为公式列，MUST保留原公式。Phase 2如有数值覆写，auto_fix_formats会报告但**不自动恢复**（公式结构可能复杂）。Agent校验并记录覆写情况，标注[待核实]（DT-151），不暂停等待。

### Step 4.3 空行删除（DT-25）

- 数据行与合计行之间的空行应使用`smart_delete_rows`删除，而非裸`ws.delete_rows`
- `smart_delete_rows`自动完成：SUM范围收缩、跨sheet引用更新、打印范围调整

```python
from excel_row_ops import smart_delete_rows, assert_result

result = smart_delete_rows(ws, start_row=empty_start, count=empty_count, 
                           total_row=total_row, wb=wb)
assert_result(result)
```

### Step 4.4 删除行后公式修复（DT-24）→ 降级为验证

> **使用`smart_delete_rows`后，SUM范围收缩和跨sheet引用更新已自动完成，本步骤降级为验证。**

- 合计行SUM范围需收缩 → ✅ `smart_delete_rows`已自动完成
- 行内自引用行号需更新 → ✅ openpyxl delete_rows已自动调整
- 跨sheet引用行号需更新 → ✅ `smart_delete_rows`已自动完成
- 残留合并单元格清理 → 需额外检查

### Step 4.5 全量格式扫描 → 降级为validate_and_fix闭环

> **G1F-1~G1F-6已封装在validate_and_fix中，自动完成验证+修复+重验。**

```python
from excel_row_ops import validate_and_fix

vf = validate_and_fix(filepath)
if not vf['passed']:
    # 处理remaining_issues
```

### Step 4.6 合计行格式统一（DT-22 + DT-163）

- 所有明细表的合计行（合　计、减：xxx减值准备等）统一合并A:C列
- 设置居中对齐和thin边框
- **DT-163 合计行B:C合并强制校验**：
  ```python
  # DT-163: 每个明细表的合计行MUST有B:C合并
  for ws in wb.worksheets:
      total_row = None
      for row in range(1, ws.max_row + 1):
          a_val = ws.cell(row=row, column=1).value
          if a_val and isinstance(a_val, str) and '合' in a_val and '计' in a_val:
              total_row = row
              break
      if total_row:
          # 检查B:C合并是否存在
          has_bc_merge = any(
              mr.min_row == total_row and mr.min_col == 2 and mr.max_col == 3
              for mr in ws.merged_cells.ranges
          )
          if not has_bc_merge:
              # 检查是否被跨行扩展覆盖（B:C合并跨多行=错误）
              for mr in list(ws.merged_cells.ranges):
                  if mr.min_col == 2 and mr.max_col == 3 and mr.max_row > mr.min_row:
                      ws.unmerge_cells(str(mr))
                      print(f'DT-163: 取消跨行B:C合并 {mr}')
              # 重建合计行B:C合并
              ws.merge_cells(f'B{total_row}:C{total_row}')
              print(f'DT-163: 重建合计行B:C合并 B{total_row}:C{total_row}')
  ```

### Step 4.7 标题行合并单元格检查（DT-34）

- 插入/删除行操作后，检查标题行合并单元格（A1:H1和A2:H2）是否丢失
- 丢失时恢复：`ws.merge_cells('A1:H1')`/`ws.merge_cells('A2:H2')`，设居中对齐

### Step 4.8 合计行唯一性验证（DT-35）

- 每个明细表只有1个合计行
- 逐行grep"合计"确认唯一
- 空数据行清空所有内容（含A列序号和SUM公式）
- 合计行SUM范围不含中间汇总行

### Step 4.9 小计/减值行位置规范（DT-42）

- "合计"（小计）、"减：减值准备"、"减：减值损失"行应紧贴最终合计行上方
- 当数据行与合计行之间有空白行时，小计/减值行需从数据行下方移至合计行上方
- 移动后中间空白行补齐格式

### Step 4.10 行内容移动后跨sheet引用更新（DT-43）

- 将明细表中的行内容移动到新行号时，除了更新本表公式外，还必须搜索所有汇总表中对旧行号的引用并更新为新行号
- ⚠️ DT-61补充：隐藏汇总表引用同样必须更新
- ⚠️ DT-62补充：辅汇总表与可见汇总表引用逻辑不同，禁止简单复制可见汇总表的引用公式

### Step 4.11 空白无格式行清理（DT-44）

- 明细表中数据行与合计行之间不允许出现空白无格式的行
- 必须使用`smart_delete_rows`删除，将合计行上移至紧贴最后数据行

### Step 4.12 打印范围精确匹配（DT-26/DT-36）→ 降级为验证

> **`smart_insert_row`/`smart_delete_rows`已自动更新打印范围，本步骤降级为验证。**

- 每个工作表的print_area必须精确覆盖至最后一行有效内容（含合计行），不能超出
- 删除行或隐藏行后必须同步调整print_area
- **DT-36**：打印区域列范围必须覆盖右侧索引列

### Step 4.13 Phase 4.5 自动化验证门控（Phase 4完成后必须执行，未通过禁止进入Phase 4）

> **本步骤为强制性验证，必须在脚本中实现并输出结果。任一检查项未通过=明细表不可交付。DT-114要求：验证失败时自动触发修复闭环。**

```python
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
import re

def validate_detail_table(wb, filepath=None, auto_fix=True, max_retries=3):
    """评估明细表质量验证，任一检查未通过=不可交付
    
    DT-114增强：auto_fix=True时，验证失败自动触发修复闭环
    """
    issues = []

    for sname in wb.sheetnames:
        ws = wb[sname]
        if ws.sheet_state == 'hidden':
            continue
        if '汇总' in sname or sname.startswith('0') or sname.startswith('2-'):
            continue

        header_row = total_row = last_data_row = None
        total_rows = []
        for r in range(1, min(ws.max_row + 1, 200)):
            a_val = ws.cell(row=r, column=1).value
            b_val = ws.cell(row=r, column=2).value
            if a_val and str(a_val).strip() == '序号':
                header_row = r
            if a_val and '合' in str(a_val):
                total_rows.append(r)
                total_row = r
            if header_row and r > header_row and b_val and str(b_val).strip() and '合' not in str(b_val):
                last_data_row = r

        if not header_row or not total_row:
            continue

        # 检查1：合计行唯一性（DT-35）
        if len(total_rows) > 1:
            issues.append(f"❌ [{sname}] 存在{len(total_rows)}个合计行: 行{total_rows}，汇总值可能翻倍!")

        # 检查2：数据行与合计行之间无空白无格式行（DT-25/DT-42/DT-44）
        if last_data_row and total_row > last_data_row + 2:
            gap = total_row - last_data_row - 1
            for r in range(last_data_row + 1, total_row):
                a_val = ws.cell(row=r, column=1).value
                if a_val is not None and str(a_val).strip():
                    if '减' not in str(a_val) and '小' not in str(a_val):
                        issues.append(f"❌ [{sname}] 数据行与合计行之间行{r}有残留内容: {a_val}")
            if gap > 3:
                issues.append(f"⚠️ [{sname}] 数据行(行{last_data_row})与合计行(行{total_row})之间有{gap}行空行，建议压缩")

        # 检查3：合计行SUM公式范围正确（DT-2）
        if total_row and last_data_row:
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=total_row, column=c)
                if not isinstance(cell, MergedCell) and isinstance(cell.value, str) and cell.value.startswith('='):
                    sum_match = re.search(r'SUM\(([A-Z])(\d+):([A-Z])(\d+)\)', cell.value)
                    if sum_match:
                        sum_start = int(sum_match.group(2))
                        sum_end = int(sum_match.group(4))
                        data_start = (header_row or 5) + 1
                        if sum_start > data_start or sum_end < last_data_row:
                            issues.append(f"❌ [{sname}] 合计行{get_column_letter(c)}{total_row}的SUM范围{sum_start}:{sum_end}未覆盖数据区{data_start}:{last_data_row}: {cell.value}")
                    refs = re.findall(r'[A-Z]+(\d+)', cell.value)
                    for row_ref in refs:
                        if int(row_ref) > total_row:
                            issues.append(f"❌ [{sname}] 合计行{get_column_letter(c)}{total_row}公式引用旧行号{row_ref}: {cell.value}")
                            break

        # 检查4：跨sheet引用行号正确（DT-24/DT-43/DT-61）
        for sum_sname in wb.sheetnames:
            if '汇总' not in sum_sname:
                continue
            ws_sum = wb[sum_sname]
            for row in ws_sum.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and sname.replace('-', '') in cell.value.replace('-', '').replace("'", "").replace("！", ""):
                        refs = re.findall(r'[A-Z]+(\d+)', cell.value)
                        for row_ref in refs:
                            if int(row_ref) > (total_row or 0) + 5:
                                issues.append(f"❌ [{sum_sname}]{cell.coordinate}引用{sname}行{row_ref}可能超出行范围: {cell.value}")

        # 检查5：打印区域精确匹配实际内容（DT-26/DT-36）
        if ws.print_area and last_data_row:
            pa = ws.print_area
            pa_match = re.search(r':([A-Z]+)(\d+)$', str(pa))
            if pa_match:
                pa_end_row = int(pa_match.group(2))
                if pa_end_row > (total_row or last_data_row) + 5:
                    issues.append(f"⚠️ [{sname}] 打印区域终止行{pa_end_row}远超合计行{total_row}，可能产生空白页")

        # 检查7：发生日期列不含文字/业务内容列不含日期序列号（DT-46）
        if header_row:
            col_meanings = {}
            for c in range(1, min(ws.max_column + 1, 15)):
                cell = ws.cell(row=header_row, column=c)
                if not isinstance(cell, MergedCell) and cell.value:
                    col_meanings[c] = str(cell.value).strip()

            for c, m in col_meanings.items():
                if '业务内容' in m or '内容' in m:
                    for r in range((header_row or 5) + 1, total_row or ws.max_row + 1):
                        cell = ws.cell(row=r, column=c)
                        if isinstance(cell.value, (int, float)) and cell.value > 40000:
                            issues.append(f"❌ [{sname}] {get_column_letter(c)}{r}业务内容列填入日期序列号{cell.value}，应为文字描述!")
                elif '发生日期' in m or '日期' in m:
                    for r in range((header_row or 5) + 1, total_row or ws.max_row + 1):
                        cell = ws.cell(row=r, column=c)
                        if isinstance(cell.value, str) and not cell.value.startswith('=') and len(cell.value) > 5 and not any(d in cell.value for d in '0123456789'):
                            issues.append(f"❌ [{sname}] {get_column_letter(c)}{r}发生日期列填入文字'{cell.value[:20]}'，应为日期!")

    # 汇总
    critical = [i for i in issues if i.startswith('❌')]
    warnings = [i for i in issues if i.startswith('⚠️')]

    # DT-114: 验证失败时自动触发修复闭环
    if (critical or warnings) and auto_fix and filepath:
        print(f"⚠️ 发现{len(critical)}个严重问题 + {len(warnings)}个警告，启动DT-114修复闭环...")
        from excel_row_ops import validate_and_fix
        vf = validate_and_fix(filepath, max_retries=max_retries)
        
        if vf['passed']:
            print(f"✅ DT-114修复闭环成功（修复{len(vf['fixes_applied'])}项）")
            return True
        else:
            print(f"🚨 DT-114修复闭环{vf['retries']}次后仍有问题:")
            for issue in vf['remaining_issues']:
                print(f"  ❌ {issue}")
            return False

    if critical:
        print(f"🚨 验证未通过！{len(critical)}个严重问题 + {len(warnings)}个警告")
        for i in critical: print(f"  {i}")
        for i in warnings: print(f"  {i}")
        return False
    elif warnings:
        print(f"⚠️ 验证通过但有{len(warnings)}个警告（建议修复）:")
        for i in warnings: print(f"  {i}")
        return True
    else:
        print("✅ 验证通过，所有检查项均合格")
        return True
```

**验证结果处理（DT-114闭环）**：
- 全部通过 → 进入Phase 4勾稽核对
- 有警告 → `auto_fix_formats`自动修复 → 重新验证 → 通过则进入Phase 4
- 有严重问题 → `validate_and_fix`自动修复闭环（最多3次）→ 仍FAIL则BLOCKED

## 输出

- 所有插入行的SUM范围已扩展覆盖
- 所有合计行引用行号已更新
- 所有跨sheet引用行号已更新（含隐藏汇总表）
- 行高全表统一 [DT-77]
- 数字格式全表校验修正 [DT-76]
- 结构行A列居中对齐 [DT-84]
- 边框完整性扫描通过 [DT-82]
- 合并单元格与模板一致 [DT-83]
- 空行已删除
- 打印范围已精确匹配
- 公式列未被覆写 [DT-67]

## 约束

> **v2.0规则编入步骤声明**：以下规则已编入对应操作步骤或由auto_fix_formats/validate_and_fix自动执行。约束区仅保留引用索引。

| 规则 | 编入步骤/自动工具 | 核心要点 |
|------|-----------------|---------|
| DT-2 插入行公式修复 | Step 4.1 → validate_and_fix | SUM范围+合计行引用+跨sheet引用 |
| DT-3 插入行复制参考行格式 | Step 4.1 → smart_insert_row | smart_insert_row自动copy格式 |
| DT-22 合计行A:C合并 | Step 4.6 | 统一合并A:C列+居中+thin边框 |
| DT-163 合计行B:C合并强制校验 | Step 4.6 | 检测跨行B:C合并扩展+重建缺失合并 |
| DT-24 删行公式修复 | Step 4.4 → smart_delete_rows | SUM收缩+跨sheet引用更新 |
| DT-25 空行直接删除 | Step 4.3 → smart_delete_rows | 禁止仅清边框 |
| DT-26 打印范围匹配 | Step 4.12 → smart_insert_row验证 | print_area精确覆盖内容 |
| DT-34 标题行合并单元格 | Step 4.7 | A1:H1/A2:H2恢复 |
| DT-35 合计行唯一性 | Step 4.8 → validate_detail_table | 每个明细表仅1个合计行 |
| DT-42 小计/减值行位置 | Step 4.9 | 紧贴合计行上方 |
| DT-43 跨sheet引用更新 | Step 4.10 → smart_insert_row | 含隐藏汇总表(DT-61) |
| DT-44 禁止空白无格式行 | Step 4.11 → smart_delete_rows | 数据行与合计行间无空白行 |
| DT-67 公式列禁止覆写 | Step 4.2.8 → auto_fix_formats检测 | J/K列MUST保留公式 |
| DT-76 增值额/增值率格式 | Step 4.2.2 → auto_fix_formats修复 | 特定数字格式强制 |
| DT-77 行高全表统一 | Step 4.2.1 → auto_fix_formats修复 | 同Sheet内数据行高一致 |
| DT-82 空白行格式补齐 | Step 4.2.4/3.2.7 | 边框+公式完整 |
| DT-83 合并单元格验证 | Step 4.2.5 | 对比模板merged_cells |
| DT-84 结构行A列居中 | Step 4.2.3 → auto_fix_formats修复 | 合计行/减值行A列居中 |
| DT-85 合计行公式引用 | Step 4.1 → validate_and_fix | 引用本行而非分隔行 |
| DT-86 跨sheet引用同步 | Step 4.10 → smart_insert_row | 汇总表+隐藏汇总表 |
| DT-112 格式集中处置 | Step 4.0~3.12 | 本Phase为格式集中处置区 |
| DT-114 验证-修复闭环 | Step 4.0/3.1/3.5 → validate_and_fix | 验证→修复→重验，最多3次 |
| DT-120 smart_insert_row强制 | Step 4.1/3.3 | 禁止裸insert_rows/delete_rows |
| DT-161 Phase 3（序时账查阅）前置Gate | Step 4.pre | 往来科目发生日期全空且有序时账→强制回退Phase 3（序时账查阅） |

**通用原则**：
- **后续复盘经验**：新增的格式修复教训直接写入Step 3.x操作段，不在约束区单独列示

## 异常处理

- 公式修复后验证不通过 → DT-114自动触发validate_and_fix修复闭环
- 格式扫描发现差异 → auto_fix_formats自动修复（G1F-1~G1F-4），重新扫描
- G1F-5合并单元格丢失 → 仅报告，需人工对照模板确认
- G1F-6公式列覆写 → 仅报告，需人工恢复公式
- 隐藏汇总表引用无法确认 → 对照原始模板确认引用目标（DT-62）
- validate_and_fix 3次重试仍FAIL → BLOCKED，输出详细诊断，禁止进入Phase 4
