# S3: 序时账查阅——发生日期确认与业务内容提取（原Phase 2e升级为独立Phase）

> **📋 Common规则适用声明**：本步骤适用 META_RULES MR-1/MR-4/MR-5/MR-6/MR-7/MR-9/MR-10 + preparation_discipline_rules G0-G13
> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2规则节+Phase 3规则节：DT-30(日期类型)、DT-46(日期/内容列序)、DT-51~55(发生日期核实6步)、DT-60(业务内容核实5步)、DT-149(业务内容自动映射)、DT-161(本步骤不可跳过)

## 🚨 Phase定位

> **核心转变**：从Phase 2的子步骤(2e)升级为独立Phase 3。原因：上海图灵项目复盘发现，2e作为2a-2e的末尾子步骤，在连续执行中被"顺带跳过"3次（河南平绿+上海图灵×2）。独立为Phase 3后：
> 1. 任务跟踪中为独立项，执行/跳过都有明确状态
> 2. G3门控前置：Phase 3未完成→Phase 4（格式修复）不得开始
> 3. 独立脚本journal_extractor.py封装完整流程，消除"手动从38656行序时账找日期"的跳过借口

## 🚨 Step 0: 脚本选择（DT-165强制，编写任何代码前MUST执行）

> **必读索引**：Read `scripts/SKILL_SCRIPT_INDEX.md`
> 本步骤MUST调用以下已有脚本：
> - **序时账查阅→`journal_extractor.py`**（本Phase新增脚本，MUST调用，禁止手写序时账解析逻辑）
> - 业务内容推断→`infer_business_content()`（business_content_map.py）
> - 列位映射→`sheet_col_map.json`（禁止硬编码列号）
> - 合计行定位→`find_header_structure()`（gate_validator.py）

### 🚨 MUST调用journal_extractor.py（DT-166新增红线）

**本Phase所有序时账查阅操作MUST通过`journal_extractor.py`脚本执行，禁止Agent自行编写序时账解析/匹配逻辑。**

| 允许 | 禁止 |
|------|------|
| `from journal_extractor import extract_dates, extract_business_contents` | 手写序时账行遍历+关键词匹配逻辑 |
| `from journal_extractor import JournalExtractor` | 手写datetime解析代码 |
| 调用extractor返回结果后写入评估明细表 | 手写GROUP BY结算对象逻辑 |

**根因**：S2_seq_verify.md中Step 2e.1~2e.4的代码是"示例代码"，需要Agent每次手动复制+调整列映射+调试关键词。三步门槛（列映射+关键词匹配+日期解析）都容易出错且费时，Agent倾向跳过。封装为脚本后，一步调用即可完成全部查阅。

## 输入

- 评估明细表（Phase 2已完成往来科目余额填写，但发生日期列和业务内容列为空或为通用文字）
- 序时账/明细账文件
- 科目余额表数据（_dt_cache/subjects.json，用于结算对象名称匹配）
- sheet_col_map.json（列位映射）
- 成本法底稿文件（可选，用于同步更新）

## 操作

### Step 3.0 执行/跳过判定与备份

> **🚨 本步骤是Phase 3的入口，必须执行。**跳过Phase 3的唯一方式是满足下方判定表中的跳过条件。

**执行/跳过判定（DT-161）**：

| 判定条件 | 动作 | 执行摘要标注 | 缓存写入(phase3_status.json) |
|---------|------|-------------|------------------------------|
| 项目有序时账数据 | **MUST执行Step 3.1~3.7** | "Phase 3: 已执行（DT-161）" | `{"executed": true, "skipped_reason": null}` |
| 使用人未提供序时账或相关替代材料 | 跳过本Phase全部后续Step | "Phase 3跳过: 未提供序时账(DT-143)" | `{"executed": false, "skipped_reason": "no_journal"}` |
| 使用人在提供项目路径时明确要求不填写发生日期 | 跳过本Phase全部后续Step | "Phase 3跳过: 使用人明确要求(DT-161②)" | `{"executed": false, "skipped_reason": "user_explicit"}` |

**判定执行代码**：

```python
import os, json

cache_dir = os.path.join(PROJECT_DIR, '_dt_cache')
manifest_path = os.path.join(cache_dir, 'file_manifest.json')

has_journal = False
if os.path.exists(manifest_path):
    with open(manifest_path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    for item in manifest.get('data', {}).get('files', []):
        fname = item.get('filename', '').lower()
        if '序时账' in fname or '明细账' in fname or 'journal' in fname:
            has_journal = True
            break

if not has_journal:
    for f in os.listdir(PROJECT_DIR):
        if '序时账' in f or '明细账' in f:
            has_journal = True
            break

if not has_journal:
    print("Phase 3跳过: 未提供序时账(DT-143)，发生日期列留空")
    with open(os.path.join(cache_dir, 'phase3_status.json'), 'w', encoding='utf-8') as f:
        json.dump({'executed': False, 'skipped_reason': 'no_journal'}, f, ensure_ascii=False)
else:
    print("Phase 3 MUST执行: 项目有序时账数据(DT-161)")
    with open(os.path.join(cache_dir, 'phase3_status.json'), 'w', encoding='utf-8') as f:
        json.dump({'executed': True, 'skipped_reason': None}, f, ensure_ascii=False)

    # 备份（DT-55）
    import shutil
    BACKUP_DIR = r'D:\workbuddy'
    os.makedirs(BACKUP_DIR, exist_ok=True)
    backup_name = os.path.basename(DETAIL_FILE).replace('.xlsx', '_pre_phase3_backup.xlsx')
    backup_path = os.path.join(BACKUP_DIR, backup_name)
    if not os.path.exists(backup_path):
        shutil.copy2(DETAIL_FILE, backup_path)
```

### Step 3.1 加载序时账数据（DT-51① + DT-54）

> **🚨 MUST调用journal_extractor.py，禁止手写序时账解析逻辑（DT-166）**

```python
import sys, os
sys.path.insert(0, os.path.join(os.path.expanduser('~'), '.codex', 'skills', 'valuation-detail-table', 'valuation-detail-table', 'scripts'))

from journal_extractor import JournalExtractor

# 初始化提取器（自动解析列映射+日期格式兼容DT-54）
extractor = JournalExtractor(SEQ_FILE)

# 验证列映射
print(f"序时账列映射: {extractor.col_map}")
print(f"序时账有效行数: {extractor.row_count}")
```

### Step 3.2 读取评估明细表往来科目待核实行

> 从评估明细表中定位所有往来科目Sheet的数据行，识别发生日期列和业务内容列为空的行。

```python
from journal_extractor import scan_empty_fields

# 扫描所有往来科目Sheet，返回待核实行列表
empty_rows = scan_empty_fields(DETAIL_FILE)

print(f"待核实发生日期: {sum(1 for r in empty_rows if r['date_empty'])}行")
print(f"待核实业务内容: {sum(1 for r in empty_rows if r['biz_empty'])}行")
print(f"业务内容为通用模板文字: {sum(1 for r in empty_rows if r['biz_generic'])}行")
```

### Step 3.3 核实发生日期（DT-51③~⑥ + DT-52 + DT-53）

> **🚨 MUST调用journal_extractor.extract_dates()，禁止手写关键词匹配逻辑（DT-166）**

```python
from journal_extractor import extract_dates

# 批量提取所有待核实行发生日期
# 内部执行：结算对象名称关键词提炼(DT-53) → 序时账匹配 → 方向筛选(资产借方/负债贷方) → 取距基准日最近的同方向凭证日期(DT-178)
date_results = extract_dates(extractor, empty_rows, subjects_path=os.path.join(CACHE_DIR, 'subjects.json'))

# 结果分类
verified = [r for r in date_results if r['status'] == 'verified']
no_match = [r for r in date_results if r['status'] == 'no_match']
ambiguous = [r for r in date_results if r['status'] == 'ambiguous']
generic_skip = [r for r in date_results if r['status'] == 'generic_skip']

print(f"已核实: {len(verified)}行")
print(f"未匹配: {len(no_match)}行 → [待核实]")
print(f"匹配歧义: {len(ambiguous)}行 → [待核实]")
print(f"泛匹配跳过(DT-52): {len(generic_skip)}行")
```

### Step 3.4 核实业务内容（DT-60）

> **🚨 MUST调用journal_extractor.extract_business_contents()，禁止手写摘要归纳逻辑（DT-166）**
> **DT-60与DT-51的区别**：DT-51核实的是发生日期（取距基准日最近的同方向凭证日期，DT-178），DT-60核实的是业务内容（归纳摘要文字），二者独立但检索逻辑前三步相同。
> **DT-149红线**：仅填科目名=无实质信息=底稿无效

```python
from journal_extractor import extract_business_contents

# 批量提取所有待核实行业务内容
# 内部执行：DT-60 5步流程（检索科目→检索结算对象→跨科目搜索→摘要归纳→映射）
biz_results = extract_business_contents(extractor, empty_rows, subjects_path=os.path.join(CACHE_DIR, 'subjects.json'))

# 结果分类
biz_updated = [r for r in biz_results if r['status'] == 'updated']
biz_inferred = [r for r in biz_results if r['status'] == 'inferred']  # 兜底推断
biz_no_match = [r for r in biz_results if r['status'] == 'no_match']

print(f"序时账摘要归纳: {len(biz_updated)}行")
print(f"兜底推断+标注[待核实]: {len(biz_inferred)}行")
print(f"未匹配: {len(biz_no_match)}行")
```

### Step 3.5 写入评估明细表

> **本步骤允许使用openpyxl直接修改单元格值（日期/业务内容），因为仅修改已有行的非金额字段，不涉及整sheet数据写入。**
> **禁止通过openpyxl写入金额数据——金额写入MUST通过fill_sheet()**

```python
import openpyxl
from datetime import datetime
from journal_extractor import write_phase3_results

# 统一写入发生日期+业务内容
write_phase3_results(DETAIL_FILE, date_results, biz_results)

# 保存
wb = openpyxl.load_workbook(DETAIL_FILE)
# write_phase3_results内部已处理，此处仅验证
wb.close()
```

### Step 3.6 同步更新成本法底稿（可选）

> 如项目有成本法底稿文件，发生日期和业务内容MUST同步更新。

```python
if COST_FILE and os.path.exists(COST_FILE):
    from journal_extractor import sync_to_cost_workpaper
    sync_to_cost_workpaper(COST_FILE, date_results, biz_results)
```

### Step 3.7 输出核实结果汇总

```python
from journal_extractor import generate_phase3_report

report = generate_phase3_report(date_results, biz_results)
print(report)
```

**核实结果汇总格式**：

| 类别 | 发生日期 | 业务内容 |
|------|---------|---------|
| **已更新** | 日期已用序时账距基准日最近同方向凭证日期更新(DT-178) | 业务内容已用序时账摘要归纳替换 |
| **兜底推断** | — | 序时账无匹配摘要，使用infer_business_content()推断并标注[待核实] |
| **跳过（泛匹配）** | DT-52泛匹配项，保留原值 | — |
| **跳过（未匹配）** | 序时账中未找到匹配关键词，保留原值 | 序时账中未找到匹配摘要 |
| **标注[待确认]** | — | 推断结果仍=科目名称，标注[待确认业务实质] |

### Step 3.8 验证更新结果（DT-55）

```python
import openpyxl
from datetime import datetime

wb_verify = openpyxl.load_workbook(DETAIL_FILE, data_only=True)
verify_ok = 0
verify_fail = 0

for r in [x for x in date_results if x['status'] == 'verified']:
    ws = wb_verify[r['sheet']]
    # 从sheet_col_map获取发生日期列号
    date_col = r.get('date_col', 5 if r['type'] == 'asset' else 4)
    cell_val = ws.cell(row=r['row'], column=date_col).value
    expected = r['verified_date']
    if isinstance(cell_val, datetime) and cell_val.date() == expected.date():
        verify_ok += 1
    else:
        verify_fail += 1
        print(f"  验证失败 | {r['sheet']} Row{r['row']} | 值={cell_val}")

print(f"\n验证结果: {verify_ok}通过, {verify_fail}失败")
wb_verify.close()
```

## 输出

- 往来科目发生日期已填入（来自序时账距基准日最近同方向凭证日期，DT-178）
- 往来科目业务内容已填入（来自序时账摘要归纳或infer_business_content()兜底推断）
- phase3_status.json缓存文件（供Phase 4 Step 4.pre读取）
- 核实结果汇总表

## 约束

> **v2.0规则编入步骤声明**：以下规则已编入对应操作步骤或由journal_extractor.py内部自动执行。约束区仅保留引用索引。

| 规则 | 编入步骤/脚本 | 核心要点 |
|------|-------------|---------|
| DT-51 | Step 3.3 → extract_dates() | 发生日期核实6步：列验证→解析→关键词匹配→方向筛选→写入→验证 |
| DT-52 | Step 3.3 → extract_dates() | 泛匹配项禁止自动核实 |
| DT-53 | Step 3.3 → extract_dates() | 关键词从摘要提炼而非使用全称 |
| DT-54 | Step 3.3 → JournalExtractor.__init__() | 序时账日期多格式兼容解析 |
| DT-55 | Step 3.0备份 + Step 3.8验证 | 修改前备份、修改后验证 |
| DT-60 | Step 3.4 → extract_business_contents() | 业务内容核实5步：检索科目→检索结算对象→跨科目搜索→摘要归纳→映射 |
| DT-149 | Step 3.4 → extract_business_contents() | 禁止仅填科目名称 |
| DT-161 | Step 3.0 | 有序时账时本Phase MUST执行，仅两种情况可跳过 |
| DT-166 | Step 0 + 全流程 | MUST调用journal_extractor.py，禁止手写序时账解析/匹配逻辑 |
| DT-30 | Step 3.5 | 发生日期必须为datetime类型 |
| DT-46 | Step 3.5 | 资产类D列=业务内容/E列=发生日期，负债类C列=发生日期/D列=业务内容 |
| DT-64 | Step 3.4 | 业务内容MUST从序时账摘要归纳，禁止凭结算对象名称猜测 |

## 异常处理

- 序时账列映射验证不通过 → JournalExtractor.__init__()自动适配，失败则停止报告
- 匹配失败（0条结果）→ 标注[待核实]+保留原始日期+继续执行（DT-151）
- 匹配歧义（>5条）→ 标注[待核实]+保留原始日期+继续执行（DT-151）
- 泛匹配项 → 跳过自动核实，保留原值，标注原因
- COM resave后datetime变int → 二次修复（在Phase 4格式修复中统一处理DT-33）
