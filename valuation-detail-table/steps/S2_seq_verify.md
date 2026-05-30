# S2_seq_verify: 序时账核实往来科目（DT-51~55/DT-60完整流程）

> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 2规则节
> **📋 特殊说明**：序时账核实仅修改已有行的日期/业务内容单元格，不涉及整sheet数据写入，因此不使用fill_sheet()管线，但仍须遵守DT-128工具选型纪律

## 🚨 Step 0: 脚本选择（DT-165强制，编写任何代码前MUST执行）

> **必读索引**：Read `scripts/SKILL_SCRIPT_INDEX.md`
> 本步骤涉及的关键已有脚本：
> - 序时账数据加载→`load_journal_data()`（data_loader.py）
> - 业务内容推断→`infer_business_content()`（business_content_map.py）
> - 列位映射→`sheet_col_map.json`（禁止硬编码列号）
> - 合计行定位→`find_header_structure()`（gate_validator.py）

## 🚨 Phase 2e在整体流程中的位置（DT-161必经步骤）

> **Phase 2e不是"可选核实"，而是Phase 2填写完成后的必经下一步骤。**
> 以下流程图明确Phase 2e的位置和执行条件：

```
Phase 2（填写往来科目）
  │
  ├── S2_fill_re / S2_fill_bs / S2_fill_liability
  │     │
  │     └── Step 4: 🚨 MUST进入Phase 2e（DT-161）
  │
  ▼
Phase 2e（序时账核实）← 本步骤文件
  │
  ├── Step 2e.0: 执行/跳过判定
  │     │
  │     ├── 有序时账 → MUST执行Step 2e.0b~2e.9（三层保障的L1入口）
  │     ├── 无序时账 → 跳过(DT-143)，发生日期留空
  │     └── 使用人明确要求 → 跳过(DT-161②)
  │
  ├── Step 2e.1~2e.4: 核实发生日期
  ├── Step 2e.5: 核实业务内容(DT-60)
  ├── Step 2e.6~2e.7: 写入评估明细表+成本法底稿
  ├── Step 2e.8~2e.9: COM重算+验证
  │
  ▼
Phase 3（格式处置）— gate_G2(G2-18)检查发生日期/业务内容非空（三层保障的L2门控）
  │
  ▼
Phase 5（交付）— Step 5.6a兜底验证（三层保障的L3拦截）
```

**三层物理保障**：

| 层级 | 物理载体 | 阻断时机 | 阻断效果 |
|------|---------|---------|---------|
| L1 脚本层 | `prepare_data_rows(has_journal=True)` → fill_sheet()内部date列非空校验 | 数据写入时 | 写入失败，Agent必须处理 |
| L2 Gate层 | gate_G2 G2-18检查发生日期+业务内容列非空 | Phase 2→Phase 3边界 | CRITICAL=禁止进入Phase 3 |
| L3 交付层 | S5_deliver Step 5.6a交付前兜底验证 | 交付前 | RuntimeError=禁止交付 |

## 🚨 工具选型纪律（DT-128+DT-160）

**本步骤特殊说明**：
- 序时账核实仅修改已有单元格的值（日期/业务内容），不涉及整sheet数据写入
- 因此**本步骤允许**使用openpyxl直接修改单元格值（`ws.cell(row=r, column=c).value = xxx`）
- 但仍然**禁止**使用xlsx skill操作评估明细表
- **禁止**通过openpyxl写入金额数据（账面价值/评估价值等）——金额写入MUST通过fill_sheet()

## 输入

- 评估明细表（已填写往来科目数据）
- 序时账/明细账文件
- 成本法底稿文件（用于同步更新）
- 往来科目结算对象清单

## 操作

> **⚠️ 必须严格遵守DT-51~55/DT-60纪律规则，违反任何一条即停止执行并报告。**

### Step 2e.0 执行判定与备份

> **🚨 本步骤是Phase 2e的入口，必须执行。**跳过Phase 2e的唯一方式是满足下方判定表中的跳过条件。DT-143（无序时账→日期留空）是跳过后的兜底规则，不是跳过的理由。

**Step 2e.0a 执行/跳过判定（DT-161）**：

| 判定条件 | 动作 | 执行摘要标注 | 缓存写入(phase2e_status.json) |
|---------|------|-------------|------------------------------|
| 项目有序时账数据 | **MUST执行Step 2e.0b~2e.9** | "Phase 2e: 已执行（DT-161）" | `{"executed": true, "skipped_reason": null}` |
| 使用人未提供序时账或相关替代材料 | 跳过本步骤全部后续Step | "Phase 2e跳过: 未提供序时账(DT-143)" | `{"executed": false, "skipped_reason": "no_journal"}` |
| 使用人在提供项目路径时明确要求不填写发生日期 | 跳过本步骤全部后续Step | "Phase 2e跳过: 使用人明确要求(DT-161②)" | `{"executed": false, "skipped_reason": "user_explicit"}` |

**判定执行代码**：

```python
import os, json

# 读取Phase -1缓存的项目材料清单
cache_dir = os.path.join(PROJECT_DIR, '_dt_cache')
manifest_path = os.path.join(cache_dir, 'file_manifest.json')

has_journal = False
if os.path.exists(manifest_path):
    with open(manifest_path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    # 检查材料清单中是否有序时账
    for item in manifest.get('files', []):
        fname = item.get('name', '').lower()
        if '序时账' in fname or '明细账' in fname or 'journal' in fname:
            has_journal = True
            break

# 如果缓存中无记录，直接检查项目目录
if not has_journal:
    for f in os.listdir(PROJECT_DIR):
        if '序时账' in f or '明细账' in f:
            has_journal = True
            break

if not has_journal:
    print("⏭️ Phase 2e跳过: 未提供序时账(DT-143)，发生日期列留空")
    print("   → 执行摘要标注: 'Phase 2e跳过: 未提供序时账(DT-143)'")
    # 写入缓存标注，供Phase 3 Step 3.pre读取
    with open(os.path.join(cache_dir, 'phase2e_status.json'), 'w', encoding='utf-8') as f:
        json.dump({'executed': False, 'skipped_reason': 'no_journal'}, f, ensure_ascii=False)
    # 退出本Step，不执行2e.0b~2e.9
    import sys; sys.exit(0)
else:
    print("✅ Phase 2e MUST执行: 项目有序时账数据(DT-161)")
    # 写入缓存标注
    with open(os.path.join(cache_dir, 'phase2e_status.json'), 'w', encoding='utf-8') as f:
        json.dump({'executed': True, 'skipped_reason': None}, f, ensure_ascii=False)
```

> **❌ 错误认知校准**（以下想法均为错误，遇到时必须纠正）：
> - "S2_fill_re.md中说序时账是可选的" → `has_journal=True`是告知fill_sheet()项目有序时账，不是触发核实的指令
> - "DT-143说无序时账时日期留空，所以跳过核实也行" → DT-143是跳过后的兜底，不是跳过的理由
> - "发生日期是锦上添花" → 发生日期是账龄分析依据，缺失=底稿不可审

**Step 2e.0b 备份（DT-55）**：

```python
import shutil
BACKUP_DIR = r'D:\workbuddy'
os.makedirs(BACKUP_DIR, exist_ok=True)
for target_file in [DETAIL_FILE, COST_FILE]:
    backup_name = os.path.basename(target_file).replace('.xlsx', '_backup.xlsx')
    backup_path = os.path.join(BACKUP_DIR, backup_name)
    if not os.path.exists(backup_path):
        shutil.copy2(target_file, backup_path)
```

### Step 2e.1 验证序时账列映射（DT-51①）

**🚨 禁止假设列号！不同导出格式列号可能不同。**

```python
wb_seq = openpyxl.load_workbook(SEQ_FILE, data_only=True)
ws_seq = wb_seq[wb_seq.sheetnames[0]]
for r in range(1, 6):
    row_data = {}
    for c in range(1, ws_seq.max_column + 1):
        val = ws_seq.cell(row=r, column=c).value
        if val is not None:
            row_data[c] = str(val)[:50]
    print(f"Row {r}: {row_data}")
```

**标准列映射（绿城熵里项目验证结果）**：

| 列号 | 列名 | 内容示例 | 说明 |
|------|------|---------|------|
| B(2) | 日期 | 2025-01-05 / datetime | 可能为字符串/datetime/数字（DT-54） |
| D(4) | 摘要 | "萍乡春风江南项目工抵房" | 关键词提炼源（DT-53） |
| F(6) | 科目名称 | "1122\\应收账款\\临沂浩然" | 含编码路径，需关键词匹配 |
| J(10) | 借方本币 | 128441444.68 | 资产类取此列末笔（DT-51④） |
| M(13) | 贷方本币 | 5000000.00 | 负债类取此列末笔（DT-51④） |

**验证不通过时**：如果列号不同，调整映射后继续。如果找不到关键列，停止并报告。

### Step 2e.2 读取并解析序时账（DT-51② + DT-54）

```python
from datetime import datetime, timedelta

seq_data = []
for r in range(5, ws_seq.max_row + 1):
    date_val = ws_seq.cell(row=r, column=2).value
    summary = ws_seq.cell(row=r, column=4).value or ''
    subject_name = ws_seq.cell(row=r, column=6).value or ''
    debit = ws_seq.cell(row=r, column=10).value
    credit = ws_seq.cell(row=r, column=13).value

    # DT-54: 日期多格式兼容解析
    dt = None
    if isinstance(date_val, datetime):
        dt = date_val
    elif isinstance(date_val, str):
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d']:
            try:
                dt = datetime.strptime(date_val.strip(), fmt)
                break
            except ValueError:
                continue
    elif isinstance(date_val, (int, float)):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=int(date_val))
        except:
            pass

    if dt and subject_name:
        seq_data.append({
            'row': r, 'date': dt, 'summary': str(summary),
            'subject_name': str(subject_name).strip(),
            'debit': float(debit) if debit else 0,
            'credit': float(credit) if credit else 0,
        })
```

### Step 2e.3 核实发生日期（DT-51③ + DT-52 + DT-53）

**每条核实规则格式**：`(sheet, row, full_name, subject_filter, keywords, type, date_column, notes)`

```python
verification_rules = [
    # === 资产类科目 (D列=发生日期, 取借方末笔) ===
    ('3-5应收账款', 6, '临沂浩然房地产开发有限公司', '应收账款', ['临沂浩然'], 'asset', 4, ''),
    ('3-7预付款项', 6, '萍乡绿盛置业有限公司', '预付', ['萍乡'], 'asset', 4,
     'DT-53: 序时账摘要为"萍乡春风江南项目工抵房"，不含"绿盛"'),
    ('3-7预付款项', 7, '河南晟通地产有限公司', '预付', ['河南', '凤湖'], 'asset', 4,
     'DT-53: 多关键词OR逻辑，摘要为"河南凤湖玫瑰园项目工抵房"'),

    # === 负债类科目 (C列=发生日期, 取贷方末笔) ===
    ('5-10-3其他应付款', 6, '江西国际汽车城投资发展有限公司', '其他应付', ['汽车城'], 'liability', 3, ''),

    # === 泛匹配项（DT-52: 禁止自动核实） ===
    ('3-8-3其他应收款', 8, '其他', '其他应收', None, 'asset', 4,
     '⚠️ DT-52泛匹配项，跳过自动核实'),
    ('5-10-3其他应付款', 9, '个人(押金)', '其他应付', None, 'liability', 3,
     '⚠️ DT-52泛匹配项，序时账中未找到"押金"关键词'),
]
```

**关键词提炼示例（DT-53）**：

| 评估明细表全称 | 关键词 | 摘要实际内容 | 提炼逻辑 |
|---------------|--------|-------------|---------|
| 临沂浩然房地产开发有限公司 | ['临沂浩然'] | "临沂浩然管理服务费" | 公司简称匹配 |
| 萍乡绿盛置业有限公司 | ['萍乡'] | "萍乡春风江南项目工抵房" | 地理关键词，摘要不含"绿盛" |
| 河南晟通地产有限公司 | ['河南', '凤湖'] | "河南凤湖玫瑰园项目工抵房" | 多关键词OR逻辑 |
| 代扣个人住房公积金 | ['公积金'] | "缴纳公积金" | 业务关键词 |

**⚠️ 合同资产特殊规则**（DT-51②补充）：合同资产实际由应收账款重分类调整而来，序时账中**没有**合同资产科目，核实合同资产发生日期时**必须匹配序时账中应收账款科目**而非合同资产科目。

### Step 2e.4 执行核实

```python
for sheet, row, full_name, subject_filter, keywords, typ, date_col, notes in verification_rules:
    # DT-52: 泛匹配项跳过自动核实
    if keywords is None:
        print(f"  ⚠️ 跳过泛匹配项 | {sheet} Row{row} | {full_name} | {notes}")
        continue

    # DT-51②: 按科目名称关键词筛选
    filtered = [s for s in seq_data if subject_filter in s['subject_name']]

    # DT-51③ + DT-53: 按提炼的关键词匹配摘要
    matched = []
    for s in filtered:
        for kw in keywords:
            if kw in s['summary']:
                matched.append(s)
                break

    # DT-51④: 方向规则
    if typ == 'asset':
        target = [m for m in matched if m['debit'] > 0]   # 资产类取借方
    else:
        target = [m for m in matched if m['credit'] > 0]   # 负债类取贷方

    # DT-56+DT-151: 匹配失败时标注[待核实]并继续
    if not matched:
        print(f"  🚫 序时账中未找到匹配 | {sheet} Row{row} | {full_name}")
        print(f"     → [待核实] 保留原始日期（DT-151）")
        continue
    if len(matched) > 5:
        print(f"  ⚠️ 匹配结果存在歧义({len(matched)}条) | {sheet} Row{row} | {full_name}")
        print(f"     → [待核实] 匹配歧义，跳过自动核实（DT-151）")
        continue

    # 取末笔（最新日期）
    if target:
        target.sort(key=lambda x: x['date'])
        last = target[-1]
        verified_date = last['date']   # datetime对象（DT-30）
```

### Step 2e.5 核实业务内容（DT-60）

> **DT-60与DT-51的区别**：DT-51核实的是**发生日期**（取末笔日期），DT-60核实的是**业务内容**（归纳摘要文字），二者是独立的核实场景，但检索逻辑前三步相同。
>
> **❌ 错误认知校准**：
> - "业务内容Phase 2已填，Phase 2e不需要再处理" → Phase 2的`infer_business_content()`仅基于科目编码+结算对象名称推断，无法访问序时账摘要，推断结果是"销售商品/提供服务""其他应收款"等通用文字；Phase 2e必须用序时账摘要归纳出真正的业务实质
> - "业务内容=科目名称即可" → DT-149红线：仅填科目名=无实质信息=底稿无效
> - "业务内容核实是可选项" → DT-60是R级红线，与DT-51同级，发生日期和业务内容MUST同时核实

**Step 2e.5a 读取评估明细表往来科目已有业务内容（定位待核实行）**：

```python
import openpyxl, re
from collections import defaultdict

wb_detail = openpyxl.load_workbook(DETAIL_FILE, data_only=True)

# 定义往来科目sheet及其业务内容列位（DT-46: 资产类D列=业务内容, 负债类E列=业务内容）
re_sheets = {
    # (sheet名关键词, 科目类型, 业务内容列号)
    # 资产类：C=结算对象, D=业务内容, E=发生日期
    '3-5':  {'type': 'asset',  'biz_col': 4, 'name_col': 3, 'date_col': 5, 'subject_code': '1122'},
    '3-7':  {'type': 'asset',  'biz_col': 4, 'name_col': 3, 'date_col': 5, 'subject_code': '1123'},
    '3-8-3':{'type': 'asset',  'biz_col': 4, 'name_col': 3, 'date_col': 5, 'subject_code': '1221'},
    '3-10': {'type': 'asset',  'biz_col': 4, 'name_col': 3, 'date_col': 5, 'subject_code': '1461'},
    # 负债类：C=结算对象, D=发生日期, E=业务内容
    '5-5':  {'type': 'liability', 'biz_col': 5, 'name_col': 3, 'date_col': 4, 'subject_code': '2202'},
    '5-6':  {'type': 'liability', 'biz_col': 5, 'name_col': 3, 'date_col': 4, 'subject_code': '2203'},
    '5-10-3':{'type':'liability', 'biz_col': 5, 'name_col': 3, 'date_col': 4, 'subject_code': '2241'},
}

# 通用模板文字黑名单（DT-149: 这些=仅填科目名=底稿无效）
GENERIC_BIZ_CONTENTS = {
    '其他应收款', '其他应付款', '其他往来', '往来款',
    '销售商品/提供服务', '采购商品/接受服务', '预付货款/服务费', '预收货款/服务费',
    '货款', '预收账款', '应付货款',
}

biz_verify_rules = []  # 待核实的业务内容行

for sname in wb_detail.sheetnames:
    for key, cfg in re_sheets.items():
        if key in sname:
            ws = wb_detail[sname]
            # 找数据起始行（A列标记"检索表头"下一行）
            data_start = None
            for r in range(1, ws.max_row + 1):
                a_val = ws.cell(row=r, column=1).value
                if a_val and '检索表头' in str(a_val):
                    data_start = r + 1
                    break
            if not data_start:
                continue

            for r in range(data_start, ws.max_row + 1):
                a_val = ws.cell(row=r, column=1).value
                if a_val and ('合计' in str(a_val) or '坏账' in str(a_val) or '预计' in str(a_val)):
                    break  # 到合计行，停止

                name_val = ws.cell(row=r, column=cfg['name_col']).value
                biz_val = ws.cell(row=r, column=cfg['biz_col']).value
                i_val = ws.cell(row=r, column=9 if cfg['type'] == 'liability' else 10).value  # 账面价值

                if not name_val or not i_val:
                    continue  # 空行或无金额行

                biz_str = str(biz_val).strip() if biz_val else ''
                # DT-149判定：业务内容为空、为通用模板文字、或=科目名称 → 需核实
                needs_verify = (
                    not biz_str
                    or biz_str in GENERIC_BIZ_CONTENTS
                    or biz_str == key.split('-', 1)[-1].replace('-', '')  # 去掉编号后=科目名
                )
                if needs_verify:
                    biz_verify_rules.append({
                        'sheet': sname,
                        'row': r,
                        'name': str(name_val).strip(),
                        'current_biz': biz_str,
                        'type': cfg['type'],
                        'biz_col': cfg['biz_col'],
                        'subject_code': cfg['subject_code'],
                    })

print(f"📋 待核实业务内容: {len(biz_verify_rules)}行")
for item in biz_verify_rules[:20]:  # 显示前20行
    print(f"  {item['sheet']} Row{item['row']}: 结算对象={item['name'][:20]} | 当前业务内容={item['current_biz'] or '(空)'}")
```

**Step 2e.5b DT-60 5步执行——从序时账摘要归纳业务内容**：

```python
from business_content_map import infer_business_content

biz_updates = []  # 需要更新的业务内容列表

for item in biz_verify_rules:
    name = item['name']
    subject_code = item['subject_code']

    # === DT-60 Step1: 先检索科目（与Step 2e.4共享seq_data） ===
    # 按结算对象名称中的关键词在序时账科目名称中筛选
    name_keywords = []
    # DT-53: 从结算对象名称提炼关键词
    # 地理关键词优先（地名+2~3字）
    geo_match = re.search(r'([\u4e00-\u9fff]{2,4}(?:省|市|区|县|镇|路|街))', name)
    if geo_match:
        name_keywords.append(geo_match.group(1)[:3])  # 取地名前3字
    # 公司简称（去掉"有限公司"等后缀，取核心部分）
    core_name = re.sub(r'(有限公司|股份有限公司|有限责任公司|公司|集团)', '', name)
    if len(core_name) >= 2:
        name_keywords.append(core_name[:4])  # 取核心名前4字

    # 在seq_data中搜索该结算对象的序时账记录
    matched_seqs = []
    for s in seq_data:
        for kw in name_keywords:
            if kw and kw in s['summary']:
                matched_seqs.append(s)
                break
            if kw and kw in s['subject_name']:
                matched_seqs.append(s)
                break

    # === DT-60 Step2: 再检索具体结算对象 ===
    # 已在Step1中通过关键词匹配摘要完成

    # === DT-60 Step3: 跨科目搜索验证业务实质 ===
    # 收集该结算对象在所有科目下的摘要（归纳全部业务关系）
    all_summaries = []
    for s in matched_seqs:
        if s['summary']:
            all_summaries.append(s['summary'])

    # === DT-60 Step4: 摘要归纳→业务内容映射 ===
    inferred_biz = None

    if all_summaries:
        # 优先从摘要归纳
        # 策略1: 高频关键词提取（出现频率最高的业务词汇）
        biz_keywords = {
            '货款': ['货款', '收货款', '发货', '出货', '采购', '进货'],
            '服务费': ['服务费', '技术服务', '咨询费', '管理费', '开发费'],
            '工程款': ['工程款', '工程', '施工', '建设', '安装', '装修'],
            '租金': ['租金', '租赁', '房租', '场地费'],
            '保证金': ['保证金', '押金', '投标保证金'],
            '报销款': ['报销', '差旅', '办公费', '通讯费'],
            '社保': ['社保', '公积金', '五险一金', '养老', '医疗'],
            '税费': ['增值税', '所得税', '税', '附加'],
            '借款': ['借款', '贷款', '融资', '利息'],
            '往来款': ['往来', '划款', '调拨', '内部'],
        }
        # 统计每个业务关键词的匹配次数
        keyword_counts = defaultdict(int)
        for summary in all_summaries:
            for biz_type, patterns in biz_keywords.items():
                for p in patterns:
                    if p in summary:
                        keyword_counts[biz_type] += 1
                        break  # 每条摘要每种业务类型只计1次

        if keyword_counts:
            # 取最高频的业务类型
            best_biz = max(keyword_counts, key=keyword_counts.get)
            inferred_biz = best_biz
        else:
            # 摘要中无明确业务关键词→取摘要前6字
            shortest_summary = min(all_summaries, key=len)
            inferred_biz = shortest_summary[:6]
    else:
        # 序时账中未找到匹配→使用infer_business_content()兜底推断
        inferred_biz = infer_business_content(subject_code, name)
        inferred_biz = f"{inferred_biz}[待核实]"

    # === DT-149: 禁止仅填科目名称 ===
    # 如果推断结果仍=科目名称→标注[待确认]
    subject_names = {'其他应收款', '其他应付款', '应收账款', '应付账款', '预付款项', '预收款项'}
    if inferred_biz in subject_names:
        inferred_biz = f"{inferred_biz}[待确认业务实质]"

    item['new_biz'] = inferred_biz
    item['source'] = 'seq_summary' if all_summaries else 'infer_fallback'
    biz_updates.append(item)

    print(f"  {'✅' if all_summaries else '⚠️'} {item['sheet']} Row{item['row']}: "
          f"'{item['current_biz'] or '(空)'}' → '{inferred_biz}' "
          f"(来源={'序时账摘要' if all_summaries else '自动推断'})")

# === DT-60 Step5: 同步更新评估明细表 ===
if biz_updates:
    wb_biz = openpyxl.load_workbook(DETAIL_FILE)
    for item in biz_updates:
        ws = wb_biz[item['sheet']]
        ws.cell(row=item['row'], column=item['biz_col']).value = item['new_biz']
    wb_biz.save(DETAIL_FILE)
    print(f"\\n✅ 已更新{len(biz_updates)}行业务内容到评估明细表")
else:
    print("\\n✅ 无需更新业务内容")
```

**约束**：

- **DT-60**：序时账核实业务内容完整5步流程——**MUST执行Step 2e.5a+2e.5b**，禁止跳过
- **DT-149**：业务内容自动映射——禁止仅填科目名称（"其他应收款"=等于没填）
- **DT-46**：业务内容与发生日期严禁混淆——业务内容列=文字，发生日期列=日期
- **DT-53**：关键词从摘要提炼而非使用全称——与Step 2e.4共用关键词提炼逻辑
- **DT-64**：上下文关联禁止——业务内容MUST从序时账摘要归纳，禁止凭结算对象名称猜测

### Step 2e.6 更新评估明细表——发生日期+业务内容（DT-30 + DT-46 + DT-55 + DT-60）

> Phase 2e修改评估明细表有两种场景：①核实发生日期（Step 2e.4输出）②核实业务内容（Step 2e.5输出）。两者都在本Step统一写入。

```python
wb_edit = openpyxl.load_workbook(DETAIL_FILE)

# === 场景1：更新发生日期（Step 2e.4输出） ===
for r in need_updates:
    sheet = r['sheet']
    row = r['row']
    verified_date = r['verified_date_obj']  # 必须是datetime对象
    date_col = 4 if r['type'] == 'asset' else 3  # DT-46: 资产类E列(5), 负债类D列(4)
    # 注意：资产类有账龄列，发生日期在E列(5)而非D列(4)
    # 实际列位由sheet_col_map.json确定，此处仅示例
    ws = wb_edit[sheet]
    ws.cell(row=row, column=date_col).value = verified_date
    ws.cell(row=row, column=date_col).number_format = 'yyyy"年"m"月"'

# === 场景2：更新业务内容（Step 2e.5输出，已在Step 2e.5b中写入） ===
# biz_updates已在Step 2e.5b中直接写入，此处无需重复写入
# 如需回验，取消下方注释：
# for item in biz_updates:
#     ws = wb_edit[item['sheet']]
#     actual = ws.cell(row=item['row'], column=item['biz_col']).value
#     print(f"  回验 {item['sheet']} Row{item['row']}: 业务内容={actual}")

wb_edit.save(DETAIL_FILE)
```

### Step 2e.7 同步更新成本法底稿——发生日期+业务内容（DT-51⑥ + DT-60⑤）

评估明细表日期和业务内容更新后，**必须**同步到成本法底稿对应过程表。

```python
wb_cost = openpyxl.load_workbook(COST_FILE)

# === 同步发生日期 ===
for r in need_updates:
    cost_sheet_name = r['sheet']
    if cost_sheet_name in wb_cost.sheetnames:
        ws_cost = wb_cost[cost_sheet_name]
        date_col = 4 if r['type'] == 'asset' else 3
        ws_cost.cell(row=r['row'], column=date_col).value = r['verified_date_obj']
        ws_cost.cell(row=r['row'], column=date_col).number_format = 'yyyy"年"m"月"'

# === 同步业务内容 ===
for item in biz_updates:
    cost_sheet_name = item['sheet']
    if cost_sheet_name in wb_cost.sheetnames:
        ws_cost = wb_cost[cost_sheet_name]
        # 成本法底稿业务内容列位与评估明细表相同
        ws_cost.cell(row=item['row'], column=item['biz_col']).value = item['new_biz']

wb_cost.save(COST_FILE)
```

**注意**：合同资产在成本法底稿中无独立过程表，其底稿数据在应收账款过程表中，故合同资产日期和业务内容更新无需同步成本法底稿。

### Step 2e.8 COM重算保存 + 二次修复datetime（DT-17/DT-29/DT-33）

```python
import win32com.client, time
app = win32com.client.DispatchEx('Excel.Application')  # DT-29: 必须用DispatchEx
app.Visible = False
app.DisplayAlerts = False
try:
    for filepath in [DETAIL_FILE, COST_FILE]:
        wb_com = app.Workbooks.Open(filepath)
        app.Calculate()
        time.sleep(2)
        wb_com.Save()
        wb_com.Close(SaveChanges=False)
finally:
    app.Quit()

# DT-33: COM resave后datetime变int，需openpyxl二次修复
for filepath in [DETAIL_FILE, COST_FILE]:
    wb_fix = openpyxl.load_workbook(filepath)
    fix_count = 0
    for r in need_updates:
        date_col = 4 if r['type'] == 'asset' else 3
        ws = wb_fix[r['sheet']]
        cell = ws.cell(row=r['row'], column=date_col)
        if isinstance(cell.value, (int, float)):
            cell.value = r['verified_date_obj']
            cell.number_format = 'yyyy"年"m"月"'
            fix_count += 1
    if fix_count > 0:
        wb_fix.save(filepath)
```

### Step 2e.9 验证更新结果（DT-55）

```python
for filepath in [DETAIL_FILE, COST_FILE]:
    wb_verify = openpyxl.load_workbook(filepath, data_only=True)
    for r in need_updates:
        ws = wb_verify[r['sheet']]
        date_col = 4 if r['type'] == 'asset' else 3
        cell_val = ws.cell(row=r['row'], column=date_col).value
        expected = r['verified_date_obj']
        if isinstance(cell_val, datetime) and cell_val.date() == expected.date():
            print(f"  ✅ 验证通过 | {r['sheet']} Row{r['row']}")
        else:
            print(f"  ❌ 验证失败 | {r['sheet']} Row{r['row']} | 值={cell_val}")
```

## 输出

核实结果汇总：

| 类别 | 发生日期 | 业务内容 |
|------|---------|---------|
| **已更新** | 日期不一致，已用序时账核实结果更新（含原值→新值） | 业务内容为通用模板文字/科目名，已用序时账摘要归纳替换 |
| **一致** | 当前日期与序时账核实结果一致，无需更新 | 业务内容已充分，无需更新 |
| **跳过（泛匹配）** | DT-52泛匹配项，无法精确核实，保留原值 | — |
| **跳过（未匹配）** | 序时账中未找到匹配关键词，保留原值 | 序时账中未找到匹配摘要，使用infer_business_content()兜底推断并标注[待核实] |
| **标注[待确认]** | — | 推断结果仍=科目名称，标注[待确认业务实质] |

## 约束

- **DT-51**：序时账核实往来科目发生日期完整流程6步
- **DT-52**：泛匹配项禁止自动核实发生日期
- **DT-53**：关键词必须从摘要提炼而非使用全称
- **DT-54**：序时账日期解析多格式兼容
- **DT-55**：修改前备份、修改后验证
- **DT-60**：🚨 序时账核实业务内容完整5步流程——Step 2e.5a+2e.5b MUST执行，禁止跳过。仅填科目名称=违反DT-149=底稿无效
- **DT-149**：🚨 业务内容自动映射——Phase 2填写时由infer_business_content()推断，Phase 2e核实时由序时账摘要归纳替换。禁止仅填科目名称
- **DT-30**：发生日期必须为datetime类型
- **DT-46**：业务内容与发生日期严禁混淆（资产D列/负债C列=日期）
- **DT-64**：上下文关联禁止——业务内容MUST从序时账摘要归纳，禁止凭结算对象名称猜测
- **DT-33**：COM resave后datetime二次修复
- **DT-32**：序时账科目编码与科目余额表编码体系不同
- **DT-161**：🚨 有序时账时本步骤MUST执行，仅两种情况可跳过

## 异常处理

- 序时账列映射验证不通过 → 调整映射后继续，找不到关键列则停止报告
- 匹配失败（0条结果）→ 标注[待核实]+保留原始日期+继续执行（DT-151）
- 匹配歧义（>5条）→ 标注[待核实]+保留原始日期+继续执行（DT-151）
- 泛匹配项 → 跳过自动核实，保留原值，标注原因
- COM resave后datetime变int → 二次修复（不再做COM resave）
