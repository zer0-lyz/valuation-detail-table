# S6: 清理与交付

> **📋 Common规则适用声明**：本步骤适用 META_RULES MR-1/MR-4/MR-5/MR-6/MR-7/MR-9/MR-10 + preparation_discipline_rules G0-G13 + G7/G8(反思+透明度)
> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 5规则节：DT-9(文件命名)、DT-17(COM重算)、DT-20(无数据Sheet隐藏)、DT-23(系统表隐藏)、DT-48(执行透明度)、DT-81(页脚一次性输出)、DT-96(公司名称统一)、DT-110(隐藏集中执行)、DT-123(汇总sheet递归可见性)、DT-151(异常不中断) + Phase 5.5规则节：DT-59(反思固化)


## 输入

- 勾稽通过+反思固化完成的评估明细表

## 操作

> **🚨 DT-182b 交付前校验**：交付前 MUST 执行 `fix_intra_sheet_total2_formulas(wb)` + `fix_all_summary_refs_batch(wb)`，确保插行导致的合计行引用偏移已全部修复。

### Step 6.1 集中隐藏操作（DT-110，交付前最后一道门控）

> **🚨 核心原则**：所有sheet隐藏操作MUST在此步骤集中执行，禁止在Phase 2-4期间分散执行。Phase 2-4期间需要读取hidden sheet的公式和数据进行勾稽验证，提前隐藏=公式引用更新遗漏=验证不完整。

**执行顺序（严格按1→2→3→4→5→6执行，不可跳步）：**

#### 6.1.1 确认设定信息sheet已填写（DT-79/DT-121确认步骤）

设定信息sheet**不属于"直接跳过隐藏"**——MUST在隐藏前**确认**关键字段已填写（Phase 0 Step 0.7已执行首次填写）：

**确认逻辑**：
```python
import openpyxl

def confirm_settings_filled(detail_path):
    """确认设定信息sheet的B6/B7已在Phase 0填写"""
    wb = openpyxl.load_workbook(detail_path, data_only=True)
    ws = wb['设定信息']
    b6 = ws['B6'].value  # 被评估单位全称
    b7 = ws['B7'].value  # 评估基准日
    wb.close()

    b6_filled = b6 is not None and str(b6).strip() != ''
    b7_filled = b7 is not None and str(b7).strip() != ''

    if not b6_filled or not b7_filled:
        return {
            'pass': False,
            'message': f'🚨 CRITICAL: 设定信息未完整填写 (B6={b6}, B7={b7})，必须在Phase 0 Step 0.7填写'
        }
    return {
        'pass': True,
        'message': f'✅ 设定信息已填写 (B6={b6}, B7={b7})'
    }
```

- B6 = 被评估单位全称（Phase 0已从BS表头/科目余额表表头提取）
- B7 = 评估基准日（Phase 0已从BS表头/科目余额表表头提取，YYYY-MM-DD格式）

**⚠️ 如果B6或B7为空**：说明Phase 0 Step 0.7自动提取失败且未人工补充 → **CRITICAL** → 禁止继续隐藏操作，必须先补充填写

**禁止**将设定信息标记为"系统辅助表"而跳过填写。

#### 6.1.2 扫描并隐藏无数据明细表（DT-20）

逐Sheet扫描所有明细表，**按以下精确步骤判定空白Sheet**：

**判定标准**：A列"合计1"所在行的账面价值列**AND**评估价值列均为0或blank → 视为空白Sheet → 隐藏

**执行步骤**：
1. 通过A列标记定位"合计1"行（`A列值=='合计1'`的行号）
2. 通过表头行（A列标记为"检索表头"/"检索表头1"/"检索表头2"的行）定位"账面价值"和"评估价值"列号
3. 读取合计1行对应的账面价值单元格值和评估价值单元格值
4. 判定：账面价值∈{0, None, blank} **AND** 评估价值∈{0, None, blank} → 空白Sheet

**不隐藏的情况**：
- 仅账面值=0而评估值≠0（或反之）——有效评估数据不应被藏起
- 汇总类Sheet（1-汇总表/2-分类汇总等结构表）——始终保留可见

隐藏方式：`ws.sheet_state = 'hidden'`（隐藏整个工作表，不是隐藏行）。

#### 6.1.2a 汇总sheet可见性递归校验（DT-123，在5.1.2之后立即执行）

> **🚨 根因**：DT-20仅检查sheet自身合计行是否有数据，不递归检查下级。河南平绿项目中4-非流动资产汇总自身合计行无直接数据，但其下级4-8-5车辆/4-8-6电子设备/4-16长期待摊费用/4-17递延所得税资产均有数据——汇总sheet被误隐藏，导致层级链断裂。

**执行步骤**：

1. **建立汇总层级树**：从各汇总sheet的公式引用中提取层级关系。评估明细表的完整层级如下（MUST作为内置参考，不需要每次动态提取）：

```
2-分类汇总
├── 3-流动资产汇总
│   ├── 3-1货币资金汇总表 → {3-1-1现金, 3-1-2银行存款}
│   ├── 3-8其他应收款汇总 → {3-8-3其他应收款}
│   ├── 3-9存货汇总 → {3-9-5产成品（库存商品）, ...}
│   ├── 3-7预付款项（直接引用）
│   └── 3-13其他流动资产（直接引用）
├── 4-非流动资产汇总
│   ├── 4-8固定资产汇总 → {4-8-5车辆, 4-8-6电子设备, ...}
│   ├── 4-16长期待摊费用（直接引用）
│   ├── 4-17递延所得税资产（直接引用）
│   └── 其他无数据汇总（4-7/4-9/4-13等）
├── 5-流动负债汇总
│   ├── 5-10其他应付款汇总表 → {5-10-3其他应付款}
│   ├── 5-5应付账款, 5-7合同负债, 5-8职工薪酬, 5-9应交税费, 5-13其他流动负债
│   └── ...
└── 6-非流动负债汇总
    └── 6-1长期借款
```

2. **标记有数据的叶子sheet**：所有在5.1.2中判定为"非空白"（未被隐藏）的明细表sheet。

3. **从叶子向上传播可见性**：
```python
# 伪代码
all_with_data = set(可见的明细表sheet名称)
should_show = set(all_with_data)
changed = True
while changed:
    changed = False
    for parent, children in hierarchy.items():
        if parent in should_show:
            continue
        for child in children:
            if child in should_show:
                should_show.add(parent)
                changed = True
                break

# 恢复汇总sheet可见性
for sn in should_show:
    if '汇总' in sn or sn in hierarchy:
        ws = wb[sn]
        if ws.sheet_state == 'hidden':
            ws.sheet_state = 'visible'
```

4. **特殊规则**：
   - 1-汇总表引用2-分类汇总，如果2-分类汇总可见则1-汇总表也必须可见
   - 净资产汇总引用2-分类汇总，同理
   - **始终隐藏的辅汇总表**（3-辅-流动资产汇总/3-9-辅-存货汇总/8-减值准备汇总表/9-非财务信息汇总表）不受此规则影响，它们由DT-61规则管辖，始终hidden

5. **验证**：输出所有可见sheet清单，确认汇总链完整

#### 6.1.3 隐藏系统辅助工作表（DT-23）

以下3个系统辅助表标记隐藏（`ws.sheet_state='hidden'`）：

| 系统辅助表名 | 隐藏前必须完成的操作 |
|-------------|-------------------|
| 设置 | 直接隐藏 |
| 0-其他方法结论 | 直接隐藏 |
| 设定信息 | **必须先完成5.1.1填写B6/B7**，再隐藏 |

**注意**：仅此3表为"始终隐藏"的辅助表，不可扩展不可遗漏。

#### 6.1.4 确认辅汇总表隐藏状态（DT-61）

以下四张辅汇总表**始终默认隐藏**，不论其子表是否有数据：
- `3-辅-流动资产汇总`
- `3-9-辅-存货汇总`
- `8-减值准备汇总表`
- `9-非财务信息汇总表`

确认其`sheet_state='hidden'`，如非hidden则设为hidden。

#### 6.1.5 输出隐藏清单供确认

```
===== Sheet隐藏操作清单 =====
系统辅助表（DT-23）:
  ✅ 设置 → hidden
  ✅ 0-其他方法结论 → hidden
  ✅ 设定信息 → hidden (B6=xxx, B7=xxxx-xx-xx已填写)

无数据明细表（DT-20）:
  ✅ 4-8-1房屋建筑物 → hidden (账面=0, 评估=0)
  ✅ 4-8-4机器设备 → hidden (账面=0, 评估=0)
  ...

始终隐藏辅汇总表（DT-61）:
  ✅ 3-辅-流动资产汇总 → hidden
  ✅ 3-9-辅-存货汇总 → hidden
  ✅ 8-减值准备汇总表 → hidden
  ✅ 9-非财务信息汇总表 → hidden

保留可见:
  ✅ 3-1-2银行存款 (账面=29,695,319.36, 评估=29,695,319.36)
  ...
=============================
```

#### 6.1.6 保存

完成所有隐藏操作后，`wb.save()`保存。后续Step 6.6将执行COM重算。

### Step 6.3 设置明细表页脚（DT-27/DT-81）

所有明细表（非汇总表）需设置页脚——左侧：`被评估单位填表人：{姓名}\n填表日期：{YYYY年MM月DD日}`，右侧：`评估人员：{姓名1}  {姓名2}`；汇总表不设页脚。需同时设置oddFooter和evenFooter。

**明细表判断规则**：标题中不含"汇总"字样且非00000000系统表。

**DT-81 页脚信息获取规则**：交付评估明细表后，MUST一次性输出页脚信息请求（填表人姓名、填表日期、评估人员姓名）。用户不回复则默认跳过，不阻塞交付，不反复追问（DT-151）。获取后补充页脚信息并重新保存。

### Step 6.4 银行存款特殊处理

- **DT-37**：银行存款步骤复核表中，函证要求应为"对银行存款账户，除金额特别小的账户外，均应函证基准日余额"
- **DT-38**：银行存款无余额调节表时，隐藏相关行/列

### Step 6.5 步骤复核表处理

- **DT-39**：步骤复核表禁止无意义标黄（已有公式自动计算的单元格不需要标黄）
- **DT-40**：每个步骤的"是"或"不适用"列必须勾选其一
- **DT-41**：过程表索引列必须完整填写，与步骤复核表互为索引

### Step 6.6 COM重算保存（DT-17/DT-29）

```python
import win32com.client, time
app = win32com.client.DispatchEx('Excel.Application')  # DT-29: 必须用DispatchEx
app.Visible = False
app.DisplayAlerts = False
try:
    wb_com = app.Workbooks.Open(DETAIL_FILE)
    app.Calculate()
    time.sleep(2)
    wb_com.Save()
    wb_com.Close(SaveChanges=False)
finally:
    app.Quit()
```

**DT-33**：如果文件中有datetime类型的发生日期，COM resave后需再次用openpyxl打开修复日期格式（COM会将datetime转成int序列号+覆盖number_format）。二次修复后不再做COM resave。

### Step 6.6a 🚨 Phase 3完整性验证（DT-161 L3兜底拦截）

> **⚠️ 本步骤为v3.48新增硬约束，上海图灵项目Phase 3被跳过复盘修复**
> **核心原则**：交付前必须验证往来Sheet的发生日期列和业务内容列完整率=100%。这是L3兜底拦截——即使L1(脚本内嵌)和L2(Gate检查)都未拦截，此步骤作为最后一道防线。

**前置条件**：`file_manifest.json`中`序时账=True`。若`序时账=False`，跳过本步骤（DT-143）。

**执行步骤**：

```python
import json, os

# 1. 读取file_manifest判断是否有序时账
cache_dir = os.path.join(PROJECT_DIR, '_dt_cache')
manifest_path = os.path.join(cache_dir, 'file_manifest.json')
has_journal = False
if os.path.exists(manifest_path):
    with open(manifest_path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    for item in manifest.get('data', {}).get('files', []):
        fname = item.get('path', '').lower()
        if '序时账' in fname or '明细账' in fname:
            has_journal = True
            break

if not has_journal:
    print("⏭️ Step 6.6a跳过: 无序时账(DT-143)")
else:
    # 2. 调用gate_G2(has_journal=True)验证
    from gate_validator import gate_G2
    passed, violations = gate_G2(DETAIL_FILE, has_journal=True)

    # 3. 筛选G2-18违规
    g18_violations = [v for v in violations if v.get('gate') == 'G2-18']
    g18_critical = [v for v in g18_violations if v['severity'] == 'CRITICAL']

    if g18_critical:
        print(f"🚨 CRITICAL: Phase 3完整性验证失败！")
        print(f"  发生日期/业务内容缺失: {len(g18_critical)}处")
        for v in g18_critical[:10]:
            print(f"    {v['sheet']} {v['cell']}: {v['message']}")
        print(f"\n  ❌ 禁止交付！请执行Phase 3（S3_journal_extract.md）后再交付。")
        # DT-161: 有序时账但发生日期/业务内容为空 = 底稿不可审 = 不可交付
        raise RuntimeError("DT-161 CRITICAL: Phase 3未执行，往来Sheet发生日期/业务内容为空，禁止交付！")
    else:
        complete_count = len([v for v in g18_violations if v['severity'] == 'INFO'])
        print(f"✅ Step 6.6a通过: Phase 3完整性验证OK ({complete_count}个往来Sheet完整率100%)")
```

**约束**：
- **DT-161**：有序时账时Phase 3 MUST执行。本步骤是L3兜底拦截，验证发生日期列+业务内容列完整率=100%
- **DT-149**：业务内容仅填科目名=等于没填=视为不完整
- G2-18 CRITICAL违规=禁止交付，必须回到Phase 3执行S3_journal_extract.md

### Step 6.7 交付

**交付文件：** `评估明细表-{公司简称}-v{版本号}-{YYYYMMDD}.xlsx`（DT-9：公司简称=去后缀取核心商号，版本号从v1.0起始，年月日=评估基准日）

**交付规则（DT-28）：**
- 文件已保存到用户指定路径（桌面/项目文件夹等workspace外路径）→ 不调用deliver_attachments
- 文件保存在workspace内 → 必须调用deliver_attachments正式交付

### Step 6.8 Skill迭代情况汇报（交付时强制）

交付成果时，必须同时向使用人汇报对应skill的迭代情况。

**汇报格式**：

| 项目 | 内容 |
|------|------|
| Skill名称 | 评估明细表填写 |
| 当前版本 | v3.39 |
| 本次是否迭代 | 是/否 |
| 新增规则/教训 | （如有，逐条列出） |
| 升级规则 | （如有，逐条列出） |
| 新增验证项 | （如有，逐条列出） |

### Step 6.10 🚨 推论映射汇报 [DT-117]

> **⚠️ 本步骤为v3.27新增硬约束，河南平绿项目"其他货币资金"映射问题复盘修复**
> **核心原则**：通过推论建立映射的填写数据，MUST在交付时以文字汇报形式说明情况，让使用人知悉并确认。推论映射非直接编码映射，使用人有权质疑和修正。

**前置条件**：Step 4.4差额推论映射已完成，数据分类清单中有"已确认映射"或"待确认映射"状态的项。

**汇报格式（MUST包含）**：

```markdown
## 推论映射汇报

以下明细表数据通过差额推论建立映射关系，非直接从科目余额表/资产负债表编码映射。
使用人确认后，该映射关系方为有效。

| # | 明细表Sheet | 金额 | 数据来源 | 推论依据 | 映射确认方式 |
|---|------------|------|---------|---------|------------|
| 1 | 3-1-3其他货币资金 | 5,870,945.58 | 保证金对账单PDF(6份) | BS货币资金中无"其他货币资金"子项，
但保证金对账单显示6笔保证金存款。BS"财务公司
存款5,329,404.56"与保证金部分对应 | 差额推论+BS子项分析 |
| 2 | ... | ... | ... | ... | ... |

**说明**：
- 推论映射的数据来源可靠（PDF对账单已提取验证），但归属关系基于推论
- 推论依据：{具体说明差额如何发现、待确认数据如何匹配}
- 使用人如对映射关系有异议，请提出修正意见
```

**DT-117-4核心约束**：
- 推论映射汇报MUST作为交付附件之一，与执行情况摘要同时提供
- 汇报中每个推论映射项MUST说明：金额、数据来源、推论依据、映射确认方式
- 使用人未确认前，推论映射数据仍为"待确认"状态，明细表备注栏保留"[推论映射]"标注
- gate G2-10检查：仍存在"待确认映射"状态项=CRITICAL=禁止交付

## 输出

- 最终交付文件：`评估明细表-{公司简称}-v{版本号}-{YYYYMMDD}.xlsx`
- 执行情况摘要表（逐Phase如实披露，DT-48）
- Skill迭代汇报

## 约束

> **v2.0规则编入步骤声明**：以下规则已编入对应操作步骤。约束区仅保留引用索引。

| 规则 | 编入步骤 | 核心要点 |
|------|---------|---------|
| DT-9 文件命名 | Step 6.7 | 原始模板不覆盖，另存为新文件 |
| DT-17 COM重算保存 | Step 6.6 | openpyxl保存后MUST用Excel COM重算 |
| DT-20 无数据Sheet隐藏 | Step 6.1.2 | 账面值AND评估值同时为0才隐藏 |
| DT-23 系统辅助表隐藏 | Step 6.1.3 | 3个系统表隐藏（设定信息须先填B6/B7） |
| DT-110 隐藏集中执行 | Step 6.1 | 交付前最后一道门控，禁止Phase 2-4期间分散执行 |
| DT-123 汇总sheet递归可见性 | Step 6.1.2a | 子表有数据→父汇总表必须可见 |
| DT-27 页脚设置 | Step 6.3 | 明细表设页脚，汇总表不设 |
| DT-28 交付路径决定 | Step 6.7 | workspace外→不调deliver；workspace内→调deliver |
| DT-29 Excel COM用DispatchEx | Step 6.6 | 禁止用Dispatch |
| DT-33 COM resave日期修复 | Step 6.6 | COM将datetime转int后需二次修复 |
| DT-37 银行存款发函要求 | Step 6.4 | 金额特别小除外，均应函证 |
| DT-39 禁止无意义标黄 | Step 6.5 | 有公式自动计算的单元格不标黄 |
| DT-48 执行透明度 | Step 6.7~5.8 | 逐Phase披露实际执行情况 |
| DT-117 推论映射汇报 | Step 6.10 | 推论映射MUST在交付时汇报 |
| DT-161 Phase 3完整性验证 | Step 6.6a | 有序时账时发生日期+业务内容完整率=100%才可交付 |

**通用原则**：
- **后续复盘经验**：新增的交付/隐藏教训直接写入Step 5.x操作段，不在约束区单独列示

## 异常处理

- COM重算保存失败 → 检查文件是否被占用，重试
- COM resave后datetime变int → 二次修复（DT-33）
- 交付路径不确定 → 确认用户偏好后选择是否调用deliver_attachments

---

## 附录C：Excel COM重算保存模板代码

```python
import win32com.client
import time

def resave_with_cached_values(filepath):
    """openpyxl保存后调用此函数，用Excel重算并写回公式缓存值，消除WPS打开后的保存提示弹窗"""
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(filepath)
        app.Calculate()
        time.sleep(2)  # 等待计算完成
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        app.Quit()

# 使用：在 openpyxl 的 wb.save(path) 之后调用
# ⚠️ 必须用DispatchEx，不用Dispatch（Dispatch会关闭用户其他WPS文件）
# ⚠️ 如果文件中有datetime类型的发生日期，COM resave后需再次用openpyxl打开
#    修复日期格式（COM会将datetime转成int序列号+覆盖number_format）
```

## 附录D：历史教训执行状态追踪

定期（每次项目交付后）审查历史教训的执行状态，确保每条都有对应的自动化验证：

| 状态 | 含义 | 占比目标 |
|------|------|---------|
| ✅ 自动化验证 | Phase 4.5门控中自动检查 | ≥60% |
| 🟡 自检清单 | Phase自检清单中列出，靠人工勾选 | ≤30% |
| ❌ 仅记录 | 记录但无任何验证机制 | ≤10% |

**目标**：每轮更新后，✅占比应提升，❌占比应下降。当❌占比>20%时，需集中补充验证。
