# S4_bs_verify: 分类汇总表财务报表数据校验

> **📋 Common规则适用声明**：本步骤适用 META_RULES MR-1/MR-4/MR-5/MR-6/MR-7/MR-9/MR-10 + preparation_discipline_rules G0-G13
> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase 4规则节：DT-4(勾稽100%覆盖)、DT-69(交叉校验)、DT-70(重分类决策树)
> **📋 v3.64 (2026-05-29)**：I列直接录入BS数值（非公式，打开即见），避免VLOOKUP匹配失败+COM重算依赖；创建隐藏Sheet "_BS对照" 备份原始BS数据供查验。

## 定位

Phase 4 勾稽核对的**前置步骤**。将资产负债表（BS）数据直接写入2-分类汇总表，使逐科目可视化校对：

- **I列（财务报表金额）**：直接录入BS期末余额数值（非公式，无需COM重算）
- **J列（校对差异）**：公式 `=I - E`，BS金额与账面价值的差异（依赖E列公式重算后显示）

## 输入

- Phase 4 格式修复完成的评估明细表
- `_dt_cache/bs_balances.json`（Phase 0 解析的BS数据）

## 🚨 前置确认：2-分类汇总表的列结构

```
R5表头行示例：
  C2=序号 | C3=项目编号 | C4=科目名称 | C5=账面价值 | C6=评估价值 | ...
                   ^^^^^^^^^^^^^   ^^^^^^^^^^
                   科目名称（D列）  账面价值（E列）
```

**关键**：科目名称在 **D列(4)** 而不是B列！B列是序号（1,2,3...），VLOOKUP用B列会全部匹配失败。账面价值在 **E列(5)**，不是D列。

## 操作

### Step 4BS.0 数据载入

```python
import json, os, openpyxl

cache_dir = os.path.join(PROJECT_DIR, '_dt_cache')
bs_path = os.path.join(cache_dir, 'bs_balances.json')

if not os.path.exists(bs_path):
    raise FileNotFoundError(f'BS数据缓存缺失: {bs_path}')

with open(bs_path, 'r', encoding='utf-8') as f:
    bs_data = json.load(f)

wb = openpyxl.load_workbook(DETAIL_FILE)
ws = wb['2-分类汇总']
```

### Step 4BS.1 创建BS对照隐藏Sheet（备查）

创建隐藏Sheet `_BS对照`，写入BS原始数据。此Sheet供人工查验，I列不引用它。

```python
AUX = '_BS对照'
if AUX in wb.sheetnames: del wb[AUX]
ws_aux = wb.create_sheet(AUX)
ws_aux.sheet_state = 'hidden'
ws_aux.cell(row=1, column=1, value='科目名称')
ws_aux.cell(row=1, column=2, value='财务报表期末余额')

# 从bs_data['items']填充
row_idx = 2
for item in bs_data['items']:
    name = item.get('name', '').strip()
    balance = item.get('balance', 0)
    if name:
        ws_aux.cell(row=row_idx, column=1, value=name)
        ws_aux.cell(row=row_idx, column=2, value=round(float(balance), 2))
        ws_aux.cell(row=row_idx, column=2).number_format = '#,##0.00'
        row_idx += 1
```

### Step 4BS.2 建立BS科目映射

```python
# BS科目名称 → 金额
bs_lookup = {}
ws_aux = wb[AUX]
for r in range(2, ws_aux.max_row + 1):
    name = ws_aux.cell(row=r, column=1).value
    bal = ws_aux.cell(row=r, column=2).value
    if name: bs_lookup[str(name).strip()] = float(bal)
```

### Step 4BS.3 逐行写入I列（硬编码数值）+ J列（差异公式）— 全行覆盖

🚨 **关键**：科目名称在 **D列(4)**，账面价值在 **E列(5)**。

🚨 **v2.0 改进**：覆盖模板中所有行，不遗漏任何科目。使用完备匹配策略（精确映射→BS直接匹配→汇总行匹配→I=0兜底）。

🚨 **模板结构**：2-分类汇总模板原始只有A~H列(8列)，I/J列为AI录入校对列。若模板已包含I/J列则直接利用，否则追加。

```python
NAME_COL = 4   # D列：科目名称
BOOK_COL = 5   # E列：账面价值
BS_COL = 9     # I列：财务报表金额
DIFF_COL = 10  # J列：差异校对

# 检查模板是否已有I/J列（已有的则保留格式，没有的需要追加表头标注）
has_header = ws.cell(row=5, column=BS_COL).value is not None
if not has_header:
    # 标注I列表头（第5行左侧已有表头）
    ws.cell(row=5, column=BS_COL).value = '财务报表金额'
    ws.cell(row=5, column=BS_COL).font = ws.cell(row=5, column=5).font  # 保持与E列同字体
    ws.cell(row=5, column=DIFF_COL).value = '差异(BS-明细)'
    ws.cell(row=5, column=DIFF_COL).font = ws.cell(row=5, column=5).font

# 预定义精确映射表：覆盖所有模板科目名称 → BS科目名称
# v2.0: 使用完整映射表覆盖模板所有行（46行+汇总行）
exact_map = {
    '货币资金': '货币资金', '交易性金融资产': '交易性金融资产',
    '衍生金融资产': '衍生金融资产', '应收票据': '应收票据',
    '应收账款': '应收账款', '应收款项融资': '应收款项融资',
    '预付款项': '预付款项', '其他应收款': '其他应收款',
    '存货': '存货', '合同资产': '合同资产',
    '持有待售资产': '持有待售资产',
    '一年内到期的非流动资产': '一年内到期的非流动资产',
    '其他流动资产': '其他流动资产',
    '债权投资': '债权投资', '其他债权投资': '其他债权投资',
    '长期应收款': '长期应收款', '长期股权投资': '长期股权投资',
    '其他权益工具投资': '其他权益工具投资',
    '其他非流动金融资产': '其他非流动金融资产',
    '投资性房地产': '投资性房地产', '固定资产': '固定资产',
    '在建工程': '在建工程', '生产性生物资产': '生产性生物资产',
    '油气资产': '油气资产', '使用权资产': '使用权资产',
    '无形资产': '无形资产', '开发支出': '开发支出', '商誉': '商誉',
    '长期待摊费用': '长期待摊费用',
    '递延所得税资产': '递延所得税资产',
    '其他非流动资产': '其他非流动资产',
    '短期借款': '短期借款', '交易性金融负债': '交易性金融负债',
    '衍生金融负债': '衍生金融负债', '应付票据': '应付票据',
    '应付账款': '应付账款', '预收款项': '预收款项',
    '合同负债': '合同负债', '应付职工薪酬': '应付职工薪酬',
    '应交税费': '应交税费', '其他应付款': '其他应付款',
    '持有待售负债': '持有待售负债',
    '一年内到期的非流动负债': '一年内到期的非流动负债',
    '其他流动负债': '其他流动负债',
    '长期借款': '长期借款', '应付债券': '应付债券',
    '租赁负债': '租赁负债', '长期应付款': '长期应付款',
    '长期应付职工薪酬': '长期应付职工薪酬',
    '预计负债': '预计负债', '递延收益': '递延收益',
    '递延所得税负债': '递延所得税负债',
    '其他非流动负债': '其他非流动负债',
    '应付股利': '应付股利', '应付利息': '应付利息',
}

# from bs_balances.json -> read _BS对照 sheet
bs_lookup = {}
if AUX in wb.sheetnames:
    ws_aux = wb[AUX]
    for r in range(2, ws_aux.max_row + 1):
        name = ws_aux.cell(row=r, column=1).value
        bal = ws_aux.cell(row=r, column=2).value
        if name:
            bs_lookup[str(name).strip()] = float(bal)

updated = 0
for r in range(6, ws.max_row + 1):
    name = ws.cell(row=r, column=NAME_COL).value
    if not name:
        continue
    ns = str(name).strip()
    
    # 跳过空行/分隔行
    if ns in ('', '-', '—'):
        continue
    
    # 处理汇总行（"一、流动资产合计", "二、非流动资产合计"等）
    bs_val = None
    
    # 策略1: 汇总行匹配（长模式优先，防止"流动资产合计"误配"非流动资产合计"）
    summary_patterns = [
        ('流动资产合计', '流动资产合计'),
        ('非流动资产合计', '非流动资产合计'),
        ('资产总计', '资产总计'),
        ('流动负债合计', '流动负债合计'),
        ('非流动负债合计', None),  # BS无独立行→0
        ('负债总计', '负债合计'),
        ('净资产', '所有者权益（或股东权益）合计'),
    ]
    for pattern, bs_key in summary_patterns:
        if pattern in ns:
            if bs_key and bs_key in bs_lookup:
                bs_val = bs_lookup[bs_key]
            else:
                bs_val = 0.0
            break
    
    # 策略2: 精确映射
    if bs_val is None:
        # 移除"一、"/"二、"/"三、"/"四、"/"五、"/"六、"/"七、" prefix for matching
        clean_name = re.sub(r'^[一二三四五六七八九十]、', '', ns).strip()
        if clean_name in exact_map:
            bk = exact_map[clean_name]
            if bk in bs_lookup:
                bs_val = bs_lookup[bk]
    
    # 策略3: 直接 BS 科目名称匹配
    if bs_val is None and ns in bs_lookup:
        bs_val = bs_lookup[ns]
    
    # 策略4: 用 clean_name 再试 BS lookup
    if bs_val is None:
        clean_name = re.sub(r'^[一二三四五六七八九十]、', '', ns).strip()
        if clean_name in bs_lookup:
            bs_val = bs_lookup[clean_name]
    
    # 策略5: 兜底 I=0（确保每行都有数据）
    if bs_val is None:
        bs_val = 0.0
    
    # I列：直接写入硬编码数值
    ws.cell(row=r, column=BS_COL).value = round(bs_val, 2)
    ws.cell(row=r, column=BS_COL).number_format = '#,##0.00'
    
    # J列：差异公式 = I - E（E列重算后自动更新）
    ws.cell(row=r, column=DIFF_COL).value = f'=IF(ISNUMBER(E{r}),I{r}-E{r},"-")'
    ws.cell(row=r, column=DIFF_COL).number_format = '#,##0.00'
    updated += 1
```

⚠️ **I列写入后，原模板第8列之后的数据会被覆盖。如果模板在第8列之后有内容，需先将这些列右移。**

### Step 4BS.4 输出校验摘要

```python
print(f"✅ I列录入了 {updated} 项BS数据（硬编码数值）")
print(f"✅ J列公式 = I - E（差异，E列重算后显示）")
print(f"✅ _BS对照 Sheet已创建，含{len(bs_lookup)}项BS原始数据（隐藏，备查）")

# 统计I=0的项数（无BS匹配的行）
zero_count = sum(1 for r in range(6, ws.max_row + 1) 
                 if ws.cell(row=r, column=BS_COL).value is not None and ws.cell(row=r, column=BS_COL).value == 0)
if zero_count > 0:
    print(f"⚠️ 有 {zero_count} 行I列=0（无对应BS科目）")
```
### Step 4BS.4 输出校验摘要

```python
print(f"✅ I列录入了 {updated} 项BS数据（硬编码数值）")
print(f"✅ J列公式 = I - E（差异，E列重算后显示）")
print(f"✅ _BS对照 Sheet已创建，含{len(bs_lookup)}项BS原始数据（隐藏，备查）")
```

## 输出

- 2-分类汇总表 **I列**：BS财务报表期末余额（直接数值）
- 2-分类汇总表 **J列**：`=I - E` 差异公式
- 隐藏Sheet `_BS对照`：BS原始数据（备查）

## 常见错误

| 错误 | 后果 | 避免方法 |
|------|------|---------|
| VLOOKUP用B列(序号)查 | 全部匹配失败→空 | ✅ 科目名称在D列，必须用D列查 |
| I列写公式而非数值 | 依赖COM重算才能显示 | ✅ 直接写`cell.value = float` |
| D列/E列搞混 | 差异计算错误 | ✅ E列=账面价值，D列=科目名称 |

## 版本

v1.2 (2026-05-29) — 邦能达项目复盘：I列改为直接录入数值，J列公式=I-E，_BS对照Sheet独立备查
