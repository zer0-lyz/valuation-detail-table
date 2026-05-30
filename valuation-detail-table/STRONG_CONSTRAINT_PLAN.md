# DT Skill 系统性强约束改进方案

> **版本**：v2.0 (已实施) | **日期**：2026-05-24
> **触发**：河南平绿项目18个问题复盘，用户要求"所有规则均应当有强约束，严禁任何跳过或不执行"
> **实施状态**：✅ Phase 1已全部实施完成

---

## 一、问题本质：规则的"知道"≠"做到"

### 1.1 当前约束力审计

对142条DT规则逐条审计，按约束力层级分类：

| 约束力层级 | 定义 | DT规则数 | 占比 | 违反后果 |
|-----------|------|---------|------|---------|
| **L1 脚本强制** | 规则已嵌入Python代码，违反则运行时崩溃/断言失败/exit(1) | ~30条 | 21% | 技术性阻断，不可能绕过 |
| **L2 Gate门控** | 规则由gate_validator.py检测，违反则流程阻断 | ~15条 | 11% | Phase间门控拦截 |
| **L3 纯文字** | 规则仅存在于Markdown（RULES.md/steps/*.md/CHECK.md），依赖Agent自觉遵守 | ~97条 | 68% | **无技术阻断，可被绕过** |

### 1.2 18个问题与约束力的对应

| 问题# | 问题描述 | 涉及DT规则 | 当前约束力 | 应有约束力 |
|-------|---------|-----------|-----------|-----------|
| 1 | 隐藏未执行 | DT-20/DT-23/DT-110 | L3(纯文字) | L1(脚本) |
| 2 | 空行 | DT-120/DT-82 | L1(gate G2-12)但未执行 | L1(强制) |
| 3 | 发生日期空着 | DT-30 | L2(G1-5) | L1(默认值) |
| 4 | 其他应收未展开末级 | DT-111 | L2(G0-3)但跳过 | L1(脚本) |
| 5 | 坏账未填至坏账行 | DT-6/DT-18 | L2(G1-4/G1-7) | L1(脚本) |
| 6 | 外币汇率不该填 | — | 无规则 | L1(默认值) |
| 7 | 存货开发成本映射 | — | 无规则 | L2(行业映射JSON) |
| 8 | 车辆未填 | DT-129 | L3(纯文字) | L1(脚本) |
| 9 | 电子设备+格式破坏 | DT-120/DT-75/DT-83 | L1(gate)但unmerge_all破坏 | L1(脚本) |
| 10 | 长摊未引用Excel | DT-129 | L3(纯文字) | L1(数据源检测) |
| 11 | 递延所得税名称 | DT-89 | L2(G2-8)仅排除零余额 | L1(脚本) |
| 12 | 应付账款格式+内容 | DT-46/DT-120/DT-60 | L3+L1混合 | L1(脚本) |
| 13 | 合同负债/职工薪酬列位错 | DT-46 | L2(G1)但跳过 | L1(脚本) |
| 14 | 职工/税费格式+内容 | DT-46/DT-126 | L2(G1)但跳过 | L1(脚本) |
| 15 | 应付利息填错列 | DT-46 | L2(G1)但跳过 | L1(脚本) |
| 16 | 其他应付格式+内容 | DT-46/DT-120 | L3+L1混合 | L1(脚本) |
| 17 | 其他流动负债格式+列 | DT-46/DT-126 | L3+L1混合 | L1(脚本) |
| 18 | 长期借款未填 | DT-129 | L3(纯文字) | L1(数据源检测) |

**关键发现**：
- 18个问题中15个的根因是"规则存在但无L1强约束"
- 3个问题（#6/7/10）是"规则本身不存在"
- L2 Gate已有但被跳过（#2/4/5/13/14/15）→ 说明L2也不够强，需L1兜底

### 1.3 核心结论

**Markdown规则 = 软约束。Agent可以"知道"规则但选择绕过，没有任何技术手段阻止。**

只有两种强约束：
1. **L1 脚本强制**：规则变成Python代码，违反就crash
2. **L2 Gate门控**：规则变成检测项，违反就阻断流程

---

## 二、系统性改进方案：四层强约束架构

### 2.1 架构总览

```
┌─────────────────────────────────────────────────────────┐
│  L4 流程层 — Phase间强制门控序列                          │
│  每个Phase完成后自动触发下一Phase的Gate检测                │
│  不通过 = 禁止继续（不是"建议修复"）                       │
├─────────────────────────────────────────────────────────┤
│  L3 规则层 — RULES.md (保留，作为规则定义的唯一来源)        │
│  每条规则标注强制等级：A=必须L1/B=必须L2/C=L3即可          │
├─────────────────────────────────────────────────────────┤
│  L2 检测层 — gate_validator.py (5级门控引擎)              │
│  G0→G1→G1F→G2→G3 逐级递进                               │
│  CRITICAL>0 → sys.exit(1) → Agent无法继续                 │
├─────────────────────────────────────────────────────────┤
│  L1 执行层 — sheet_filler.py (Agent唯一写入接口)           │
│  Agent只做：准备数据 → 调用fill_sheet() → 结束             │
│  Agent禁止：直接import openpyxl、直接操作ws对象             │
│  违反 = import时检测 + 运行时断言                          │
└─────────────────────────────────────────────────────────┘
```

### 2.2 L1 核心模块：`sheet_filler.py`

**设计原则**：Agent唯一的Excel写入接口，物理上封堵绕过路径。

```python
# valuation-common/scripts/sheet_filler.py
"""
评估明细表统一写入引擎 — Agent唯一允许调用的Excel写入接口

强制执行：
1. DT-136 动态列位映射 — 从表头读取col_map，硬编码列号→assert失败
2. DT-46  C/D列序校验 — 资产/负债列序自动切换
3. DT-120 smart_insert_row — 插入行只能走这个函数
4. DT-125 MergedCell检测 — 自动unmerge数据行，保护表头
5. DT-116 A列标记体系 — 自动识别行类型
6. DT-18  减值行自动填写 — 传入坏账准备金额即可
7. DT-30  日期类型强制 — 自动转datetime
8. DT-66  全列位校验 — 写入前校验列含义
9. DT-97  写入后回读 — assert数据已写入
10. 行业映射 — 从industry_mapping.json读取
11. 业务内容映射 — 自动从科目编码+结算对象推断
12. 外币默认值 — 无外币数据时不填币种/汇率列
"""

def fill_sheet(ws, sheet_id, data_rows, settings, wb=None):
    """
    统一写入函数
    
    Args:
        ws: openpyxl worksheet对象
        sheet_id: Sheet编号（如"3-1-2", "5-5"）
        data_rows: list of dict, 每行数据
        settings: dict, 含company_name/base_date等
        wb: workbook对象（用于跨sheet引用更新）
    
    Returns:
        dict: {'success': bool, 'rows_written': int, 'warnings': list}
    
    Raises:
        AssertionError: 任何DT规则违反时
        ValueError: 数据类型错误时
    """
    # Step 1: 读取A列标记，识别行结构
    struct = _find_header_structure(ws)
    
    # Step 2: 动态列位映射（DT-136）
    col_map = _build_col_map(ws, struct, sheet_id)
    
    # Step 3: 列序校验（DT-46）
    is_asset = sheet_id.startswith('3-') or sheet_id.startswith('4-')
    _validate_col_sequence(col_map, is_asset)
    
    # Step 4: 插入行（DT-120，强制smart_insert_row）
    rows_needed = len(data_rows)
    result = smart_insert_rows_for_data(ws, data_count=rows_needed, wb=wb)
    
    # Step 5: 逐行写入（DT-125 MergedCell检测 + DT-30日期类型 + DT-0数据来源）
    for i, row_data in enumerate(data_rows):
        row_num = struct['data_start_row'] + i
        _write_data_row(ws, row_num, row_data, col_map, is_asset, settings)
    
    # Step 6: 减值行填写（DT-18/DT-6）
    if settings.get('bad_debt_amount'):
        _fill_bad_debt_row(ws, struct, settings['bad_debt_amount'])
    
    # Step 7: 写入后回读验证（DT-97）
    _assert_data_written(ws, struct, data_rows, col_map)
    
    # Step 8: 列位回读assert（DT-136）
    _assert_col_map_correct(ws, struct, col_map, data_rows)
    
    return {'success': True, 'rows_written': len(data_rows), 'warnings': []}
```

**关键约束机制**：

| 绕过路径 | L1如何封堵 |
|---------|-----------|
| Agent自行写openpyxl代码 | SKILL.md + steps/*.md中删除所有裸openpyxl代码模板，只保留`from sheet_filler import fill_sheet` |
| Agent跳过fill_sheet直接操作ws | 步骤文件中只允许"准备data_rows字典"→调用fill_sheet，无其他写入代码可复制 |
| Agent硬编码列号 | fill_sheet内部强制DT-136 col_map，无参数可传入列号 |
| Agent用unmerge_all() | fill_sheet内部只在数据行unmerge，表头区域受保护 |
| Agent跳过行业映射 | fill_sheet从industry_mapping.json读取，Agent无法绕过 |
| Agent不填业务内容 | fill_sheet的_write_data_row自动推断业务内容 |
| Agent填错C/D列序 | fill_sheet根据sheet_id自动判断资产/负债列序 |
| Agent填外币汇率 | fill_sheet在无外币数据时自动跳过币种/汇率列 |

### 2.3 L2 门控增强：Gate强制执行机制

**当前问题**：gate_validator.py存在且功能完备，但Agent可以选择不调用。

**改进方案**：Phase间Gate强制触发

```python
# valuation-common/scripts/phase_gate.py
"""
Phase间强制门控 — 每个Phase完成后自动触发，Agent无法跳过

设计原理：
- gate_validator.py是"检测器"，phase_gate.py是"触发器"
- 检测器可以存在但被跳过，触发器不行——因为触发器嵌入fill_sheet的返回路径
- fill_sheet()成功返回后自动调用对应级别的gate
"""

PHASE_GATE_MAP = {
    'Phase-1': 'G0',    # 数据源完整性
    'Phase0':  'G0',    # 数据源完整性
    'Phase2':  'G1',    # 每个Sheet写入级
    'Phase2-end': 'G2', # Phase 2完成后
    'Phase3':  'G1F',   # 格式门控
    'Phase4':  'G3',    # 勾稽级
}

def run_phase_gate(xlsx_path, phase, **kwargs):
    """强制运行Phase对应Gate，CRITICAL>0则sys.exit(1)"""
    gate = PHASE_GATE_MAP.get(phase)
    if not gate:
        return True
    
    result = subprocess.run(
        [sys.executable, GATE_VALIDATOR_PATH, xlsx_path, '--gate', gate],
        capture_output=True, text=True, **kwargs
    )
    
    if result.returncode != 0:
        print(f"🚨 GATE {gate} FAILED for {phase}")
        print(result.stdout)
        sys.exit(1)  # 强制阻断
    
    return True
```

**嵌入方式**：在`fill_sheet()`末尾自动调用：

```python
def fill_sheet(ws, sheet_id, data_rows, settings, wb=None, filepath=None):
    # ... 数据写入 ...
    
    # 写入后自动触发G1门控（DT-74即时门控）
    if filepath:
        run_phase_gate(filepath, 'Phase2', sheet=sheet_id)
    
    return result
```

### 2.4 行业映射数据：`industry_mapping.json`

```json
{
  "real_estate": {
    "name": "房地产开发",
    "subject_to_sheet": {
      "5002": {"sheet": "3-9-6", "row_name": "开发成本", "category": "存货-在产品"},
      "5001": {"sheet": "3-9-6", "row_name": "开发产品", "category": "存货-产成品"},
      "1405": {"sheet": "3-9-1", "row_name": "原材料", "category": "存货-原材料"},
      "1601": {"sheet": "4-8", "split_by": "card_register", "sub_sheets": ["4-8-5","4-8-6"]}
    },
    "auto_detect_keywords": ["置业", "房地产", "开发", "建设集团"]
  },
  "manufacturing": {
    "name": "制造业",
    "subject_to_sheet": {
      "1405": {"sheet": "3-9-1", "row_name": "原材料", "category": "存货-原材料"},
      "1406": {"sheet": "3-9-2", "row_name": "库存商品", "category": "存货-产成品"},
      "1601": {"sheet": "4-8", "split_by": "card_register", "sub_sheets": ["4-8-1","4-8-5","4-8-6"]}
    },
    "auto_detect_keywords": ["制造", "科技", "工业", "电子"]
  }
}
```

**强制执行**：
1. Phase 0解析BS时自动识别行业类型
2. `fill_sheet()`从JSON读取映射，Agent无法跳过
3. 新行业只需在JSON中添加条目，无需改规则文档

### 2.5 业务内容自动映射：`business_content_map.py`

```python
# valuation-common/scripts/business_content_map.py
"""
业务内容自动推断 — 从科目编码+结算对象名→业务实质

根因：18个问题中#12/#16均涉及"业务内容填应付账款=没填"
Agent倾向填入科目名称而非业务实质，需脚本强制推断
"""

SUBJECT_CONTENT_MAP = {
    '2202': {  # 应付账款
        'default': '货款',
        'keywords': {
            '工程': '工程款', '施工': '工程款', '建设': '工程款',
            '材料': '材料款', '钢材': '材料款', '水泥': '材料款',
            '设备': '设备款', '机械': '设备款',
            '劳务': '劳务费', '人力': '劳务费',
            '服务': '服务费', '咨询': '咨询费', '设计': '设计费',
            '物业': '物业费', '租赁': '租金',
            '暂估': '暂估款', '估': '暂估款',
        }
    },
    '2203': {  # 预收款项
        'default': '预收货款',
        'keywords': {
            '房': '预售房款', '购房': '预售房款', '定金': '定金',
            '货': '预收货款', '服务': '预收服务费',
        }
    },
    '2241': {  # 其他应付款
        'default': '往来款',
        'keywords': {
            '保证金': '保证金', '押金': '押金', '担保': '保证金',
            '社保': '代扣社保', '公积金': '代扣公积金', '医保': '代扣医保',
            '工资': '代发工资', '薪': '代发工资',
            '利息': '利息', '股利': '股利', '分红': '股利',
            '集团': '集团往来', '内部': '内部往来', '结算中心': '结算中心借款',
            '工会': '工会经费', '职工': '职工往来',
            '税': '代扣税款', '增值税': '代扣增值税',
        }
    },
    '1122': {  # 应收账款
        'default': '货款',
        'keywords': {
            '物业': '物业费', '管理': '管理费', '服务': '服务费',
            '租赁': '租金', '租金': '租金',
            '工程': '工程款', '施工': '工程款',
        }
    },
    '1221': {  # 其他应收款
        'default': '往来款',
        'keywords': {
            '保证金': '保证金', '押金': '押金', '投标': '投标保证金',
            '备用金': '备用金', '借款': '个人借款',
            '社保': '代垫社保', '公积金': '代垫公积金',
            '集团': '集团往来', '内部': '内部往来',
            '退款': '退款', '退': '退款',
            '税': '退税', '返还': '税费返还',
        }
    },
}

def infer_business_content(subject_code, settlement_name):
    """从科目编码+结算对象名推断业务内容"""
    code_prefix = subject_code[:4]  # 取4位科目编码前缀
    content_config = SUBJECT_CONTENT_MAP.get(code_prefix)
    
    if not content_config:
        return settlement_name  # 无映射规则时返回原名
    
    name_str = str(settlement_name)
    for keyword, content in content_config['keywords'].items():
        if keyword in name_str:
            return content
    
    return content_config['default']
```

### 2.6 DT-141 assert链全面覆盖

**当前状态**：DT-141定义了"脚本自检assert=规则执行确认"，但仅覆盖了S2_fill_bs.md中的部分规则。

**改进方案**：所有Phase 2步骤文件中，每条DT规则必须有对应的assert断言

```python
# ══════════════════════════════════════════════════════════
# DT-141: 规则校验清单（本脚本强制执行）
# ══════════════════════════════════════════════════════════
#
# [DT-0]  零幻觉 → assert all(v.get('source') for v in data_rows)
# [DT-46] C/D列序 → assert col_map['业务内容'] < col_map['发生日期'] if is_asset else col_map['发生日期'] < col_map['业务内容']
# [DT-66] 全列位校验 → assert len(col_map) >= 4, f"col_map仅{len(col_map)}列，表头未完整读取"
# [DT-90] 表头禁写 → assert first_data_row > header_row + (1 if has_sub_header else 0)
# [DT-97] 写入后回读 → assert ws.cell(row=first_data_row, column=col_map['账面价值']).value is not None
# [DT-116] A列标记 → assert struct['total_row'] is not None, "未找到合计1行"
# [DT-120] smart_insert → assert 'smart_insert' in str(result), "未使用smart_insert_row"
# [DT-125] MergedCell → 已在_write_data_row中自动处理
# [DT-129] 非往来展开 → assert len(data_rows) >= len(sub_subjects), "子科目未逐项展开"
# [DT-136] 动态列位 → assert 'col_map' in dir(), "未建立col_map"
# [DT-138] gate强制 → 已在fill_sheet末尾自动调用
# [DT-18]  减值行 → assert bad_debt_row_written, "坏账准备行未填写"
# [DT-30]  日期类型 → assert isinstance(date_val, datetime), f"日期非datetime: {type(date_val)}"
# [DT-111] 辅助余额展开 → assert len(data_rows) >= aux_count, f"行数{len(data_rows)}<辅助余额{aux_count}"
# ══════════════════════════════════════════════════════════
```

---

## 三、18个问题的逐项强约束方案

| 问题# | 根因 | L1脚本强制 | L2 Gate | L3规则补充 |
|-------|------|-----------|---------|-----------|
| 1 | 隐藏未执行 | `sheet_filler.py`内置`hide_empty_sheets()`，Phase 5自动调用 | G3-11已有，增强为自动触发 | 无需补充 |
| 2 | 空行 | `smart_insert_row`已修复此问题，`sheet_filler`强制使用 | G2-12已有检测 | 无需补充 |
| 3 | 发生日期 | `sheet_filler._write_data_row()`自动设`None`而非跳过 | G1-5检测datetime | 新增DT-143：无序时账时发生日期留空(填None) |
| 4 | 其他应收未展开 | `sheet_filler.fill_sheet()`强制从auxiliary_balance_all.json读取 | G0-3已有检测 | 无需补充 |
| 5 | 坏账未填 | `sheet_filler._fill_bad_debt_row()`自动填写 | G1-4/G1-7已有 | 无需补充 |
| 6 | 外币汇率 | `sheet_filler._write_data_row()`无外币数据时跳过E/F列 | — | 新增DT-144：无外币时不填币种/汇率列 |
| 7 | 存货映射 | `industry_mapping.json`内置5002→3-9-6映射 | G2新增行业映射检测 | 新增DT-145：行业特殊科目映射规则 |
| 8 | 车辆未填 | `sheet_filler`根据卡片台账自动拆分4-8-x子表 | G2新增固定资产子表完整性 | 无需补充 |
| 9 | 电子设备+格式 | `sheet_filler`消灭`unmerge_all()`，仅数据行unmerge | G1F-5检测合并 | 无需补充 |
| 10 | 长摊未引用Excel | `sheet_filler`强制从科目余额表+辅助数据读取 | G2新增数据源检测 | 新增DT-146：长期待摊费用须引用辅助明细 |
| 11 | 递延所得税名称 | `sheet_filler._write_data_row()`自动填充具体名称 | G2-8增强 | 无需补充 |
| 12 | 应付格式+内容 | `sheet_filler`消灭unmerge_all + business_content_map自动推断 | G1已有 | 无需补充 |
| 13 | 列位错 | `sheet_filler`强制DT-136 col_map | G1-6已有 | 无需补充 |
| 14 | 职工/税费格式+内容 | `sheet_filler`消灭unmerge_all + 列位强制 | G1已有 | 新增DT-147：应交税费逐税种填写+征税机关格式 |
| 15 | 应付利息列位错 | `sheet_filler`强制col_map | G1-6已有 | 无需补充 |
| 16 | 其他应付内容 | `business_content_map.py`自动推断 | G1已有 | 无需补充 |
| 17 | 其他流动负债格式+列 | `sheet_filler`消灭unmerge_all + 列位强制 | G1已有 | 无需补充 |
| 18 | 长期借款未填 | `sheet_filler`强制从D1/D2/D3映射读取 | G2新增 | 新增DT-148：长期借款必须填写(6-1) |

---

## 四、实施路径

### Phase 1：L1核心模块（P0，消除61%问题）

| 步骤 | 内容 | 预计工作量 |
|------|------|-----------|
| 1.1 | 创建`valuation-common/scripts/sheet_filler.py` | 2天 |
| 1.2 | 创建`valuation-common/scripts/business_content_map.py` | 0.5天 |
| 1.3 | 创建`valuation-common/scripts/industry_mapping.json` | 0.5天 |
| 1.4 | 创建`valuation-common/scripts/phase_gate.py` | 0.5天 |
| 1.5 | 重写S2_fill_*.md步骤文件：删除裸openpyxl代码模板，只保留`fill_sheet()`调用 | 1天 |
| 1.6 | 消灭`unmerge_all()`：`sheet_filler`内部实现`_safe_write_cell()` | 0.5天 |

### Phase 2：L2 Gate增强（P1，阻止错误累积）

| 步骤 | 内容 | 预计工作量 |
|------|------|-----------|
| 2.1 | gate_validator.py新增行业映射检测(G2-13) | 0.5天 |
| 2.2 | gate_validator.py新增数据源完整性强制检测(G0-6) | 0.5天 |
| 2.3 | gate_validator.py新增长期借款/固定资产子表检测(G2-14) | 0.5天 |
| 2.4 | Phase间Gate自动触发机制嵌入fill_sheet返回路径 | 0.5天 |

### Phase 3：L3规则补充（P2，补缺）

| 步骤 | 内容 | 预计工作量 |
|------|------|-----------|
| 3.1 | 新增DT-143~DT-148规则 | 0.5天 |
| 3.2 | RULES.md所有规则标注强制等级(A/B/C) | 0.5天 |
| 3.3 | 步骤文件全面替换为fill_sheet调用 | 0.5天 |

---

## 五、强约束验证标准

### 5.1 每条规则必须回答的问题

| 问题 | 合格答案 | 不合格答案 |
|------|---------|-----------|
| 违反时Agent会怎样？ | 脚本crash/断言失败/exit(1) | "应该遵守"/"建议执行" |
| Agent能否绕过？ | 不能，物理封堵 | 可能忘记/跳过 |
| 绕过的后果？ | 不可能绕过 | 产出有错误 |
| 检测方式？ | 自动运行(脚本/Gate) | 人工review |

### 5.2 规则强制等级分类标准

| 等级 | 含义 | 实现方式 | 适用规则类型 |
|------|------|---------|------------|
| **A** | 必须L1脚本强制 | 嵌入sheet_filler.py/gate_validator.py，违反crash | 写入类、格式类、列位类、类型类 |
| **B** | 必须L2 Gate门控 | gate_validator.py检测，CRITICAL>0→exit(1) | 完整性类、勾稽类、映射类 |
| **C** | L3纯文字可接受 | Markdown规则+CHECK.md校验项 | 流程指导类、命名规范类、交付信息类 |

### 5.3 142条DT规则的强制等级目标分配

| 强制等级 | 当前 | 目标 | 变化 |
|---------|------|------|------|
| A (L1脚本) | ~30条(21%) | ~95条(67%) | +65条 |
| B (L2 Gate) | ~15条(11%) | ~30条(21%) | +15条 |
| C (L3文字) | ~97条(68%) | ~17条(12%) | -80条 |

**核心变化**：68%的纯文字规则降至12%，67%的规则升级为脚本强制。

---

## 六、风险与缓解

| 风险 | 影响 | 缓解措施 |
|------|------|---------|
| sheet_filler.py过于刚性，不适应新模板 | 新模板需改代码 | 设计插件式col_map + industry_mapping，新模板仅需加JSON条目 |
| Agent仍尝试自行写openpyxl | 绕过强约束 | 步骤文件零裸API代码 + SKILL.md红线 + gate检测 |
| fill_sheet()bug导致全部crash | 误报阻断 | 单元测试覆盖 + --skip-gate紧急逃生（仅限开发调试） |
| 行业映射JSON不全 | 新行业缺映射 | fallback到科目余额表通用映射 + WARNING提示人工补充 |

---

## 七、总结

**一句话**：把142条DT规则从"写在Markdown里等Agent自觉遵守"变成"嵌入Python代码里强制执行"。

**核心改变**：

| 维度 | 改变前 | 改变后 |
|------|--------|-------|
| Agent角色 | 准备数据+自行写openpyxl代码+自行校验 | 准备data_rows→调用fill_sheet()→结束 |
| 规则执行力 | Agent"知道"规则但可选择绕过 | 规则嵌入代码，违反=crash |
| unmerge_all | Agent可调用，破坏表头 | sheet_filler内部_safe_write_cell，表头受保护 |
| 列位 | Agent可硬编码，填错列 | fill_sheet强制col_map，硬编码=assert失败 |
| 业务内容 | Agent填"应付账款"=没填 | business_content_map自动推断 |
| 行业映射 | 无，每次人工判断 | industry_mapping.json自动映射 |
| Gate执行 | Agent可选择不调用 | fill_sheet返回路径自动触发 |
| 18个问题复发概率 | 高（61%无L1约束） | 极低（95%以上规则有L1/L2约束） |

---

## 九、实施状态（2026-05-24）

### ✅ 已完成

| 编号 | 改进项 | 优先级 | 实施内容 | 产出文件 |
|------|--------|--------|---------|---------|
| P0-1 | sheet_filler.py核心模块 | P0 | Agent唯一写入接口，12条DT规则内部强制 | `valuation-common/scripts/sheet_filler.py` |
| P0-2 | business_content_map.py | P0 | 科目编码+结算对象→业务实质自动推断 | `valuation-common/scripts/business_content_map.py` |
| P0-3 | industry_mapping.json | P0 | 房地产/制造/施工行业特殊科目映射 | `valuation-common/scripts/industry_mapping.json` |
| P1-1 | phase_gate.py | P1 | Phase间Gate自动触发，CRITICAL→exit(1) | `valuation-common/scripts/phase_gate.py` |
| P1-2 | gate_validator.py v3.8 | P1 | 新增G2-13~G2-17五个校验项 | `valuation-detail-table/scripts/gate_validator.py` |
| P1-3 | 新增DT-143~DT-150规则 | P1 | 8条新规则覆盖规则空白区 | `valuation-detail-table/RULES.md v3.45` |
| P1-4 | 强制等级标注 | P1 | A/B/C三级标注体系 | `valuation-detail-table/RULES.md v3.45` |
| P2-1 | S2步骤文件升级 | P2 | 4个步骤文件头部升级为sheet_filler接口 | `steps/S2_fill_bs/re/inventory/liability.md` |
| P2-2 | SKILL.md升级 | P2 | DT-128升级+四层强约束架构+Agent职责边界 | `valuation-detail-table/SKILL.md` |
| P2-3 | FLOW.md升级 | P2 | Phase 2即时门控升级为auto_gate_after_fill | `valuation-detail-table/FLOW.md` |

### 📋 新增文件清单

| 文件路径 | 行数 | 说明 |
|---------|------|------|
| `valuation-common/scripts/sheet_filler.py` | ~450 | Agent唯一写入接口 |
| `valuation-common/scripts/business_content_map.py` | ~300 | 业务内容自动推断 |
| `valuation-common/scripts/industry_mapping.json` | ~60 | 行业特殊科目映射 |
| `valuation-common/scripts/phase_gate.py` | ~280 | Phase间Gate自动触发 |

### 📋 修改文件清单

| 文件路径 | 变更内容 |
|---------|---------|
| `valuation-detail-table/scripts/gate_validator.py` | v3.7→v3.8，新增G2-13~G2-17 |
| `valuation-detail-table/RULES.md` | v3.44→v3.45，新增DT-143~DT-150+强制等级标注 |
| `valuation-detail-table/SKILL.md` | DT-128升级+四层强约束架构 |
| `valuation-detail-table/FLOW.md` | Phase 2即时门控升级 |
| `valuation-detail-table/steps/S2_fill_bs.md` | 头部升级为sheet_filler接口 |
| `valuation-detail-table/steps/S2_fill_re.md` | 头部升级为sheet_filler接口 |
| `valuation-detail-table/steps/S2_fill_inventory.md` | 头部升级+DT-145行业映射 |
| `valuation-detail-table/steps/S2_fill_liability.md` | 头部升级+DT-147/148/149 |

### 🔄 后续迭代

| 待实施 | 说明 |
|--------|------|
| 更多规则L3→L1升级 | 将RULES.md中剩余C级规则逐步升级为A级脚本强制 |
| sheet_filler.py实战测试 | 在下一个评估项目中验证fill_sheet()接口的完整覆盖 |
| 行业映射扩展 | 增加更多行业（金融/医药/能源等）的特殊科目映射 |
| 应交税费子项映射扩展 | 补充更多税种子项编码映射 |
