# 评估明细表填写 — 流程总览 (FLOW.md)

> 本文件为L1流程层，只描述"先做什么、后做什么"，纯流转逻辑。
> 操作细节详见 `steps/S{N}.md`，校验规则详见 `CHECK.md`。

---

## 🚨 前置判断0：断点恢复检测 [DT-131]

> **⚠️ 本步骤为v3.42新增，DT-130中间数据持久化的配套恢复机制**
> **核心目的**：新对话中，用户只需提供项目文件夹路径（如 `C:\Users\Administrator\Desktop\1-河南平绿`），Agent自动检测缓存并从断点恢复，无需手动描述"继续填写XX项目"。

**触发条件**：用户消息中包含一个**项目文件夹路径**，且该路径下存在 `_dt_cache/` 子目录。

**检测流程**：

```
Skill被触发
  │
  ├── 用户消息中包含文件夹路径？
  │     │
  │     ├── YES → 检查 {路径}/_dt_cache/ 是否存在
  │     │     │
  │     │     ├── _dt_cache/ 存在且有JSON文件 → 🔄 断点恢复模式
  │     │     │     │
  │     │     │     ├── 1. 读取 _dt_cache/ 中所有JSON文件
  │     │     │     ├── 2. 对比 DT-130 缓存清单（见下表），判断缓存完整度
  │     │     │     ├── 3. 输出恢复状态报告（已缓存步骤 vs 待执行步骤）
  │     │     │     ├── 4. 从第一个缺失缓存对应的步骤开始继续执行
  │     │     │     └── 5. 不需要用户确认，直接继续
  │     │     │
  │     │     └── _dt_cache/ 不存在或为空 → ⬇️ 正常流程（前置判断1）
  │     │
  │     └── NO → ⬇️ 正常流程（前置判断1）
  │
  └── 正常流程继续
```

**DT-131缓存完整度判定表**：

| 缓存文件 | 对应步骤 | 缓存存在=可跳过 | 缺失=需执行 |
|----------|---------|---------------|-----------|
| `file_manifest.json` | Step -1.3 | 文件遍历 | 执行文件遍历+持久化 |
| `pdf_extraction_*.json` | Step -1.5 | PDF预提取 | 执行PDF提取+持久化 |
| `pdf_completeness_report.json` | Step -1.6 | PDF完整性报告 | 执行完整性检查+持久化 |
| `subjects.json` | Step 0.2 | 科目余额表解析 | 执行科目余额表解析+持久化 |
| `bs_balances.json` | Step 0.3 | 资产负债表解析 | 执行资产负债表解析+持久化 |
| `data_source_mapping.json` | Step 0.4 | 数据源映射 | 执行映射+持久化 |
| `auxiliary_balance_*.json` | Step 0.5 | 辅助余额表提取 | 执行辅助余额表提取+持久化 |
| `d1d2d3_mapping.json` | Step 0.5a | 三级递进映射 | 执行映射+持久化 |
| `reclassification.json` | Step 0.5a | 重分类映射 | 执行重分类+持久化 |
| `data_classification.json` | Step 0.6 | 数据分类 | 执行分类+持久化 |
| `settings_info.json` | Step 0.7 | 设定信息 | 执行设定信息填写+持久化 |

**恢复状态报告示例**：

```
🔄 断点恢复 — 河南平绿2025评估明细表

项目文件夹：C:\Users\Administrator\Desktop\1-河南平绿
缓存目录：_dt_cache/ (8个JSON文件)

| 步骤 | 缓存文件 | 状态 |
|------|---------|------|
| Step -1.3 文件遍历 | file_manifest.json | ✅ 已缓存 |
| Step -1.5 PDF预提取 | pdf_extraction_*.json (8个) | ✅ 已缓存 |
| Step -1.6 完整性报告 | pdf_completeness_report.json | ✅ 已缓存 |
| Step 0.2 科目余额表 | subjects.json | ✅ 已缓存 |
| Step 0.3 资产负债表 | bs_balances.json | ✅ 已缓存 |
| Step 0.5 辅助余额表 | auxiliary_balance_*.json (6个) | ✅ 已缓存 |
| Step 0.5a 三级映射 | d1d2d3_mapping.json | ❌ 缺失 |
| Step 0.5a 重分类 | reclassification.json | ❌ 缺失 |
| Step 0.6 数据分类 | data_classification.json | ❌ 缺失 |
| Step 0.7 设定信息 | settings_info.json | ❌ 缺失 |

✅ Phase -1 完成，Phase 0 部分完成
⏩ 从 Step 0.5a 继续执行...
```

**DT-131核心约束**：
- **DT-131-1**：Agent检测到项目文件夹下存在 `_dt_cache/` 且含JSON文件时，MUST自动进入断点恢复模式，不输出"是否继续"之类的确认问题，直接加载缓存并继续。
- **DT-131-2**：断点恢复模式下，已缓存的步骤MUST从磁盘加载而非重新执行。每个步骤开始前先检查对应缓存文件是否存在，存在则 `load_cache()` + 跳过，不存在则执行+ `save_cache()`。
- **DT-131-3**：用户只需提供项目文件夹路径一条消息，Agent自动完成全部恢复+继续。**不需要**用户额外输入"继续""恢复""接着做"等指令。
- **DT-131-4**：缓存文件时间戳与源文件修改时间不一致时（源文件比缓存更新），MUST重新执行该步骤并覆盖缓存。检测方式：比较JSON的 `_meta.created_at` 与源文件的 `mtime`。
- **DT-131-5**：Phase 2-5的中间产物（已填写的评估明细表xlsx、gate验证结果等）不纳入缓存检测范围。这些文件在项目文件夹中直接存在，Agent按DT-57文件版本承接原则处理。

---

## 前置判断1：确认执行模式

```
输入文件检查
  │
  └── 有科目余额表/账套？ ── YES ──→ 完整模式（Phase -1 → Phase 0~5）
```

---

## Phase -1: 材料准备与操作提示

| 项目 | 内容 |
|------|------|
| **输入** | 用户发起的填写明细表请求、用户已提供的文件路径（如有） |
| **操作** | 详见 → [S-1_prep.md](steps/S-1_prep.md) |
| **输出** | 材料文件夹路径、完整文件清单、PDF/图片预提取结果、PDF提取完整性报告(DT-108)、必需材料完整性检查结果、**中间数据持久化文件(DT-130)** |
| **校验** | 详见 → [CHECK.md](CHECK.md) C-1 |
| **流转** | 材料齐全 → Phase 0；必需材料缺失 → 输出缺失清单+标注[待核实]+继续执行 |

**核心目的**：确保所有填写所需材料集中到位，从源头杜绝数据源遗漏。

**本阶段自动执行**：材料路径已知时，自动扫描文件夹+提取PDF+校验完整性，无需用户确认。必需材料缺失时输出缺失清单+标注[待核实]+继续执行。

**关键规则**：DT-105/106/107/108/132/133/130（详见CHECK.md C-1）

> **确认点**：材料已集中、文件已识别、PDF/图片已预提取、**PDF提取完整性报告已生成且未提取数=0**、**中间数据已持久化到_dt_cache/（DT-130）**。完整校验项见CHECK.md C-1。

> **🚨 DT-108硬Gate脚本化验证（v3.48新增）**：Phase -1完成后、进入Phase 0前，MUST执行以下命令验证PDF提取完整性：
> ```bash
> python -c "import json; r=json.load(open('_dt_cache/pdf_completeness_report.json')); unextracted=[f for f in r['files'] if f.get('status')!='PASS']; assert len(unextracted)==0, f'DT-108 FAIL: {len(unextracted)}个PDF未成功提取'; print('DT-108 PASS')"
> ```
> **断言失败 → 禁止进入Phase 0。** 此门控为L1层级约束，脚本crash即阻断。

---

## 完整模式流程

### Phase 0: 输入确认与数据源解析

| 项目 | 内容 |
|------|------|
| **输入** | 用户提供的文件列表（评估明细表模板、科目余额表、资产负债表、可选：序时账/收发存/固定资产台账） |
| **操作** | 详见 → [S0_input.md](steps/S0_input.md) |
| **输出** | 执行模式判定、科目余额表末级科目清单、资产负债表各科目期末余额、模板Sheet列表、**科目→数据源映射表(DT-109)**、**科目→结算对象清单(DT-111)**、**数据分类清单(DT-117)**、**D1/D2/D3三级递进映射表(DT-119)**、**重分类映射表(DT-118)**、**设定信息已填写(DT-121)**、**中间数据持久化文件(DT-130)** |
| **校验** | 详见 → [CHECK.md](CHECK.md) C0 |
| **流转** | 完整模式 → Phase 1 |

- [ ] **CCEP-1**：本Phase执行符合C0合规。**完整校验项见CHECK.md C0。**

> **确认点**：模式判定、必需文件齐全、**辅助余额表已提取(DT-111)**、**D1/D2/D3映射已建立(DT-119)**、**重分类映射已生成(DT-118)**、**设定信息已填写(DT-121)**、**中间数据已持久化(DT-130)**。完整校验项见CHECK.md C0。

> **🚨 G0强制门控（v3.48新增）**：Phase 0完成后、进入Phase 1前，MUST执行以下命令验证数据源完整性：
> ```bash
> python3 valuation-detail-table/scripts/gate_validator.py <xlsx_path> --gate G0 --sb-path <科目余额表路径> --bs-path <资产负债表路径>
> ```
> **CRITICAL>0 → 禁止进入Phase 1，必须回退修复数据源问题。** 此门控为L2层级约束，Agent不可绕过。

---

### Phase 1: 结构解析与科目映射

| 项目 | 内容 |
|------|------|
| **输入** | 评估明细表模板、Phase 0解析结果 |
| **操作** | 详见 → [S1_structure.md](steps/S1_structure.md) |
| **输出** | 所有Sheet列定义、科目映射表（科目编码→Sheet→列）、各科目填写策略、**D1/D2/D3映射策略(DT-119)** |
| **校验** | 详见 → [CHECK.md](CHECK.md) C1 |
| **流转** | → Phase 2 |

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C1。

> **确认点**：科目映射表完整性、减值行位置标注、**D1/D2/D3递进映射策略已融入科目映射表(DT-119)**

---

### Phase 2: 数据填写（逐科目逐Sheet）

Phase 2包含多个并行子步骤，按科目类别拆分：

| 子步骤 | 操作文件 | 说明 |
|--------|---------|------|
| 2a | [S2_fill_bs.md](steps/S2_fill_bs.md) | 货币资金+固定资产+无形资产+长期待摊费用等 |
| 2b | [S2_fill_re.md](steps/S2_fill_re.md) | 往来科目（应收/应付/预收/预付/其他应收应付），含DT-46业务内容与日期严禁混淆 |
| 2c | [S2_fill_inventory.md](steps/S2_fill_inventory.md) | 存货科目填写 |
| 2d | [S2_fill_liability.md](steps/S2_fill_liability.md) | 负债类科目填写 |

**执行顺序**：2a → 2b → 2c → 2d（按科目依赖关系串行）

**⚠️ Phase 2即时门控（v3.48+：sheet_filler + 即时格式微调 + auto_gate_after_fill）**：
每填写完一个Sheet后MUST通过`auto_gate_after_fill()`自动触发G1门控。sheet_filler内部已执行12条DT规则强制校验 + 即时格式微调（DT-152/DT-153）。详细门控项见CHECK.md C2。

```python
# v3.45+: Phase 2填写后自动Gate触发
from sheet_filler import fill_sheet, prepare_data_rows
from phase_gate import auto_gate_after_fill

result = fill_sheet(ws=ws, sheet_id=sheet_id, data_rows=data_rows, wb=wb)
gate_result = auto_gate_after_fill(filepath, sheet_id, result)
# gate_result.passed=False → sys.exit(1) → 流程阻断
```

**往来科目特殊子步骤**（2b完成后按需执行）：

| 子步骤 | 操作文件 | 说明 |
|--------|---------|------|
| 2e | [S2_seq_verify.md](steps/S2_seq_verify.md) | 序时账核实往来科目发生日期与业务内容（DT-51-T55/DT-60完整流程） |

| 项目 | 内容 |
|------|------|
| **输入** | 科目余额表明细数据、序时账（可选）、收发存明细（可选）、固定资产台账（可选） |
| **输出** | 各科目Sheet数据已填入、减值准备已填入对应"减："行、差异项已标注 |
| **校验** | 详见 → [CHECK.md](CHECK.md) C2 |
| **流转** | → Phase 3 |

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C2。
- [ ] **[DT-120] 🚨 G2-12 smart_insert_row工具调用检测**：所有明细表Sheet已完成G2-12检测，CRITICAL=0

---

### Phase 3: 公式修复与格式修复（格式集中处置区，DT-112 v1.4优化后）

| 项目 | 内容 |
|------|------|
| **输入** | 已填写数据的评估明细表（Phase 2已完成数据写入+即时格式微调+Step3.5前置格式统一） |
| **操作** | 详见 → [S3_format.md](steps/S3_format.md) |
| **输出** | 所有公式修复完成、所有深度格式修复完成、空行删除、打印范围精确 |
| **校验** | 详见 → [CHECK.md](CHECK.md) C3 |
| **流转** | → Phase 3.5（格式门控）→ Phase 4.5（自动化验证门控） |

> **🎯 DT-112格式集中处置原则（v1.4优化后）**：Phase 2已通过smart_insert_row()即时格式微调（行高/数字格式/对齐）+ fill_sheet() Step3.5前置格式统一，3项高频格式问题在Phase 2已解决。Phase 3仅需处理深度格式修复（边框扫描DT-82/合并单元格验证DT-83/结构行A列居中DT-84/合计行下方清理DT-78②/公式列覆写检查DT-67）和公式修复。Phase 3从"修复+验证(2-3次闭环)"降级为"纯验证(1次)"。

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C3。
- [ ] **[DT-112] 格式集中处置已执行**

> **确认点**：公式修复即时验证（DT-2/DT-24）、格式修复即时验证（DT-3）、**深度格式全量扫描通过**

---

### Phase 3.5: 格式门控（Phase 3完成后必须执行，未通过禁止进入Phase 4）

> **本步骤为Phase 3格式集中处置的验证门控（DT-112）。** 检查Phase 3中应完成的所有深度格式修复是否已执行。

| 项目 | 内容 |
|------|------|
| **输入** | Phase 3完成后的评估明细表 |
| **操作** | 运行 `gate_validator.py <xlsx_path> --gate G1-Format` 格式门控验证 |
| **输出** | 格式验证报告（通过/警告/严重问题） |
| **校验** | 格式门控检查项 |
| **流转** | 全部通过 → Phase 4.5；有严重问题 → 回退Phase 3修复 |

**格式门控检查项（从Phase 2即时门控拆分而来）：**

- [ ] **[DT-76] 增值额/增值率列number_format已全表扫描修正**
- [ ] **[DT-77] 所有数据行行高与模板默认值一致**
- [ ] **[DT-82①] 数据行首行无空白跳过**
- [ ] **[DT-82②] 空白数据行保留thin边框+公式**
- [ ] **[DT-83] 多行表头合并单元格与模板一致**
- [ ] **[DT-84] 合计/减值/小计行A列居中对齐**
- [ ] **[DT-67] 公式列（J/K）未被数值覆写**

> **确认点**：格式门控全部通过才允许进入Phase 4.5

### Phase 4.5: 自动化验证门控

| 项目 | 内容 |
|------|------|
| **输入** | Phase 3完成后的评估明细表 |
| **操作** | 运行 `validate_detail_table()` + `gate_validator.py` 自动化验证脚本 |
| **输出** | 验证报告（通过/警告/严重问题） |
| **校验** | 详见 → [CHECK.md](CHECK.md) C4.5 |
| **流转** | 全部通过 → Phase 4；有警告 → Phase 4（建议修复）；有严重问题 → 回退Phase 3修复后重新验证 |

**v3.17 gate_validator.py增强**：
```bash
# G2科目级校验（Phase 2完成后，需bs-path和sb-path参数）
python3 valuation-detail-table/scripts/gate_validator.py <xlsx_path> --gate G2 --bs-path <bs_path> --sb-path <sb_path>
# G2新增：DT-87其他流动资产行数门控/DT-89递延所得税零余额排除/DT-103数据源完整性
# G2新增[DT-120]：G2-12 smart_insert_row工具调用检测（4项格式特征，CRITICAL=禁止进入Phase 3）

# G3勾稽级校验（Phase 4完成后）
python3 valuation-detail-table/scripts/gate_validator.py <xlsx_path> --gate G3 --bs-path <bs_path>
# G3新增：DT-18减值准备行评估值方向/DT-86汇总表跨sheet引用行号校验
```

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C4.5。

> **确认点**：自动化验证100%通过才允许继续

---

### Phase 4a: BS数据校验（2-分类汇总表 I/J 列比对）

| 项目 | 内容 |
|------|------|
| **输入** | Phase 4.5验证通过的评估明细表、`_dt_cache/bs_balances.json` |
| **操作** | 详见 → [S4_bs_verify.md](steps/S4_bs_verify.md) |
| **输出** | 2-分类汇总表 I 列（BS财务报表期末余额）、J 列（差异校对公式）、差异清单 |
| **校验** | ~~详见 → [CHECK.md](CHECK.md) C4a~~ (暂无，后续补充) |
| **流转** | 全部写入成功 → Phase 4；BS数据缓存缺失 → 回退 Phase 0 Step 0.3 |

**核心逻辑**：
1. 从 `_dt_cache/bs_balances.json` 加载 BS 各科目期末余额
2. 遍历 `2-分类汇总` 每一行，根据科目名称匹配 BS 数据
3. 将BS期末余额写入隐藏`_BS对照`结构表；I列写入链接`_BS对照`的公式，J列写入差异公式（自动计算差异）
4. 差异 > 1 元的科目输出差异清单，供 Phase 4 勾稽核对使用

- [ ] **CCEP-1**：BS数据已写入隐藏`_BS对照`，2-分类汇总I列公式链接完整
- [ ] **CCEP-2**：J列差异公式已写入，COM重算后差异值正确
- [ ] **CCEP-3**：差异清单已输出

> **确认点**：I列和J列数据完整性确认，差异清单供Phase 4勾稽分析

---

### Phase 4: 勾稽核对

| 项目 | 内容 |
|------|------|
| **输入** | Phase 4.5验证通过的评估明细表、资产负债表 |
| **操作** | 详见 → [S4_reconcile.md](steps/S4_reconcile.md) |
| **输出** | 三级勾稽核对结果（明细表→汇总表→分类汇总表→资产负债表）、差异处理说明、**差额推论映射结果汇总(DT-117)**、**重分类映射表(DT-118)** |
| **校验** | 详见 → [CHECK.md](CHECK.md) C4 |
| **流转** | 勾稽100%通过 → Phase 4链接检查 → Phase 5；勾稽不符 → 修正后重新勾稽 |

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C4。

> **确认点**：勾稽核对必须100%覆盖（DT-4），不符=禁止交付

---

### Phase 4 链接检查: 隐藏汇总表联动检查

| 项目 | 内容 |
|------|------|
| **输入** | Phase 4勾稽通过的评估明细表 |
| **操作** | 详见 → [S4_linkage.md](steps/S4_linkage.md) |
| **输出** | 隐藏汇总表引用验证结果、辅汇总表vs可见汇总表逻辑确认 |
| **校验** | 详见 → [CHECK.md](CHECK.md) C4 |
| **流转** | 通过 → Phase 5.5 |

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C4.5。

> **确认点**：DT-61/DT-62隐藏汇总表引用正确性

---

### Phase 5.5: 反思固化门控

| 项目 | 内容 |
|------|------|
| **输入** | 本次执行中遇到的所有问题、用户反馈、纠错过程 |
| **操作** | 问题回溯→归类判断→固化T规则或Phase步骤→写入SKILL.md |
| **输出** | 反思固化自检清单（全部✅） |
| **校验** | 详见 → [CHECK.md](CHECK.md) C5.5 |
| **流转** | 全部✅ → Phase 5 |

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C4.5。

> **确认点**：所有问题已固化，无用户指出但未固化的反馈。未通过=禁止交付

---

### Phase 5: 清理与交付

| 项目 | 内容 |
|------|------|
| **输入** | 勾稽通过+反思固化完成的评估明细表 |
| **操作** | 详见 → [S6_deliver.md](steps/S6_deliver.md) |
| **输出** | 最终交付文件：`{项目名}{年度}_评估明细表_已填写.xlsx` + **推论映射汇报(DT-117)** |
| **校验** | 详见 → [CHECK.md](CHECK.md) C5 |
| **硬门控** | **🚨 G3-11（DT-110隐藏校验）不通过=禁止交付；G2-10（DT-117待确认映射未解决）不通过=禁止交付；G2-12（DT-120 smart_insert_row工具调用检测）不通过=禁止交付** |

- [ ] **CCEP-1**：本Phase执行符合C0合规。完整校验项见CHECK.md C5。
- [ ] **CCEP-2**：G3-11隐藏操作校验已通过，G2-10/G2-12门控已通过
- [ ] **CCEP-3**：推论映射汇报已输出(DT-117)

> **确认点**：隐藏操作已集中执行(DT-110)、COM重算保存已完成、G3-11 Gate已通过

> **🚨 硬约束**：交付前**必须**执行 `gate_validator.py --gate G3` 并确认G3-11项全部通过（CRITICAL违规=0），否则**禁止交付**

---

## 交付物定义

| 模式 | 交付物 | 格式 |
|------|--------|------|
| **完整模式** | 评估明细表 | Excel (.xlsx) |
