---
name: valuation-detail-table
description: 根据科目余额表和资产负债表数据，自动填写资产评估明细表。触发场景：(1)用户提供评估明细表模板和科目余额表要求填写；(2)需要对已填写的评估明细表进行勾稽核对；(3)需要用序时账核实往来科目发生日期并同步更新成本法底稿；(4)用户提供项目文件夹路径且该路径下存在_dt_cache/子目录时自动触发断点恢复(DT-131)。覆盖范围：流动资产（货币资金、应收账款、预付款项、其他应收款、存货等）、非流动资产（固定资产、无形资产、长期待摊费用等）、流动负债（短期借款、应付账款、预收款项、其他应付款等）、非流动负债（长期借款等）、减值准备。输出：评估明细表Excel。
allowed-tools:
disable: false
---

# 评估明细表填写Skill

## ⚠️ 执行模式

```bash
# Phase 0: 输入确认与数据源解析（自动编排）
python scripts/dt_runner.py --phase 0 --project <项目文件夹路径>

# 缓存完整性检查
python scripts/dt_runner.py --phase cache --project <项目文件夹路径>
```

### 三层读取架构

| 层级 | 何时读取 | 读取内容 | 字符量 |
|------|---------|---------|--------|
| **L0 启动层** | 每次执行一次 | 本文件（SKILL.md概要） | ~2K |
| **L1 Phase层** | 进入新Phase时 | 当前Phase对应的Step文件 | 10-48K |
| **L2 按需层** | 遇到特定科目 | RULES.md中该科目专属规则 | 0-20K |

### Phase→Step文件映射

| Phase | 名称 | Step文件 | 核心DT规则 | 可跳过？ |
|-------|------|---------|-----------|---------|
| **Phase -1** | 材料准备 | S-1_prep.md | 8条 | ❌ 不可跳过 |
| Phase 0 | 输入确认 | S0_input.md | 11条 | ❌ 不可跳过 |
| Phase 1 | 结构映射 | S1_structure.md | 5条 | ❌ 不可跳过 |
| Phase 2a | 资产填写 | S2_fill_bs.md | 24条 | ❌ 不可跳过 |
| Phase 2b | 往来填写 | S2_fill_re.md | 15条 | ❌ 不可跳过 |
| Phase 2c | 存货填写 | S2_fill_inventory.md | 8条 | ❌ 不可跳过 |
| Phase 2d | 负债填写 | S2_fill_liability.md | 24条 | ❌ 不可跳过 |
| **Phase 3** | **序时账查阅** | **S3_journal_extract.md** | **12条** | **⚠️ 仅两种情况可跳过** |
| Phase 4 | 格式修复 | S4_format.md | 23条 | ❌ 不可跳过 |
| **Phase 4a** | **BS数据校验** | **S4_bs_verify.md** | **6条(v1.2)** | **❌ 不可跳过** |
| Phase 5 | 勾稽核对 | S5_reconcile.md | 9条 | ❌ 不可跳过 |
| **Phase QA** | **自动验收质检** | **quality_assurance.py** | **6项检查** | **❌ 不可跳过（`--phase all`自动触发）** |
| Phase 6 | 清理交付 | S6_deliver.md | 14条 | ❌ 不可跳过 |

**Phase 3 跳过条件（DT-161）**：①未提供序时账 ②用户明确要求不填写发生日期。
**Phase 4a 核心价值**：将BS期末余额写入2-分类汇总表I列(硬编码)+J列公式 =I-E（E=账面价值），实现可视化逐科目校对。
**Phase QA 核心价值**：在全部Phase执行完成后自动启动6维度验收检查（报表校对/字段齐全/汇总校验/空白表隐藏/格式完整性/固定资产分类），最多3轮自动修复+重检，超限转人工。

### 手动执行模式（备选）

如dt_runner.py不可用，Agent MUST按Phase逐个Read对应Step文件，不得全量加载。

## 核心原则

阶段划分是内部质量控制手段，不是外部交互节点。全部Phase必须依次执行，中间不等待用户确认。

## 🚨 最高优先级红线（违反=禁止交付）

| 红线 | 规则 | 后果 |
|------|------|------|
| 🚨R1 | **数据不匹配即停**：禁止AI自行处理不明数据 | 零幻觉底线突破 |
| 🚨R2 | **勾稽核对100%覆盖**（DT-4） | 数据不可信 |
| 🚨R26 | **禁止下派执行**（DT-122） | 产出无效 |
| 🚨R35 | **禁止硬编码列号**（DT-136）：MUST通过sheet_col_map.json | 列偏移=合计为0 |
| 🚨R38 | **BS解析后强制自校验**（DT-139）：资产=负债+权益 | 全链路污染 |
| 🚨R40 | **Poppler环境前置检查**（DT-142） | 扫描件无法提取 |
| 🚨R41 | **汇总表禁止录入数据**（DT-182）：汇总Sheet数据区域MUST仅含公式 | 硬编码=公式链断裂=底稿作废 |

> 完整红线列表（R1~R40）详见 [RULES.md](RULES.md)。

## 四层强约束架构

| 层级 | 物理载体 | 约束力 | Agent可绕过？ |
|------|---------|--------|-------------|
| **L1 脚本强制** | sheet_filler.py / gate_validator.py / phase_gate.py | 违反→crash/raise | ❌ 不可能 |
| **L2 Gate门控** | gate_validator.py (17项校验) + phase_gate.py | 违反→流程阻断 | ❌ 不可能 |
| **L3 规则文字** | RULES.md (DT-0~DT-213) | 依赖Agent自觉 | ⚠️ 可绕过 |
| **L4 流程硬卡** | Phase间强制Gate序列 | 违反→sys.exit(1) | ❌ 不可能 |

## Agent职责边界

| Agent被允许 | Agent被禁止 |
|-----------|-----------|
| 准备data_rows数据 | 直接import openpyxl写入ws（DT-160） |
| 调用fill_sheet()写入 | 调用unmerge_all() |
| 调用prepare_data_rows()组织数据 | 硬编码列号写入（DT-136） |
| 调用auto_gate_after_fill()触发Gate | 跳过Gate校验 |
| 读取JSON缓存文件 | 绕过行业映射 |
| 用openpyxl加载/保存/获取ws对象 | 用openpyxl直接写金额数据到ws（DT-160） |

## 执行透明度披露（交付时必须包含）

```markdown
## 执行情况摘要
| Phase | 执行状态 | 跳过步骤（如有） | 关键发现 |
|-------|---------|-----------------|---------|
| Phase 0 | ✅/❌ | — | — |
| Phase 1 | ✅/❌ | — | — |
| Phase 2 | ✅/❌ | — | — |
| Phase 3 | ✅/❌ | 跳过原因 | — |
| Phase 4 | ✅/❌ | — | — |
| Phase 4a | ✅/❌ | — | 差异N项 |
| Phase 5 | ✅/❌ | — | 勾稽通过/不符 |
| Phase QA | ✅/❌ | — | 验收通过/失败项 |
| Phase 6 | ✅/❌ | — | — |
```

## 版本

v3.65

## 相关文件

| 文件 | 用途 |
|------|------|
| [FLOW.md](FLOW.md) | 流程总览、流转逻辑 |
| [RULES.md](RULES.md) | 完整DT规则库、引用映射 |
| [CHECK.md](CHECK.md) | 校验清单、红线校验 |
| [CHANGELOG.md](CHANGELOG.md) | 变更记录 |
| [字段定义总表.md](字段定义总表.md) | 50个Sheet字段内容规范（2026-05-29） |
| [lessons_learned.md](lessons_learned.md) | 复盘经验 |
| [STRONG_CONSTRAINT_PLAN.md](STRONG_CONSTRAINT_PLAN.md) | 强约束架构设计 |
