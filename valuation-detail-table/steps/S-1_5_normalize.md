# S-1.5: 财务资料标准化 (DT-NORM)

> **📋 新增 v1.0 (2026-05-29)**：整合 financial-normalizer 工具，将原始财务文件标准化后再进入 Phase 0。
> **📋 DT规则引用 (RULES.md)**：DT-130(中间数据持久化)、DT-131(断点恢复)

## 定位

Phase -1.5 位于 **Phase -1（材料准备）之后、Phase 0（输入解析）之前**，作为预处理步骤，
使用 financial-normalizer 的 guess+apply 管道将原始财务文件标准化为统一格式。

**核心价值**：
- 消除科目余额表、资产负债表中因列名/格式差异导致的解析失败
- 标准化后的 JSON 数据直接供给 Phase 0，替代内置解析器
- 标准化失败不阻断流程 — Phase 0 自动降级到内置解析器

## 输入

- 项目文件夹路径（含已在 Phase -1 集中放置的财务文件）
- `_dt_cache/`（如有上一个对话的缓存）

## 操作

### Step -1.5.1 发现财务文件

自动扫描项目文件夹，按文件名关键词识别财务资料类型：

| 文件类型 | 关键词模式 | 优先级 |
|---------|-----------|-------|
| 科目余额表 | `科目余额`、`余额表`、`试算平衡` | 🔴 最高 |
| 资产负债表 | `资产负债表`、`财务报表` | 🔴 最高 |
| 序时账 | `序时账`、`明细账`、`凭证一览表` | 🟡 可选 |
| 固定资产台账 | `固定资产`、`资产台账` | 🟢 一般 |

### Step -1.5.2 检查 normalizer 可用性

检测 `financial-normalizer/` 核心模块是否可导入：
- ✅ 可用 → 执行标准化
- ❌ 不可用 → 跳过 Phase -1.5，Phase 0 使用内置解析器

### Step -1.5.3 标准化科目余额表

1. 调用 `detector.guess(file)` → 自动检测 Sheet、表头行、列映射
2. 自动确认映射（`confirmed: true`）
3. 调用 `mapper.apply(mapping)` → 执行标准化转换
4. 格式转换 → valuation-detail-table 的 subject 格式
5. 保存到 `_dt_cache/subjects_normalized.json`
6. **同时写入** `_dt_cache/subjects.json` 别名（Phase 0 兼容）

### Step -1.5.4 标准化资产负债表

1. 同上流程，调用 normalizer 对 BS 文件执行标准化
2. 格式转换 → valuation-detail-table 的 items 格式
3. 保存到 `_dt_cache/bs_normalized.json`
4. **同时写入** `_dt_cache/bs_balances.json` 别名（Phase 0 兼容）

### Step -1.5.5 序时账标准化（预留）

当前序时账处理仍在 Phase 3 按原流程执行，此步骤为预留扩展点。

## 输出

| 文件 | 内容 | 消费方 |
|------|------|-------|
| `_dt_cache/subjects_normalized.json` | 标准化科目余额表 | Phase 0 Step 0.2 |
| `_dt_cache/subjects.json` | 别名（与旧缓存兼容） | Phase 0 Step 0.2 |
| `_dt_cache/bs_normalized.json` | 标准化资产负债表 | Phase 0 Step 0.3 |
| `_dt_cache/bs_balances.json` | 别名（与旧缓存兼容） | Phase 0 Step 0.3 |

## 异常处理

- normalizer 不可用 → 跳过，Phase 0 降级到内置解析器
- 标准化部分失败 → 成功部分写入缓存，失败部分由 Phase 0 降级处理
- 未发现财务文件 → 跳过，Phase 0 按原流程执行

## 流转

- 标准化成功 → Phase 0（优先使用 normalizer 输出）
- 标准化失败/跳过 → Phase 0（使用内置解析器降级）

## 数据质量门禁

Phase -1.5 内置自动质量检测，防止错误数据污染缓存：

| 检测项 | 触发条件 | 行为 |
|-------|---------|------|
| 零余额率检测 | >99% 科目余额为零 且 所有ending_debit/ending_credit为零 | 拒绝缓存, 输出错误原因, Phase 0 降级 |
| BS 有效余额检测 | 所有条目 ending_balance 为零/NaN | 拒绝缓存, Phase 0 降级 |
| 双行表头检测 | 上述条件触发 + 提示"疑似双行表头未识别" | 建议使用 source_header_parser |

**常见场景**：
- 双行合并表头（如"期初余额"在第5行，"借方/贷方"在第6行）→ normalizer 可能漏识方向列 → 质量门禁拒绝
- 多Sheet文件（如资产负债表+利润表在同一xlsx）→ normalizer 可能选错Sheet → 质量门禁拒绝

Phase 0 对 `subjects.json` 旧缓存也会做同样的质量检查。
