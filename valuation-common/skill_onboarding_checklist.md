# 评估Skill准入检查清单（New Skill Onboarding Checklist）

> **目的**：确保新生成的评估skill从一开始就遵守共享架构，不再出现"遗漏引用共享规则"的问题。
> 版本：v1.0 | 创建日期：2026-05-20

---

## 根因回顾

declaration-table skill 曾因未引用META_RULES.md，导致：
- 备份版本堆积（MR-4备份纪律未生效）
- 多版本副本共存（MR-6单一保存路径未生效）
- 截图文件残留（无清理规则）

**根因**：新skill创建时无强制准入检查，共享规则靠各skill自行引用，遗漏即失效。

---

## 准入检查清单（强制）

每个新建/重大修改的评估skill，**MUST**在SKILL.md写入前逐项确认：

### A. 元规则引用（L0层）

| # | 检查项 | 必须内容 | 验证方法 |
|---|--------|---------|---------|
| A1 | META_RULES引用声明 | SKILL.md标题后第一行：`[元规则] 参照 valuation-common/META_RULES.md（MR-1至MR-15），不可违反` | grep "META_RULES" SKILL.md |
| A2 | 共享纪律规则引用 | 编制类：`> **G0-G13** 共享规则详见 valuation-common/preparation_discipline_rules.md`；审核类：`> **T0-T20** 共享规则详见 valuation-common/audit_discipline_rules.md` | grep "discipline_rules" SKILL.md |
| A3 | 规则编号命名空间注册 | 在META_RULES.md命名空间表中注册skill前缀 | grep "skill名" valuation-common/META_RULES.md |

### B. 关键规则显式声明（即使已引用共享规则，以下规则MUST在SKILL.md中显式出现）

| # | 规则 | 对应共享规则 | SKILL.md中的编号 | 原因 |
|---|------|------------|-----------------|------|
| B1 | 修改前备份≤2版本 | MR-4/G5 | 各skill自定义 | 防止备份堆积 |
| B2 | 覆盖保存+单一保存路径 | MR-6/G6 | 各skill自定义 | 防止多版本副本 |
| B3 | 验证截图用完即删 | MR（无直接对应，新增） | 各skill自定义 | 防止截图残留 |
| B4 | 文件版本承接 | MR-5/G5 | 各skill自定义 | 防止用旧版覆盖新版 |

> **为什么需要显式声明？** 因为Agent在执行时主要读取SKILL.md，如果关键规则仅在引用的外部文件中，Agent可能"知道但未执行"。显式声明确保规则在上下文窗口内。

### C. 文件结构（L1-L3四层架构）

| # | 检查项 | 要求 |
|---|--------|------|
| C1 | SKILL.md | ≤300行，仅含元规则引用+FLOW引用+领域特定规则+版本号 |
| C2 | FLOW.md | L1流程层：纯流转逻辑（输入→操作→输出→校验→流转） |
| C3 | CHECK.md | L3校验层：每Phase后必触发的校验规则 |
| C4 | steps/ | L2操作层：按需加载，每步独立文件 |
| C5 | lessons_learned.md | 踩坑记录 |
| C6 | references/ | 参考文档 |

### D. CHECK绑定（MR-14强制）

| # | 检查项 | 要求 |
|---|--------|------|
| D1 | D型规则有对应CHECK项 | 在对应Phase的C{N}校验中有一行 |
| D2 | R型规则有Phase门控级CHECK | 不通过=禁止流转 |
| D3 | 新增规则无CHECK=advisory | 不作为强制检查项 |

---

## 准入流程

```
新skill创建
  │
  ├─ 1. 确定skill类型（编制类/审核类/辅助类）
  │
  ├─ 2. 选择共享规则引用
  │     编制类 → preparation_discipline_rules.md (G0-G13)
  │     审核类 → audit_discipline_rules.md (T0-T20)
  │     辅助类 → 仅META_RULES.md (MR-1至MR-15)
  │
  ├─ 3. 注册命名空间前缀
  │     在META_RULES.md命名空间表中添加新条目
  │
  ├─ 4. 按模板创建SKILL.md
  │     包含：元规则引用 + 共享规则引用 + B1-B4显式声明 + 领域特定规则
  │
  ├─ 5. 创建L1-L3文件
  │     FLOW.md + CHECK.md + steps/ + lessons_learned.md
  │
  ├─ 6. 执行准入检查清单
  │     逐项确认A1-A3, B1-B4, C1-C6, D1-D3
  │
  └─ 7. 更新skill_index.md
        在编制类/审核类/辅助类表格中添加新skill
```

---

## SKILL.md最小模板（评估类）

```markdown
---
name: valuation-{skill-name}
description: {触发场景描述}
allowed-tools:
disable: false
agent_created: true
---

# {Skill中文名} Skill

[元规则] 参照 valuation-common/META_RULES.md（MR-1至MR-15），不可违反

> **{G0-G13 或 T0-T20}** 共享规则详见 `valuation-common/{对应文件}`，以下仅列出本skill的领域特定规则。

## 触发条件
{描述}

## 输入/输出
{描述}

## 执行纪律

| 编号 | 类别 | 规则 | 适用阶段 | 违反后果 |
|------|------|------|----------|----------|
| **{前缀}-1** | R | 修改前备份（≡MR-4/G5）：备份到`D:\workbuddy`，**最多保留2个备份版本**，超量删除最旧 | 全流程 | 备份堆积 |
| **{前缀}-2** | R | 覆盖保存+单一保存路径（≡MR-6/G6）：只保存到原路径，覆盖保存，禁止创建多版本副本 | 全流程 | 版本混乱 |
| **{前缀}-3** | R | 验证截图用完即删：验证完成后立即删除截图文件，或最多保留最终版1张 | 验证阶段 | 截图残留 |
| **{前缀}-4** | R | 文件版本承接（≡MR-5/G5）：承接上文最近修改版本，不确定时向用户确认 | 全流程 | 用旧版覆盖新版 |

{...领域特定规则...}

## 流程
参照 [FLOW.md](FLOW.md)

## 版本
| 版本 | 日期 | 变更内容 |
|------|------|---------|
| v1.0 | {日期} | 初始版本 |
```

---

## 版本记录

| 版本 | 日期 | 变更 |
|------|------|------|
| v1.0 | 2026-05-20 | 初始版本：基于declaration-table遗漏共享规则的根因分析，建立准入检查清单 |
