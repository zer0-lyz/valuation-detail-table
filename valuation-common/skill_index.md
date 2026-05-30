# 评估Skill索引与数据流向

> 本文档记录各评估skill之间的数据流向和调用关系，便于理解skill间协作逻辑。
> 2026-05-18 更新：全部skill已完成L1-L3分层改造（FLOW.md + steps/ + CHECK.md + SKILL.md精简）

---

## L0元规则

所有评估skill共享 `valuation-common/META_RULES.md`（MR-1至MR-16），不可违反。

---

## Skill总览

### 编制类Skill（生成底稿/文档）

| Skill | 版本 | 输入 | 输出 | 关键依赖 | L1-L3文件 |
|-------|------|------|------|---------|----------|
| valuation-detail-table | v3.2 | 科目余额表+资产负债表 | 评估明细表(xlsx) | 无 | FLOW+11steps+CHECK |
| valuation-asset-based-workpaper | v2.7 | 评估明细表+成本法底稿模板 | 成本法底稿(xlsx) | ← detail-table | FLOW+14steps+CHECK |
| valuation-vouching-extract | v3.4 | 评估明细表+序时账 | 抽凭底稿(xlsx) | ← detail-table | FLOW+9steps+CHECK |
| valuation-vouching-journal | v1.0 | 成本法底稿+序时账 | 抽凭序时账明细(xlsx) | ← vouching-extract | SKILL.md |
| valuation-supplemental-checklist | v1.1 | 评估明细表 or 资产负债表 | 待补资料清单(xlsx) | ← detail-table | FLOW+9steps+CHECK |
| valuation-declaration-table | v1.1 | 评估明细表 | 评估申报表(xlsx) | ← detail-table | SKILL.md |
| valuation-company-basic-info | v3.1 | 企查查MCP+用户提供资料 | 评估报告章节(docx) | 无 | FLOW+3steps+CHECK(三模块) |
| valuation-file-organize | v1.0 | 评估项目文件 | 归档目录结构 | ← all | FLOW+steps+CHECK |

### 审核类Skill（审核已有底稿）

| Skill | 版本 | 审核对象 | 评估方法 | 共享规则 | L1-L3文件 |
|-------|------|---------|---------|---------|----------|
| valuation-dcf-workpaper | v1.0 | 收益法底稿(xlsx) | DCF/FCF/WACC | audit+META | FLOW+11steps+CHECK |
| valuation-equipment-workpaper | v1.0 | 设备评估底稿(xlsx) | 成本法(重置成本) | audit+META | FLOW+11steps+CHECK |
| valuation-land-use-right-workpaper | v1.0 | 土地使用权底稿(xlsx) | 市场法/基准地价法/成本逼近法 | audit+META | FLOW+12steps+CHECK |
| valuation-market-workpaper | v1.0 | 市场法底稿(xlsx) | 上市公司比较法/VM指数法 | audit+META | FLOW+11steps+CHECK |
| valuation-real-estate-workpaper | v1.0 | 房地产底稿(xlsx) | 收益法/市场法/成本法 | audit+META | FLOW+11steps+CHECK |
| asset-report-reviewer | v1.0 | 评估报告全文 | 全方法 | audit+META | FLOW+9steps+CHECK |

---

## Skill文件结构（统一四层架构）

```
{skill-name}/
├── SKILL.md             ← 精简版（≤300行）：元规则引用+FLOW引用+领域特定规则+版本号
├── FLOW.md              ← L1流程层：纯流转逻辑（输入→操作→输出→校验→流转）
├── CHECK.md             ← L3校验层：每Phase后必触发的校验规则
├── steps/               ← L2操作层：按需加载，每步独立文件
│   ├── S0_xxx.md
│   ├── S1_xxx.md
│   └── ...
├── scripts/             ← 独立脚本（从SKILL.md中抽取）
│   └── README.md
├── references/          ← 参考文档
├── assets/              ← 资源文件
└── lessons_learned.md   ← 踩坑记录
```

---

## 数据流向图

```
                    ┌─────────────────┐
                    │  科目余额表       │
                    │  资产负债表       │
                    └────────┬────────┘
                             │
                    ┌────────▼────────┐
                    │  valuation-      │
                    │  detail-table   │
                    │  (评估明细表填写) │
                    └────────┬────────┘
                             │
              ┌──────────────┼──────────────┐
              │              │              │
     ┌────────▼───────┐ ┌───▼──────────┐ ┌──▼──────────────────┐
     │  valuation-     │ │  valuation-  │ │  valuation-          │
     │  asset-based-   │ │  vouching-   │ │  supplemental-       │
     │  workpaper      │ │  extract     │ │  checklist           │
     │ (成本法底稿编制) │ │ (抽凭底稿)   │ │ (待补资料清单)        │
     └────────┘        └────────┘ └──────────────────────────────┘

     ┌─────────────────────────────────────────────────────────┐
     │                  审核类Skill（独立并行）                    │
     │                                                         │
     │  valuation-dcf-workpaper ←── 收益法底稿(xlsx)           │
     │  valuation-equipment-workpaper ←── 设备评估底稿(xlsx)     │
     │  valuation-land-use-right-workpaper ←── 土地底稿(xlsx)   │
     │  valuation-market-workpaper ←── 市场法底稿(xlsx)          │
     │  valuation-real-estate-workpaper ←── 房地产底稿(xlsx)     │
     │  asset-report-reviewer ←── 评估报告(docx/pdf)            │
     └─────────────────────────────────────────────────────────┘

     ┌─────────────────────────────────────────────────────────┐
     │               valuation-company-basic-info               │
     │  模块A: 公司基本信息（docx-js → 建议迁移python-docx）    │
     │  模块B: 宏观经济与行业分析（python-docx）                  │
     │  模块C: 业务与财务分析（python-docx）                     │
     └─────────────────────────────────────────────────────────┘
```

---

## 关键数据传递规范

### 评估明细表 → 成本法底稿
- **传递内容**：评估明细表中各科目的账面价值、评估价值
- **关键列位**：账面价值列（通常E/F列）、评估价值列（通常G/J列）
- **注意事项**：明细表插入行后须检查隐藏汇总表引用（T61/T62规则）

### 评估明细表 → 抽凭底稿
- **传递内容**：往来科目结算对象、发生金额
- **关键匹配**：明细表科目名称 → 序时账科目编码（T32规则：编码体系不同）

### 评估明细表 → 待补资料清单
- **传递内容**：科目列表、往来结算对象
- **联动规则**：已有抽凭底稿时简化函证/抽凭描述（CL12规则）

### 成本法底稿 → 索引号一致性
- **传递内容**：步骤复核表索引号 → 过程表索引号 → 抽凭底稿sheet名
- **规则**：XXXX/2=凭证，XXXX/3=合同，XXXX/4=函证（W17规则）

---

## 共享资源

| 资源 | 路径 | 用途 |
|------|------|------|
| L0元规则 | `valuation-common/META_RULES.md` | MR-1至MR-16，所有skill通用 |
| 审核类共享纪律规则 | `valuation-common/audit_discipline_rules.md` | T0-T20通用审核执行纪律 |
| 审核类共享工作流 | `valuation-common/audit_workflow_framework.md` | Phase 0-7通用框架 |
| 编制类共享纪律规则 | `valuation-common/preparation_discipline_rules.md` | G0-G13通用编制执行纪律 |
| 共享脚本目录 | `valuation-common/scripts/` | 可复用的通用Python脚本 |
| company-basic-info迁移计划 | `valuation-common/company_basic_info_migration_plan.md` | docx-js→python-docx方案 |
| 新Skill准入检查清单 | `valuation-common/skill_onboarding_checklist.md` | 新建skill必须通过的检查清单 |

---

_本文件由skill优化自动生成，最后更新：2026-05-18_
