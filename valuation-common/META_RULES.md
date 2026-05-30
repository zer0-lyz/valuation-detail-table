# META_RULES.md — 评估Skill元规则（L0层）

> 所有评估类skill必须遵守，不可违反。
> 创建日期：2026-05-18 | 版本：v2.5

---

## MR-1 数据不匹配即停

明细表科目在序时账无法匹配时，**MUST**暂停提请使用人判断，**禁止**AI凭幻觉自行处理。

## MR-2 联动检查

修改明细表后**MUST**同步检查三张隐藏汇总表：
- `3-辅-流动资产汇总`
- `3-9-辅-存货汇总`
- `8-减值准备汇总表`

## MR-3 格式规范

所有输入内容默认字体 **Times New Roman**，字号 **10号**（10pt）。
中文字段可使用宋体，但仍以10号为默认字号。
写入数据时须显式设置 `Font(name='Times New Roman', size=10)`。

## MR-4 备份纪律

修改前**MUST**备份到 `D:\workbuddy`，最多保留2个版本。

## MR-5 文件版本承接

更新文件时**MUST**承接上文最近一个修改版本（即本次会话中已被操作/保存的文件），不可重新从其他路径检索文件作为修改基础。如果不确定哪个版本是最新版，**MUST**向用户确认后再操作。

## MR-6 单一保存路径

更新后的文件只保存到原路径（即文件被读取时的路径），**禁止**在多个位置保存副本。

## MR-7 零幻觉原则

每一条判断/填写/审核意见**MUST**有源文件中的具体数据支撑（单元格位置、数值、公式），**禁止**基于推测、假设、"应该是"做出判断。不确定的数据标注"待核实"，不编造。

## MR-8 反思固化强制门控

每次完成修改/交付成果前，**MUST**主动回顾本次操作中遇到的所有问题，固化为具有强执行力的操作规范写入SKILL.md和lessons_learned.md。**禁止**等用户提醒，禁止以"记录了"代替"更新了"。未完成反思固化 = 禁止交付。

## MR-9 异常即停原则

- 匹配失败 → 停止 → 报告 → 等待指令
- 数据矛盾 → 停止 → 分析根因 → 提出修复计划 → 等待确认
- 前置条件缺失 → 停止 → 列出缺失项 → 等待补充

## MR-10 保密脱敏

skill目录下所有文件及子目录中，**严禁**出现可识别项目/客户的信息：
1. 项目/客户/被评估单位真实名称 → 使用"项目A/B/C"编号
2. 底稿真实文件名 → 使用"某项目-底稿.xlsx"
3. 编制人姓名/工号 → 使用"某评估师"
4. 文件完整路径 → 使用"某路径/某文件.xlsx"
5. 脚本输出文件及目录名不得包含客户/项目名称
6. 可比公司真实名称在经验教训中使用"某可比公司"替代

**写入时即脱敏**，完成后必须删除scripts/下所有非.py运行产出文件，仅保留脚本本身。

## MR-11 抽凭规则

- 应交税费/职工薪酬期末余额为0但有发生额**MUST**抽凭
- 选取距基准日较近凭证
- 个人所得税可免抽凭

## MR-12 步骤复核表未执行步骤隐藏

步骤复核表中未执行的步骤行**MUST**隐藏（`ws.row_dimensions[r].hidden = True`），不可仅标记"不适用"而保留可见。

---

## MR-13 规则注册协议（Rule Registration Protocol, RRP）

> **目的**：确保新知识点/规则可系统性地融入原有架构而不打乱结构。未注册的规则=建议，已注册的规则=纪律。

每条新增领域特定规则**MUST**填写以下注册模板，方可生效：

| 字段 | 必填 | 说明 |
|------|------|------|
| **Rule ID** | ✅ | 格式：`{skill前缀}-{编号}`，如 DT-46、EQ-23、RE-22 |
| **分类** | ✅ | D=纪律型(必须遵守)/O=操作型(实现细节→迁移至step文件)/R=红线型(违反=禁止交付) |
| **规则内容** | ✅ | 具体执行纪律描述 |
| **适用Phase** | ✅ | 如"Phase 2-3" |
| **关联step文件** | ✅ | O型必填，如`steps/S2_seq_verify.md` |
| **关联CHECK项** | ✅ | D/R型必填，如"C2-3"；R型必须有Phase门控级CHECK |
| **来源** | ✅ | 项目经验/错误纠正/标准更新/使用人反馈 |
| **影响skill** | ⚠️ | 如需同步更新其他skill，列出skill名和对应规则编号 |

### 规则编号命名空间

各skill的领域特定规则使用唯一前缀，**禁止跨skill重号**：

| Skill | 前缀 | 共享规则引用 |
|-------|------|------------|
| asset-based-workpaper | **AB** | G0-G13 |
| detail-table | **DT** | G0-G13 |
| dcf-workpaper | **DC** | T0-T13,T15-T20 |
| market-workpaper | **MK** | T0-T13,T15-T20 |
| real-estate-workpaper | **RE** | T0-T13,T15-T20 |
| equipment-workpaper | **EQ** | T0-T13,T15-T20 |
| land-use-right-workpaper | **LU** | T0-T13,T15-T20 |
| supplemental-checklist | **CL** | G0-G13 |
| vouching-extract | **VE** | G0-G13 |
| company-basic-info | **CI** | G0-G13 |
| asset-report-reviewer | **AR** | T0-T13,T15-T20 |
| file-organize | **FO** | G0-G13 |
| declaration-table | **DCR** | G0-G13 |
| vouching-journal | **VJ** | G0-G13 |

> 共享规则保持原编号（G0-G13, T0-T20），它们在各自文件中已具唯一性。
> 域规则旧编号用括号标注过渡期，如"EQ-21(原T21)"。

### 规则添加位置

- **D/R型规则**：添加到SKILL.md的"执行纪律"表格末尾
- **O型规则**：直接写入对应step文件，SKILL.md不再列出
- **references知识点**：添加到对应references文件，并更新INDEX.md索引

### 跨skill规则同步

新增规则影响其他skill时，**MUST**在注册模板的"影响skill"字段声明。被影响的skill**MUST**在下一版更新中：
1. 创建对应规则或引用声明
2. 补充对应CHECK项
3. 如涉及数据流传递，在step文件中补充同步操作步骤

---

## MR-14 强制CHECK绑定（Rule-CHECK Binding）

> **目的**：消除"写了规则但不执行"的空转。每条规则必须有执行保障机制。

### 三层执行保障

| 层级 | 机制 | 适用分类 | 保障强度 |
|------|------|---------|---------|
| **L1 声明层** | SKILL.md规则表格有"分类"列(D/O/R) | D/O/R全部 | 明确规则强度 |
| **L2 校验层** | 对应Phase的C{N}校验中有一行CHECK项 | D/R型必须 | 执行时逐项勾选 |
| **L3 门控层** | Phase门控级CHECK项（不通过=禁止流转） | R型必须 | 阻断性保障 |

### 具体要求

1. **D型规则**→至少在对应Phase的C{N}校验中有一行检查项，标注`[DT-46]`等规则编号
2. **R型规则**→必须有Phase门控级CHECK项（在C4.5/C5.5/C6.5等门控阶段），不通过=禁止流转
3. **O型规则**→迁移至step文件，step文件内部自行验证，无需独立CHECK
4. **新增规则无对应CHECK项**→视为"advisory"（建议），不作为强制检查项

### C5.5/C6.5 反思固化门控增强

在反思固化门控中增加校验：

- [ ] **本次新增的每条D/R规则均有对应CHECK项** — 无CHECK的新规则=未完成固化
- [ ] **本次新增的R型规则均在Phase门控中有阻断性CHECK** — 仅有一般性CHECK的R规则=降级为D型

**未完成CHECK绑定的新规则 = 未完成反思固化 = 禁止交付**

---

## 引用方式

在各评估skill的SKILL.md头部使用以下引用：

```markdown
[元规则] 参照 valuation-common/META_RULES.md（MR-1至MR-17），不可违反
```

---

## 与共享执行纪律的关系

| 文件 | 定位 | 适用范围 |
|------|------|---------|
| `META_RULES.md`（本文件） | L0 元规则 | 所有评估skill通用 |
| `audit_discipline_rules.md` | 审核类L1执行纪律 | 仅审核类skill（T0-T20） |
| `preparation_discipline_rules.md` | 编制类L1执行纪律 | 仅编制类skill（G0-G13） |

**去重原则**：本文件只放"两类skill都适用"的规则。审核类独有规则留在audit_discipline_rules；编制类独有规则留在preparation_discipline_rules。

---

## MR-15 跨skill主动同步

> **目的**：避免"声明了跨skill影响但不执行"的空转。反思固化或规则新增时，若存在跨skill影响声明，MUST主动执行同步，无需逐次询问使用人。

### 触发条件

满足以下**全部**条件时，MUST主动执行跨skill同步：

1. 在反思固化（Phase 5.5）或规则新增过程中，提出了新的领域特定规则
2. 该规则在注册模板（MR-13 RRP）的"影响skill"字段中声明了受影响的skill
3. 同步内容可提升受影响skill的工作质量（即：受影响skill存在同类风险场景）

### 执行要求

| 步骤 | 内容 | 说明 |
|------|------|------|
| 1 | 在受影响skill中创建等效规则 | 按受影响skill的前缀编号，内容适配该skill的场景 |
| 2 | 补充对应CHECK项 | D/R型规则必须有CHECK绑定（MR-14） |
| 3 | 更新版本号和变更记录 | 标注"跨skill同步自XXX vN.N" |
| 4 | 如涉及数据流传递，补充step文件 | 在对应step文件中增加同步操作步骤 |

### 同步时机

- **即时同步**：在本次反思固化的同一轮中完成，不推迟到"下一版"
- **不同步的条件**：仅当受影响skill不存在同类风险场景时，可跳过同步，但MUST在注册模板中注明"已评估，无需同步（原因：XXX）"

### 与MR-13的关系

MR-13定义了"声明跨skill影响"的格式要求；MR-15定义了"声明后必须执行"的执行力保障。两者共同确保跨skill影响不是"纸上声明"。

---

## MR-16 新Skill准入检查（Onboarding Gate）

> **目的**：确保新生成的评估skill从一开始就遵守共享架构，不再出现"遗漏引用共享规则"的问题。

### 根因

declaration-table skill曾因未引用META_RULES.md，导致备份版本堆积（MR-4失效）和多版本副本共存（MR-6失效）。根因：新skill创建时无强制准入检查，共享规则靠各skill自行引用，遗漏即失效。

### 强制要求

每个新建或重大修改的评估skill，**MUST**在SKILL.md定稿前通过准入检查：

1. **元规则引用声明**（A1）：SKILL.md标题后第一行**MUST**包含 `[元规则] 参照 valuation-common/META_RULES.md（MR-1至MR-15），不可违反`
2. **共享纪律规则引用**（A2）：编制类**MUST**引用 `preparation_discipline_rules.md`；审核类**MUST**引用 `audit_discipline_rules.md`
3. **命名空间注册**（A3）：**MUST**在本文件命名空间表中注册skill前缀
4. **关键规则显式声明**（B1-B4）：备份≤2版本、覆盖保存、截图用完即删、文件版本承接——即使已在共享规则中定义，**MUST**在SKILL.md中显式声明域规则版本
5. **准入检查清单**：详见 `valuation-common/skill_onboarding_checklist.md`

### 不通过=禁止使用

未通过准入检查的skill，**禁止**正式使用。Agent在执行评估任务时，**MUST**检查skill是否有元规则引用声明，无声明=降级为advisory建议，不可作为正式工作流使用。

---

## MR-17 硬约束关卡机制（Hard Gate System）

> **目的**：解决"规则写了但Agent不执行"的系统性问题。将软性规则转化为结构性依赖，使跳过步骤在物理上不可行。

### 根因

5月20日审核源奇商誉减值测试收益法底稿时，v1版报告遗漏3项🔴严重问题（SY4占比当增长率、SY9/SY10 $O列绝对引用），审核执行透明度附注与5月9日版本严重不一致。根因分析：

| 层级 | 根因 | 具体表现 |
|------|------|---------|
| L1 执行层 | 脚本"✅无问题"被直接采信 | check_formulas.py对SY4/SY9/SY10返回0 findings，Agent未质疑 |
| L2 关卡层 | G-P23关卡未强制执行 | 规则写了"必须通过"，但Agent自行跳过进入Phase 3 |
| L3 报告层 | 模板要求未强制检查 | S7要求读取report_template.md但Agent跳过 |
| L4 规则层 | 软规则无物理阻断力 | FLOW.md/CHECK.md写"禁止""必须"，但仅是文字声明 |

### 三重硬约束（HC-1/HC-2/HC-3）

| # | 约束 | 机制 | 物理阻断方式 |
|---|------|------|------------|
| **HC-1** | 每Phase结束必须写入checkpoint.json | Agent在Phase末尾用Python写入`phase{N}_checkpoint.json`（含gate项状态+findings+spot_checks+scripts_executed+self_check_completed） | checkpoint缺失→下Phase验证失败→BLOCKED |
| **HC-2** | 每Phase开始前必须验证前序checkpoint | Agent运行`gate_validator.py`验证前序checkpoint | 验证返回exit code 1→必须暂停补完 |
| **HC-3** | 脚本"无问题"必须手动抽检 | check1/check2返回0 findings时，Agent必须手动openpyxl提取2行公式验证，结果写入spot_checks | Phase 2无spot_checks→gate_validator判定BLOCKED |

### checkpoint.json schema（最小必填字段）

```json
{
  "phase": 2,
  "phase_name": "Excel链接错误检查",
  "timestamp": "ISO8601",
  "gate": {
    "gate_id": "G-P23",
    "items": [
      {"id": "G-P23-1", "description": "描述", "status": "pass|skip|fail", "evidence": "脚本输出文件名+结论", "reason": "仅skip时必填"}
    ],
    "verdict": "pass|BLOCKED"
  },
  "findings": {"critical": 0, "important": 0, "suggestion": 0, "details": ["P01: ..."]},
  "spot_checks": [{"sheet": "SY4", "trigger": "check1返回0 findings", "rows_checked": ["F32"], "method": "openpyxl data_only=False", "result": "描述"}],
  "scripts_executed": ["check_formulas.py", "check_hidden.py"],
  "self_check_completed": true,
  "common_compliance": {
    "ccep_items": [
      {"id": "CCEP-1", "desc": "本Phase执行符合C0合规（MR-1+MR-7+MR-9）", "status": "pass"},
      {"id": "CCEP-2", "desc": "本Phase输出数据/结论均有源文件支撑（MR-7零幻觉）", "status": "pass"}
    ],
    "c0_status": {
      "MR-1_数据不匹配即停": "pass",
      "MR-4_备份纪律": "pass",
      "MR-5_文件版本承接": "pass",
      "MR-6_单一保存路径": "pass",
      "MR-7_零幻觉原则": "pass",
      "MR-9_异常即停": "pass",
      "MR-10_保密脱敏": "pass"
    }
  }
}
```

### 适用范围

- **审核类skill**：DCF收益法、设备、土地使用权、市场法、房地产 — 全部Phase
- **编制类skill**：暂不强制，但建议参照HC-1（过程记录）

### 与MR-14的关系

- MR-14（强制CHECK绑定）解决"规则→CHECK"的映射
- MR-17（硬约束关卡）解决"CHECK→执行"的保障
- 两者互补：MR-14确保每条规则有CHECK项，MR-17确保每个CHECK项会被执行

---

## MR-18 Common合规强制协议（Common Compliance Enforcement Protocol, CCEP）

> **目的**：解决"Common文件被引用但不被执行"的系统性断层。确保所有评估skill在每一步执行中，Common规则不仅是"知道"，而是"必须遵守+必须验证+违反即停"。

### 根因

2026-05-21审计发现：
1. **编制类8个skill**的CHECK.md/FLOW.md几乎不引用Common规则（L2/L3层普遍缺失）
2. **审核类6个skill**的steps/仅S2/S7有Common引用，其他Phase的步骤文件无任何Common声明
3. **gate_validator.py**为6个skill各自的物理副本，修复需同步6处（违反DRY）
4. **declaration-table**缺失preparation_discipline_rules.md引用
5. Agent按step文件顺序执行时，SKILL.md中的Common声明"看不见"→执行时遗忘

### 5层强制执行架构

| 层级 | 机制 | 实施位置 | 物理阻断方式 |
|------|------|---------|------------|
| **L1 声明层** | SKILL.md引用Common文件 | 各skill SKILL.md | 已有（MR-16准入检查） |
| **L2 校验层** | CHECK.md C0段：Common合规检查项 | 各skill CHECK.md | C0不通过→Phase无法关闭 |
| **L3 门控层** | FLOW.md每个Phase追加"Common合规"关卡项 | 各skill FLOW.md | 关卡未过→不可流转下一Phase |
| **L4 步骤层** | 每个step文件开头声明适用Common规则 | 各skill steps/*.md | 步骤未声明→Agent不知规则→执行歧义 |
| **L5 自动化层** | common_compliance.py验证+gate_validator共享化 | valuation-common/scripts/ | 脚本验证不通过→BLOCKED |

### L2实施规格：CHECK.md C0段模板

每个skill的CHECK.md**MUST**在C1之前添加C0段：

```markdown
## C0: Common合规检查（跨Phase强制·引自valuation-common）

| 编号 | 检查项 | 对应Common规则 | 违反后果 |
|------|--------|--------------|---------|
| **C0-1** | 数据不匹配即停（MR-1） | META_RULES MR-1 | 强行处理=数据质量事故 |
| **C0-2** | 备份纪律（MR-4） | META_RULES MR-4 | 修改前未备份=版本丢失 |
| **C0-3** | 文件版本承接（MR-5） | META_RULES MR-5 | 用旧版覆盖新版=数据丢失 |
| **C0-4** | 单一保存路径（MR-6） | META_RULES MR-6 | 多位置副本=版本混乱 |
| **C0-5** | 零幻觉原则（MR-7/G0/T0） | META_RULES MR-7 | 无来源判断=审核造假 |
| **C0-6** | 反思固化强制门控（MR-8/G7） | META_RULES MR-8 | 未完成=禁止交付 |
| **C0-7** | 异常即停（MR-9） | META_RULES MR-9 | 异常后继续=错误放大 |
| **C0-8** | 保密脱敏（MR-10/T15） | META_RULES MR-10 | 真实名称=保密违规 |
| **C0-9** | [审核类]执行纪律遵守 | audit_discipline_rules T0-T20 | 审核类专用 |
| **C0-9** | [编制类]执行纪律遵守 | preparation_discipline_rules G0-G13 | 编制类专用 |
```

> **C0段性质**：C0项为跨Phase强制检查项，不与任何特定Phase绑定。Agent在每个Phase自检时**MUST**逐项确认C0合规状态。

### L3实施规格：FLOW.md Common合规关卡模板

每个skill的FLOW.md每个Phase的GATE关卡项中**MUST**追加：

```markdown
- [ ] **CCEP-1**：本Phase执行的每项操作均符合C0合规检查（MR-1/MR-7/MR-9：数据不匹配即停+零幻觉+异常即停）
- [ ] **CCEP-2**：本Phase输出的数据/结论均有源文件支撑（MR-7/G0/T0：零幻觉原则）
```

> **CCEP**（Common Compliance Enforcement Point）编号统一，不按skill区分，因为所有skill通用。

### L4实施规格：步骤文件Common声明模板

每个step文件**MUST**在标题后第一行添加：

```markdown
> **📋 Common规则适用声明**：本步骤适用 META_RULES MR-{适用编号} + {audit_discipline_rules T-适用编号 | preparation_discipline_rules G-适用编号}
```

### L5实施规格：自动化验证

1. **gate_validator.py共享化**：从6个skill本地副本迁移至`valuation-common/scripts/gate_validator.py`，各skill通过`sys.path.insert`引用
2. **common_compliance.py**：新增通用合规验证脚本，检查：
   - checkpoint.json中是否包含`common_compliance`字段
   - CCEP关卡项是否已勾选
   - 零幻觉原则：findings中每条问题是否有单元格引用

### 适用范围

- **所有15个评估skill**（审核类6个+编制类8个+报告审核1个）
- **每个Phase**的执行过程中
- **不限于Phase 2**：MR-1/MR-5/MR-7/MR-9等跨Phase规则在所有Phase都适用

---

## 版本记录

| 版本 | 日期 | 变更 |
|------|------|------|
| v1.0 | 2026-05-18 | 初始版本，MR-1至MR-12 |
| v2.0 | 2026-05-18 | 新增MR-13规则注册协议(RRP)、MR-14强制CHECK绑定；建立规则编号命名空间、三层执行保障机制、反思固化门控增强 |
| v2.1 | 2026-05-19 | 新增MR-15跨skill主动同步：声明跨skill影响后MUST主动执行同步，无需逐次询问使用人；与MR-13配合确保跨skill影响不是"纸上声明" |
| v2.2 | 2026-05-20 | 命名空间注册补全：新增declaration-table(DCR)和vouching-journal(VJ)前缀；vouching-journal补加元规则引用 |
| v2.3 | 2026-05-20 | 新增MR-16新Skill准入检查：新建skill必须通过准入检查（元规则引用+共享规则引用+命名空间注册+关键规则显式声明），未通过=禁止使用 |
| v2.4 | 2026-05-21 | 新增MR-17硬约束关卡机制：三重硬约束HC-1/HC-2/HC-3（checkpoint文件+关卡验证脚本+Post-OK抽检协议），解决"规则写了但Agent不执行"的系统性问题 |
| v2.5 | 2026-05-21 | 新增MR-18 Common合规强制协议（CCEP）：5层强制执行架构（L1声明→L2校验C0段→L3门控CCEP关卡→L4步骤声明→L5自动化），解决"Common文件被引用但不被执行"的系统性断层 |
