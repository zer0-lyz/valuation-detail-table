# 编制类Skill共享执行纪律规则

> 本文件汇总了所有编制类skill（成本法底稿编制、评估明细表填写、待补资料清单、抽凭底稿生成）共用的执行纪律规则。
> 各skill的SKILL.md中对应编号规则统一引用本文件，仅保留各skill独有的领域特定规则。

---

## 共享规则总表

| # | 执行纪律 | 适用范围 | 违反后果 |
|---|---------|---------|---------|
| **G0** | **零幻觉原则**：填入底稿/明细表的每一个数值必须有源文件中的具体数据支撑，禁止基于推测填写。不确定的数据标注"待核实"，不编造 | 全Phase | 无来源数据 = 数据质量事故 |
| **G1** | **Python脚本保存为.py文件执行**，避免PowerShell/Bash中f-string引号转义冲突 | 全Phase | 内联脚本 = 编码问题 |
| **G2** | **openpyxl加载时保留公式**：默认`data_only=False`加载，写入数值时不破坏已有公式结构 | Phase 2-3 | 公式丢失 = 后续计算断裂 |
| **G3** | **openpyxl保存后必须用Excel COM重算保存**：openpyxl保存Excel文件会清除公式缓存值（`<v>`元素），导致WPS/Excel打开后重新计算并标记文件为"已修改"。正确做法：openpyxl保存后，通过`win32com.client.DispatchEx('Excel.Application')`打开文件→`app.Calculate()`→`wb.Save()`→关闭 | Phase 5 | 未重算保存 = 每次打开弹"是否保存更改" |
| **G4** | **Excel COM必须使用DispatchEx**：调用Excel COM时，**必须**使用`win32com.client.DispatchEx('Excel.Application')`创建独立实例，而非`Dispatch`或`GetActiveObject`。`Dispatch`会复用用户已打开的WPS/Excel进程 | Phase 5 | 使用Dispatch = 关闭用户其他文件 |
| **G5** | **🚨 文件版本承接原则（最高优先级）**：更新文件时，①必须承接上文最近一个修改版本（即本次会话中已被操作/保存的文件），不可重新从其他路径检索文件作为修改基础；②修改前必须备份到`D:\workbuddy`（最多2个版本）；③如果不确定哪个版本是最新版，必须向用户确认后再操作 | 全Phase | 用旧版本覆盖新版本 = 数据丢失事故 |
| **G6** | **🚨 单一保存路径原则**：更新后的文件只保存到原路径（即文件被读取时的路径），禁止在多个位置保存副本 | 全Phase | 多版本副本 = 版本混乱 |
| **G7** | **🚨 交付前反思固化原则（强制门控）**：每次完成修改/交付成果前，必须主动回顾本次操作中遇到的所有问题，固化为具有强执行力的操作规范写入SKILL.md和lessons_learned.md。**禁止等用户提醒**，禁止以"记录了"代替"更新了"。未完成反思固化 = 禁止交付 | 全Phase | 问题复发 = 需用户反复提醒 = 信任崩塌 |
| **G8** | **执行透明度强制披露**：最终交付时必须附执行情况摘要表，逐Phase如实披露已执行/跳过/关键发现。禁止笼统表述"全部通过"掩盖跳过步骤 | Phase 5 | 隐瞒未执行步骤 = 过程不可追溯 |
| **G9** | **文件交付规则：按保存路径决定是否调用deliver_attachments**：①文件已通过`wb.save()`/`shutil.copy`保存到**用户指定路径**（WPS云盘、桌面、项目文件夹等workspace外路径）→ **不调用deliver_attachments**；②文件保存在**workspace内**→ **必须调用deliver_attachments**正式交付 | Phase 5 | 已存到用户路径又deliver=每次弹窗；workspace内文件未deliver=用户拿不到 |
| **G10** | **🚨 输入内容格式统一规范（字体+填充+边框）**：所有输入内容默认字体为 Times New Roman，字号10号（10pt）。中文字段可使用宋体但仍以10号为默认字号。写入数据时须显式设置`Font(name='Times New Roman', size=10)`，同步复制参考行格式（fill+border） | Phase 2-4 | 字体字号填充边框不统一 = 底稿排版混乱 |
| **G11** | **步骤复核表未执行内容应隐藏**：步骤复核表中未执行的步骤行应在索引和文字工作完成后隐藏（`ws.row_dimensions[r].hidden = True`），不可仅标记"不适用"而保留可见 | Phase 2 | 未执行步骤可见 = 底稿不严谨 |
| **G12** | **合并单元格检测**：openpyxl写入前必须检查`isinstance(cell, MergedCell)`，合并单元格不能直接赋值 | Phase 2-3 | 写入报错 = 脚本中断 |
| **G13** | **文件被占用时使用临时文件策略**：先copy2到临时文件→修改→os.replace替换原文件 | 全Phase | PermissionError = 保存失败 |

---

## 引用方式

在各编制类skill的SKILL.md"执行纪律"章节中，使用以下引用格式替代逐条重复：

```markdown
> **G0-G13** 共享规则详见 `valuation-common/preparation_discipline_rules.md`，以下仅列出本skill的领域特定规则。
```

> **命名空间规则**（MR-13）：域规则编号须带skill前缀（AB-/DT-/CL-/VE-/CI-），**禁止跨skill重号**。旧编号用括号标注过渡期，如"DT-46(原T46)"。

> **强制CHECK绑定**（MR-14）：每条D/R型域规则必须有对应CHECK项。R型规则必须有Phase门控级CHECK。

---

_本文件由skill优化自动生成，最后更新：2026-05-18_
