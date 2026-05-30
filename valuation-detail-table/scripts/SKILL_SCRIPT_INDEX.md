# SKILL_SCRIPT_INDEX.md — 脚本调用索引 (DT-165)

> **📋 DT-165强制要求**：Agent编写任何涉及评估明细表/成本法底稿的Python脚本时，
> **MUST先读本索引**，优先调用已有脚本，禁止从零重写同等功能。

---

## 一、核心写入接口（红线级，禁止绕过）

| 脚本 | 路径 | 功能 | 调用方式 | 对应规则 |
|------|------|------|---------|---------|
| **fill_sheet()** | `valuation-common/scripts/sheet_filler.py` | 评估明细表数据写入唯一入口，内含7步管线（结构识别→列位映射→插行判断→合计行保护断言→数据写入→回读验证→即时勾稽） | `from sheet_filler import fill_sheet; result = fill_sheet(ws=ws, sheet_id='3-7', data_rows=data, ...)` | DT-160/152/153/164/164.1 |
| **fill_sheet_safe()** | `valuation-common/scripts/sheet_filler.py` | fill_sheet()的安全封装，catch所有异常返回result而非crash。Agent优先调用此函数，避免因crash驱使绕过管道 | `from sheet_filler import fill_sheet_safe; result = fill_sheet_safe(ws, sheet_name='3-7', data_rows=data, ...)` | DT-160/152/153 |
| **smart_insert_row()** | `valuation-common/scripts/excel_row_ops.py` | 安全插入行（修复SUM/合并/打印范围/行高/数字格式/跨sheet引用/条件格式范围），fill_sheet内部自动调用。v1.5升级：支持三行结构+条件格式(ISFORMULA)底色同步 | `from excel_row_ops import smart_insert_row` | DT-2/113/120/163 |
| **smart_delete_rows()** | `valuation-common/scripts/excel_row_ops.py` | 安全删除行 | `from excel_row_ops import smart_delete_rows` | DT-113 |
| **fix_three_row_structure()** | `valuation-common/scripts/fix_three_row_structure.py` | 修复明细表中三行结构遗漏问题（B:C合并/公式/跨sheet链接/打印范围/序号格式/字体），可独立运行 | `from fix_three_row_structure import main; main(filepath)` | DT-163升级 |
| **fix_format_issues()** | `valuation-common/scripts/fix_format_issues.py` | 修复明细表8类格式问题（打印范围B列起/序号TNR+0格式/C列12pt→11pt/公式下拉/多余边框清理/序号自动填写/条件格式下拉/SUM范围自适应修复），可独立运行 | `from fix_format_issues import fix_workbook; fix_workbook(filepath)` | DT-162/163升级 |
| **fix_round3()** | `valuation-common/scripts/fix_round3.py` | 修复明细表第3轮问题：1)增值额/增值率/账龄公式列格式固化(会计格式) 2)打印范围右至备注列 3)减值类科目贷方金额→填正数 4)条件格式(ISFORMULA)范围扩展→公式列浅灰底色同步 | `from fix_round3 import fix_workbook; fix_workbook(filepath)` | DT-162/163升级 |

---

## 二、数据加载接口

| 脚本 | 路径 | 功能 | 调用方式 |
|------|------|------|---------|
| **load_subject_data()** | `valuation-detail-table/scripts/data_loader.py` | 从科目余额表提取科目数据（名称/金额/结算对象） | `from data_loader import load_subject_data` |
| **load_auxiliary_balance()** | `valuation-detail-table/scripts/data_loader.py` | 从辅助余额表提取往来科目结算对象明细 | `from data_loader import load_auxiliary_balance` |
| **load_journal_data()** | `valuation-detail-table/scripts/data_loader.py` | 从序时账提取结算对象发生日期/业务内容（Phase 3用） | `from data_loader import load_journal_data` |
| **load_bank_statement()** | `valuation-common/scripts/bank_statement_extract.py` | 从银行对账单PDF提取账户余额 | `from bank_statement_extract import extract_bank_statement_pdfplumber` |
| **_extract_pdf_sources()** | `valuation-detail-table/scripts/dt_runner.py` | DT-211: Phase 0 Step 0.4自动扫描项目PDF，按关键词分类并提取（银行对账单/卡片台账/辅助余额/通用），结果存pdf_extractions.json，扫描件存multimodal_tasks.json | dt_runner内部调用，Phase 0自动执行 |
| **batch_extract_bank_statements()** | `valuation-common/scripts/bank_statement_extract.py` | 批量提取银行对账单PDF（混合方案：pdfplumber→多模态兜底），含9家银行适配器+通用适配器+去重+DT-133基准日倒序查找 | `from bank_statement_extract import batch_extract_bank_statements` |
| **extract_pdf()** | `valuation-common/scripts/pdf_extract.py` | 通用PDF三级提取（pdfplumber→PyMuPDF→OCR兜底），含银行对账单/资产台账/辅助余额场景化提取 | `from pdf_extract import extract_pdf` |
| **_extract_generic_bank()** | `valuation-common/scripts/bank_statement_extract.py` | DT-213: 通用银行适配器——对未专门适配的银行，用pdfplumber.extract_tables()提取结构化数据，匹配关键词定位列位 | bank_statement_extract内部自动调用 |
| **subject_classification.json** | `valuation-common/scripts/subject_classification.json` | DT-213: 科目代码前缀→分类桶映射配置（115条），支持项目级覆盖（项目目录下同名文件优先） | dt_runner._classify_data()自动加载 |

---

## 二-A、序时账查阅接口（Phase 3专用，DT-166红线）

> **🚨 DT-166红线**：Phase 3所有序时账查阅操作MUST通过以下脚本执行，禁止Agent自行编写序时账解析/匹配逻辑。

| 脚本 | 路径 | 功能 | 调用方式 | 对应规则 |
|------|------|------|---------|---------|
| **JournalExtractor** | `valuation-detail-table/scripts/journal_extractor.py` | 序时账数据提取器，封装列映射验证+日期解析+数据加载+关键词查询。v1.1: 三级模糊匹配降级（精确→摘要→编码前缀），超50条按金额排序取TOP20 | `from journal_extractor import JournalExtractor; ext = JournalExtractor(seq_file)` | DT-51①/54 |
| **extract_dates()** | `valuation-detail-table/scripts/journal_extractor.py` | 批量提取往来科目发生日期（末笔日期） | `from journal_extractor import extract_dates; results = extract_dates(ext, empty_rows)` | DT-51③~⑥/52/53 |
| **extract_business_contents()** | `valuation-detail-table/scripts/journal_extractor.py` | 批量提取往来科目业务内容（DT-60 5步流程） | `from journal_extractor import extract_business_contents; results = extract_business_contents(ext, empty_rows)` | DT-60/149 |
| **scan_empty_fields()** | `valuation-detail-table/scripts/journal_extractor.py` | 扫描评估明细表中往来科目的空字段行 | `from journal_extractor import scan_empty_fields; empty = scan_empty_fields(detail_file)` | DT-46 |
| **write_phase3_results()** | `valuation-detail-table/scripts/journal_extractor.py` | 将Phase 3核实结果写入评估明细表（仅修改日期/业务内容列） | `from journal_extractor import write_phase3_results; write_phase3_results(detail_file, date_res, biz_res)` | DT-30/46 |
| **generate_phase3_report()** | `valuation-detail-table/scripts/journal_extractor.py` | 生成Phase 3核实结果汇总报告 | `from journal_extractor import generate_phase3_report; report = generate_phase3_report(date_res, biz_res)` | DT-48 |

---

## 二-B、强制自检接口（P1层防护，DT-200执行后MUST运行）

> **🚨 P1防护**：Phase 5内置post_execution_audit.py，检测6类已知问题。绕过管线后自检是兜底检测。

| 脚本 | 路径 | 功能 | 调用方式 | 对应规则 |
|------|------|------|---------|---------|
| **run_audit()** | `valuation-detail-table/scripts/post_execution_audit.py` | 6类强制自检（序号/坏账正数/名称非空/日期格式/边框/SUM范围） | `from post_execution_audit import run_audit; result = run_audit(xlsx_path, cache_dir)` | DT-160/153/18/166/167/46/82/2 |
| **print_audit_report()** | `valuation-detail-table/scripts/post_execution_audit.py` | 打印审计报告 | `from post_execution_audit import print_audit_report; print_audit_report(result)` | DT-160 |

---

## 三、验证接口

| 脚本 | 路径 | 功能 | 调用方式 | 对应规则 |
|------|------|------|---------|---------|
| **gate_G2()** | `valuation-detail-table/scripts/gate_validator.py` | Phase 2门控校验（G2-1~G2-19） | `from gate_validator import gate_G2; passed, violations = gate_G2(filepath, has_journal=True)` | DT-138/161/164.1 |
| **validate_sheet_after_fill()** | `valuation-detail-table/scripts/validate_sheet_after_fill.py` | 单Sheet填写后验证 | `from validate_sheet_after_fill import validate` | DT-97 |
| **find_header_structure()** | `valuation-detail-table/scripts/gate_validator.py` | 识别Sheet表头结构（data_start_row/total_row/bad_debt_row/total2_row） | `from gate_validator import find_header_structure` | 通用 |

---

## 四、列位映射接口

| 脚本/数据 | 路径 | 功能 | 调用方式 | 对应规则 |
|-----------|------|------|---------|---------|
| **sheet_col_finder.py** | `valuation-common/scripts/sheet_col_finder.py` | 列位动态查找公共模块——运行时扫描Row5/Row6表头文字，按语义关键词匹配列号。零配置、自动适配所有模板 | `from sheet_col_finder import find_header_cols, get_formula_cols, get_amount_cols, find_data_start_row, find_last_print_col, SheetColFinder` | DT-153升级 |
| **sheet_col_map.json** | `valuation-detail-table/assets/sheet_col_map.json` | 112个Sheet的精确列位映射（预生成） | `from sheet_filler import _load_col_map; col_map = _load_col_map(sheet_name)` | DT-153 |
| **_build_col_map()** | `valuation-common/scripts/sheet_filler.py` | 运行时从表头行读取列位映射 | `from sheet_filler import _build_col_map` | DT-136 |
| **build_col_map()** | `valuation-detail-table/scripts/code_templates/step_templates.py` | 简化版列位映射（step_templates） | `from code_templates.step_templates import build_col_map` | DT-136 |

---

## 五、业务逻辑接口

| 脚本 | 路径 | 功能 | 调用方式 |
|------|------|------|---------|
| **infer_business_content()** | `valuation-common/scripts/business_content_map.py` | 从行业映射+科目名推断业务内容 | `from business_content_map import infer_business_content` |
| **prepare_data_rows()** | `valuation-common/scripts/sheet_filler.py` | 将原始数据组织为fill_sheet()接受的data_rows格式 | `from sheet_filler import prepare_data_rows` |

---

## 六、代码模板

| 脚本 | 路径 | 功能 | 调用方式 |
|------|------|------|---------|
| **standard_fill_pipeline()** | `valuation-detail-table/scripts/code_templates/step_templates.py` | 标准填写流水线（列映射→列序→回读→规则汇总→门控） | `from code_templates.step_templates import standard_fill_pipeline` |
| **run_gate_validator()** | `valuation-detail-table/scripts/code_templates/step_templates.py` | DT-138门控调用封装 | `from code_templates.step_templates import run_gate_validator` |

---

## 七、格式处置接口

| 脚本 | 路径 | 功能 | 调用方式 |
|------|------|------|---------|
| **_apply_direct_format()** | `valuation-common/scripts/sheet_filler.py` | 数据行标准格式直接定义（11pt/thin/千分位） | fill_sheet()内部自动调用 |
| **hide_empty_sheets()** | `valuation-common/scripts/hide_empty_sheets.py` | 隐藏空白Sheet | `from hide_empty_sheets import hide_empty_sheets` |

---

## 🚨 禁止做的事（DT-160 + DT-165）

| 禁止操作 | 应调用 |
|---------|-------|
| 直接`ws.cell(row=r, column=9).value = xxx`写数据 | `fill_sheet(ws=ws, ...)` |
| 自己写`fill_sheet_data()`函数 | `fill_sheet()` |
| 硬编码`column=9`为账面价值列 | `sheet_col_map.json`或`_build_col_map()` |
| 直接`ws.insert_rows()` | `smart_insert_row()` |
| 手动写SUM公式修复 | `smart_insert_row()`内部自动处理 |
| 手动处理B:C合并 | `smart_insert_row()`内部自动处理 |
| 自己写回读验证逻辑 | `fill_sheet()` Step2f自动回读 |
| 自己写勾稽比对逻辑 | `fill_sheet()` Step2g即时勾稽 |
| 跳过Phase 3序时账查阅 | `S3_journal_extract.md` Step 3.0~3.8 + `journal_extractor.py` |

---

## 📋 DT-165 脚本选择流程

```
Agent需要编写脚本处理评估明细表/成本法底稿
  │
  ├── Step 0: Read 本索引 (SKILL_SCRIPT_INDEX.md)
  │     │
  │     └── 确认已有脚本能否覆盖需求
  │
  ├── Step 1: 能覆盖 → MUST调用已有脚本（import+调用）
  │     │                        禁止从零重写同等功能
  │     │
  ├── Step 2: 不能完全覆盖 → MUST在已有脚本基础上增量扩展
  │     │                        禁止另起炉灶写独立脚本
  │     │
  └── Step 3: 确需新写 → MUST在脚本头部注释原因+更新本索引
        # DT-165 NEW SCRIPT: 原因=XXX, 无法复用=XXX
```
