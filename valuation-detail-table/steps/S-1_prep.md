# S-1: 材料准备与操作提示

> **📋 Common规则适用声明**：本步骤适用 META_RULES MR-1/MR-4/MR-5/MR-9 + preparation_discipline_rules G0/G7
> **📋 DT规则引用（RULES.md）**：执行前MUST Read RULES.md → Phase -1规则节：DT-105(材料集中)、DT-106(全量读取)、DT-107(预提取)、DT-108(提取完整性硬Gate)、DT-115(提取工具化)、DT-134(提取质量维度)、DT-142(Poppler环境前置检查)

## 定位

Phase -1 是正式填写流程（Phase 0~5）之前的**操作提示阶段**，核心目的是确保所有填写所需材料集中到位，避免Phase 0执行时因材料分散、遗漏导致数据不完整。

**关键原则**：材料不全=填写不全=勾稽差异。本阶段通过交互式提示，将材料准备环节显式化，从源头杜绝数据源遗漏。

## 输入

- 用户发起的填写明细表请求
- 用户已提供的文件路径（如有）

## 操作

### Step -1.1 检查材料集中状态

**⚠️ 优先检查断点恢复 [DT-131]**：如果用户提供的路径下已存在 `_dt_cache/` 子目录且含JSON文件，直接跳转FLOW.md"前置判断0：断点恢复检测"，不执行本Step后续内容。

**判断逻辑（强制执行）：**

```
用户发起"填写明细表"请求
  │
  ├── 用户已将全部材料集中到一个文件夹？
  │     │
  │     ├── YES → 跳至 Step -1.3（确认材料清单）
  │     │
  │     └── NO 或 不确定 → 进入 Step -1.2（操作提示）
  │
  └── 用户提供了零散的文件路径（多个@文件）？
        │
        ├── YES → 建议用户集中到一个文件夹后继续
        │
        └── NO → 进入 Step -1.2（操作提示）
```

### Step -1.2 操作提示：材料集中放置

向用户输出以下操作提示：

---

**📋 填写评估明细表 — 材料准备提示**

为确保明细表填写完整准确，请将以下材料**集中放置到同一个文件夹**内：

| 序号 | 材料 | 必需/可选 | 说明 |
|------|------|----------|------|
| 1 | **科目余额表** | ✅ 必需 | 包含末级科目编码、名称、期末余额 |
| 2 | **资产负债表** | ✅ 必需 | 用于勾稽核对 |
| 3 | 序时账/明细账 | ⚠️ 有则必执行Phase 2e | 往来科目核实发生日期与业务内容（DT-161：有序时账时Phase 2e MUST执行） |
| 4 | 辅助余额表 | 可选 | 往来科目结算对象明细 |
| 5 | 银行对账单 | 可选 | 银行存款校对（PDF/图片/Excel均可） |
| 6 | 固定资产卡片台账 | 可选 | 固定资产逐项明细（PDF/图片/Excel均可） |
| 7 | 收发存明细表 | 可选 | 存货分类明细 |
| 8 | 评估明细表模板 | 可选 | 未提供时使用内置默认模板v1.90 |
| 9 | 其他相关材料 | 可选 | 合同、发票、权属证明等（PDF/图片均可） |

**⚠️ 重要说明：**

1. **文件夹内的所有文件将被自动读取和识别**，包括PDF、图片（PNG/JPG）、Excel、Word等格式
2. 系统会自动识别文件用途（如"科目余额表""银行对账单"等），无需手动标注
3. PDF和图片材料将通过多模态识别提取其中的数据
4. 放置完成后，请告知文件夹路径，系统将自动开始处理

---

### Step -1.3 确认材料清单与文件遍历

自动扫描材料文件夹（无需用户确认），执行以下操作：

1. **遍历目标文件夹**：递归扫描所有子目录，生成完整文件清单
2. **文件类型识别**：
   - `.xlsx`/`.xls` → Excel文件，openpyxl读取
   - `.pdf` → PDF文件，Read工具多模态识别（DT-73）
   - `.png`/`.jpg`/`.jpeg`/`.bmp`/`.tiff` → 图片文件，Read工具多模态识别
   - `.docx`/`.doc` → Word文件，Read工具读取
   - `.csv` → CSV文件，直接解析
3. **用途预判**：根据文件名关键词预判用途（规则同DT-103 Step 0.1.5）
4. **输出材料清单**：向用户展示识别结果，确认无遗漏

**输出格式示例：**

```
📁 材料文件夹：D:/项目/平绿/明细表填写资料/

| 序号 | 文件名 | 类型 | 预判用途 | 状态 |
|------|--------|------|---------|------|
| 1 | 科目余额表2025.xlsx | Excel | 科目余额表（主数据源） | ✅ 已识别 |
| 2 | 资产负债表.pdf | PDF | 勾稽对照 | ✅ 已识别 |
| 3 | 银行对账单.pdf | PDF | 银行存款校对 | ✅ 已识别 |
| 4 | 序时账.xlsx | Excel | 往来核实 | ✅ 已识别 |
| 5 | 合同扫描件.jpg | 图片 | 用途待确认 | ⚠️ 需确认 |

必需材料：科目余额表✅ 资产负债表✅
可选材料：已识别3份
```

### Step -1.4 必需材料完整性检查

**检查必需材料是否齐全：**

- ✅ 科目余额表 — 有 → 继续
- ❌ 科目余额表 — 无 → 输出缺失清单+标注[待核实]+继续执行（DT-151异常不中断原则）
- ✅ 资产负债表 — 有 → 继续
- ❌ 资产负债表 — 无 → 提示，可继续（勾稽时标注"无BS对照"）

**检查可选材料缺失告警：**

- 银行存款科目有余额但无对账单 → WARNING（银行存款校对受限DT-65）
- 往来科目有余额但无序时账 → WARNING（往来科目发生日期核实受限DT-51，适用DT-143日期留空，Phase 2e可跳过DT-161①）
- 往来科目有余额且有序时账 → **INFO：Phase 2e MUST执行（DT-161）**
- 固定资产科目有余额但无卡片台账 → WARNING（固定资产明细受限DT-88）

> **注意**：此处的"余额"信息在Phase -1阶段可能无法获取（尚未解析科目余额表），WARNING仅为预判。Phase 0解析科目余额表后，会执行更精确的DT-103数据源完整性门控。

### Step -1.5 文件内容预提取

对文件夹中的**非Excel文件**（PDF、图片），在Phase -1阶段即开始预提取：

**🚨 [DT-142] Poppler环境前置检查（强约束）**：

> **背景**：PDF提取的四级策略中，第3级（OCR）和DT-132多模态Read兜底均依赖 `pdf2image` 库将PDF转为图片，而 `pdf2image` 底层依赖 `poppler` 的 `pdftoppm` 命令。若poppler未安装/未配置PATH，`pdf_to_images()` 会静默返回空列表，导致DT-132多模态Read兜底流程无法生成图片文件，Agent无法用Read工具识别扫描件内容。

**强制规则**：

1. **Step -1.5执行前MUST检查poppler是否可用**：调用 `pdf_to_images()` 前先验证环境
2. **poppler不可用时MUST配置PATH**：
   - 如本机存在可选Windows Release缓存，可使用`scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin/`（含pdftoppm.exe）；该大体积依赖不随Git仓库分发
   - 另有源码包 `scripts/poppler-26.05.0/`（需编译，Linux/macOS可用）
   - 配置步骤详见 `scripts/README.md` → "依赖工具 → Poppler" 节
3. **禁止在poppler不可用时跳过PDF转图片步骤**

**检查与配置脚本**：

```bash
# 快速检查poppler是否可用
pdftoppm -v 2>/dev/null && echo "✅ poppler OK" || echo "❌ poppler NOT FOUND"

# 若不可用，执行以下配置：
SKILL_DIR="$HOME/.codex/skills/valuation-detail-table/valuation-detail-table"

# 优先使用预编译Windows二进制包（Release-26.02.0-0）
POPLER_BIN="$SKILL_DIR/scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin"
if [ -d "$POPLER_BIN" ]; then
    export PATH="$POPLER_BIN:$PATH"
    echo "✅ 已添加 $POPLER_BIN 到 PATH"
else
    # 回退：尝试源码包编译后的bin目录
    for bindir in "$SKILL_DIR/scripts/poppler-26.05.0/Library/bin" "$SKILL_DIR/scripts/poppler-26.05.0/bin"; do
        if [ -d "$bindir" ]; then
            export PATH="$bindir:$PATH"
            echo "✅ 已添加 $bindir 到 PATH"
            break
        fi
    done
fi

# 验证
pdftoppm -v 2>/dev/null && echo "✅ poppler配置成功" || echo "❌ poppler配置失败，检查scripts/Release-26.02.0-0/目录"
```

```python
# Python中自动配置poppler PATH（推荐在脚本开头调用）
import os, shutil

def ensure_poppler():
    """确保poppler可用，自动配置PATH。优先使用Release预编译包。"""
    if shutil.which('pdftoppm'):
        return True
    
    skill_dir = os.path.join(os.path.expanduser('~'), '.codex', 'skills', 'valuation-detail-table', 'valuation-detail-table')
    
    # 优先级1: Release预编译Windows二进制（直接可用）
    candidates = [
        os.path.join(skill_dir, 'scripts', 'Release-26.02.0-0', 'poppler-26.02.0', 'Library', 'bin'),
        os.path.join(skill_dir, 'scripts', 'poppler-26.05.0', 'Library', 'bin'),
        os.path.join(skill_dir, 'scripts', 'poppler-26.05.0', 'bin'),
    ]
    for poppler_bin in candidates:
        if os.path.isdir(poppler_bin):
            os.environ['PATH'] = poppler_bin + os.pathsep + os.environ.get('PATH', '')
            if shutil.which('pdftoppm'):
                return True
    
    print('[WARNING] poppler未安装或未配置PATH。PDF转图片功能不可用。')
    print(f'[HINT] 预编译包路径: {candidates[0]}')
    return False
```

**🚨 [DT-134] 场景化提取分发规则（强约束）**：

> **背景**：河南平绿项目Phase -1对所有PDF统一调用`extract_pdf()`通用提取，产出raw_text而非结构化数据。银行对账单的关键产出是`{bank_name, account_no, ending_balance}`结构化字段，而非57KB的raw_text。通用提取导致Phase 2无可用结构化数据，Agent走捷径用汇总数填充，违反DT-104/DT-65/DT-109三条红线。

**强制分发逻辑**：根据Step -1.3的用途预判结果，MUST调用对应的场景化提取函数：

| 用途预判 | MUST调用的函数 | 产出结构化字段 | 禁止使用 |
|---------|--------------|-------------|---------|
| 银行对账单 | `extract_bank_statement()` | bank_name, account_no, ending_balance, currency | ❌ `extract_pdf()` |
| 保证金对账单 | `extract_bank_statement()` | bank_name, account_no, ending_balance, currency | ❌ `extract_pdf()` |
| 固定资产卡片台账 | `extract_asset_register()` | items[{name, original_value, net_value, quantity}] | ❌ `extract_pdf()` |
| 辅助余额表PDF | `extract_auxiliary_balance()` | counterparties[{name, debit_balance, credit_balance}] | ❌ `extract_pdf()` |
| 其他PDF/图片 | `extract_pdf()` | raw_text（通用场景无结构化字段要求） | — |

**执行流程**：

```
Step -1.3用途预判结果
  │
  ├── 文件用途 = "银行对账单" / "保证金对账单"
  │     │
  │     ├── MUST调用 extract_bank_statement()
  │     │     │
  │     │     ├── 提取成功(含bank_name+account_no+ending_balance) → PASS → 持久化
  │     │     │
  │     │     ├── 提取失败/结果缺结构化字段 → 回退到extract_pdf()获取raw_text
  │     │     │     │
  │     │     │     ├── raw_text非空 → 从raw_text中正则解析bank_name+account_no+ending_balance
  │     │     │     │     ├── 解析成功 → PASS → 持久化
  │     │     │     │     └── 解析失败 → 触发DT-132多模态Read兜底 → 持久化
  │     │     │     │
  │     │     │     └── raw_text为空(<100字符) → 直接触发DT-132多模态Read兜底 → 持久化
  │     │     │
  │     │     └── 🚨 绝对禁止：仅调用extract_pdf()产出raw_text就标记"extracted"
  │     │
  │     └── 持久化格式MUST包含结构化字段：
  │           {
  │             "bank_name": "中国建设银行",
  │             "account_no": "41050172560800001218",
  │             "account_name": "河南平煤神马平绿置业有限公司",
  │             "currency": "CNY",
  │             "ending_balance": 12345678.90,
  │             "balance_date": "2025-12-31",
  │             "raw_text": "...(保留原始文本供回溯)"
  │           }
  │
  ├── 文件用途 = "固定资产卡片台账"
  │     └── MUST调用 extract_asset_register() → 类似回退逻辑
  │
  ├── 文件用途 = "辅助余额表"
  │     └── MUST调用 extract_auxiliary_balance() → 类似回退逻辑
  │
  └── 文件用途 = "其他"
        └── 调用 extract_pdf() → raw_text非空即PASS
```

**🚨 关键约束**：
- **DT-134-1**：银行对账单**禁止**仅调用`extract_pdf()`产出raw_text就标记"extracted"。MUST经过场景化提取（extract_bank_statement或从raw_text解析）产出bank_name+account_no+ending_balance结构化字段，才能标记"extracted"
- **DT-134-2**：场景化提取失败时，允许回退到extract_pdf()获取raw_text，但MUST额外执行"从raw_text解析结构化字段"步骤。仅产出raw_text而不尝试解析=提取质量不足=PARTIAL
- **DT-134-3**：所有回退提取的结果MUST持久化到_dt_cache/，结构化字段+raw_text同时保存

**🚨 [DT-133] 银行对账单基准日倒序查找规则（强约束）**：

> **背景**：银行对账单通常为多页PDF，包含整个月份甚至多个月份的交易记录。评估基准日（如2025-12-31）的账户余额是核心数据，顺序查找效率低且可能提取大量无关数据。

**强制规则**：

1. **按基准日定位对账单**：优先查找基准日当月（如12月）的对账单文件
2. **页内倒序查找余额**：多页对账单从**最后一页向前**查找，因为账户余额/可用余额通常在对账单末尾
3. **多月份对账单按月份倒序**：若文件夹含多月份对账单（10月/11月/12月），优先处理基准日当月（12月），次之基准日上月（11月），以此类推
4. **🚨 找到即停（Early Stop）**：**找到基准日当天最后的余额信息后，该对账单PDF前面页面的交易明细不需要继续提取**。评估明细表只需要基准日余额，不需要完整交易流水
5. **余额确认逻辑**：
   - 对账单末尾有"账户余额"/"余额"行 → 直接取该数值
   - 对账单末尾有"可用余额"+"冻结余额" → 账户余额 = 可用余额 + 冻结余额
   - 对账单末尾只有"上页余额"无期末余额 → 向前翻页查找期末余额行
   - 基准日当天无交易 → 取基准日前最后一个交易日的余额
6. **提取效率**：每个对账单PDF只需提取：银行名称、账号、户名、币种、**基准日余额**。不需要逐笔交易明细

**DT-133操作流程**：

```
对账单PDF文件（多页）
  │
  ├── 1. 从最后一页开始提取/OCR
  │     │
  │     ├── 最后一页找到"账户余额"行？
  │     │     ├── YES → 提取余额数据 → ✅ 完成（不提取前面页面）
  │     │     └── NO → 继续向前翻页
  │     │
  │     ├── 倒数第二页找到"账户余额"行？
  │     │     ├── YES → 提取余额数据 → ✅ 完成
  │     │     └── NO → 继续向前翻页
  │     │
  │     └── ... 依次向前，找到即停
  │
  └── 输出：{bank_name, account_no, account_name, currency, balance_date, balance}
       balance_date = 基准日或基准日前最近交易日
       ❌ 不输出：逐笔交易明细（除非用户明确要求）
```

```python
# ⚠️ MUST使用 pdf_extract.py 共享工具（DT-115），禁止每次重写提取逻辑
from pdf_extract import (
    extract_pdf, extract_bank_statement, extract_asset_register,
    extract_auxiliary_balance, batch_extract, validate_extraction,
    save_extraction_json, load_extraction_json
)

# 方式1：批量提取（推荐，自动判断文件类型+生成完整性报告）
result = batch_extract(pdf_image_paths, output_dir='提取结果目录')
print(result['report_text'])  # 输出DT-108完整性报告

# 方式2：单文件场景化提取
# 银行对账单
bank = extract_bank_statement('银行对账单.pdf')
# → bank.bank_name / bank.account_no / bank.ending_balance

# 固定资产卡片台账
asset = extract_asset_register('卡片台账.pdf')
# → asset.items [{name, original_value, net_value, quantity}]

# 辅助余额表
aux = extract_auxiliary_balance('辅助余额表.pdf')
# → aux.counterparties [{name, debit_balance, credit_balance}]

# 通用PDF/图片
general = extract_pdf('任意文件.pdf')  # 三级策略：pdfplumber→PyMuPDF→OCR

# 加载已保存的提取结果（Phase 0~5直接引用，无需重复提取）
loaded = load_extraction_json('提取结果目录/银行对账单.json')
```

**四级提取策略**（按优先级自动执行，前三级为Python脚本，第四级为Agent多模态）：
1. pdfplumber（文本PDF首选，支持表格提取）
2. PyMuPDF（文本PDF备选，速度快）
3. OCR兜底（图片型PDF/扫描件，使用pytesseract+pdf2image；**若Tesseract未安装则此级不可用**）
4. **🚨 [DT-132] Agent多模态Read兜底（强制最终手段）**：当前三级Python脚本全部提取失败或结果为空时，**MUST使用Read工具直接读取PDF/图片文件**，由Agent的多模态视觉能力识别内容。**禁止以前三级失败为由跳过提取、标注"扫描件无法提取"或"需人工处理"**

**DT-132 Agent多模态Read兜底规则（强约束）**：

```
Python脚本提取（pdfplumber→PyMuPDF→OCR）
  │
  ├── 提取成功（有实质数据） → 直接使用（完成）
  │
  └── 提取失败/结果为空/扫描件
        │
        ├── 🚨 MUST立即执行以下步骤（禁止跳过）：
        │     │
        │     ├── 1. 将PDF转为图片（使用bank_statement_extract.py的pdf_to_images()，DPI=300）
        │     │     → 生成图片文件路径列表
        │     │
        │     ├── 2. 对每个图片，使用Read工具读取（Read工具支持多模态识别）
        │     │     → Read(filepath) 返回图片的视觉识别结果
        │     │
        │     ├── 3. 从识别结果中提取结构化数据（银行名/账号/余额/资产名/原值等）
        │     │     → 按文件用途执行对应提取逻辑
        │     │
        │     └── 4. 将提取结果持久化到 _dt_cache/（DT-130）
        │
        └── 🚨 绝对禁止的行为：
              ├── ❌ 以"扫描件"为由标注"PDF数据源缺失"并跳过
              ├── ❌ 以"需人工处理"为由回避提取
              ├── ❌ 以"OCR未安装"为由放弃提取
              ├── ❌ 以"Python脚本全部失败"为由放弃提取
              └── ❌ 在未执行Read多模态识别前就进入Phase 0
```

**多模态Read提取的操作示例**：

```python
# Step A: Python脚本提取失败后，PDF转图片
import sys
sys.path.insert(0, os.path.expanduser('~/.codex/skills/valuation-detail-table/valuation-common/scripts'))
from bank_statement_extract import pdf_to_images, generate_multimodal_prompt

images = pdf_to_images('工行12月份.pdf', output_dir='工行12月份_images', dpi=300)
# → ['工行12月份_images/page_001.png', '工行12月份_images/page_002.png']

# Step B: Agent使用Read工具逐页读取图片（Agent在下一步执行）
# Read(filepath='工行12月份_images/page_001.png')
# → Agent视觉模型自动识别图片内容

# Step C: Agent从视觉识别结果提取结构化数据
# 例：识别到 "中国工商银行 账号1703020109200023279 余额6,185,060.35"
# → bank_name='工商银行', account_no='1703020109200023279', ending_balance=6185060.35

# Step D: 将结果持久化到_dt_cache/（DT-130）
import json
result = {
    "_meta": {"rule": "DT-132", "created_at": "...", "source_step": "Step -1.5"},
    "data": {
        "filepath": "工行12月份.pdf",
        "strategy": "agent_multimodal_read",
        "status": "extracted",
        "bank_name": "工商银行",
        "account_no": "1703020109200023279",
        "ending_balance": 6185060.35
    }
}
```

**预提取的价值**：
- Phase 0/2填写时可直接引用已提取数据，无需临时读取
- 提前发现PDF/图片质量问题（模糊、扫描歪斜），及时告知用户
- 固定资产卡片台账PDF等大型文件提前解析，避免Phase 2时阻塞
- 提取结果JSON持久化，跨Phase复用，避免重复提取

### Step -1.6 🚨 PDF提取完整性硬Gate [DT-108] + 提取质量验证 [DT-134]

> **⚠️ v3.14新增DT-108硬约束，v3.42新增DT-134提取质量维度**
> **DT-108根因**：河南平绿项目Phase -1材料清单中24份银行对账单PDF标注了"✅已识别"，但实际从未提取数据。Phase 2填写银行存款科目时，因PDF未提取，直接用科目余额表汇总数填写，导致银行存款明细缺失16个账户。
> **DT-134根因**：v3.14修复后，Phase -1确实执行了extract_pdf()提取raw_text，完整性报告显示25/25="extracted"。但银行对账单提取结果仅含raw_text，缺少bank_name/account_no/ending_balance结构化字段——"文本已提取"≠"结构化数据已提取"。Phase 2因无可用结构化数据而用汇总数填充，违反DT-104/DT-65/DT-109三条红线。

**核心问题演进**：
- **v3.14前**：DT-107只要求"预提取"但不要求"提取结果非空"→"已识别"≠"已提取"
- **v3.14后**：DT-108要求"提取结果非空"→raw_text非空即PASS→**"文本已提取"≠"结构化数据已提取"**
- **v3.42修复**：DT-134增加"提取质量"维度→银行对账单必须含结构化字段才算PASS

**硬Gate规则（DT-108 + DT-134合并）**：

```
Step -1.5 预提取执行完毕后
  │
  ├── 第一层：DT-108 基础完整性检查（提取结果非空）
  │     │
  │     ├── 对每个已预提取的PDF/图片文件，检查提取结果
  │     │     │
  │     │     ├── 提取结果有实质数据（金额/账号/名称等） → 进入第二层
  │     │     │
  │     │     ├── 提取结果为空或仅含页码/标题 → RETRY → 重新提取（最多2次）
  │     │     │     │
  │     │     │     ├── 重试后仍为空 → 🚨 MUST执行DT-132多模态Read兜底
  │     │     │     │     │
  │     │     │     ├── 多模态Read提取成功 → 进入第二层
  │     │     │     │
  │     │     │     └── 多模态Read也失败 → 标注"提取失败"+原因，输出WARNING
  │     │     │
  │     │     └── 文件打开失败 → 标注"打开失败"+原因，输出WARNING
  │     │
  │     └── raw_text字符数<100 → 等同"提取结果为空" → RETRY
  │           （河南平绿项目工行12月份.pdf仅提取1字符，被判为"extracted"）
  │
  ├── 第二层：DT-134 提取质量检查（结构化字段完整性）
  │     │
  │     ├── 仅对"有场景化提取要求"的文件类型执行：
  │     │
  │     ├── 银行对账单/保证金对账单：
  │     │     MUST包含 bank_name + account_no + ending_balance
  │     │     │
  │     │     ├── 三个字段全部存在且ending_balance>0 → PASS(质量合格)
  │     │     ├── 部分字段缺失 → PARTIAL(提取质量不足) → 回退执行Step -1.5分发逻辑中的回退流程
  │     │     │     ├── 从raw_text解析结构化字段 → 解析成功 → PASS
  │     │     │     ├── 从raw_text解析失败 → DT-132多模态Read兜底 → 成功 → PASS
  │     │     │     └── 全部回退失败 → FAIL
  │     │     └── 无结构化字段(raw_text only) → PARTIAL → 同上回退流程
  │     │
  │     ├── 固定资产卡片台账：
  │     │     MUST包含 items数组(至少1项，每项含name+original_value或net_value)
  │     │     → 类似银行对账单的PARTIAL回退逻辑
  │     │
  │     └── 其他PDF：无结构化字段要求 → raw_text非空即PASS
  │
  └── 生成PDF提取完整性报告（含提取质量列）
```

**⚠️ DT-132强制补充**：原DT-108-3条款"提取失败不阻塞流程"已被DT-132部分覆盖。新逻辑为：
- Python脚本提取失败 → **MUST先执行多模态Read兜底** → 仍然失败 → 才可标注"提取失败"并继续
- **禁止在未执行多模态Read兜底的情况下，直接以"扫描件""需人工处理"为由跳过**
- 只有在多模态Read也确认无法识别（图片严重模糊/损坏/内容与评估无关）时，才允许标注"提取失败"

**提取结果非空判定标准 + 提取质量标准**：

| PDF用途 | 非空标准（至少满足一项） | 提取质量标准（DT-134） | 质量不足时 |
|---------|----------------------|---------------------|-----------|
| 银行对账单 | 检测到银行名称/账号/余额数字 | MUST含bank_name+account_no+ending_balance | 缺任一字段=PARTIAL→回退提取 |
| 保证金对账单 | 检测到银行名称/账号/余额数字 | MUST含bank_name+account_no+ending_balance | 同上 |
| 固定资产卡片台账 | 检测到资产名称/原值/净值数字 | MUST含items数组(≥1项) | items为空=PARTIAL→回退提取 |
| 合同扫描件 | 检测到合同编号/金额/签约方名称 | 无结构化字段要求 | raw_text非空即PASS |
| 其他PDF | 检测到任意非空白文本内容 | 无结构化字段要求 | raw_text非空即PASS |

**完整性报告输出格式（v3.42新增"提取质量"列）**：

```
📊 PDF提取完整性报告

| 类别 | 文件数 | PASS | PARTIAL | FAIL | 未提取 |
|------|--------|------|---------|------|--------|
| 银行对账单 | 6 | 4 | 2 | 0 | 0 |
| 保证金对账单 | 4 | 4 | 0 | 0 | 0 |
| 卡片台账 | 1 | 1 | 0 | 0 | 0 |
| 合计 | 11 | 9 | 2 | 0 | 0 |

⚠️ PARTIAL文件清单：
  - 中行对账单.pdf: 缺ending_balance（raw_text有余额信息但未解析）
  - 工行12月份.pdf: 缺bank_name+account_no+ending_balance（扫描件，需DT-132兜底）

PARTIAL项数 > 0 → 🚨 CRITICAL → 禁止进入Phase 0
-- 或 --
✅ 全部PDF提取完毕且质量合格，可进入Phase 0
```
|------|--------|------|------|--------|
| 银行对账单 | 24 | 24 | 0 | 0 |
| 保证金对账单 | 4 | 4 | 0 | 0 |
| 卡片台账 | 1 | 1 | 0 | 0 |
| 合计 | 29 | 29 | 0 | 0 |

✅ 全部PDF提取完毕，可进入Phase 0
-- 或 --
🚨 有 N 份PDF未提取，禁止进入Phase 0
```

**关键约束**：
- **DT-108-1**：PDF"已识别"≠"已提取"。"已识别"仅表示文件存在且格式正确。"已提取"表示文件内容已被读取并产生了可引用的数据。材料清单状态列MUST区分这两个状态。
- **DT-108-2**：Phase -1结束时，未提取PDF数必须=0。>0 则CRITICAL，禁止进入Phase 0。
- **DT-108-3**：提取失败（重试2次后仍为空）的PDF不阻塞流程，但MUST在Phase 2填写对应科目时标注"PDF数据源缺失"，并在勾稽核对时输出差异WARNING。
- **DT-134-1**：银行对账单"文本已提取"≠"结构化数据已提取"。raw_text非空但缺少bank_name/account_no/ending_balance结构化字段=提取质量不足=PARTIAL。PARTIAL项数>0=CRITICAL，禁止进入Phase 0。
- **DT-134-2**：银行对账单raw_text字符数<100（如扫描件仅1字符）=等同"提取结果为空"，必须触发DT-132多模态Read兜底。
- **DT-134-3**：PARTIAL文件MUST按Step -1.5分发逻辑中的回退流程执行回退提取，回退成功→PASS；回退失败→FAIL。禁止将PARTIAL状态带入Phase 0。

### Step -1.7 🚨 中间数据强制持久化 [DT-130]

> **⚠️ 本步骤为v3.42新增硬约束，解决评估明细表填写任务上下文空间耗尽导致中间数据丢失的问题**
> **根因**：河南平绿项目Phase -1/Phase 0执行过程中产生大量中间数据（科目余额表解析结果、辅助余额表提取结果、PDF提取结果、重分类映射等），这些数据仅在Python变量/对话上下文中存在。当对话上下文耗尽截断时，所有中间数据丢失，后续Phase无法继续，必须从头重新提取。整个Phase -1+0的执行时间约30-50次工具调用，重新提取浪费大量时间。

**核心问题**：评估明细表填写任务涉及100+个Sheet、20+个PDF文件、14个辅助余额表、6家银行对账单，数据提取产生的中间数据量大。对话上下文空间有限（约60-70%被Phase -1/0消耗），Phase 1-5（约30-50次工具调用）大概率在Phase 2中途触及上下文上限。上下文截断=中间数据全丢=必须重新开始。

**DT-130强制规则**：

```
Phase -1/Phase 0 每个关键数据提取步骤完成后
  │
  ├── 将提取结果持久化到项目文件夹内的 _dt_cache/ 子目录
  │     │
  │     ├── 项目文件夹 = 用户提供的财务资料所在文件夹
  │     │     例：C:\Users\Administrator\Desktop\1-河南平绿\
  │     │
  │     └── JSON文件保存路径 = {项目文件夹}/_dt_cache/{文件名}.json
  │           例：C:\Users\Administrator\Desktop\1-河南平绿\_dt_cache\subjects.json
  │
  └── 后续Phase引用数据时，从磁盘JSON读取而非依赖上下文变量
        │
        ├── JSON文件存在 → 直接加载，跳过提取步骤
        │
        └── JSON文件不存在 → 执行提取+持久化
```

**必须持久化的数据清单（Phase -1部分）**：

| 序号 | JSON文件名 | 内容 | 产生步骤 | 消费阶段 |
|------|-----------|------|---------|---------|
| 1 | `file_manifest.json` | 完整文件清单（路径+类型+用途+提取状态） | Step -1.3 | Phase 0 DT-103 |
| 2 | `pdf_extraction_{文件名}.json` | 每个PDF的提取结果 | Step -1.5 | Phase 0 DT-109, Phase 2 |
| 3 | `pdf_completeness_report.json` | PDF提取完整性报告 | Step -1.6 | Phase 0 门控 |

**JSON文件格式规范**：

```json
{
  "_meta": {
    "rule": "DT-130",
    "created_at": "2026-05-23T20:00:00",
    "project_dir": "C:\\Users\\Administrator\\Desktop\\1-河南平绿",
    "source_step": "Step -1.3"
  },
  "data": {
    // 实际数据内容
  }
}
```

**DT-130核心约束**：
- **DT-130-1**：Phase -1/Phase 0的每个关键数据提取步骤完成后，MUST立即将结果持久化到 `{项目文件夹}/_dt_cache/` 目录下的JSON文件。**提取完成≠持久化完成**——Python变量中有数据不等于数据已安全保存。
- **DT-130-2**：JSON文件统一保存在用户提供的财务资料所在文件夹内的 `_dt_cache/` 子目录。不保存到桌面、临时目录或其他位置。项目资料自包含，新对话只需项目文件夹路径即可恢复全部上下文。
- **DT-130-3**：后续Phase（1-5）引用Phase -1/0的数据时，MUST优先从磁盘JSON加载。如果JSON存在且内容完整，跳过提取步骤直接使用缓存。如果JSON不存在或内容不完整，执行提取+持久化。
- **DT-130-4**：新对话恢复时，Agent只需知道项目文件夹路径，执行 `ls _dt_cache/` 检查已有缓存文件，按需加载即可从断点继续。**无需重新执行Phase -1/0的提取步骤**。
- **DT-130-5**：`_dt_cache/` 目录下的JSON文件为过程文件，不属于交付物。Phase 5交付时不需要删除，但MUST在交付清单中注明"过程缓存文件未交付"。

## 输出

- 材料文件夹路径
- 完整文件清单（含文件名、路径、类型、预判用途、**提取状态**[已识别/已提取/PARTIAL/提取失败]、**提取质量**[PASS/PARTIAL/FAIL]）
- PDF/图片预提取结果（如有）
- PDF提取完整性报告（DT-108）+ 提取质量报告（DT-134）
- 必需材料完整性检查结果
- 可选材料缺失告警（如有）
- **中间数据持久化文件**（DT-130，保存于 `{项目文件夹}/_dt_cache/`）

## 约束

> **v2.0规则编入步骤声明**：以下规则已编入对应操作步骤，Agent执行到该步即自动生效，无需额外记忆。约束区仅保留引用索引，不再重复操作段内容。

| 规则 | 编入步骤 | 核心要点 |
|------|---------|---------|
| DT-105 材料集中 | Step -1.1~-1.2 | 用户未集中材料时MUST输出提示，不跳过 |
| DT-106 全量读取 | Step -1.3 | 所有文件MUST尝试读取，禁止跳过 |
| DT-107 非Excel预提取 | Step -1.5 | PDF/图片MUST预提取，失败标WARNING |
| DT-108 PDF完整性硬Gate | Step -1.6 | 未提取PDF>0=CRITICAL=禁止进入Phase 0 |
| DT-134 提取质量维度 | Step -1.6 | 银行对账单MUST含结构化字段，仅raw_text=PARTIAL |
| DT-132 多模态Read兜底 | Step -1.5 | 脚本失败后MUST用Read工具，禁止以"扫描件"跳过 |
| DT-142 Poppler前置检查 | Step -1.5开头 | 执行前MUST检查poppler，不可用时配置PATH |
| DT-130 中间数据持久化 | Step -1.7 | 提取完成后MUST立即持久化到_dt_cache/ |
| DT-133 倒序查找+找到即停 | Step -1.5银行对账单 | 从最后一页向前查找余额，找到即停 |

**通用原则**：
- 必需材料缺失=输出缺失清单+标注[待核实]+继续执行（DT-151），不暂停等待
- 用户明确表示"材料已全"后进入Phase 0（材料路径已知时自动扫描，无需确认）
- **后续复盘经验**：新增的提取教训/规则直接写入Step -1.5/-1.6操作段，不在约束区单独列示

## 异常处理

- 文件夹路径不存在 → 输出提示+标注[待核实]+继续
- 文件夹为空 → 输出提示+标注[待核实]+继续
- PDF/图片读取失败 → 标注"预提取失败"+原因，输出WARNING，Phase 2时再次尝试
- 文件加密/密码保护 → 标注"需人工解锁"，不跳过

## 流转

- 材料齐全 → Phase 0（输入确认与数据源解析）
- 必需材料缺失 → 输出缺失清单+标注[待核实]+继续执行（DT-151）
