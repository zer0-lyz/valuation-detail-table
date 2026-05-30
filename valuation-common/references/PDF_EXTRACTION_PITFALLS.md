# PDF提取踩坑经验汇总

> 来源: 河南平煤神马平绿置业评估项目 (2026-05)
> 适用: 银行对账单PDF提取、评估明细表编制中的PDF数据处理
> 配套脚本: `scripts/bank_statement_extract.py` (v2.0 混合方案)

---

## 一、核心教训：pdfplumber不是万能的

### 1.1 扫描件PDF是硬伤

**问题**: 工商银行12月份对账单PDF，pdfplumber提取文本为空（零输出），包括 `extract_text()` 和 `extract_tables()` 均返回空。

**根因**: 该PDF为扫描件（图片型PDF），没有底层文本层。pdfplumber只能解析有文本层的PDF。

**对策**: 
- v2.0混合方案：pdfplumber失败后自动切换到PDF转图片+多模态识别
- 多模态识别天然支持扫描件，因为视觉模型直接"看"图片
- **[DT-132] 强制兜底**：Python脚本（pdfplumber/PyMuPDF/OCR）全部失败后，MUST使用Read工具直接读取PDF/图片进行多模态视觉识别。**禁止以"扫描件""需人工处理"为由跳过提取**

**判断扫描件的方法**:
```python
# pdfplumber提取文本长度为0 → 大概率扫描件
with pdfplumber.open(pdf_path) as pdf:
    text = pdf.pages[0].extract_text()
    if not text or len(text.strip()) == 0:
        # 扫描件，需要多模态
```

### 1.2 文本跨行拆分

**问题**: 建设银行对账单PDF中，账号和余额被换行符拆分到不同文本行。

**实例**:
- 账号: 第10行 "41050172" → 第12行 "56080000" → 第14行 "1913"
  - 实际账号: 41050172560800001913
- 余额: 第30行 "298,335.8" → 第31行 "4 人民币"
  - 实际余额: 298,335.84

**根因**: PDF排版时单元格宽度不足，长数字被自动换行。pdfplumber按文本层逐行提取，不感知单元格边界。

**对策（优先级排序）**:
1. **从文件名提取账号**（最可靠）: 建行单账户PDF文件名即账号，如 `41050172560800001913.pdf`
2. **合并文本后正则匹配**: `text.replace('\n', ' ')` 后搜索完整账号模式
3. **逐行重建**: 识别 prefix1 + prefix2 + suffix，手动拼接
4. **clean_number()去除换行**: 余额拆行时 `replace('\n', '')` 可修复

### 1.3 多格式变体

**问题**: 同一家银行（中国银行）的PDF存在至少3种格式:
- **BOCCA**: 综合汇总表，第1页竖线分隔的表格
- **BOCVC**: 明细对账单，含"账号/账户类型/本对账期末余额"字段
- **BOCCC**: 贷款明细，无期末余额（需跳过）

**对策**: 文件名前缀路由，不依赖内容判断:
```python
if basename.lower().startswith('bocca'):
    return extract_boc_summary(pdf_path)
elif basename.lower().startswith('bocvc'):
    return extract_boc_detail(pdf_path)
elif basename.lower().startswith('boccc'):
    return []  # 贷款明细跳过
```

---

## 二、各银行格式特征速查

| 银行 | 文本提取可靠性 | 常见问题 | 推荐策略 |
|------|:---:|------|------|
| 建行(CCB) | ⚠️ 中 | 账号/余额跨行拆分 | 文件名提账号 + 多策略提余额 |
| 中行(BOC) | ✅ 高 | 3种格式变体 | 文件名前缀路由 |
| 工行(ICBC) | ❌ 低 | 常见扫描件 | 多模态识别 |
| 交行(BOCOM) | ✅ 高 | 格式规范 | pdfplumber直接提取 |
| 中原银行 | ✅ 高 | 格式规范 | pdfplumber直接提取 |
| 财务公司 | ✅ 高 | "本月合计"行 | pdfplumber直接提取 |

---

## 三、混合方案设计理念

### 3.1 为什么用混合方案而不是纯多模态

| 维度 | pdfplumber | 多模态识别 |
|------|------|------|
| 速度 | 快（纯文本解析） | 慢（图片渲染+模型推理） |
| 成本 | 零（本地执行） | 有（每页消耗模型token） |
| 数字精度 | 高（文本直取） | 可能有OCR误读（0/O, 3/5） |
| 扫描件支持 | ❌ | ✅ |
| 格式兼容性 | 需逐家写适配器 | 通用prompt |
| 维护成本 | 高（每家银行写代码） | 低（prompt模板） |

**结论**: pdfplumber处理80%的常规PDF（快速+精确），多模态处理20%的扫描件/格式混乱PDF（兜底）。

### 3.2 混合方案流程

```
PDF文件
  ↓
pdfplumber文本提取
  ↓
有有效结果？ ──是──→ 直接使用（完成）
  ↓ 否
PDF转图片 (pdf2image, 300dpi)
  ↓
生成多模态识别prompt
  ↓
Agent读取图片 + 视觉模型识别
  ↓
解析识别结果 → 余额合理性校验
  ↓
完成
```

### 3.3 多模态识别注意事项

1. **DPI建议300**: 200dpi以下数字可能模糊，300dpi平衡清晰度和文件大小
2. **poppler依赖**: pdf2image需要系统安装poppler
   - Windows: 下载 [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)，解压后将bin目录添加到PATH
   - Linux: `apt install poppler-utils`
   - macOS: `brew install poppler`
3. **prompt要强调**: 账号完整拼接、期末余额≠中间余额、千分位逗号去除
4. **后处理校验**: 对多模态识别结果仍需 `validate_balance()` 合理性检查

---

## 四、常见错误与修复

### 4.1 变量未定义

**错误**: v1版 `extract_ccb_balance_from_text()` 末尾 `return balance`，但函数内无 `balance` 变量定义。当所有4个策略都没匹配时，触发 `NameError`。

**修复**: v2.0中所有提取函数的 `balance` 初始化为 `None`，返回时使用 `None` 而非未定义变量。

### 4.2 余额跨行解析错误

**错误**: `clean_number("298,335.8\n4")` 直接 `float()` 会抛 ValueError。

**修复**: 在 `clean_number()` 中先 `replace('\n', '')` 再转浮点:
```python
s = s.replace('\n', '').replace('\r', '')
# "298,335.8\n4" → "298,335.84" → 298335.84
```

### 4.3 BOCCC贷款明细误提取

**错误**: BOCCC格式PDF含贷款信息但无期末余额，如果统一按明细格式提取会得到 `balance=None`。

**修复**: 文件名前缀判断直接跳过:
```python
if basename.lower().startswith('boccc'):
    return []  # 贷款明细跳过
```

### 4.4 去重逻辑导致覆盖

**问题**: 同一账号出现在综合对账单和单账户PDF中，综合对账单的余额可能不完整（多页中间余额），如果保留综合对账单的记录会覆盖更准确的单账户记录。

**修复**: 去重时优先保留来自单账户PDF（文件名含账号）的记录:
```python
is_single = key in existing['source_file']
new_is_single = key in r['source_file']
if new_is_single and not is_single:
    seen[key] = r  # 单账户优先
```

---

## 五、银行存款与评估明细表勾稽

### 5.1 总额差异处理

**问题**: 银行对账单提取的余额合计（~3700万）≠ 科目余额表银行存款余额（~3814万）≠ 资产负债表银行存款（~4320万）。

**根因**:
1. 工行扫描件无法提取（~619万缺失）
2. 其他货币资金中的保证金与银行存款分类差异（~370万）
3. 部分账户可能未提供对账单

**处理方式**: 在评估明细表中添加调整行:
- 银行存款调整行: +6,185,060.35（工行缺失+未提供对账单部分）
- 其他货币资金调整行: -3,701,445.58（保证金重复分类）

### 5.2 账户类型分类

银行存款 vs 其他货币资金的判断:
- **银行存款**: 活期、一般账户
- **其他货币资金**: 保证金、资金监管

同一银行同一账号在不同目录下的对账单（基准日对账单/保证金对账单）可能分别归属不同科目，需要根据"账户类型"字段区分。

---

## 六、扩展到其他PDF提取场景

### 6.1 固定资产卡片台账

- 通常有文本层，pdfplumber可提取
- 表格结构较规范，但列名可能有变体（"原值"/"账面原值"/"入账价值"）
- 已有: `scripts/pdf_extract.py` → `extract_asset_register()`

### 6.2 辅助余额表

- 需提取: 科目名称 + 结算对象 + 借方余额/贷方余额
- 表格可能有合并单元格，pdfplumber提取可能丢失
- 已有: `scripts/pdf_extract.py` → `extract_auxiliary_balance()`

### 6.3 通用建议

1. **先试pdfplumber**: 80%的PDF可以直接提取
2. **检查提取结果**: 文本为空 → 扫描件，需多模态
3. **数字交叉校验**: 提取的余额合计 vs 科目余额表 vs 资产负债表
4. **保留原始文本**: 方便人工核验和调试

---

## 七、依赖安装备忘

```bash
# pdfplumber (首选文本提取)
pip install pdfplumber>=0.11

# pdf2image (PDF转图片，多模态前置)
pip install pdf2image>=1.17

# poppler (pdf2image的系统依赖)
# Windows: 下载 https://github.com/oschwartz10612/poppler-windows/releases
#          解压后将bin目录添加到PATH环境变量
# Linux:   sudo apt install poppler-utils
# macOS:   brew install poppler

# Pillow (图片处理)
pip install Pillow
```

验证poppler安装:
```bash
pdftoppm -h  # 能显示帮助即安装成功
```

---

_文档维护: 当遇到新的PDF提取问题，请追加到本文档对应章节。_
_最后更新: 2026-05-23_
