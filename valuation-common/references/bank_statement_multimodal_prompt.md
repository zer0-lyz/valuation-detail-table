# 银行对账单多模态识别Prompt模板

> 用途: Agent读取银行对账单PDF图片时，使用此prompt引导视觉模型提取结构化数据
> 配套: `scripts/bank_statement_extract.py` → `generate_multimodal_prompt()`

---

## 一、通用银行对账单识别Prompt

```
请识别这张银行对账单图片，提取以下关键信息：

请按以下JSON格式输出提取结果（每页一个对象）:

```json
[{
  "page": 1,
  "bank_name": "银行全称",
  "account_no": "完整账号（连续数字，不要拆分）",
  "account_name": "账户名称/户名",
  "account_type": "活期/保证金/资金监管/贷款",
  "ending_balance": "期末余额（纯数字，含小数，不含逗号）",
  "currency": "CNY/USD/HKD",
  "period": "对账期间，如2025年12月",
  "notes": "任何识别疑问或特殊情况说明"
}]
```

**特别注意:**
1. 账号必须完整连续，不要因为换行而拆分。如果账号跨行显示，请拼接为完整号码。
2. 期末余额是最终余额（通常是最后一行的余额），不是中间发生额。
3. 如果同一页有多个账户，请分别提取每个账户的信息。
4. 保证金账户请在account_type中标注"保证金"。
5. 如果是扫描件/图片模糊，请在notes中说明。
6. 数字中的逗号是千分位分隔符，提取时去除。

如果无法识别某些字段，请填null并说明原因。
```

---

## 二、建行专项Prompt

建行对账单特点是账号和余额跨行拆分，需要额外强调:

```
请识别这张建设银行对账单图片。

**建行PDF特殊格式提醒:**
- 账号可能跨行显示，如第1行显示"41050172"，第2行显示"56080000"，第3行显示"1913"
  请将这些片段拼接为完整账号: 41050172560800001913
- 余额可能跨行显示，如"298,335.8"在第1行，"4"在第2行
  请拼接为完整数字: 298335.84
- "保证金"字样可能出现在户名行或账户类型中

请提取:
1. 完整账号（拼接所有行片段）
2. 期末余额（拼接跨行数字，注意小数点位置）
3. 账户类型（活期/保证金）
4. 户名

按以下JSON格式输出:
```json
[{
  "account_no": "完整20位账号",
  "account_name": "户名",
  "account_type": "活期或保证金",
  "ending_balance": 123456.78,
  "notes": ""
}]
```
```

---

## 三、中行专项Prompt

中行存在多种格式，需要先判断类型:

```
请识别这张中国银行对账单图片，首先判断其类型:

**格式A - 综合汇总表**: 页面含竖线分隔的表格，多行多列，每行一个账户
**格式B - 明细对账单**: 页面顶部有"账号""账户类型""账户名称"等字段
**格式C - 贷款明细**: 页面含"贷款""LPR"等字样

请先判断类型，然后按对应格式提取:

格式A提取: 序号、账号、产品类型、余额、户名
格式B提取: 账号、账户类型、户名、本对账期末余额
格式C提取: 标记为"贷款明细，无期末余额"，跳过

按以下JSON格式输出:
```json
[{
  "format_type": "A/B/C",
  "account_no": "完整账号",
  "account_name": "户名",
  "account_type": "活期/保证金/资金监管/贷款",
  "ending_balance": 123456.78,
  "notes": ""
}]
```
```

---

## 四、工行扫描件专项Prompt

工行对账单常见扫描件，图片可能倾斜/模糊:

```
请识别这张工商银行对账单扫描件图片。

由于是扫描件，可能出现:
- 图片倾斜: 请仔细辨认文字
- 数字模糊: 重点关注余额数字的精确度
- 印章遮挡: 如有印章遮挡关键信息，请在notes中说明

请提取:
1. 账号（通常19位）
2. 户名
3. 期末余额（精确到分）
4. 对账期间

如果因扫描质量问题无法确认某个数字，请在该字段填null并在notes中说明。

按以下JSON格式输出:
```json
[{
  "account_no": "完整账号或null",
  "account_name": "户名或null",
  "ending_balance": 123456.78或null,
  "period": "对账期间",
  "confidence": "high/medium/low",
  "notes": "识别疑问说明"
}]
```
```

---

## 五、使用方法

### 方式1: Agent直接使用（推荐）

```python
from bank_statement_extract import extract_bank_statement, pdf_to_images, generate_multimodal_prompt

# Step 1: 尝试pdfplumber提取
result = extract_bank_statement('建行对账单.pdf')

# Step 2: 如果需要多模态
if result['needs_multimodal']:
    # Agent读取图片并使用prompt
    images = result['multimodal_images']
    prompt = result['multimodal_prompt']
    # → Agent用Read工具读取图片 + 视觉模型分析
```

### 方式2: CLI批量生成图片

```bash
# 将PDF转为图片（供手动上传到多模态模型）
python bank_statement_extract.py 工行12月份.pdf --images-only --dpi 300

# 仅pdfplumber提取（禁用多模态）
python bank_statement_extract.py 对账单文件夹/ --no-multimodal

# 完整混合提取
python bank_statement_extract.py 对账单文件夹/ --output-dir ./output/
```

### 方式3: 编程调用

```python
from bank_statement_extract import (
    pdf_to_images, 
    generate_multimodal_prompt,
    parse_multimodal_result
)

# 转图片
images = pdf_to_images('工行12月份.pdf', dpi=300)

# 生成prompt
prompt = generate_multimodal_prompt('工行12月份.pdf', len(images), bank_hint='工商银行')

# Agent识别后，解析结果
records = parse_multimodal_result(agent_response_text, source_file='工行12月份.pdf')
```

---

## 六、识别结果校验

多模态识别后，务必进行以下校验:

1. **账号位数校验**: 建行20位、中行12-16位、工行19位、交行17-21位
2. **余额合理性**: 不会为极端值（负1亿以上或正100亿以上）
3. **账号去重**: 同一账号出现在多个PDF中时，只保留一条记录
4. **与科目余额表交叉验证**: 提取总额 vs 科目余额表银行存款余额

```python
from bank_statement_extract import validate_balance, deduplicate_accounts

# 校验余额
check = validate_balance(1234567.89)
if not check['valid']:
    print(f"余额异常: {check['warnings']}")

# 去重
unique = deduplicate_accounts(all_records)
```

---

_最后更新: 2026-05-23_
