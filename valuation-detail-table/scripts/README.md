# 评估明细表填写 - 脚本索引

## 脚本列表

| 脚本 | 用途 | 调用时机 |
|------|------|---------|
| `validate_sheet_after_fill.py` | Phase 2 per-sheet即时验证（DT-74） | 每填完一个Sheet后立即运行 |
| `validate_sheet_after_fill.py` (全量模式) | Phase 4.5自动化验证门控 | Phase 3完成后 |
| `validate_numbering.py` | 编号唯一性校验（R/教训/DT编号+虚空引用检测） | 版本迭代后/Skill文件修改后 |

### validate_sheet_after_fill.py

**调用方式**：
```bash
# 验证单个Sheet（Phase 2即时门控）
python validate_sheet_after_fill.py <xlsx_path> <sheet_name>

# 验证所有明细表Sheet（Phase 4.5门控）
python validate_sheet_after_fill.py <xlsx_path>
```

**检查项**：
| 检查 | 级别 | 规则 | 说明 |
|------|------|------|------|
| 1 | R(红线) | DT-67 | 公式列未被数值覆写 |
| 2 | R(红线) | DT-46 | 发生日期列不含文字/业务内容列不含日期序列号 |
| 3 | R(红线) | DT-66 | 列位校验（实际列数与已知结构对比） |
| 4 | W(警告) | DT-0 | 数据行无"待补充"/"待核对"占位符 |
| 5 | R(红线) | DT-35 | 合计行唯一性（无重复合计行） |
| 6 | W(警告) | DT-44 | 数据行与合计行之间无空白无格式行 |
| 7 | W(警告) | DT-3 | 关键数值列格式检查 |
| 8 | R(红线) | DT-76 | 增值额/增值率列格式（禁止General） |
| 9 | R/W | DT-77 | 行高统一性（众数占比<60%为红线，<80%为警告） |
| 10 | W(警告) | DT-78 | 合计行下方无残留边框/格式 |
| 11 | R(红线) | DT-82① | 数据行首行无空白跳过（A/B列均空=首行被跳过） |
| 12 | W(警告) | DT-82② | 数据区空白行应有thin边框和公式 |
| 13 | R(红线) | DT-84 | 合计/减值/小计行A列必须center对齐 |

**返回码**：
- 0 = 全部通过
- 1 = 有红线问题，必须修复
- 2 = 有警告，建议修复

### validate_numbering.py

**调用方式**：
```bash
# 默认自动查找Skill目录
python validate_numbering.py

# 指定Skill目录
python validate_numbering.py --skill-dir <path>

# 输出修复建议
python validate_numbering.py --fix-hints
```

**检查项**：
| 检查 | 说明 |
|------|------|
| R编号重复 | 同一R编号被多个DT规则使用 |
| R编号乱序 | R编号未按递增顺序排列 |
| 教训编号重复 | 同一教训编号出现多次 |
| 教训编号间隙 | 教训编号跳号（已知33/34不报告） |
| DT编号重复 | 同一DT编号被多次定义 |
| DT编号大间隙 | 连续缺失>5个DT编号 |
| 虚空文件引用 | 引用了不存在的.md/.py文件 |
| 虚空DT引用 | Step/CHECK文件引用了SKILL.md中无定义的DT编号 |
| Step文件引用 | FLOW.md/SKILL.md引用了不存在的Step文件 |

**返回码**：
- 0 = 全部通过
- 1 = 有错误（重复/虚空引用）
- 2 = 仅有警告（间隙/乱序）

## 依赖工具

### Poppler（PDF转图片前置依赖）

**用途**：`pdf2image` 库的底层依赖，用于将PDF转为PNG图片（DT-132多模态Read兜底、OCR提取均需要）  
**影响范围**：`valuation-common/scripts/bank_statement_extract.py` 的 `pdf_to_images()` 和 `valuation-common/scripts/pdf_extract.py` 的OCR策略

**Skill内提供的两个包**：

| 包 | 类型 | 路径 | 说明 |
|---|------|------|------|
| **Release-26.02.0-0** | ✅ 预编译Windows二进制（直接可用） | `scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin/` | 含 `pdftoppm.exe` 及所有DLL依赖，**Windows首选** |
| **poppler-26.05.0** | ⚠️ 源码包（需编译） | `scripts/poppler-26.05.0/` | 含CMakeLists.txt/cpp源码，Linux/macOS可从此编译，Windows不建议 |

**快速配置（Windows）**：

```bash
# 1. 将预编译poppler的bin目录加入PATH（每次会话需执行，或写入profile永久生效）
SKILL_DIR="$HOME/.workbuddy/skills/valuation-detail-table"
export PATH="$SKILL_DIR/scripts/Release-26.02.0-0/poppler-26.02.0/Library/bin:$PATH"

# 2. 验证可用
pdftoppm -v 2>&1
# 应输出：pdftoppm version 26.02.0 ...

# 3. 验证pdf2image可用
python -c "from pdf2image import convert_from_path; print('pdf2image OK')"
```

**在Python脚本中自动配置（推荐）**：

```python
import os, shutil

def ensure_poppler():
    """确保poppler可用，自动配置PATH。优先使用Release预编译包。"""
    if shutil.which('pdftoppm'):
        return True
    
    skill_dir = os.path.join(os.path.expanduser('~'), '.workbuddy', 'skills', 'valuation-detail-table')
    
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

**常见问题**：

| 问题 | 原因 | 解决 |
|------|------|------|
| `PDFInfoNotInstalledError` | poppler未加入PATH | 执行 `export PATH=...` 或使用Python `ensure_poppler()` |
| `pdftoppm` not found | 同上 | 同上 |
| pdf2image导入失败 | Python包未安装 | `pip install pdf2image` |
| poppler-26.05.0目录无bin/ | 该目录为源码包，非预编译 | 使用 `Release-26.02.0-0` 预编译包代替 |

**注意事项**：
- `Release-26.02.0-0` 为预编译Windows x64二进制包，**不适用于Linux/macOS**
- Linux/macOS需从 `poppler-26.05.0` 源码包编译，或通过系统包管理器安装（`apt install poppler-utils` / `brew install poppler`）
- Python脚本调用 `pdf_to_images()` 前，MUST确保poppler PATH已配置（参见S-1_prep.md Step -1.5 DT-142）
- 预编译包含 `pdftoppm.exe`、`pdftotext.exe`、`pdfinfo.exe` 等工具，以及所需DLL

## 共享脚本

- `../../valuation-common/scripts/` 下的通用脚本可复用

## 内嵌代码位置

| 功能 | 位置 |
|------|------|
| 科目余额表解析 | steps/S0_input.md |
| 科目映射表建立 | steps/S1_structure.md |
| 往来科目填写 | steps/S2_fill_re.md |
| 序时账核实（T51-T55/T60） | steps/S2_seq_verify.md |
| 公式修复与格式修复 | steps/S3_format.md |
| 隐藏汇总表联动检查（T61/T62） | steps/S4_linkage.md |
| COM重算保存+datetime修复 | steps/S5_deliver.md |
| 自动化验证门控 `validate_detail_table()` | steps/S3_format.md → Phase 4.5 |
