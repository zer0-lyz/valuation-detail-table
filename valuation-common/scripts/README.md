# 评估Skill共享脚本索引

> 本目录存放所有评估skill可复用的通用脚本。

---

## 脚本列表

| 脚本 | 用途 | 调用skill |
|------|------|----------|
| `shared_checks.py` | 通用检查函数库v1.0：check1列递推（含半绝对列引用检测）、check2辅助语义（含AVERAGE占比检测）、工具函数、语义判定 | DCF/设备/房地产/市场法/土地 |
| `gate_validator.py` | MR-17硬约束关卡验证 v2.0（含CCEP合规检查MR-18） | 所有审核类skill |
| `common_compliance.py` | MR-18 CCEP合规验证脚本 | 所有评估skill |
| `shared_utils.py` | 通用工具函数库v1.0：get_sheet_prefix（sheet前缀提取）、compact_process_table（过程表空白行清除） | DT/AB |

### shared_checks.py 导入方式

```python
# 从各skill的scripts/目录导入（check_formulas.py / check_*_formulas.py 已内置路径）
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'valuation-common', 'scripts'))
from shared_checks import check1_column_progression, check2_auxiliary_semantics, determine_semantic
```

### shared_utils.py 导入方式

```python
# 从各skill导入共享工具函数
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/valuation-common/scripts'))
from shared_utils import get_sheet_prefix, compact_process_table
```

### shared_utils.py 核心能力

| 函数 | 说明 | 来源 |
|------|------|------|
| `get_sheet_prefix(sheet_name)` | 从sheet名提取前缀编码（如'3-5应收账款'→'3-5'） | DT Skill gate_validator.py / validate_sheet_after_fill.py 去重提取 |
| `compact_process_table(ws)` | 过程表空白行彻底清除+合计行上移+公式重写+打印区域设置 | AB Skill compact_process_table.py 去重迁移 |

### gate_validator.py 导入方式

```python
# 从各skill调用共享gate_validator
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'valuation-common', 'scripts'))
from gate_validator import validate_gate, validate_all
```

### gate_validator.py v2.0 新增能力

| 能力 | 说明 |
|------|------|
| CCEP合规检查 | 自动检查checkpoint中是否包含`common_compliance`字段，缺失=BLOCKED |
| 共享化改造 | 各skill不再需要物理副本，通过`sys.path.insert`引用 |

### shared_checks.py 核心能力

| 函数 | 说明 | 关键修复 |
|------|------|---------|
| `check1_column_progression` | 检查预测列基底值递推是否锁定 | v5.0新增`semi_abs_base_usage`追踪`$O15`半绝对列引用 |
| `check2_auxiliary_semantics` | 检查辅助列语义（占比vs增长率） | v5.0新增AVERAGE占比检测+`wb_val_full`跨Sheet值读取 |
| `determine_semantic` | 语义判定（ratio/growth/ambiguous） | v5.0新增数值范围启发式(0.05-0.50)+成本/费用标签关键词 |

## 待开发脚本

| 脚本 | 用途 | 优先级 |
|------|------|--------|
| `format_fix.py` | 格式修复（字体/字号/边框） | 中 |

## 使用规范

1. 先读取本README判断是否可用已有脚本（对应T18/G1规则）
2. 新脚本开发后必须更新本索引
3. 脚本文件头部必须包含：输入/输出/依赖/使用示例
4. 脚本中不得包含客户/项目真实信息（MR-10）
5. 修改shared_checks.py后必须对所有调用skill做回归测试
6. 修改gate_validator.py后需验证所有审核类skill的关卡定义兼容性

---

_本文件最后更新：2026-05-23_
