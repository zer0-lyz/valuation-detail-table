# Phase 4a：BS数据校验

> 适用规则：DT-182、DT-216、DT-218。
> 本步骤采用低自由度执行：MUST调用统一函数，禁止另写脚本直接回填`2-分类汇总`金额。

## 目标

构建可追踪、可重算的BS校对链：

```text
_dt_cache/bs_balances.json
  -> 隐藏Sheet _BS对照（结构化金额与来源）
  -> 2-分类汇总 I列（公式链接）
  -> 2-分类汇总 J列（差异公式）
```

`2-分类汇总`属于汇总表。根据DT-182，其数据区域禁止直接写入硬编码金额。

## 执行

正常流程由`dt_runner.py`自动调用：

```python
from dt_runner import _fill_classification_summary_I_column

_fill_classification_summary_I_column(wb, cache_dir)
```

函数职责：

1. 从`_dt_cache/bs_balances.json`读取BS科目。
2. 重建隐藏`_BS对照`表，写入标准科目、科目类型、年初余额、期末余额和来源科目。
3. 为`2-分类汇总`明细科目行写入I列`VLOOKUP`公式和J列差异公式。
4. 为汇总、总计、净资产行保留模板公式。
5. 将分类汇总未覆盖的BS原始科目追加到`_BS对照`并标记为`[新增]`，避免静默遗漏。
6. 将未匹配科目在`_BS对照`来源列标记为`[未匹配]`。

## 禁止事项

- 禁止向`2-分类汇总`I列或J列直接写入金额数值。
- 禁止复制历史版本中的“逐行硬编码I列”脚本。
- 禁止跳过`G-DT182`门控。
- 禁止把`_BS对照`取消隐藏后作为交付可见表。

## 验证

```bash
python3 valuation-detail-table/scripts/dt_runner.py \
  --phase gate \
  --gate G-DT182 \
  --project <项目文件夹路径> \
  --xlsx-path <评估明细表路径>
```

交付前还需确认：

- `2-分类汇总`明细科目行I列均为公式。
- `2-分类汇总`J列差异公式完整。
- `_BS对照`存在且保持隐藏。
- `_BS对照`中的`[未匹配]`和`[新增]`项目已纳入差异说明或待确认清单。

## 版本

v2.0 (2026-06-01)：改为`_BS对照`结构表 + 公式链接，消除与DT-182冲突的历史硬编码路径。
