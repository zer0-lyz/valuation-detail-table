# valuation-company-basic-info 优化方案

## 现状问题

1. **技术栈不一致**：模块A使用docx-js（Node.js），模块B/C使用python-docx
2. **单文件过大**：1150行，3个独立模块混在一个SKILL.md中
3. **维护困难**：修改任一模块需在巨大文件中定位

## 推荐方案：统一为python-docx + 保持单skill

### 理由
- 拆分为3个skill会破坏现有触发逻辑和用户习惯
- python-docx是项目标准工具链，docx-js是唯一例外
- 统一后可复用docx公共工具函数

### 迁移步骤（待执行）
1. **Phase 1**：将模块A的docx-js代码转为python-docx实现
2. **Phase 2**：提取python-docx公共工具函数到`references/docx_utils.py`
3. **Phase 3**：验证三个模块的输出格式一致性

### 风险控制
- 迁移前备份现有skill
- 逐模块迁移并验证输出
- 保留docx-js版本作为回退方案直到python-docx版本验证通过

## 不推荐方案：拆分为3个独立skill

- 破坏用户习惯：用户已习惯"写基本信息"触发全部3个模块
- 增加配置复杂度：3个skill需独立配置MCP连接
- 数据流断裂：模块A的企查查数据可能被B/C引用

---

_本方案待用户确认后执行。记录于2026-05-18。_
