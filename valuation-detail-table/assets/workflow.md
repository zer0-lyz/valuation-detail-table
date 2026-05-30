# 评估明细表填写 Skill — 完整工作流程图

```mermaid
---
title: 评估明细表自动填写 — 完整工作流
---
flowchart TD
    %% ==================== 样式定义 ====================
    classDef startEnd fill:#1a237e,color:#fff,stroke:#1a237e,stroke-width:2px,rx:12
    classDef phase   fill:#e3f2fd,color:#1a237e,stroke:#1565c0,stroke-width:2px,rx:8
    classDef subphase fill:#f3e5f5,color:#4a148c,stroke:#7b1fa2,stroke-width:1px,rx:6
    classDef gate    fill:#fff3e0,color:#e65100,stroke:#f57c00,stroke-width:2px,rx:10
    classDef script  fill:#e8f5e9,color:#1b5e20,stroke:#2e7d32,stroke-width:1px,rx:6
    classDef rule    fill:#ffebee,color:#b71c1c,stroke:#c62828,stroke-width:1px,rx:6,stroke-dasharray: 5 5
    classDef deliver fill:#1b5e20,color:#fff,stroke:#1b5e20,stroke-width:2px,rx:10

    %% ==================== 触发 ====================
    TRIGGER(("🧠 Skill 被触发\n用户提供项目路径 或 请求填写")):::startEnd

    %% ==================== 前置判断0：断点恢复 ====================
    subgraph CHECKPOINT["前置判断 0：断点恢复检测 (DT-131)"]
        direction TB
        CK_START["检查项目路径下 _dt_cache/ 是否存在"]:::gate
        CK_EXISTS{"_dt_cache/ 存在且含 JSON？"}:::gate
        CK_LOAD["读取缓存 JSON\n对比缓存完整度\n生成恢复状态报告"]:::gate
        CK_CONTINUE["从第一个缺失缓存的 Phase 开始继续执行\n（无需用户确认）"]:::gate
        CK_FRESH["缓存无过期\n直接继续"]:::gate
    end

    %% ==================== Phase -1 ====================
    subgraph PHASE_M1["Phase -1：材料准备 (S-1_prep.md)\n8条规则 | ❌ 不可跳过"]
        direction TB
        M1_1["Step -1.1\n检查材料集中状态"]:::phase
        M1_2["Step -1.2\n材料准备操作提示\n列出必需/可选材料清单"]:::phase
        M1_3["Step -1.3\n文件遍历 → 生成file_manifest\n类型识别 + 用途预判"]:::phase
        M1_5["Step -1.5\nPDF 预提取\n（多模态/poppler）"]:::phase
        M1_6["Step -1.6\nPDF 完整性报告确认"]:::phase
    end

    %% ==================== Phase 0 ====================
    subgraph PHASE_0["Phase 0：输入确认与数据源解析 (S0_input.md)\n11条规则 | ❌ 不可跳过"]
        direction TB
        P0_1["Step 0.1\n确认输入文件 & 模式判断"]:::phase
        P0_1_5["Step 0.1.5 [DT-103]\n目录递归遍历 → 逐科目对应数据源\n→ 未使用文件告警"]:::phase
        P0_2["Step 0.2\n解析科目余额表\n末级科目提取 + 银行存款子科目计数"]:::phase
        P0_3["Step 0.3 [DT-139]\n解析资产负债表\n🚨 BS 解析后强制自校验\n资产=负债+权益 ±1元"]:::phase
        P0_4["Step 0.4\n数据源映射\n→ data_source_mapping.json"]:::phase
        P0_5["Step 0.5\n辅助余额表提取\n→ auxiliary_balance_*.json"]:::phase
        P0_5A["Step 0.5a [DT-119]\nD1→D2→D3 三级递进映射\n+ 重分类检测"]:::phase
        P0_6["Step 0.6\n数据分类\n→ data_classification.json"]:::phase
        P0_7["Step 0.7 [DT-79]\n设定信息填写\n（被评估单位/基准日）"]:::phase
    end

    %% ==================== Phase 1 ====================
    subgraph PHASE_1["Phase 1：结构映射 (S1_structure.md)\n5条规则 | ❌ 不可跳过"]
        direction TB
        P1_1["Step 1.1\n解析模板 Sheet 结构"]:::phase
        P1_2["Step 1.2\n建立科目编码→Sheet 映射表\n+ column_map 列位映射"]:::phase
        P1_2A["Step 1.2a [DT-119]\nD1/D2/D3 三级映射策略"]:::phase
    end

    %% ==================== Phase 2 ====================
    subgraph PHASE_2["Phase 2：数据写入 (S2_fill_*.md)\nDT-160 禁止裸 openpyxl 写入"]
        direction TB
        P2_A["Phase 2a [S2_fill_bs.md]\n资产类科目填写\n（货币资金/应收/预付/其他应收/存货）"]:::phase
        P2_B["Phase 2b [S2_fill_re.md]\n往来科目填写\n（应收/应付/预收/其他应付）"]:::phase
        P2_C["Phase 2c [S2_fill_inventory.md]\n存货明细填写"]:::phase
        P2_D["Phase 2d [S2_fill_liability.md]\n负债类科目填写\n（短期借款/应付/预收/长期借款）"]:::phase
        P2_PIPE["🔄 三步管线（每 Sheet 执行）\n① prepare_data_rows() 组织数据\n② fill_sheet() 写入（强制接口）\n③ auto_gate_after_fill() Gate 校验"]:::script
    end

    %% ==================== Phase 3 ====================
    subgraph PHASE_3["Phase 3：序时账查阅 (S3_journal_extract.md)\n12条规则 | ⚠️ 仅2种情况可跳过"]
        direction TB
        P3_0["Step 3.0 [DT-161]\n执行/跳过判定\n有序时账→MUST执行\n无序时账/用户明确跳过→跳过"]:::gate
        P3_EXE["Step 3.1~3.7\n调用 journal_extractor.py\n① 核查发生日期 (DT-51~55)\n② 归纳业务内容 (DT-60)\n③ 同步更新成本法底稿"]:::phase
        P3_SKIP["⏭ 跳过\n→ 发生日期留空 (DT-143)\n→ 业务内容=推断文字"]:::rule
    end

    %% ==================== Phase 3 旧 ====================
    subgraph PHASE_3F["Phase 3（旧）：格式修复 (S3_format.md)\n23条规则 | 已迁移至 Phase 4"]
        direction TB
        P3F_ERR["（Phase 3 旧功能已合并到 Phase 4 格式集中处置区）"]:::rule
    end

    %% ==================== G4 Gate ====================
    subgraph GATE_G4["G4 门控：Phase 3→Phase 4"]
        G4_JUDGE{"Phase 3 执行状态？"}:::gate
        G4_PASS["✅ 已执行 → 进入 Phase 4"]:::gate
        G4_BLOCK["❌ 未执行且无序时账/用户跳过\n→ 回退 Phase 3"]:::gate
    end

    %% ==================== Phase 4 ====================
    subgraph PHASE_4["Phase 4：格式集中修复 (S4_format.md)\n23条规则 | ❌ 不可跳过"]
        direction TB
        P4_PRE["Step 4.pre [DT-161]\nPhase 3 前置 Gate 检查\n（往来科目发生日期是否已填写）"]:::gate
        P4_0["Step 4.0~4.12\n格式集中修复 + 即时验证\n- 数字格式/行高/合并单元格\n- 公式列保护/合计行公式\n- 空白行边框保留\n- smart_insert_row 强制校验"]:::phase
    end

    %% ==================== Phase 4.5 ====================
    subgraph PHASE_45["Phase 4.5 自动化验证门控"]
        P45_RUN["运行 gate_validator.py\n--gate G2 / --gate G3"]:::script
        P45_CHK{"G2/G3 全部通过？"}:::gate
        P45_FAIL["❌ 有严重问题 → 回退修复"]:::rule
        P45_WARN["⚠️ 有警告 → 建议修复"]:::rule
        P45_PASS["✅ 全部通过 → 进入勾稽"]:::gate
    end

    %% ==================== Phase 4 Reconcile ====================
    subgraph PHASE_4R["Phase 4（续）：勾稽核对 (S4_reconcile.md)\n9条规则+字段完整性"]
        direction TB
        P4R_1["Step 5.1\n三级勾稽核对\n明细表→汇总表→分类汇总→BS"]:::phase
        P4R_2["Step 5.2\n差异处理 → 备注栏标注"]:::phase
        P4R_2A["Step 5.2a [DT-69]\n科目归属交叉校验\n检查差异是否出现在其他 Sheet"]:::phase
        P4R_2B["Step 5.2b [DT-70]\nBS 重分类决策树\n（禁止统一标"BS重分类"）"]:::phase
        P4R_2C["Step 5.2c [DT-118]\n重分类映射表 → 7种常见场景"]:::phase
    end

    %% ==================== Phase 4 Linkage ====================
    subgraph PHASE_4L["Phase 4 链接检查：隐藏汇总表联动 (S4_linkage.md)"]
        P4L_CHK["DT-61/DT-62\n隐藏汇总表引用验证\n辅汇总表 vs 可见汇总表逻辑确认"]:::phase
    end

    %% ==================== Phase 5.5 ====================
    subgraph PHASE_55["Phase 5.5：反思固化门控"]
        P55_REV["问题回溯 → 归类判断\n→ 固化 DT 规则或 Phase 步骤\n→ 写入 SKILL.md"]:::phase
        P55_CHK{"自检清单全部 ✅？"}:::gate
    end

    %% ==================== Phase 5 ====================
    subgraph PHASE_5["Phase 5：清理与交付 (S5_deliver.md)\n14条规则 | ❌ 不可跳过"]
        direction TB
        P5_1["Step 5.1 [DT-110]\n集中隐藏操作\n① 确认设定信息已填写\n② 隐藏无数据空白 Sheet (DT-20)\n③ 递归校验汇总 Sheet 可见性 (DT-123)"]:::phase
        P5_2["Step 5.2\n页脚一次性输出 (DT-81)"]:::phase
        P5_3["Step 5.3 [DT-17]\nCOM 重算保存"]:::phase
        P5_4["Step 5.4 [DT-110]\n隐藏操作集中执行确认"]:::phase
        P5_5["Step 5.5\n最终 Gate 验证 (G3-11)"]:::gate
        P5_6["Step 5.6 [DT-59]\nSKILL.md / lessons_learned.md 更新"]:::phase
        P5_7["Step 5.7\n生成执行情况摘要"]:::phase
    end

    %% ==================== 交付 ====================
    DELIVER(("📦 交付：{项目名}{年度}_评估明细表_已填写.xlsx\n+ 推论映射汇报 (DT-117)")):::deliver

    %% ==================== 连接线 ====================
    TRIGGER --> CK_START
    CK_START --> CK_EXISTS

    CK_EXISTS -->|"是"| CK_LOAD
    CK_LOAD --> CK_FRESH
    CK_FRESH -->|"从缺失 Phase 恢复"| PHASE_M1

    CK_EXISTS -->|"否"| PHASE_M1

    M1_1 --> M1_2 --> M1_3 --> M1_5 --> M1_6

    M1_6 --> PHASE_0
    P0_1 --> P0_1_5 --> P0_2 --> P0_3 --> P0_4 --> P0_5 --> P0_5A --> P0_6 --> P0_7

    P0_7 --> PHASE_1
    P1_1 --> P1_2 --> P1_2A

    P1_2A --> PHASE_2
    P2_A --> P2_PIPE
    P2_B --> P2_PIPE
    P2_C --> P2_PIPE
    P2_D --> P2_PIPE

    P2_PIPE --> PHASE_3

    P3_0 -->|"有序时账"| P3_EXE
    P3_0 -->|"无序时账/用户明确跳过"| P3_SKIP
    P3_EXE --> GATE_G4
    P3_SKIP --> GATE_G4

    G4_JUDGE -->|"已执行或合法跳过"| G4_PASS
    G4_JUDGE -->|"应执行未执行"| G4_BLOCK
    G4_BLOCK -.->|"回退"| P3_0

    G4_PASS --> PHASE_4
    P4_PRE --> P4_0

    P4_0 --> PHASE_45
    P45_RUN --> P45_CHK
    P45_CHK -->|"通过"| P45_PASS
    P45_CHK -->|"警告"| P45_WARN
    P45_CHK -->|"严重问题"| P45_FAIL
    P45_FAIL -.->|"回退修复"| P4_0

    P45_PASS --> PHASE_4R
    P4R_1 --> P4R_2 --> P4R_2A --> P4R_2B --> P4R_2C

    P4R_2C --> PHASE_4L
    P4L_CHK --> PHASE_55

    P55_REV --> P55_CHK
    P55_CHK -->|"✅"| PHASE_5
    P55_CHK -->|"❌"| P55_REV

    P5_1 --> P5_2 --> P5_3 --> P5_4 --> P5_5 --> P5_6 --> P5_7

    P5_7 --> DELIVER

    %% ==================== 规则注记 ====================
    R160["🚨 DT-160 红线：禁止裸 openpyxl 写入\n数据写入 MUST 通过 sheet_filler.fill_sheet()"]:::rule
    R139["🚨 DT-139 红线：BS 解析后强制自校验\n资产≠负债+权益 → 禁止进入 Phase 1"]:::rule
    R161["🚨 DT-161 红线：Phase 3 不可跳过\n有序时账却不执行=发生日期全空+账龄失真"]:::rule
    R110["🚨 DT-110 红线：隐藏操作集中执行\nPhase 2-4 期间禁止提前隐藏 Sheet"]:::rule

    P2_PIPE -.-> R160
    P0_3 -.-> R139
    P3_0 -.-> R161
    P5_1 -.-> R110
```

## 流程图说明

| 图例 | 含义 |
|------|------|
| 🟦 深蓝圆角 | 触发/交付节点 |
| 🟦 浅蓝圆角 | Phase 执行步骤 |
| 🟪 紫色圆角 | 子步骤 |
| 🟧 橙色圆角 | Gate 门控（阻断/判定节点） |
| 🟩 绿色圆角 | 脚本执行 |
| 🟥 红色虚线圆角 | 规则/红线注记 |
| ❌ 不可跳过 | 该 Phase 必须执行，不存在跳过路径 |
| ⚠️ 可跳过 | 满足特定条件时可跳过 |

## 核心架构总结

**四层约束架构**（自底向上）：

| 层级 | 载体 | 约束力 |
|------|------|--------|
| L1 脚本强制 | sheet_filler.py / gate_validator.py / phase_gate.py | ❌ 不可绕过（crash/raise） |
| L2 Gate 门控 | 17 项 G2/G3 校验 + Phase 间 Gate 序列 | ❌ 不可绕过（流程阻断） |
| L3 规则文字 | RULES.md (DT-0~DT-213) | ⚠️ 依赖 Agent 自觉 |
| L4 流程硬卡 | Phase 间强制 Gate + sys.exit(1) | ❌ 不可绕过 |

**三层读取架构**：

| 层级 | 何时读取 | 内容 | 字符量 |
|------|----------|------|--------|
| L0 启动层 | 每次执行 | SKILL.md 概要 | ~3K |
| L1 Phase 层 | 进入新 Phase | 对应 Step 文件 | 10~48K |
| L2 按需层 | 遇到特定科目 | RULES.md 该科目专属规则 | 0~20K |
