"""
business_content_map.py — 业务内容自动推断模块 v1.1

设计目的:
  解决"业务内容未填写/填写=科目名/描述太冗长"的问题（复盘问题12/16+DT-172）。
  Agent自行填写业务内容时，往往只填科目名称（如"应付账款"）或拼接全称（如"工程款-中建三局"）。
  本模块根据科目编码+结算对象名称自动推断业务实质，默认只返回精简关键词。

强制等级: A级(L1脚本强制) — sheet_filler.py内部自动调用
对应规则: DT-149 业务内容自动映射规则 + DT-172 业务内容精简原则

使用:
  from business_content_map import infer_business_content
  content = infer_business_content('2202', '中建三局第一分公司')
  # → '工程款'（DT-172精简原则：只返回业务类型关键词，不拼接结算对象全称）

  content = infer_business_content('222106', None, '应交所得税')
  # → '企业所得税'

v1.1 (2026-05-25): DT-172业务内容精简原则——默认只返回业务类型关键词，不拼接结算对象全称
  - 旧版: infer_business_content('2202', '中建三局') → '工程款-中建三局第一分公司'
  - v1.1: infer_business_content('2202', '中建三局') → '工程款'
  - 结算对象全称已在"结算对象名称"列单独列示，业务内容列不应重复
v1.0 (2026-05-24): 初始版本
  - infer_business_content(): 主推断函数
  - _map_by_subject_code(): 科目编码→业务类型映射
  - _extract_from_counterparty(): 结算对象名称→业务实质提取
"""

import re


# ============================================================
# 科目编码→业务类型映射表
# ============================================================

SUBJECT_BUSINESS_MAP = {
    # 资产类
    '1122': {'prefix': '货款', 'default': '销售货款'},
    '1123': {'prefix': '预付', 'default': '预付货款'},
    '1221': {'prefix': '', 'default': '其他往来'},     # 其他应收需从结算对象推断
    '1231': {'prefix': '', 'default': '坏账准备'},     # 坏账准备行专用
    '1401': {'prefix': '', 'default': '材料采购'},
    '1403': {'prefix': '', 'default': '原材料'},
    '1405': {'prefix': '', 'default': '库存商品'},
    '1408': {'prefix': '', 'default': '委托加工物资'},
    '5001': {'prefix': '', 'default': '生产成本'},
    '5002': {'prefix': '开发成本', 'default': '开发成本'},
    '5101': {'prefix': '', 'default': '制造费用'},
    '1501': {'prefix': '', 'default': '持有至到期投资'},
    '1511': {'prefix': '', 'default': '长期股权投资'},
    '1601': {'prefix': '', 'default': '固定资产'},
    '1602': {'prefix': '', 'default': '累计折旧'},
    '1603': {'prefix': '', 'default': '固定资产减值准备'},
    '1604': {'prefix': '', 'default': '在建工程'},
    '1605': {'prefix': '', 'default': '工程物资'},
    '1606': {'prefix': '', 'default': '固定资产清理'},
    '1801': {'prefix': '', 'default': '长期待摊费用'},
    '1811': {'prefix': '', 'default': '递延所得税资产'},
    '1821': {'prefix': '', 'default': '其他长期资产'},

    # 负债类
    '2201': {'prefix': '', 'default': '短期借款'},
    '2202': {'prefix': '工程款', 'default': '应付货款'},
    '2203': {'prefix': '预收', 'default': '预收账款'},
    '2211': {'prefix': '', 'default': '应付职工薪酬'},
    '2221': {'prefix': '', 'default': '应交税费'},
    '2231': {'prefix': '', 'default': '应付利息'},
    '2232': {'prefix': '', 'default': '应付股利'},
    '2241': {'prefix': '', 'default': '其他往来'},     # 其他应付需从结算对象推断
    '2401': {'prefix': '', 'default': '递延收益'},
    '2501': {'prefix': '', 'default': '长期借款'},
    '2502': {'prefix': '', 'default': '应付债券'},
    '2701': {'prefix': '', 'default': '长期应付款'},

    # 所有者权益类
    '4001': {'prefix': '', 'default': '实收资本'},
    '4002': {'prefix': '', 'default': '资本公积'},
    '4101': {'prefix': '', 'default': '盈余公积'},
    '4103': {'prefix': '', 'default': '本年利润'},
    '4104': {'prefix': '', 'default': '利润分配'},
}


# ============================================================
# 结算对象名称→业务实质提取规则
# ============================================================

COUNTERPARTY_PATTERNS = [
    # (正则模式, 提取的业务类型, 优先级)
    # DT-212: 业务类型统一为2-5字名词，无符号拼接
    (r'(工程|施工|建设|基建|装饰|装修|安装)', '工程款', 10),
    (r'(材料|钢材|水泥|建材|管材|设备|采购|供应)', '材料设备', 9),
    (r'(设计|咨询|监理|勘察|测绘)', '设计咨询', 8),
    (r'(物业|保洁|保安|绿化|维护|维修)', '物业维修', 7),
    (r'(租金|租赁|房租|场地)', '租赁费', 7),
    (r'(工资|薪酬|社保|公积金|福利|绩效|奖金)', '职工薪酬', 8),
    (r'(税|增值税|所得税|附加|印花|房产税|土地使用)', '税费', 8),
    (r'(利息|贷款|借款|融资|银行|信托)', '借款利息', 7),
    (r'(保证金|押金|担保|质保)', '保证金', 6),
    (r'(报销|差旅|办公|通讯|交通|会议)', '费用报销', 5),
    (r'(保险|车险|社保)', '保险费', 6),
    (r'(广告|宣传|推广|营销)', '广告宣传', 5),
    (r'(运输|物流|货运|快递)', '运输物流', 5),
    (r'(水电|燃气|供暖|能源)', '水电能源', 5),
    (r'(结算中心|集团|内部|往来)', '内部往来', 9),
    (r'(政府|财政局|税务局|国土|规划|住建)', '政府往来', 7),
    (r'(股利|分红|利润)', '利润分配', 6),
]


# ============================================================
# 特殊科目处理：应交税费逐税种
# ============================================================

TAX_SUBJECT_MAP = {
    '222101': {'name': '应交增值税', 'tax_type': '增值税', 'authority': '税务局'},
    '22210101': {'name': '进项税额', 'tax_type': '增值税-进项', 'authority': '税务局'},
    '22210102': {'name': '销项税额', 'tax_type': '增值税-销项', 'authority': '税务局'},
    '22210105': {'name': '转出未交增值税', 'tax_type': '增值税-转出', 'authority': '税务局'},
    '222102': {'name': '未交增值税', 'tax_type': '增值税', 'authority': '税务局'},
    '222103': {'name': '应交营业税', 'tax_type': '营业税', 'authority': '税务局'},
    '222104': {'name': '应交消费税', 'tax_type': '消费税', 'authority': '税务局'},
    '222106': {'name': '应交所得税', 'tax_type': '企业所得税', 'authority': '税务局'},
    '222108': {'name': '应交城市维护建设税', 'tax_type': '城建税', 'authority': '税务局'},
    '222109': {'name': '应交房产税', 'tax_type': '房产税', 'authority': '税务局'},
    '222110': {'name': '应交土地使用税', 'tax_type': '土地使用税', 'authority': '税务局'},
    '222111': {'name': '应交车船使用税', 'tax_type': '车船使用税', 'authority': '税务局'},
    '222112': {'name': '应交个人所得税', 'tax_type': '个人所得税', 'authority': '税务局'},
    '222113': {'name': '应交教育费附加', 'tax_type': '教育费附加', 'authority': '税务局'},
    '222114': {'name': '应交地方教育附加', 'tax_type': '地方教育附加', 'authority': '税务局'},
    '222115': {'name': '应交印花税', 'tax_type': '印花税', 'authority': '税务局'},
    '222116': {'name': '应交土地增值税', 'tax_type': '土地增值税', 'authority': '税务局'},
    '222117': {'name': '应交资源税', 'tax_type': '资源税', 'authority': '税务局'},
    '222118': {'name': '应交环保税', 'tax_type': '环保税', 'authority': '税务局'},
}

# 征税机关默认值映射
TAX_AUTHORITY_MAP = {
    '增值税': '国家税务总局',
    '企业所得税': '国家税务总局',
    '个人所得税': '国家税务总局',
    '城建税': '国家税务总局',
    '教育费附加': '国家税务总局',
    '地方教育附加': '地方税务局',
    '房产税': '地方税务局',
    '土地使用税': '地方税务局',
    '土地增值税': '地方税务局',
    '印花税': '地方税务局',
    '车船使用税': '地方税务局',
    '资源税': '地方税务局',
    '环保税': '地方税务局',
    '消费税': '国家税务总局',
    '营业税': '国家税务总局',
}


# ============================================================
# 递延所得税资产名称披露
# ============================================================

DEFERRED_TAX_NAMES = {
    '1811': {
        'default': '递延所得税资产',
        'disclosure_items': [
            '资产减值准备差异',
            '公允价值变动差异',
            '固定资产折旧差异',
            '无形资产摊销差异',
            '长期待摊费用摊销差异',
            '预计负债差异',
            '可抵扣亏损',
            '租赁负债差异',
        ],
    }
}


# ============================================================
# 职工薪酬子目映射表（DT-174）
# ============================================================

PAYROLL_SUBJECT_MAP = {
    # 2211末级科目编码 → 业务内容关键词
    # 格式: 子编码后缀 → {'name': 子目名称, 'biz_content': 业务内容关键词}
    '01': {'name': '工资、奖金、津贴和补贴', 'biz_content': '工资'},
    '02': {'name': '职工福利费', 'biz_content': '职工福利'},
    '03': {'name': '社会保险费', 'biz_content': '社保'},
    '04': {'name': '住房公积金', 'biz_content': '公积金'},
    '05': {'name': '工会经费', 'biz_content': '工会经费'},
    '06': {'name': '职工教育经费', 'biz_content': '职工教育'},
    '07': {'name': '非货币性福利', 'biz_content': '非货币性福利'},
    '08': {'name': '辞退福利', 'biz_content': '辞退福利'},
    '09': {'name': '股份支付', 'biz_content': '股份支付'},
    '10': {'name': '设定提存计划', 'biz_content': '设定提存'},
    '11': {'name': '设定受益计划', 'biz_content': '设定受益'},
}

# 社保子目细化映射
PAYROLL_SOCIAL_INSURANCE_MAP = {
    '养老保险': '养老保险',
    '医疗': '医疗保险',
    '失业': '失业保险',
    '工伤': '工伤保险',
    '生育': '生育保险',
    '社保': '社保',
}


def infer_payroll_content(subject_code, subject_name=None):
    """职工薪酬专用：推断末级科目对应的业务内容关键词（DT-174）。

    2211科目的末级科目名称格式多样，需统一映射：
    - '221101' → '工资'
    - '221103' → '社保'
    - '22110301' → '养老保险' (如果科目名含具体险种)
    - '221104' → '公积金'

    Args:
        subject_code: 科目编码（如'221101'）
        subject_name: 科目名称（如'工资、奖金、津贴和补贴'）

    Returns:
        str: 业务内容关键词
    """
    code = str(subject_code).strip()
    name = str(subject_name).strip() if subject_name else ''

    # 优先从末级编码匹配（2211后面的2位）
    if code.startswith('2211') and len(code) >= 6:
        sub_code = code[4:6]  # 取2211后的2位
        payroll_info = PAYROLL_SUBJECT_MAP.get(sub_code)
        if payroll_info:
            # 社保子目细化：检查科目名中是否含具体险种
            if sub_code == '03':  # 社会保险费
                for kw, biz in PAYROLL_SOCIAL_INSURANCE_MAP.items():
                    if kw in name and kw != '社保':
                        return biz
            return payroll_info['biz_content']

    # 从科目名称关键词匹配
    payroll_keywords = [
        ('工资', '工资'), ('奖金', '工资'), ('津贴', '工资'), ('补贴', '工资'),
        ('福利', '职工福利'), ('社保', '社保'), ('养老', '养老保险'),
        ('医疗', '医疗保险'), ('失业', '失业保险'), ('工伤', '工伤保险'),
        ('生育', '生育保险'), ('公积金', '公积金'), ('住房', '公积金'),
        ('工会', '工会经费'), ('教育', '职工教育'), ('培训', '职工教育'),
        ('辞退', '辞退福利'), ('股份支付', '股份支付'),
        ('非货币', '非货币性福利'),
    ]
    for kw, biz in payroll_keywords:
        if kw in name:
            return biz

    # 兜底
    return name if name else '职工薪酬'


# ============================================================
# 辅助函数
# ============================================================

# DT-175: 交易方向前缀——不是业务实质，应剥离
_DIRECTION_PREFIXES = ['应付', '应收', '预付', '预收', '支付', '付款', '收款', '收到', '收回']
_DIRECTION_ONLY = {'应付', '应收', '预付', '预收', '支付', '付款', '收款', '收到', '收回'}


def _strip_direction(text):
    """DT-175: 剥离交易方向前缀（应付/应收/支付等非业务实质词）。

    规则：
    - 如果文本以方向前缀开头，剥离后返回剩余部分
    - 如果剥离后为空（文本本身就是方向词），返回空串
    - 如果文本不以方向前缀开头，原样返回
    """
    for prefix in _DIRECTION_PREFIXES:
        if text.startswith(prefix):
            remaining = text[len(prefix):]
            return remaining
    return text


def _smart_truncate(summaries):
    """DT-173: 智能截断——先净化再在词语边界截断，替代硬切6字。

    处理逻辑：
    1. 净化：去除连字符后缀/地名/数字/公司简称残片
    2. 在净化后文本上重跑关键词匹配
    3. 无匹配则在标点/空格处截断，或取4字
    """
    _SUFFIX_NOISE = re.compile(
        r'[—\-\s].*$'
        r'|\d+周[年月日].*$'
        r'|\d+月\d*日?$'
        r'|第[一二三四五六七八九十\d]+[季度期]$'
    )

    cleaned = []
    for s in summaries:
        s = str(s).strip()
        s_clean = _SUFFIX_NOISE.sub('', s).strip()
        s_clean = s_clean.rstrip('—-').strip()
        if s_clean:
            cleaned.append(s_clean)

    if not cleaned:
        cleaned = [str(s).strip() for s in summaries if str(s).strip()]

    # 二次关键词匹配（DT-174: 暂估保留前缀+主体词; DT-175: 剥离方向前缀）
    summary_biz_keywords_local = {
        '暂估款': ['暂估'],  # DT-174: 最前优先匹配
        '货款': ['货款', '收货款', '发货', '出货', '采购', '进货', '购', '买'],
        '服务费': ['服务费', '技术服务', '咨询费', '管理费', '开发费'],
        '外协费': ['外协', '加工', '代工', '委外', '外包'],
        '工程款': ['工程款', '工程', '施工', '建设', '安装'],
        '租金': ['租金', '租赁', '房租', '场地费'],
        '员工福利': ['入职', '礼品', '福利', '慰问', '节日', '员工'],
        '税费': ['增值税', '所得税', '税', '附加', '销项', '进项'],
        '往来款': ['往来', '划款', '调拨', '打款', '转账'],  # DT-175: 移除"付款"（方向词）
        '油费': ['汽油', '柴油', '加油', '油费'],
        '餐费': ['餐费', '餐饮', '伙食'],
    }
    for s in cleaned:
        for biz_type, patterns in summary_biz_keywords_local.items():
            for p in patterns:
                if p in s:
                    # DT-174: "暂估"是重要会计标识，保留"暂估+主体词"
                    if biz_type == '暂估款' and p == '暂估':
                        m_za = re.match(r'暂估(.+)', s)
                        if m_za:
                            rest = m_za.group(1).strip()
                            if rest:
                                return _strip_direction('暂估' + rest)
                    # DT-175: 剥离方向前缀后再返回
                    return _strip_direction(biz_type)

    # 词语边界截断
    shortest = min(cleaned, key=len) if cleaned else ''
    if not shortest:
        return '往来款'
    # DT-175: 剥离方向前缀
    if len(shortest) <= 4:
        stripped = _strip_direction(shortest)
        return stripped if stripped else '往来款'

    m = re.search(r'[，。、；：！？\s—\-/\\]', shortest)
    if m and m.start() > 0:
        result = _strip_direction(shortest[:m.start()])
        return result if result else '往来款'
    result = _strip_direction(shortest[:4])
    return result if result else '往来款'


# ============================================================
# 主推断函数
# ============================================================

def infer_business_content(subject_code, counterparty_name=None, subject_name=None, summaries=None):
    """根据科目编码+结算对象名称+序时账摘要推断业务实质内容。

    DT-172精简原则：默认只返回业务类型关键词，不拼接结算对象全称。
    结算对象全称已在"结算对象名称"列单独列示，业务内容列不应重复。

    推断策略（按优先级）：
    1. 特殊科目处理（应交税费→逐税种、递延所得税→具体内容）
    2. 序时账摘要高频关键词归纳（Phase 2e传入summaries时启用，DT-60）
    3. 结算对象名称模式匹配（工程/材料/设计/物业等）→只返回关键词
    4. 科目编码→业务类型映射→只返回prefix/default
    5. 兜底：科目名称

    Args:
        subject_code: 科目编码（如 '2202', '22210106'）
        counterparty_name: 结算对象名称（如 '中建三局第一分公司'）
        subject_name: 科目全名（兜底使用）
        summaries: 序时账摘要列表（Phase 2e核实时传入，如 ['萍乡春风江南项目工抵房', '工程款']）

    Returns:
        str: 推断的业务内容（精简关键词）

    用法：
        infer_business_content('2202', '中建三局第一分公司')
        # → '工程款'（DT-172精简：只返回业务类型，不拼接结算对象）

        infer_business_content('2202', '中建三局第一分公司', summaries=['支付工程款', '工程结算'])
        # → '工程款'（从摘要归纳）

        infer_business_content('222106', None, '应交所得税')
        # → '企业所得税'

        infer_business_content('1811', None, '递延所得税资产')
        # → '递延所得税资产-资产减值准备差异(待确认具体事项)'
    """
    code = str(subject_code).strip()
    name = str(subject_name).strip() if subject_name else ''

    # ---- 1. 特殊科目：应交税费 → 逐税种 ----
    if code.startswith('2221'):
        tax_info = TAX_SUBJECT_MAP.get(code)
        if tax_info:
            return tax_info['tax_type']
        # 子编码不在映射表中→取科目名称中的税种部分
        if '税' in name or '附加' in name:
            # 去掉"应交"前缀
            return name.replace('应交', '').strip()
        return name or '税费'

    # ---- 2. 特殊科目：递延所得税 → 披露具体内容 ----
    if code.startswith('1811'):
        dt_info = DEFERRED_TAX_NAMES.get(code[:4])
        if dt_info:
            # 如果科目名已经是具体差异类型，直接用
            for item in dt_info['disclosure_items']:
                if item in name:
                    return f'递延所得税资产-{item}'
            # 否则标注待确认
            return f'递延所得税资产-(待确认具体事项: {", ".join(dt_info["disclosure_items"][:3])}等)'
        return name or '递延所得税资产'

    # ---- 2.5 序时账摘要高频关键词归纳（Phase 2e传入summaries时启用，DT-60） ----
    if summaries and len(summaries) > 0:
        # 从摘要中提取业务类型关键词
        # DT-173: 扩展关键词表，覆盖更多业务场景
        # DT-174: "暂估"从"货款"独立为"暂估款"，匹配时保留"暂估+主体词"
        summary_biz_keywords = {
            '暂估款': ['暂估'],  # DT-174: 最前优先匹配
            '货款': ['货款', '收货款', '发货', '出货', '采购', '进货', '购货', '购', '买'],
            '服务费': ['服务费', '技术服务', '咨询费', '管理费', '开发费', '软件费', '检测费', '平台费', '订阅'],
            '工程款': ['工程款', '工程', '施工', '建设', '安装', '装修', '装饰', '基建'],
            '外协费': ['外协', '加工', '代工', '委外', '外包', '协作'],
            '租金': ['租金', '租赁', '房租', '场地费', '物业费'],
            '保证金': ['保证金', '押金', '投标保证金', '履约保证金'],
            '报销款': ['报销', '差旅', '办公费', '通讯费', '交通费'],
            '员工福利': ['入职', '礼品', '福利', '慰问', '节日', '员工', '周年'],
            '社保': ['社保', '公积金', '五险一金', '养老', '医疗', '失业'],
            '税费': ['增值税', '所得税', '税', '附加', '印花', '城建', '销项', '进项', '开票', '抵扣'],
            '借款': ['借款', '贷款', '融资', '利息', '归还'],
            '往来款': ['往来', '划款', '调拨', '内部', '结算', '打款', '转账'],  # DT-175: 移除"付款""支付"
            '销售款': ['销售', '销售款', '出售', '出卖', '开票'],
            '退税款': ['退税', '出口退税', '留抵退税'],
            '代垫款': ['代垫', '代扣', '代缴', '代付'],
            '油费': ['汽油', '柴油', '加油', '油费', 'ETC', '过路'],
            '餐费': ['餐费', '餐饮', '伙食', '饭费'],
            '保险费': ['保险', '车险', '财险'],
            '运输费': ['运输', '物流', '货运', '快递', '发货'],
        }

        from collections import Counter
        keyword_counts = Counter()
        for summary in summaries:
            summary_str = str(summary)
            for biz_type, patterns in summary_biz_keywords.items():
                for p in patterns:
                    if p in summary_str:
                        keyword_counts[biz_type] += 1
                        break  # 每条摘要每种业务类型只计1次

        if keyword_counts:
            # 取最高频的业务类型
            best_biz = keyword_counts.most_common(1)[0][0]
            # 如果最高频=1且有多个同频→组合
            top_count = keyword_counts.most_common(1)[0][1]
            top_items = [k for k, v in keyword_counts.items() if v == top_count]
            if len(top_items) > 1:
                best_biz = top_items[0]  # DT-212: 只取最高频，不拼接
            # DT-174: "暂估"保留前缀+主体词（如"暂估半成品"而非"暂估款"）
            if best_biz == '暂估款':
                for summary in summaries:
                    m_za = re.match(r'暂估(.+)', str(summary).strip())
                    if m_za and m_za.group(1).strip():
                        return _strip_direction('暂估' + m_za.group(1).strip())
            # DT-175: 剥离方向前缀
            stripped = _strip_direction(best_biz)
            return stripped if stripped else best_biz

        # DT-173: 摘要净化+词语边界截断（替代硬切6字）
        return _smart_truncate(summaries)

    # ---- 3. 结算对象名称模式匹配 ----
    # DT-172: 业务内容只返回业务类型关键词，不拼接结算对象全称
    # 结算对象全称已在"结算对象名称"列单独列示，业务内容列不应重复
    if counterparty_name:
        cp = str(counterparty_name).strip()
        best_match = None
        best_priority = -1

        for pattern, biz_type, priority in COUNTERPARTY_PATTERNS:
            if re.search(pattern, cp):
                if priority > best_priority:
                    best_priority = priority
                    best_match = biz_type

        if best_match:
            # DT-172: 只返回业务类型关键词，不拼接结算对象全称
            # 旧版: return f'{best_match}-{cp}' → 太冗长，结算对象名称列已有全称
            return best_match

    # ---- 4. 科目编码→业务类型映射 ----
    # 先精确匹配4位编码
    code4 = code[:4] if len(code) >= 4 else code
    subject_info = SUBJECT_BUSINESS_MAP.get(code4)

    if subject_info:
        prefix = subject_info.get('prefix', '')
        default = subject_info.get('default', '')

        # DT-172: 只返回业务类型(prefix或default)，不拼接结算对象全称
        if prefix:
            return prefix
        return default or name

    # ---- 5. 兜底：科目名称 ----
    return name or f'科目{code}'


def infer_tax_details(subject_code, subject_name=None):
    """应交税费专用：返回税种名称和征税机关。

    对应规则: DT-147 应交税费逐税种填写+征税机关格式

    Args:
        subject_code: 税费科目编码
        subject_name: 科目名称

    Returns:
        dict: {'tax_type': str, 'authority': str, 'business_content': str}
    """
    code = str(subject_code).strip()
    name = str(subject_name).strip() if subject_name else ''

    tax_info = TAX_SUBJECT_MAP.get(code)
    if tax_info:
        tax_type = tax_info['tax_type']
        authority = TAX_AUTHORITY_MAP.get(tax_type, '税务局')
        return {
            'tax_type': tax_type,
            'authority': authority,
            'business_content': tax_type,
        }

    # 不在映射表→从名称提取
    tax_type = name.replace('应交', '').strip() if name else '税费'
    authority = '税务局'

    # 根据名称推断征税机关
    for kw, auth in [('增值税', '国家税务总局'), ('所得税', '国家税务总局'),
                     ('附加', '国家税务总局'), ('房产税', '地方税务局'),
                     ('土地', '地方税务局'), ('印花税', '地方税务局'),
                     ('车船', '地方税务局'), ('环保', '地方税务局')]:
        if kw in tax_type:
            authority = auth
            break

    return {
        'tax_type': tax_type,
        'authority': authority,
        'business_content': tax_type,
    }


def infer_counterparty_from_name(subject_code, subject_name):
    """从科目余额表名称中提取结算对象名称（DT-166强制）。

    科目余额表中的末级科目名称格式多样，需统一解析：
    - '应收账款_斑马网络技术有限公司' → '斑马网络技术有限公司'
    - '公司_上海闵行交大科技园运营有限公司' → '上海闵行交大科技园运营有限公司'
    - '个人_徐磊' → '徐磊'
    - '预付账款_蓝星光域航天科技' → '蓝星光域航天科技'
    - '应付账款_上海胜崇智能科技有限公司' → '上海胜崇智能科技有限公司'
    - 无下划线(如'坏账准备') → 返回科目全名

    对应规则: DT-166 结算对象名称非空断言

    Args:
        subject_code: 科目编码
        subject_name: 科目余额表中的末级科目名称

    Returns:
        str: 结算对象名称（不可为空字符串）
    """
    name = str(subject_name).strip() if subject_name else ''

    if '_' in name:
        # 取最后一个下划线后部分作为结算对象
        # 修复：'其他应付款_个人_徐磊' 应返回 '徐磊' 而非 '个人_徐磊'
        # 策略：连续去除已知前缀词，直到剩余部分不是前缀词
        known_prefixes = {'应收账款', '应付账款', '预付账款', '预收账款',
                         '其他应收款', '其他应付款', '公司', '个人',
                         '坏账准备', '长期待摊费用', '递延所得税资产'}
        parts = name.split('_')
        # 从后往前找第一个非前缀词的部分
        counterparty = parts[-1].strip() if parts else name
        # 如果最后一个部分仍是前缀词，返回拼接的剩余部分
        if counterparty in known_prefixes and len(parts) > 1:
            # 全部拼接（去掉第一个前缀词）
            counterparty = '_'.join(parts[1:]).strip()
        # 如果结果仍为空，返回全名
        return counterparty if counterparty else name

    # 无下划线 → 返回科目全名
    return name


def infer_deferred_tax_content(subject_code, subject_name=None, balance_amount=None):
    """递延所得税资产专用：推断具体披露内容。

    对应规则: DT-150 递延所得税名称披露具体内容

    Args:
        subject_code: 科目编码
        subject_name: 科目名称
        balance_amount: 期末余额（用于判断是否有子项）

    Returns:
        str: 披露内容
    """
    return infer_business_content(subject_code, None, subject_name)
