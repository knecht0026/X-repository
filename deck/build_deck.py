"""Fill in the Anthropic core-pool deck on top of the analyst's draft.

Usage: python3 deck/build_deck.py
Reads  deck/template/Anthropic申请加入境外核心库_原稿.pptx (titles + notes written by the analyst)
Writes Anthropic申请加入境外核心库.pptx (repo root)

All figures come from the Anthropic S-1/A No.2 (filed 2026-11-10) and the IPO FWP unless a
different source is named on the slide. Items marked 测算 are computed here from those figures.
"""
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import (BLUE, CYAN, DARK, GRAY, LIGHT, NAVY, PALE, PALE2, PALECYAN, RED, WHITE, YELLOW,
                          arrow, bar_chart, box, bullets, ctitle, doughnut, kpi, line, line_chart, move_slide,
                          picture, remove_shape, source, table, tb)

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "template" / "Anthropic申请加入境外核心库_原稿.pptx"
OUT = HERE.parent / "Anthropic申请加入境外核心库.pptx"
ASSETS = HERE / "assets"

S1 = "公司招股书（S-1/A，2026/11/10）"
QUARTERS = ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26"]

# ---- S-1 figures (USD millions unless noted)
REV_Q = [376, 787, 1382, 2040, 4731, 11555, 16956]
COMPUTE_Q = [944, 1390, 2047, 2945, 3387, 5558, 9863]
ADJ_MARGIN_Q = [-2.21, -1.29, -0.93, -0.84, -0.13, 0.19, 0.15]
PERIODS = ["2023", "2024", "2025", "9M25", "9M26"]
REVENUE = [41, 386, 4585, 2545, 33242]
COMPUTE = [433, 2522, 7326, 4381, 18808]
COST_OF_REVENUE = [26, 752, 2773, 1796, 4812]
RND = [422, 2275, 6338, 3594, 19240]
CFO = [-1627, -5692, -2637, -2024, 6569]

# ---- share counts (post-offering, S-1 capitalization tables)
FD_SHARES = 7_430_143_944 + 7_804_580_273 + 1_704_567_865 + 1_073_701_297
BASIC_SHARES = 13_086_672_593
IPO_MID = 74.0


def content_ph(slide):
    for shp in slide.placeholders:
        if shp.placeholder_format.idx == 10:
            return shp
    return None


def bullet_box(slide, items, x, y, w, h, size=9, **kw):
    """Bulleted text: reuse the draft's content placeholder when the slide has one."""
    shp = content_ph(slide) or tb(slide, x, y, w, h, "")
    return bullets(shp, items, x, y, w, h, size=size, **kw)


def kpi_row(slide, cards, x, y, w, h, gap=0.1, **kw):
    cw = (w - gap * (len(cards) - 1)) / len(cards)
    for i, (big, small) in enumerate(cards):
        kpi(slide, x + i * (cw + gap), y, cw, h, big, small, **kw)


# ------------------------------------------------------------------ slide 2
def slide_conclusion(s):
    mcap = BASIC_SHARES * IPO_MID / 1e8
    kpi_row(s, [
        ("170亿美元", "Q3'26单季收入，同比+1,127%"),
        ("~810亿美元", "Run-Rate收入（2026/10/31）"),
        ("15%", "Q3'26调整后经营利润率（9M26为12%）"),
        (">11,900家", "TTM收入≥10万美元客户，NRR约900%"),
        (f"~{mcap:,.0f}亿美元", "发行市值中枢（$74/股；全面摊薄约1.33万亿美元）"),
    ], 0.35, 1.0, 9.3, 0.74, big_size=14)
    bullet_box(s, [
        "**结论**：申请将Anthropic（拟于11/19登陆纳斯达克，代码ANT）纳入境外核心库。公司是当前模型能力最领先、"
        "商业化增速最快的AI研究型公司，人才密度、自上而下的研发组织与RSI共同构成其持续领先的基础。",
        "**公司简介**：2021年1月由7位前OpenAI成员创立于旧金山（特拉华州公益公司PBC），产品为Claude系列模型"
        "（Mythos/Fable、Opus、Sonnet、Haiku）、Claude Platform及Claude Code、Claude Cowork等应用；员工约5,300人。",
        "**股东结构**：7位联合创始人全部在任，合计经济权益约9%，通过Founder LLC持有的F类股锁定50.1%投票权；"
        "长期利益信托（LTBT）选举7席董事中的4席；Amazon、Google为前两大外部股东（持无投票权N类股）。",
        "**投资逻辑**：①人才密度与使命驱动的文化（遗憾离职率约1%，两年留存率80% vs OpenAI 67%）；"
        "②集中押注的研发体系带来高研发效率（AA智能指数第一；2023–9M26累计研发283亿美元，累计经营现金净流出仅34亿美元）；"
        "③企业客户反馈飞轮（NRR约900%，财富500强过半为≥10万美元客户）；"
        "④RSI加速（9月>90%代码由Claude编写），中长期有望形成“模型领先→收入/融资→训练资源”的正循环。",
        "**财务与估值**：9M26收入332亿美元（同比+1,206%），调整后经营利润40.8亿美元，经营现金流65.7亿美元；"
        "GAAP净亏损1,040亿美元主要来自可转债公允价值变动（非现金，943亿美元）。发行区间$70–78对应市值9,160–10,210亿美元"
        "（发行后总股本130.9亿股），中枢约12x Run-Rate收入；盈利预测与目标估值待补充。",
        "**风险提示**：算力承诺约9,009亿美元（未来12个月到期707亿美元）；模型领先优势被追赶、用户迁移成本低；"
        "监管与诉讼（战争部“供应链风险”认定、FTC调查、版权诉讼）；PBC+LTBT的特殊治理结构；锁定期后大额解禁。",
    ], 0.35, 1.9, 9.3, 3.0, size=9.5, space_after=5)
    source(s, f"{S1}、Artificial Analysis、SignalFire；市值及估值倍数为测算")


# ------------------------------------------------------------------ slide 3
def slide_company(s):
    milestones = [
        ("2021.01", "公司成立\n7位联创，特拉华州PBC"),
        ("2023.03", "Claude商业化发布"),
        ("2025.02", "Claude Code\n研究预览"),
        ("2026.01", "Claude Cowork\n研究预览"),
        ("2026.05–06", "H轮融资652亿美元\n累计融资1,411亿美元"),
        ("2026.11", "纳斯达克IPO（ANT）\n发行区间$70–78"),
    ]
    n, x0, w = len(milestones), 0.35, 9.3 / len(milestones)
    for i, (date, text) in enumerate(milestones):
        shape = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        box(s, x0 + i * w, 1.0, w + 0.06, 0.3, date, fill=NAVY if i < n - 1 else CYAN, size=9,
            color=WHITE, bold=True, shape=shape)
        tb(s, x0 + i * w + 0.05, 1.33, w - 0.05, 0.4, text, size=8, align="c")

    # product stack
    ctitle(s, 0.35, 1.84, 4.5, "核心业务：模型—平台—应用三层架构")
    lx, cx, cw_total = 0.35, 1.02, 3.83
    rows = [
        ("应用", ["Chat\n对话协作", "Claude Code\n软件开发Agent", "Claude Cowork\n知识工作Agent", "Security / Design\n垂直产品"]),
        ("平台", ["Claude Platform：扩展思考、工具调用、长上下文、MCP、Agent SDK、Managed Agents、批处理/缓存"]),
        ("模型", ["Mythos / Fable\n前沿智能", "Opus\n强大智能", "Sonnet\n通用智能", "Haiku\n高性价比"]),
    ]
    fills = [PALE, PALECYAN, PALE]
    for r, (label, chips) in enumerate(rows):
        y = 2.12 + r * 0.56
        box(s, lx, y, 0.6, 0.5, label, fill=NAVY, color=WHITE, bold=True, size=9.5, radius=0.1)
        gap = 0.05
        chip_w = (cw_total - gap * (len(chips) - 1)) / len(chips)
        for i, chip in enumerate(chips):
            lines = chip.split("\n")
            text = [(lines[0], 8.5, NAVY, True)] + [(l, 7.5) for l in lines[1:]] if len(lines) > 1 else chip
            box(s, cx + i * (chip_w + gap), y, chip_w, 0.5, text, fill=fills[r], size=8, radius=0.1)
    bullet_box(s, [
        "**成立与定位**：2021年1月成立，7位联合创始人至今全部在任；自我定位“首先是一家研究实验室”，使命为负责任地开发先进AI",
        "**五大核心优势**：使命与专注、人才与文化、前沿与安全研究、算力策略（多芯片可互换算力池）、企业战略",
        "**财务**：9M26收入332亿美元（+1,206%），按量计费收入占79%、订阅21%；调整后经营利润率12%，经营现金流65.7亿美元",
    ], 0.35, 3.86, 4.55, 1.05, size=8, space_after=2)

    # financial charts
    ctitle(s, 5.1, 1.84, 4.5, "单季收入（亿美元）：Q3'26同比+1,127%")
    bar_chart(s, 5.05, 2.05, 4.6, 1.7, QUARTERS, [("收入", [v / 100 for v in REV_Q])], [NAVY],
              fmt="#,##0.0", gap=45, point_colors=[BLUE] * 6 + [NAVY])
    ctitle(s, 5.1, 3.72, 4.5, "调整后经营利润率：Q2'26起转正")
    line_chart(s, 5.05, 3.93, 4.6, 0.98, QUARTERS, [("调整后经营利润率", ADJ_MARGIN_Q)], [CYAN],
               fmt="0%", labels_low=True)
    source(s, f"{S1}、IPO Factsheet（FWP）")


# ------------------------------------------------------------------ slide 4
def slide_shareholders(s):
    ctitle(s, 0.35, 0.96, 5.2, "控制权结构：联合创始人、LTBT、普通股东三方分权")
    bw, gap, y0 = 1.62, 0.17, 1.24
    holders = [
        (NAVY, WHITE, ["Founder LLC", "7位联合创始人，一人一票", "持F类股（1股）"], "固定50.1%投票权\n与A类股共选3席董事"),
        (BLUE, WHITE, ["长期利益信托 LTBT", "受托人5席（现任4位），无经济利益", "持T类股（1股）"], "选举并可罢免\n7席董事中的4席"),
        (PALECYAN, DARK, ["A类普通股", "创始人、员工、投资人、公众", "1股1票"], "合计49.9%投票权"),
    ]
    for i, (fill, color, lines, label) in enumerate(holders):
        x = 0.35 + i * (bw + gap)
        text = [(lines[0], 9, color, True), (lines[1], 7.5, color), (lines[2], 7.5, color)]
        box(s, x, y0, bw, 0.7, text, fill=fill, radius=0.1)
        arrow(s, x + 0.3, y0 + 0.7, x + 0.3, 2.52, color=NAVY)
        tb(s, x + 0.36, y0 + 0.75, bw - 0.36, 0.5, label, size=7.5, color=NAVY, bold=True, anchor="m")
    box(s, 0.35, 2.52, 5.2, 0.62, [
        ("Anthropic, PBC（特拉华州公益公司，纳斯达克“受控公司”）", 9.5, NAVY, True),
        ("董事会7席：T类董事4席（Hastings、Liddell、Narasimhan、1席空缺）+ A/F类董事3席（Dario、Daniela、Razavi）", 7.5),
    ], fill=PALE, line=NAVY, radius=0.08)
    arrow(s, 2.95, 3.42, 2.95, 3.14, color=GRAY, dash=True)
    tb(s, 3.02, 3.16, 2.0, 0.24, "无投票权，经济权利与A类相同", size=7.5, color=GRAY, anchor="m")
    box(s, 0.35, 3.42, 5.2, 0.4, [
        ("N类普通股：Amazon 19.2亿股 ｜ Google 17.9亿股 ｜ 同步私募战略投资者 ｜ 公益基金8.5亿股", 8, DARK, True),
    ], fill=PALE2, line=GRAY, dash=True, radius=0.1)

    # fully diluted ownership (post-offering, at IPO midpoint)
    founders = 237_259_602 + 238_030_562 + 240_794_662 + 238_893_462 + 238_972_792 + 240_005_472 + 239_504_042
    amazon_held, amazon_notes = 1_920_511_218, 1_563_617_444
    google = 1_789_598_269
    pbf = 846_840_960
    matching = 553_441_297 + 27_991_850
    ipo = 825_000_000
    strategic = (26.0e9 - 10.0e9) / IPO_MID
    awards = 1_704_567_865 + 1_073_701_297
    parts = [
        ("7位联合创始人", founders, NAVY),
        ("Amazon（已持有）", amazon_held, BLUE),
        ("Amazon（待转股票据）", amazon_notes, CYAN),
        ("Google", google, YELLOW),
        ("公益基金（PBF）", pbf, RED),
        ("员工捐赠配捐", matching, GRAY),
        ("期权及RSU", awards, DARK),
        ("IPO公众股东", ipo, PALECYAN),
        ("其他同步私募投资者", strategic, LIGHT),
    ]
    other = FD_SHARES - sum(p[1] for p in parts)
    parts.append(("其他（财务投资人、员工等）", other, PALE))
    ctitle(s, 5.75, 0.96, 3.9, f"全面摊薄股本结构（约{FD_SHARES / 1e8:.0f}亿股，测算）")
    doughnut(s, 5.6, 1.45, 1.95, 1.95, [p[0] for p in parts], [p[1] for p in parts], [p[2] for p in parts],
             hole=58, legend=None)
    tb(s, 6.2, 2.2, 0.75, 0.45, [(f"{FD_SHARES / 1e8:.0f}亿股", 9, NAVY, True), ("全面摊薄", 7, GRAY)], align="c", anchor="m")
    for i, (name, val, color) in enumerate(parts):
        y = 1.24 + i * 0.255
        box(s, 7.62, y + 0.06, 0.11, 0.11, fill=color, line=LIGHT if color in (PALE, PALECYAN, LIGHT) else None,
            shape=MSO_SHAPE.RECTANGLE, line_w=0.5)
        tb(s, 7.78, y, 1.45, 0.23, name, size=7.5, anchor="m")
        tb(s, 9.15, y, 0.5, 0.23, f"{val / FD_SHARES:.1%}", size=7.5, bold=True, color=NAVY, align="r", anchor="m")

    bullet_box(s, [
        "**Dario的控制权**：个人持股约2.4亿股（发行后占A类股3.2%、投票权1.6%），并非单一控制人；控制权来自7位联合创始人共同持有的"
        "Founder LLC（F类股固定50.1%投票权，多数决，Dario与Daniela兄妹占7票中的2票）。持续创始人不足3人满90天后进入450天过渡期，届时F类股赎回",
        "**LTBT**：受托人与公司无经济利益，可选举并罢免董事会多数席位，对公司使命及AI安全相关重大事项形成制衡",
        "**慈善安排**：已向全资子公司Anthropic Public Benefit Fund发行8.47亿股N类股（占6/30全面摊薄股本5%）；此后股价每较前一里程碑上涨50%，"
        "董事会可再发行至多2%。7位创始人各自承诺捐出个人股权的80%，员工捐赠由公司配捐（已归属配捐义务5.53亿股）",
    ], 0.35, 3.9, 9.3, 1.02, size=8, space_after=2)
    source(s, f"{S1}；股本结构按发行后口径、发行价中枢$74测算，创始人持股含60日内可行权期权")


# ------------------------------------------------------------------ slide 5
def slide_management(s):
    ph = content_ph(s)
    if ph is not None:
        remove_shape(ph)
    execs = [
        ("Dario Amodei（43岁）", "联合创始人、CEO、董事", [
            "**职责**：CEO、董事；每两周面向全员公开讲述正在考虑的决策与行业判断",
            "**背景**：前OpenAI研究副总裁（2016–2020），此前任Google Brain高级研究科学家；Scaling Laws论文核心作者之一",
            "**教育**：普林斯顿大学生物物理博士，斯坦福大学物理学士及医学院博士后",
        ]),
        ("Daniela Amodei（39岁）", "联合创始人、总裁、董事会主席", [
            "**职责**：总裁（President）兼董事会主席、提名与治理委员会成员；与Dario为兄妹",
            "**背景**：前OpenAI安全与政策副总裁、人事副总裁、工程经理（2018–2020）；此前在Stripe任风险经理等职（2013–2018）",
            "**教育**：加州大学圣克鲁兹分校英语文学学士",
        ]),
        ("Krishna Rao（43岁）", "首席财务官（2024年4月加入）", [
            "**职责**：首席财务官，2024年4月起任职",
            "**背景**：前Fanatics Commerce CFO、Cedar CFO；Airbnb全球企业与业务发展负责人；黑石私募股权Principal；贝恩咨询",
            "**教育**：耶鲁法学院JD，哈佛学院经济学学士",
        ]),
    ]
    w, gap = 3.0, 0.15
    for i, (name, title, lines) in enumerate(execs):
        x = 0.35 + i * (w + gap)
        box(s, x, 1.0, w, 1.78, fill=PALE2, line=LIGHT, radius=0.04)
        box(s, x, 1.0, w, 0.44, [(name, 10, WHITE, True), (title, 8, WHITE)], fill=NAVY,
            shape=MSO_SHAPE.RECTANGLE, align="l", margin=0.08)
        tb(s, x + 0.06, 1.48, w - 0.12, 1.28, lines, size=7.5, space_after=3, margin=0.04)

    ctitle(s, 0.35, 2.84, 6.0, "其他联合创始人（均为Founder LLC成员）")
    cofounders = [
        ("Tom Brown", "首席算力官", "负责算力的获取、扩容与调度；前OpenAI GPT-3工程负责人，与Dario共同发明deep RLHF"),
        ("Jared Kaplan", "首席科学官", "理论物理学家，Scaling Laws主要作者之一，参与开创Constitutional AI；哈佛物理博士"),
        ("Sam McCandlish", "首席架构师", "专注大规模模型训练与系统架构，Scaling Laws工作为GPT-3奠基；斯坦福物理博士"),
        ("Christopher Olah", "可解释性研究负责人", "机制可解释性领域奠基人；前OpenAI可解释性负责人、Google Brain研究科学家"),
        ("Jack Clark", "Anthropic Institute负责人", "向外部伙伴与公众分享公司研究成果；前OpenAI政策总监，Import AI作者"),
    ]
    w2, gap2 = 1.74, 0.15
    for i, (name, title, desc) in enumerate(cofounders):
        x = 0.35 + i * (w2 + gap2)
        box(s, x, 3.1, w2, 1.0, fill=WHITE, line=LIGHT, radius=0.05)
        box(s, x, 3.1, w2, 0.05, fill=CYAN, shape=MSO_SHAPE.RECTANGLE)
        tb(s, x + 0.05, 3.17, w2 - 0.1, 0.92, [(name, 9, NAVY, True), (title, 8, DARK, True), (desc, 7)],
           space_after=1, margin=0.03)

    box(s, 0.35, 4.2, 0.75, 0.68, [("董事会", 9.5, WHITE, True), ("7席", 8, WHITE)], fill=NAVY, radius=0.1)
    tb(s, 1.18, 4.2, 8.47, 0.34,
       "**T类董事（LTBT选举，4席）**：Reed Hastings（首席独立董事，Netflix联合创始人）｜Christopher Liddell（前微软、通用汽车CFO，"
       "前白宫副幕僚长）｜Vas Narasimhan（诺华CEO）｜1席空缺",
       size=8, anchor="m", bold_color=NAVY)
    tb(s, 1.18, 4.54, 8.47, 0.34,
       "**A/F类董事（3席）**：Dario Amodei｜Daniela Amodei（主席）｜Yasmin Razavi（Spark Capital普通合伙人）",
       size=8, anchor="m", bold_color=NAVY)
    source(s, f"{S1}、创始人致股东信")


# ------------------------------------------------------------------ slide 6
def slide_market(s):
    ctitle(s, 0.35, 0.96, 4.6, "TAM：知识工作40万亿美元+，生命科学2万亿美元+")
    box(s, 0.45, 1.22, 2.85, 2.85, fill=PALE, line=BLUE, shape=MSO_SHAPE.OVAL, line_w=1.25)
    tb(s, 0.75, 1.72, 2.25, 1.1, [
        ("知识工作", 11, NAVY, True), ("40万亿美元+", 20, NAVY, True),
        ("2025年全球知识工作者创造的经济价值", 7.5),
    ], align="c", anchor="m")
    box(s, 1.02, 3.1, 0.16, 0.16, fill=RED, shape=MSO_SHAPE.OVAL)
    tb(s, 1.22, 2.98, 1.75, 0.4, [("当前Run-Rate收入约810亿美元", 7.5, RED, True), ("≈知识工作价值的0.2%", 7.5, RED)],
       anchor="m")
    box(s, 3.2, 1.2, 1.5, 1.5, fill=PALECYAN, line=CYAN, shape=MSO_SHAPE.OVAL, line_w=1.25)
    tb(s, 3.25, 1.53, 1.4, 0.85, [("生命科学", 9.5, NAVY, True), ("2万亿美元+", 13, NAVY, True), ("2026E全球药品支出", 7)],
       align="c", anchor="m")
    tb(s, 3.3, 2.74, 1.6, 0.2, "更多新市场（招股书未量化）", size=7.5, color=GRAY, bold=True, align="c")
    tb(s, 0.4, 3.95, 1.6, 0.18, "示意图，非等比例", size=6.5, color=GRAY)
    for i, name in enumerate(["机器人", "药物发现", "材料科学", "先进制造"]):
        x = 3.42 + (i % 2) * 0.7
        y = 2.95 + (i // 2) * 0.6
        box(s, x, y, 0.62, 0.56, name, fill=WHITE, line=GRAY, dash=True, shape=MSO_SHAPE.OVAL, size=7.5, margin=0.0)
    bullet_box(s, [
        "**招股书口径**：知识工作>40万亿美元为公司基于BLS/ILO数据的测算；全球药品支出2026年>2万亿美元（IQVIA）",
        "**渗透仍处早期**：当前Run-Rate收入约810亿美元，仅相当于知识工作经济价值的约0.2%",
    ], 0.35, 4.18, 4.6, 0.72, size=8, space_after=2)

    ctitle(s, 5.15, 0.96, 4.5, "增长故事：从Coding到白领知识工作，再到科学与物理世界")
    stages = [
        (GRAY, "已验证", "2025", "**Coding Agent（Claude Code）**：约一年内从补全单行代码发展到在生产代码库中自主运行Agent集群；同期公司收入由2025年46亿美元升至9M26的332亿美元"),
        (NAVY, "短期", "当下第二曲线", "**Claude Cowork渗透白领市场**：把Claude Code的Agent架构扩展到全部知识工作者；2026年1月研究预览，采用速度超过Claude Code同期；9月起与Claude.ai合并"),
        (CYAN, "中长期", "新市场", ["**制药**：前10大药企中9家已是≥10万美元客户；Claude Science", "**网安**：Claude Security、OSS Scanner；ExploitBench约4个月40%→91%", "**机器人**：Model Hardware Standard可操控实验室与制造设备"]),
    ]
    for i, (fill, tag, sub, text) in enumerate(stages):
        y = 1.24 + i * 0.93
        box(s, 5.15, y, 0.95, 0.82, [(tag, 10, WHITE, True), (sub, 7.5, WHITE)], fill=fill, radius=0.1)
        box(s, 6.15, y, 3.5, 0.82, text, fill=PALE2, size=7.5, align="l", radius=0.06, margin=0.08)
        if i < 2:
            arrow(s, 5.62, y + 0.82, 5.62, y + 0.93, color=GRAY)
    kpi_row(s, [
        ("22.5万亿美元", "IDC：2025–2031年AI累计全球经济影响"),
        ("1.2→4.3万亿美元", "全球IT支出2006→2025，云由0增至1.9万亿美元"),
    ], 5.15, 4.08, 4.5, 0.78, big_size=12, small_size=7)
    source(s, f"{S1}、IPO Factsheet（FWP）；占比为测算")


# ------------------------------------------------------------------ new slide after 6
def slide_market_2(s):
    ctitle(s, 0.35, 0.96, 3.0, "TTM收入≥10万美元客户数（家）")
    bar_chart(s, 0.3, 1.18, 3.0, 2.05, ["2023", "2024", "2025", "2026/9"], [("客户数", [20, 200, 1500, 11900])], [NAVY],
              gap=50, custom_labels={0: ["~20", "~200", "~1,500", ">11,900"]}, point_colors=[BLUE, BLUE, BLUE, NAVY])
    ctitle(s, 3.5, 0.96, 3.0, "大企业渗透率（≥10万美元客户占比）")
    bar_chart(s, 3.45, 1.18, 3.0, 2.05, ["财富10强", "财富500强", "全球2000强"], [("渗透率", [1.0, 0.5, 0.3])], [NAVY],
              horizontal=True, fmt="0%", gap=60, vmax=1.25,
              custom_labels={0: ["100%", ">50%", ">30%"]}, point_colors=[NAVY, BLUE, CYAN])
    ctitle(s, 6.65, 0.96, 3.0, "钱包份额仍低（Q3'26）")
    kpi(s, 6.65, 1.24, 3.0, 0.6, "<0.1%", "财富500强客户的Claude支出占其总费用（平均）", big_size=14)
    kpi(s, 6.65, 1.92, 3.0, 0.6, "<1%", "前10大财富500强客户（为平均水平的8倍以上）", big_size=14)
    kpi(s, 6.65, 2.6, 3.0, 0.6, ">50%", "≥10万美元客户同时使用平台与第一方应用（Q4'25为30%）", big_size=14)

    ctitle(s, 0.35, 3.3, 9.3, "形态演进：Cowork把Agent从开发者扩展到全部白领")
    steps = [
        (GRAY, "Chatbot", "单轮问答，回答窄问题"),
        (BLUE, "Copilot", "嵌入单个应用或任务内辅助"),
        (NAVY, "Virtual Collaborator（Claude Code / Cowork）", "接受目标后长时间自主工作：读文件、操作应用、自我纠错，交付文档/表格/PPT"),
    ]
    widths, x = [1.9, 2.2, 5.2], 0.35
    for i, ((fill, name, desc), w) in enumerate(zip(steps, widths)):
        shape = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        box(s, x, 3.56, w + 0.05, 0.5, [(name, 9, WHITE, True), (desc, 7.5, WHITE)], fill=fill, shape=shape, margin=0.12)
        x += w
    bullet_box(s, [
        "**短期第二曲线**：Cowork于2026年1月开启研究预览，采用速度超过Claude Code同期；非技术团队正越来越多地使用Claude Code与Cowork管理自身工作流",
        "**大客户覆盖广、花费浅**：财富500强中前10大软件/IT服务、互联网公司、银行全部覆盖，前10大药企与半导体公司各覆盖9家；≥100万美元客户>1,900家，≥1,000万美元约250家",
        "**提升空间**：财富500强客户平均钱包份额不足0.1%，头部客户已达其8倍以上，客户内扩张（NRR约900%）是主要增长来源",
    ], 0.35, 4.12, 9.3, 0.8, size=8, space_after=1.5)
    source(s, f"{S1}、IPO Factsheet（FWP）；客户口径为TTM收入≥10万美元")


# ------------------------------------------------------------------ slide 7
def slide_talent(s):
    kpi_row(s, [
        ("7 / 7", "联合创始人全部在任（主要AI公司中唯一）"),
        ("~1%", "TTM遗憾离职率（截至Q3'26）"),
        (">75%", "前100名员工仍在职"),
        ("19%", "遗憾离职者中后来回流的比例"),
        (">660万美元", "人均TTM收入（一年前为140万美元）"),
    ], 0.35, 1.0, 9.3, 0.72, big_size=14)

    ctitle(s, 0.35, 1.8, 3.0, "顶级AI Lab人才双向流动比（x:1）")
    bar_chart(s, 0.3, 2.02, 3.05, 1.85, ["DeepMind→Anthropic", "OpenAI→Anthropic", "DeepMind→OpenAI"],
              [("流动比", [10.6, 8.2, 2.4])], [NAVY], horizontal=True, fmt='0.0":1"', gap=60, vmax=13.5,
              point_colors=[NAVY, NAVY, GRAY])
    ctitle(s, 3.5, 1.8, 3.0, "入职两年后员工留存率")
    bar_chart(s, 3.45, 2.02, 3.05, 1.85, ["Anthropic", "DeepMind", "OpenAI", "Cohere", "Mosaic"],
              [("留存率", [0.80, 0.78, 0.67, 0.64, 0.63])], [GRAY], fmt="0%", gap=55, vmin=0.4, vmax=0.9,
              point_colors=[NAVY, GRAY, GRAY, GRAY, GRAY])

    box(s, 6.65, 1.8, 3.0, 2.07, fill=PALE2, line=LIGHT, radius=0.04)
    tb(s, 6.72, 1.84, 2.86, 2.0, [
        ("愿景与共同价值观（招股书 / 创始人信）", 8.5, NAVY, True),
        ("“Talent density matters more than talent mass.”", 8, DARK, True),
        ("“公司文化建立在公共利益、诚实与信任、低自我（low-ego）协作，以及专注于做最简单有效的事之上……"
         "技术再强、不具备这些特质的候选人，我们也愿意拒绝。”", 7.5),
        ("“一个由数据中心里的天才组成的国家（a country of geniuses in a data center）”", 7.5),
        ("七条价值观：Act for the Global Good｜Hold Light and Shade｜Be Good to Our Users｜Race to the Top on Safety｜"
         "Do the Simple Thing That Works｜Be Helpful, Honest, and Harmless｜Put the Mission First", 6.5, GRAY),
    ], space_after=3, margin=0.02)

    bullet_box(s, [
        "**人才密度高于OpenAI**：SignalFire统计（2021–23年入职者）两年留存率Anthropic 80% vs OpenAI 67%，每8.2人从OpenAI流向Anthropic才有1人反向流动；"
        "Meta高薪挖角时OpenAI据称流失数十人，Anthropic仅2人",
        "**员工特质**：Ego小、追求真相、使命驱动（重视AI安全、相信AI应造福人类）；内部自称“Ants”，多家明星公司CTO甘愿以普通MTS身份加入，7位联合创始人股权均等",
        "**文化机制**：全员须通过culture-fit面试（1小时15–20道情景题）；Dario称约1/3–40%的时间用于文化建设；员工可承诺捐出部分股权，公司以N类股配捐",
    ], 0.35, 3.95, 9.3, 0.97, size=8.5, space_after=2)
    source(s, f"{S1}、创始人致股东信、SignalFire（转引自海外独角兽《拆解Anthropic》，2026/5）")


# ------------------------------------------------------------------ slide 8
def slide_org(s):
    ctitle(s, 0.35, 0.96, 4.9, "组织形态示意：Anthropic自上而下 vs OpenAI自下而上")
    picture(s, str(ASSETS / "org_anthropic_vs_openai.png"), 0.35, 1.22, w=4.9)
    tb(s, 0.35, 3.3, 4.9, 0.18, "图片credit：@odysseus0z，转引自海外独角兽", size=6.5, color=GRAY, align="r")

    table(s, 5.45, 1.0, [0.72, 1.74, 1.74], [
        ["维度", "Anthropic", "OpenAI"],
        ["组织方式", "模型**Top-down** + 产品Bottom-up", "Bottom-up，团队各自推进"],
        ["战略重心", "All in Coding + ToB企业；不做图像/视频生成", "业务线繁多、多点开花（内部项目一度约300个）"],
        ["研发方式", "算力集中于coding、agentic能力与安全研究；重预训练与数据质量", "依靠明星研究员个人突破，重0到1创新"],
        ["创始团队", "7位联创全部在任，股权均等", "11位创始成员仅余3位"],
        ["招聘文化", "严格culture-fit面试", "规模扩大后已取消专门文化面试"],
    ], fs=7.5, rowh=0.41, aligns=["c", "l", "l"], bold_first_col=True)

    ctitle(s, 0.35, 3.5, 9.3, "内部高度坦诚：自上而下同步判断，让分布式决策保持一致")
    steps = [
        ("Dario Vision Quest", "每两周一次全员会：Dario讲约1小时并配3–4页文档，公开说明正在考虑的决策与行业判断，现场答疑"),
        ("Slack公开笔记", "Dario与员工在notebook频道持续记录所思所虑，全员可订阅、讨论和质疑"),
        ("全员对齐", "人人知道决策如何做出、何为最高优先级，资源集中投向最关键的方向"),
    ]
    w, gap = 2.9, 0.3
    for i, (name, desc) in enumerate(steps):
        x = 0.35 + i * (w + gap)
        box(s, x, 3.76, w, 0.6, [(name, 8.5, NAVY, True), (desc, 7)], fill=PALE, radius=0.08, margin=0.06)
        if i < len(steps) - 1:
            arrow(s, x + w + 0.03, 4.06, x + w + gap - 0.03, 4.06, color=NAVY)
    bullet_box(s, [
        "**自上而下**：公司“首先是一家研究实验室”，4位联创为Scaling Laws论文核心作者，研究方向由创始团队集中判断，资源长期聚焦语言模型scaling与coding，专注探索智能上界",
        "**高度坦诚**：创始人信提到“每两周，Dario通常会站在全公司面前，公开讲述他正在思考的决策和行业现状”，其他联创在各自领域同样如此",
    ], 0.35, 4.4, 9.3, 0.52, size=8, space_after=1)
    source(s, f"海外独角兽《拆解Anthropic》（2026/5）、{S1}、创始人致股东信")


# ------------------------------------------------------------------ slide 9
def slide_rnd(s):
    ctitle(s, 0.35, 0.96, 3.2, "AA智能指数（2026/11/12，各模型最高档）")
    models = ["Claude Fable 5.2", "Gemini 4 Carbon", "Claude Opus 5.5", "GPT-6 Astra", "GPT-6.1 Sol",
              "Claude Sonnet 5.5", "GPT-6.1 Terra", "Gemini 4 Argon"]
    scores = [64.6, 60.0, 59.1, 56.9, 54.6, 54.5, 51.3, 50.4]
    colors = [NAVY if m.startswith("Claude") else GRAY for m in models]
    bar_chart(s, 0.3, 1.16, 3.25, 2.3, models, [("AA Intelligence Index", scores)], [GRAY], horizontal=True,
              fmt="0.0", gap=40, vmin=0, vmax=82, point_colors=colors)

    cum_rnd = sum(RND[:3]) + RND[4]
    cum_cfo = sum(CFO[:3]) + CFO[4]
    ctitle(s, 3.7, 0.96, 3.0, "融资几乎未被消耗（亿美元）")
    bar_chart(s, 3.65, 1.16, 3.0, 2.3, ["累计融资", "期末现金\n及短期投资", "累计研发费用\n2023–9M26", "累计经营现金\n净流出"],
              [("金额", [1411, 1318, cum_rnd / 100, -cum_cfo / 100])], [NAVY], gap=45,
              point_colors=[NAVY, BLUE, CYAN, GRAY], fs=7.5, label_fs=8)

    ctitle(s, 6.8, 0.96, 2.85, "与OpenAI对比")
    table(s, 6.8, 1.22, [0.85, 1.0, 1.0], [
        ["", "Anthropic", "OpenAI"],
        ["员工人数", "约5,300", "约11,100*"],
        ["累计融资", "1,411亿美元", "约1,860–1,910亿美元*"],
        ["AA最高分", "64.6（第1）", "56.9（第4）"],
        ["人均TTM收入", ">660万美元", "未披露"],
    ], fs=7.5, rowh=0.37, aligns=["l", "c", "c"], bold_first_col=True)
    tb(s, 6.8, 3.1, 2.85, 0.36, "*OpenAI为第三方口径（Tracxn、Airframe，2026/9–11），非公司披露，待核实",
       size=6.5, color=GRAY)

    ctitle(s, 0.35, 3.5, 4.3, "算力效率：单季收入 / 算力及基础设施费用（x）")
    line_chart(s, 0.3, 3.7, 4.4, 1.2, QUARTERS, [("收入/算力", [r / c for r, c in zip(REV_Q, COMPUTE_Q)])], [NAVY],
               fmt='0.00"x"', vmin=0, vmax=2.7)
    bullet_box(s, [
        "**模型领先**：AA智能指数前三占两席，Fable 5.2以64.6分居首，领先Gemini 4 Carbon约5分、GPT-6 Astra约8分（Mythos 5.2仅向受信机构开放，未参评）",
        f"**研发投入产出高**：2023–9M26累计研发费用{cum_rnd / 100:,.0f}亿美元、累计经营现金净流出仅{-cum_cfo / 100:,.0f}亿美元；"
        "期末现金及短期投资1,318亿美元，相当于累计融资额的93%",
        "**人效**：员工约5,300人（约为OpenAI第三方口径人数的一半），人均TTM收入>660万美元；2024年底以来人数增长5倍多，单季收入增长86倍",
    ], 4.85, 3.5, 4.8, 1.4, size=8, space_after=2)
    source(s, f"Artificial Analysis、{S1}、Tracxn、Airframe；累计值为测算（招股书仅披露2023年以来数据）")


# ------------------------------------------------------------------ slide 10
def slide_flywheel(s):
    ctitle(s, 0.35, 0.96, 4.6, "招股书：两个自我强化的飞轮")
    picture(s, str(ASSETS / "s1_self_reinforcing_loops.jpg"), 0.35, 1.2, w=4.6)
    tb(s, 0.35, 3.66, 4.6, 0.34, [
        "**改进飞轮**：更好的模型 → 更多应用场景 → 更多客户反馈 → 更好的模型",
        "**投资飞轮**：更高增长 → 更多投入 → 更强能力 → 更高增长",
    ], size=7.5, bold_color=NAVY)

    ctitle(s, 5.15, 0.96, 2.5, "≥10万美元客户贡献收入（亿美元）")
    bar_chart(s, 5.1, 1.18, 2.55, 1.72, ["2023", "2024", "2025", "TTM 26/9"], [("收入", [0.37, 2.2, 35, 261])], [NAVY],
              fmt="#,##0.0", gap=45, point_colors=[BLUE, BLUE, BLUE, NAVY])
    kpi(s, 7.8, 1.0, 1.85, 0.6, "~900%", "净收入留存率（NRR）", big_size=14)
    kpi(s, 7.8, 1.66, 1.85, 0.6, ">11,900家", "≥10万美元客户（一年前>950家）", big_size=13)
    kpi(s, 7.8, 2.32, 1.85, 0.6, "74%", "TTM收入来自≥10万美元客户", big_size=14)
    kpi_row(s, [
        (">80%", "Q3'26收入来自近6个月内发布的模型"),
        ("250项", "9M26发布的新模型版本、平台功能与应用"),
        ("-40%", "Opus 5.5较Opus 5的客户单任务成本"),
    ], 5.15, 3.0, 4.5, 0.66, big_size=13, small_size=7)

    bullet_box(s, [
        "**飞轮机制**：企业在生产环境中的部署让公司更早看到难题、边界情形与行业特有需求，据此改进模型、平台与应用，进而带来更广的部署；"
        "Coding/Agent任务结果可验证、反馈回路短，是最适合形成该闭环的场景（海外独角兽：公司2021年的内部文件已论证过聚焦coding的理由）",
        "**为什么是26年的护城河**：≥10万美元客户一年内由950多家增至11,900多家、NRR约900%，真实agentic工作流的反馈规模同行难以复制；研究与产品团队并肩工作，客户反馈直接影响研究优先级",
        "**需跟踪**：飞轮依赖模型持续领先，Q3'26超80%收入来自近6个月发布的模型，说明客户会快速切换到最新最强的模型",
    ], 0.35, 4.02, 9.3, 0.9, size=7.5, space_after=1.5)
    source(s, f"{S1}、IPO Factsheet（FWP）、海外独角兽《拆解Anthropic》")


# ------------------------------------------------------------------ slide 11
def slide_rsi(s):
    ctitle(s, 0.35, 0.96, 3.9, "Closing the loop：每一代Claude帮助构建下一代")
    steps = ["Claude的智能处于前沿", "越来越能够辅助前沿AI研究", "加速下一代模型的构建", "实现更大的智能跃迁"]
    bx, bw, bh = 0.5, 2.9, 0.4
    ys = [2.98, 2.4, 1.82, 1.24]
    fills = [PALE, PALECYAN, BLUE, NAVY]
    for i, (text, y) in enumerate(zip(steps, ys)):
        box(s, bx, y, bw, bh, text, fill=fills[i], size=9, bold=True, color=WHITE if i >= 2 else NAVY, radius=0.15)
        if i < 3:
            arrow(s, bx + bw / 2, y, bx + bw / 2, ys[i + 1] + bh, color=NAVY)
    rx = bx + bw
    line(s, rx, ys[3] + bh / 2, rx + 0.32, ys[3] + bh / 2, color=CYAN, width=1.5)
    line(s, rx + 0.32, ys[3] + bh / 2, rx + 0.32, ys[0] + bh / 2, color=CYAN, width=1.5)
    arrow(s, rx + 0.32, ys[0] + bh / 2, rx + 0.02, ys[0] + bh / 2, color=CYAN, width=1.5)
    tb(s, rx + 0.36, 2.0, 0.45, 0.6, "循环\n加速", size=8, color=CYAN, bold=True, anchor="m")

    kpi(s, 4.55, 1.0, 2.45, 0.7, ">90%", "2026年9月公司代码库贡献由Claude编写", big_size=16)
    kpi(s, 7.2, 1.0, 2.45, 0.7, ">250倍", "内部Claude用量价值（按目录价，2026/1→9），4月起超过月度人员费用（不含SBC）",
        big_size=16, small_size=7)
    ctitle(s, 4.55, 1.8, 2.45, "工程师人均日代码量（2024=1）")
    bar_chart(s, 4.5, 2.02, 2.5, 1.4, ["2024年", "Q3'26"], [("倍数", [1, 45])], [NAVY], gap=90, vmin=0, vmax=56,
              custom_labels={0: ["1x", ">45x"]}, point_colors=[GRAY, NAVY])
    ctitle(s, 7.2, 1.8, 2.45, "高难度开放式工程问题成功率")
    bar_chart(s, 7.15, 2.02, 2.5, 1.4, ["2025/12", "2026/9"], [("成功率", [0.4, 0.9])], [NAVY], fmt="0%", gap=90,
              vmin=0, vmax=1.1, custom_labels={0: ["~40%", "~90%"]}, point_colors=[GRAY, NAVY])

    ctitle(s, 0.35, 3.5, 3.3, "ExploitBench得分：约4个月内40%→91%")
    bar_chart(s, 0.3, 3.7, 3.3, 1.2, ["Opus 4.8", "Mythos 5", "Opus 5.5"], [("得分", [0.40, 0.78, 0.91])], [NAVY],
              fmt="0%", gap=70, vmin=0, vmax=1.1, point_colors=[GRAY, BLUE, NAVY])
    bullet_box(s, [
        "**RSI已在发生**：研究、训练、测试的诸多环节（跑实验、生成和处理训练数据、自动化评测与安全审计、写代码与审代码）已由Claude参与，招股书称之为“closing the loop”",
        "**直接体现**：“算力乘数”（训练与推理的算法改进）产出在加速——更强的模型帮助研究员更快找到改进，改进又让同样的算力训练出更强的模型",
        "**边界**：公司表示不预期在缺乏有意义的人类监督下让Claude完全自主设计下一代前沿模型，并考虑更新RSP，RSI节奏受安全政策约束",
    ], 3.8, 3.5, 5.85, 1.4, size=8, space_after=2)
    source(s, f"{S1}、IPO Factsheet（FWP）；图表按公开披露口径重绘，成功率起点由“9个月提升约50个百分点”倒推")


# ------------------------------------------------------------------ slide 12
def slide_rsi_2(s):
    bullet_box(s, [
        "**客观来讲，模型这门生意目前未见强壁垒**：用户迁移难度低，只有持续领先才能持续留住用户（Q3'26超80%收入来自近6个月内发布的模型）",
        "**RSI将成为公司在中长期的壁垒**：RSI加速 → 模型领先 → 收入/估值/融资扩张 → 更多的训练资源 → 模型领先",
    ], 0.35, 0.93, 9.3, 0.56, size=8.5, space_after=2)

    nodes = [
        ("RSI加速", "9月>90%代码由Claude编写\n人均代码量为2024年的45倍以上"),
        ("模型持续领先", "AA智能指数第1（64.6分）\n9M26发布250项新模型与功能"),
        ("收入 / 估值 / 融资扩张", "Run-Rate收入约810亿美元\n2026年股权融资1,104亿美元+IPO"),
        ("更多训练资源", "算力承诺约9,009亿美元\n9M26训练及研发算力占比约74%"),
    ]
    w, gap, y = 1.98, 0.46, 1.52
    for i, (head, body) in enumerate(nodes):
        x = 0.35 + i * (w + gap)
        box(s, x, y, w, 0.28, head, fill=NAVY, color=WHITE, bold=True, size=9.5, shape=MSO_SHAPE.RECTANGLE)
        box(s, x, y + 0.28, w, 0.46, body, fill=PALE, size=7.5, shape=MSO_SHAPE.RECTANGLE)
        if i < len(nodes) - 1:
            arrow(s, x + w + 0.04, y + 0.37, x + w + gap - 0.04, y + 0.37, color=NAVY, width=1.75)
    x_last = 0.35 + 3 * (w + gap) + w / 2
    x_first = 0.35 + w / 2
    yb = y + 0.74
    line(s, x_last, yb, x_last, yb + 0.17, color=CYAN, width=1.5)
    line(s, x_last, yb + 0.17, x_first, yb + 0.17, color=CYAN, width=1.5)
    arrow(s, x_first, yb + 0.17, x_first, yb + 0.01, color=CYAN, width=1.5)
    box(s, 3.3, yb + 0.08, 3.4, 0.18, "更多训练资源带来更强的模型，进一步加速RSI", fill=WHITE, size=7.5, color=CYAN, bold=True,
        shape=MSO_SHAPE.RECTANGLE, margin=0.0)

    ctitle(s, 0.35, 2.58, 5.4, "训练 / 推理算力费用拆分（百万美元，测算）")
    infer_share = [c / t for c, t in zip(COST_OF_REVENUE, COMPUTE)]
    fmt_m = lambda v: f"{v:,.0f}"
    rows = [
        [""] + PERIODS,
        ["算力及基础设施费用"] + [fmt_m(v) for v in COMPUTE],
        ["　推理：营业成本（服务客户的算力）"] + [fmt_m(v) for v in COST_OF_REVENUE],
        ["　训练及研发：其余（含内部使用）"] + [fmt_m(t - c) for c, t in zip(COST_OF_REVENUE, COMPUTE)],
        ["推理占比"] + [f"{v:.0%}" for v in infer_share],
        ["训练及研发占比"] + [f"{1 - v:.0%}" for v in infer_share],
        ["收入"] + [fmt_m(v) for v in REVENUE],
        ["毛利率（1−营业成本/收入）"] + [f"{1 - c / r:.0%}" for c, r in zip(COST_OF_REVENUE, REVENUE)],
    ]
    table(s, 0.35, 2.83, [2.1, 0.66, 0.66, 0.66, 0.66, 0.66], rows, fs=7.5, rowh=0.24,
          aligns=["l", "r", "r", "r", "r", "r"], highlight_rows=(5,))

    ctitle(s, 5.95, 2.58, 3.7, "历轮融资价格（美元/股，拆股调整后）")
    bar_chart(s, 5.9, 2.8, 3.75, 2.1, ["D轮\n24H1", "E轮\n25Q1", "F轮\n25Q3", "G轮\n26Q1", "H轮\n26Q2", "IPO中枢\n26/11"],
              [("每股价格", [3.00, 5.61, 14.10, 25.91, 58.90, 74.00])], [NAVY], fmt="0.0", gap=45,
              point_colors=[BLUE] * 5 + [NAVY], fs=7.5)
    source(s, f"{S1}（附注8）、Artificial Analysis；训推拆分为测算（营业成本含少量人员及分摊费用，推理占比略有高估）；融资价格按10:1拆股调整")


# ------------------------------------------------------------------ slide 15
def slide_compute(s):
    kpi_row(s, [
        ("~9,009亿美元", "算力及基础设施承诺总额（截至9/30/2026及期后）"),
        ("~4,516亿美元", "预计2030年底前发生，其余主要在2031–2045年"),
        ("707亿美元", "未来12个月内到期"),
        ("1,862亿美元", "附可行使终止权或生效前或有条件"),
        ("1,500–1,750亿美元", "无固定金额安排，未计入总额"),
    ], 0.35, 1.0, 9.3, 0.66, big_size=12.5, small_size=7)

    ctitle(s, 0.35, 1.72, 5.9, "主要对手方（亿美元）")
    table(s, 0.35, 1.95, [1.0, 1.2, 1.05, 0.95, 1.75], [
        ["供应商", "芯片 / 形式", "承诺金额", "期限", "备注"],
        ["Google Cloud", "TPU（云托管）", "≥1,213", "2026/4–2033/7", "另有追加TPU约1,250–1,500亿，未计入"],
        ["Broadcom", "TPU设备融资租赁", "~1,612", "5年期主租约", "不含利息；2026年起上线"],
        ["Amazon AWS", "Trainium", "≥1,100", "2026/5–2036/4", "实际支出不足需补差"],
        ["SpaceXAI", "NVIDIA GPU", "≤1,205", "至2029年", "满产约31亿/月；2027年起可解约"],
        ["Microsoft Azure", "NVIDIA GPU", "472–492", "2026/11–2033/6", "不可撤销"],
        ["Nscale", "NVIDIA GPU", "≤446", "至2033年", "按交付验收计费，对方需落实融资"],
        ["Lambda", "NVIDIA GPU", "~413", "至2034年", "其中349亿协议目前双方均可终止"],
        ["AMD", "Helios机柜系统", "无固定金额", "2027年首批", "预计>200亿，未计入"],
    ], fs=7.5, rowh=0.29, aligns=["l", "l", "c", "c", "l"], bold_first_col=True)

    ctitle(s, 6.45, 1.72, 3.2, "9,009亿美元承诺构成（亿美元）")
    bar_chart(s, 6.4, 1.92, 3.25, 1.4, ["尚未起租的租赁", "合同义务表内", "期后新增承诺", "短期承诺", "期后新增短期"],
              [("金额", [4512, 3200, 1057, 174, 66])], [NAVY], horizontal=True, gap=40, vmax=5600,
              point_colors=[NAVY, BLUE, CYAN, GRAY, GRAY], fs=7.5)
    ctitle(s, 6.45, 3.32, 3.2, "表内采购承诺到期分布（亿美元）")
    bar_chart(s, 6.4, 3.52, 3.25, 1.38, ["26Q4", "2027", "2028", "2029", "2030", "2031+"],
              [("采购承诺", [51, 284, 435, 469, 442, 1458])], [BLUE], gap=40, fs=7.5,
              point_colors=[BLUE] * 5 + [GRAY])
    ph = content_ph(s)
    if ph is not None:
        remove_shape(ph)
    tb(s, 0.35, 4.74, 5.95, 0.18,
       "多芯片策略：Trainium、TPU、NVIDIA GPU、AMD Instinct构成可互换调度的算力池；算力通常提前12–18个月锁定",
       size=7, color=DARK, anchor="m")
    source(s, f"{S1}（MD&A—合同义务、基础设施服务安排）；表中为已披露的主要对手方，合计小于承诺总额")


# ------------------------------------------------------------------ slide 16
def slide_risk(s):
    ctitle(s, 0.35, 0.96, 5.2, "资产负债表排查（截至2026年9月30日）")
    table(s, 0.35, 1.2, [1.1, 1.15, 2.95], [
        ["项目", "金额", "排查结论"],
        ["总资产", "1,543亿美元", "现金及短期投资1,318亿美元，占85%"],
        ["总负债", "1,467亿美元", "其中可转债（公允价值）1,305亿美元，非付息；剔除后负债162亿美元，负债率约10%"],
        ["流动比率", "12.1x", "流动资产1,410亿美元 vs 流动负债117亿美元，短期没有偿债压力"],
        ["应收账款", "79亿美元", "按Q3收入测算周转天数约43天；单一客户最高占13%，前五大客户占Q3收入约21%"],
        ["预付及其他流动资产", "13亿美元", "占总资产0.8%，没有大额其他应收款"],
        ["商誉", "6亿美元", "占总资产0.4%，减值压力可忽略"],
        ["有息负债", "无借款", "150亿美元循环贷款无提款，仅开立2亿美元备用信用证；最低流动性约束50亿美元"],
        ["股东权益", "-1,499亿美元", "累计亏损主要来自可转债公允价值变动；IPO后备考净资产约1,081亿美元"],
    ], fs=7.5, rowh=0.335, aligns=["l", "c", "l"], bold_first_col=True)
    box(s, 0.35, 4.3, 5.2, 0.56, [
        ("结论：表内没有明显的“雷”，主要风险在表外算力承诺", 9, NAVY, True),
        ("9M26经营现金流已转正（+65.7亿美元）；未来12个月到期承诺707亿美元，可由现有流动性及IPO募资覆盖", 7.5),
    ], fill=PALE, radius=0.08, align="l", margin=0.08)

    ctitle(s, 5.75, 0.96, 3.9, "流动性 vs 12个月内到期承诺（亿美元）")
    bar_chart(s, 5.7, 1.16, 3.95, 1.45, ["现金及\n短期投资", "未动用\n循环贷款", "IPO及私募\n净募资（预计）", "12个月内\n到期承诺"],
              [("金额", [1318, 148, 846, 707])], [NAVY], gap=50, fs=7.5, label_fs=8,
              point_colors=[NAVY, BLUE, CYAN, RED])

    ctitle(s, 5.75, 2.66, 3.9, "风险提示")
    bullet_box(s, [
        "**算力承诺刚性**：总额约9,009亿美元，收入增速若放缓将显著压制现金流；Q3'26含50亿美元高价短期算力，收入/算力比由2.08x回落至1.72x",
        "**领先优势被追赶**：用户迁移成本低，AA指数上领先第二名仅约5分",
        "**监管与诉讼**：战争部“供应链风险”认定（D.C.巡回法院9月裁定不利，已申请全席复审）；FTC于10月下发调查令；版权诉讼（Bartz案15亿美元和解）",
        "**治理与稀释**：PBC+LTBT+F类股，公众股东话语权有限；公益发行、Amazon待转股票据（15.6亿股）及期权/RSU带来持续稀释",
        "**收入质量**：9M26约5–7%收入来自违反使用政策的客户；按量计费收入占79%，存在季节性波动",
        "**解禁与股份支付**：锁定期180/365天；IPO时一次性确认约142亿美元股份支付费用",
    ], 5.75, 2.9, 3.9, 2.02, size=7.5, space_after=1.5, line_spacing=1.0)
    source(s, f"{S1}；周转天数、剔除可转债后的负债率为测算")


def main():
    prs = Presentation(str(TEMPLATE))
    slides = list(prs.slides)

    slide_conclusion(slides[1])
    slide_company(slides[2])
    slide_shareholders(slides[3])
    slide_management(slides[4])
    slide_market(slides[5])
    slide_talent(slides[6])
    slide_org(slides[7])
    slide_rnd(slides[8])
    slide_flywheel(slides[9])
    slide_rsi(slides[10])
    slide_rsi_2(slides[11])
    # slides[12], slides[13]: 财务数据及盈利预测 — left as in the draft on the analyst's instruction
    slide_compute(slides[14])
    slide_risk(slides[15])

    # the analyst allowed 市场空间 to be split over two pages
    extra = prs.slides.add_slide(slides[5].slide_layout)
    extra.shapes.title.text = "市场空间（续）"
    slide_market_2(extra)
    move_slide(prs, len(prs.slides) - 1, 6)

    prs.save(str(OUT))
    print("saved", OUT)


if __name__ == "__main__":
    main()
