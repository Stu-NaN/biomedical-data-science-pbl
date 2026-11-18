from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "第6讲_多重检验校正_课堂展示讲稿.docx"
FONT = "Microsoft YaHei"
INK = "20292D"
MUTED = "56666A"
TEAL = "176D66"

SLIDES = [
    (
        "多重检验校正：显著的结果，值得相信吗？",
        "50 秒",
        "大家好，我们今天讨论多重检验校正。做组学分析时，研究者常同时检验成千上万个探针，单看每个 p 值会遇到偶然发现。今天沿着“提出问题、定义检验家族、选择错误控制、分析真实数据、判断效应和验证边界”这条线展开。先给出结果：完整检验 7,129 个探针后，BH 筛出 535 项；预先锁定的 10 个候选里，有 7 个通过留出样本判据。这支持继续验证，不等于证明它们是真实疾病基因。",
    ),
    (
        "1,000 次比较与 53 个显著结果",
        "60 秒",
        "先看一个假设情境：同时检验 1,000 个基因，报告了 53 个 raw p 小于 0.05 的结果。p 值是在零假设和统计模型成立时，观察到当前或更极端统计量的概率，不是零假设为真的概率。若 1,000 个零假设全都成立，且每项检验恰以 0.05 的概率错误拒绝，期望假发现数就是 50；若检验只是保守有效，期望不超过 50。期望不表示每次刚好有 50 个，也不能断言 53 个里有 50 个是假的。若额外假设检验独立，至少出现一次假阳性的概率会接近 1。这个 53 是课程假设案例，不是本次 CSV 的结果。",
    ),
    (
        "研究数据与预先固定的分析协议",
        "60 秒",
        "本次使用 Golub 白血病表达数据，共 72 份样本、7,129 个探针。发现阶段用 train 的 38 份样本，ALL 27 份、AML 11 份；test 的 34 份留到候选锁定后再看。每个探针做一次双侧 Welch t 检验，完整家族预先定为 7,129 项。两组样本量不相等，Welch 检验不要求方差相等；它仍依赖单项检验有效和样本独立等条件。数据是含负值的处理后表达强度，所以保留原尺度，不把均值差说成倍数变化。比较对象是两种白血病亚型，不是患者和健康人。",
    ),
    (
        "发现集合中的错误：FWER 与 FDR",
        "45 秒",
        "这两个概念关注的错误不同。单次研究如果宣布发现 20 项，其中 2 项是假发现，那么这次的 FDP，也就是发现集合中的假发现比例，是 2 除以 20，等于 0.10。FWER 关注重复做很多次研究时，至少出现一次假发现的概率；FDR 则是每次 FDP 的长期平均。FDR 控制在 0.05，不代表本次发现集合最多只有 5% 是假的，更不代表每个发现有 95% 的概率为真。",
    ),
    (
        "实验 A：没有真信号，也会出现显著结果",
        "55 秒",
        "我们先看一个知道真相的模拟：1,000 个变量其实都没有效应，重复模拟 1,000 次。未经校正的 raw p 平均发现约 49.944 项，FWER 和 FDR 都是 1，因为全零时只要发现一项，那项就是假发现。这一次模拟里 raw 恰好发现了 42 项，而 Bonferroni 和 BH 都没有发现。校正方法在 1,000 次重复中的错误率约为 0.051。要点是，单次模拟的发现数会波动，长期错误率才体现方法的控制目标。",
    ),
    (
        "实验 B：错误控制与检出力的取舍",
        "60 秒",
        "再看部分真信号的情形：1,000 个变量中前 100 个有真实效应。效应偏移为 1.0 时，BH 平均 power 约 28.65%，Bonferroni 约 5.96%；BH 的 FWER 估计为 0.754，FDR 约为 0.046。BH 控制的是发现集合的 FDR，不是至少犯一次错的概率。效应减弱到 0.5 后，BH 和 Bonferroni 的 power 分别约 0.34% 和 0.23%。没有检出不代表已知真信号不存在，只说明在当前条件下不容易检出。",
    ),
    (
        "实验 C：真实数据的三种发现集合",
        "55 秒",
        "回到真实数据。对 train 的 7,129 个探针检验后，raw p≤0.05 有 1,694 项；Bonferroni 有 26 项；BH 有 535 项。Bonferroni 的 raw p 阈值约为 7.01×10⁻⁶。报告的是 Bonferroni adjusted p 和 BH adjusted p，没有计算 Storey q-value。三种方法发现数不同，是因为控制目标不同。数据没有逐探针真值标签，因此不能知道 535 项里有多少是假发现，也不能叫 535 个真基因。",
    ),
    (
        "先挑最小 p，再按 20 项校正，会漏掉选择机会",
        "60 秒",
        "这里展示一个常见的事后选择问题：先在 7,129 个探针里挑出最小的 20 个 p 值，再把检验家族当成只有 20 项。以首位候选为例，按完整家族计算的 BH adjusted p 是 5.98 乘以 10 的负 5 次方；错误地按 20 项计算会变成 1.68 乘以 10 的负 7 次方，数值小了约 356 倍。问题在于，前面已经看过 7,129 项并挑选了最小值。只有在看这批数据之前，依据独立信息预先指定那 20 项，家族大小为 20 才有相应解释。",
    ),
    (
        "候选探针：统计差异还要看方向与尺度",
        "55 秒",
        "再看首位候选 U22376_cds2_s_at。发现组里 AML 平均值约 702，ALL 平均值约 3,863，AML 减 ALL 的均值差约为负 3,161；95% 逐项 Welch 区间约为负 3,978 到负 2,344，BH adjusted p 约为 5.98 乘以 10 的负 5 次方。负号表示这个探针在 AML 组的测量值较低。由于这是表达强度，差值不能解释为倍数变化；区间也只是这个探针的逐项区间，没有同时覆盖全部 7,129 项。",
    ),
    (
        "预先锁定的 10 个候选：7 个通过留出判据",
        "60 秒",
        "在查看 test 结果前，我们按 train 中 BH 显著、raw p 排序的规则锁定 10 个候选。之后只在 test 检验这 10 项，要求 Bonferroni adjusted p≤0.05，也就是 raw p≤0.005，并且 AML−ALL 方向与 train 一致。结果 10 项方向都一致，7 项通过，3 项未达到阈值。未达到不等于证明没有差异，也不能事后换掉。test 是同一研究的历史留出样本，因此这是留出验证，不是独立外部验证。",
    ),
    (
        "方法选择取决于错误代价",
        "55 秒",
        "方法选择要看研究目的。探索性组学筛选常用 BH 控制 FDR，以较高检出力形成候选；确认性主要终点若任何假阳性的代价都很高，Bonferroni 控制 FWER 可能更合适，但会更严格。Bonferroni 在单项 p 值有效时不要求检验独立；BH 的经典保证依赖独立或适当正依赖。微阵列探针可能相关，一个基因也可能对应多个探针，我们没有充分排除批次和样本独立性问题。因此这些结果适合继续研究，不适合直接做因果或临床结论。",
    ),
    (
        "结论：可以支持什么主张？",
        "55 秒",
        "把结论按主张、证据、推理和边界串起来。我们的主张是，这份数据提供了 AML 与 ALL 之间一些值得继续验证的探针差异。证据是：完整家族中有 535 项 BH 显著，预先锁定的 10 项里有 7 项通过留出判据。完整校正减少了只报告有利结果的风险，锁定后的 test 结果提供了同一研究内的重复证据。但这些结果不能证明 535 项都是独立真实基因，不能证明疾病导致了表达变化，也不能说明候选已有临床价值或能在外部人群复现。",
    ),
    (
        "课堂追问备答",
        "60 秒",
        "最后用几个问题检查概念。第一，BH 显著是不是每个探针有 95% 概率为真？不是，FDR 不是单项后验概率。第二，为什么 Bonferroni 只有 26 项，而 BH 有 535 项？因为前者控制家族至少一次假阳性的概率，后者控制发现集合的错误比例期望，目标不同。第三，7 项通过 test 是不是外部验证？不是，test 来自同一研究。第四，535 项能不能叫 535 个真基因？也不能，因为数据没有逐项真值标签，而且探针与基因不是一一对应。回答追问时，我们应该回到具体数据和方法边界。",
    ),
    (
        "数据来源与分析说明",
        "35 秒",
        "本次分析使用课程压缩包提供的 CSV 和元数据，数据集来源及统计方法文献列在这一页。AI 帮助整理结构、编写核验程序和校对文字；分析程序直接读取数据，并重算真实数据结果和固定种子的模拟。我会对报告中的数字、方法和结论负责。我的汇报到这里，谢谢大家，欢迎提问。",
    ),
]

QUESTIONS = [
    (
        "p-value 是什么？是不是 H0 为真的概率？",
        "不是。p-value 是在 H0 和统计模型成立时，观察到当前或更极端统计量的概率。它既不是 H0 为真的概率，也不表示效应大小。本分析的 H0 是 AML 与 ALL 的总体平均表达相等。",
    ),
    (
        "1,000 个零假设全真时，为什么期望有 50 个假发现？这个结论要求独立吗？",
        "若每项检验恰以 α=0.05 的概率错误拒绝，E[V]=m0×α=1,000×0.05=50；若只知道检验保守有效，则 E[V]≤50。线性期望不要求独立。只有把 FWER 写成 1−(1−0.05)^1000 时，才额外用到了独立假设。",
    ),
    (
        "FDP、FWER 和 FDR 有什么区别？",
        "FDP 是一次研究中假发现占发现数的比例；FWER 是重复研究中至少出现一次假发现的概率；FDR 是重复研究中 FDP 的期望。本次 BH 模拟里 FWER=0.754、FDR=0.046，说明两种目标不能互换。",
    ),
    (
        "为什么 Bonferroni 只有 26 项，BH 有 535 项？",
        "两者控制目标不同。Bonferroni 控制整个家族至少出现一次假阳性的概率，通常更严格；BH 控制发现集合的 FDR，在探索性筛选中通常保留更多候选。本次 26 和 535 都来自同一个 7,129 项 train 家族。",
    ),
    (
        "实际分析中的检验家族 m 是多少？为什么不能只算前 20 项？",
        "发现阶段 m=7,129，因为分析前把 train 中全部探针作为检验家族。前 20 项是看过同一批数据后按 p 值挑出来的，不能删掉前面的搜索机会。test 阶段的 m=10，是在看 test 前锁定候选后定义的验证家族。",
    ),
    (
        "BH adjusted p-value 和 q-value 是一回事吗？",
        "不能不加说明地当成一回事。本次算的是 BH adjusted p-value，没有使用 Storey 方法估计 π0，也没有计算 Storey q-value。汇报时按表里的名称说“BH adjusted p”。",
    ),
    (
        "为什么用 Welch t 检验？它的前提是什么？",
        "两组样本量分别为 27 和 11，Welch 检验不要求两组方差相等。它仍需要单项 p 值有效，并依赖样本独立及适当的抽样模型。我们没有充分排除批次或样本依赖，因此结果按探索性证据解释。",
    ),
    (
        "535 个 BH 显著探针是不是有 5% 假的，或 535 个真基因？",
        "都不能这样说。FDR 0.05 是在条件满足、重复研究意义下对发现集合错误比例期望的控制，不保证本次恰有 5% 是假。数据没有逐项真值，且探针可能相关、一个基因也可能有多个探针。",
    ),
    (
        "为什么 7/10 不能叫外部验证？",
        "这 34 份 test 样本是同一研究中的历史留出样本，只能提供同研究内的留出检查。外部验证需要独立队列，并在分析前固定候选、测量和判据。",
    ),
    (
        "test 中有 3 项没通过，是否说明它们没有差异？",
        "不能。只能说它们没有达到预先规定的候选家族 Bonferroni 阈值。样本量、效应大小和噪声都会影响 power；未拒绝 H0 不等于证明 H0 为真，也不能事后替换候选。",
    ),
    (
        "p 值很小是否表示效应很大？这份数据的效应怎么解释？",
        "不表示。首位探针的 AML−ALL 均值差约 −3,161，95% 逐项 Welch CI 约为 [−3,978, −2,344]，BH adjusted p 约为 5.98×10⁻⁵。差值单位是处理后的表达强度，不是倍数变化；区间也不是 7,129 项的同时区间，筛选后的效应可能偏大。",
    ),
    (
        "可以说 AML 导致了这些表达差异，或这些探针有临床价值吗？",
        "不能。当前分析是亚型间的观察性比较，没有识别因果，也没有提供机制或临床效用证据。需要机制研究、独立队列和适当的临床评价。",
    ),
    (
        "如果显著性阈值或校正方法变了，结论会变吗？",
        "发现数量会变。本次 Bonferroni 得到 26 项，BH 得到 535 项，说明集合对控制目标敏感。总体结论仍应写成“提供值得进一步验证的探索性证据”，而不是把任一数量说成真基因数或因果结论。",
    ),
]


def set_font(run, size=10.5, bold=None, color=INK, italic=None):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    return run


def add_page_number(paragraph):
    run = paragraph.add_run()
    set_font(run, 8, color=MUTED)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run._r.addnext(field)


def configure_document(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.68)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.18

    for name, size in (("Heading 1", 15), ("Heading 2", 11.5)):
        style = styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(INK)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(9 if name == "Heading 1" else 6)
        style.paragraph_format.space_after = Pt(3)

    title_properties = styles["Title"]._element.get_or_add_pPr()
    title_border = title_properties.find(qn("w:pBdr"))
    if title_border is not None:
        title_properties.remove(title_border)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font(header.add_run("生物医学数据科学  |  第 6 讲"), 8, color=MUTED)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(footer.add_run("多重检验校正课堂展示讲稿  ·  "), 8, color=MUTED)
    add_page_number(footer)


def make_document():
    doc = Document()
    configure_document(doc)

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(4)
    set_font(title.add_run("多重检验校正课堂展示讲稿"), 22, True)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(9)
    set_font(subtitle.add_run("第 6 讲  |  建议讲述约 13 分钟，不含老师追问"), 10, color=MUTED)

    intro = doc.add_paragraph()
    intro.paragraph_format.space_after = Pt(8)
    set_font(intro.add_run("使用方式  "), 9.5, True, TEAL)
    set_font(
        intro.add_run("按 PPT 页码逐页讲述。方括号内容为提示，不需要读出来；时间是建议值，可按课堂安排略作删减。"),
        9.5,
        color=MUTED,
    )

    for index, (title_text, duration, speech) in enumerate(SLIDES, 1):
        heading = doc.add_paragraph(style="Heading 2")
        heading.paragraph_format.keep_with_next = True
        set_font(heading.add_run(f"第 {index} 页  {title_text}"), 11.5, True)
        duration_run = heading.add_run(f"  [{duration}]")
        set_font(duration_run, 9, color=TEAL)
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.keep_together = True
        set_font(paragraph.add_run(speech), 10.3)

    doc.add_paragraph("老师可能追问", style="Heading 1")
    method = doc.add_paragraph()
    method.paragraph_format.space_after = Pt(5)
    set_font(method.add_run("答问顺序  "), 9.5, True, TEAL)
    set_font(
        method.add_run("先直接回答判断，再引用本次数据，接着解释对应概念，最后补上结论边界。每题优先用 2 至 4 句话回答。"),
        9.5,
        color=MUTED,
    )
    for question, answer in QUESTIONS:
        q = doc.add_paragraph()
        q.paragraph_format.keep_with_next = True
        q.paragraph_format.space_before = Pt(5)
        q.paragraph_format.space_after = Pt(2)
        set_font(q.add_run(question), 10.3, True, TEAL)
        a = doc.add_paragraph()
        a.paragraph_format.left_indent = Inches(0.16)
        a.paragraph_format.space_after = Pt(4)
        set_font(a.add_run(answer), 10)

    doc.add_page_break()
    doc.add_paragraph("上台前自查", style="Heading 1")
    checklist = [
        ("A1 概念讲对", "能不用稿说明 p-value、检验家族、FDP、FWER、FDR、Bonferroni 与 BH 的区别；不把 FDR 说成单项为真的概率。"),
        ("A2 讲出用途", "能解释为什么先定 m=7,129，为什么探索性筛选报告 BH，确认性高代价终点可考虑 FWER 控制。"),
        ("A3 主线连贯", "能按“问题与假设 → 定义家族 → 单项检验 → 错误控制 → 效应和验证 → 结论边界”串起来。"),
        ("B1 数据与方法", "能说出 72 份样本、train 38 份、test 34 份，比较 ALL 与 AML，使用双侧 Welch 检验。"),
        ("B2 概念落在例子上", "解释 FDR 时引用模拟的 FWER=0.754、FDR=0.046；解释校正时引用真实数据的 26 与 535。"),
        ("B3 报具体结果", "至少能准确说出 raw 1,694、Bonferroni 26、BH 535，以及锁定 10 项中 7 项通过。"),
        ("C1 用结果回答", "遇到追问先给结论，再说数字来自哪一页、对应哪个家族和阈值，不只重复定义。"),
        ("C2 说清边界", "主动说明探针数不等于真基因数，留出样本不等于外部验证，统计关联不等于因果或临床价值。"),
    ]
    for label, text in checklist:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.08)
        p.paragraph_format.first_line_indent = Inches(-0.08)
        p.paragraph_format.space_after = Pt(3)
        set_font(p.add_run(f"{label}  "), 9.5, True, TEAL)
        set_font(p.add_run(text), 9.5)

    doc.add_paragraph("三个通用追问示范", style="Heading 1")
    quick_answers = [
        (
            "这是从哪个数看出来的？",
            "看第 7 页：在完整 m=7,129 的 train 家族中，BH adjusted p≤0.05 的探针有 535 个。这个数是按预设规则得到的发现数，不是真基因数。",
        ),
        (
            "这个数反过来，结论会变吗？",
            "换控制方法后发现数会变，本次 Bonferroni 是 26，BH 是 535；但两者都只支持继续验证的探索性证据，不能推出因果或临床价值。",
        ),
        (
            "这个结论不能说明什么？",
            "不能说明 535 项都是真实独立基因，不能说明 AML 导致表达差异，也不能把同研究留出样本称为外部验证。还需要机制研究和独立队列。",
        ),
    ]
    for question, answer in quick_answers:
        q = doc.add_paragraph()
        q.paragraph_format.keep_with_next = True
        q.paragraph_format.space_before = Pt(3)
        q.paragraph_format.space_after = Pt(1)
        set_font(q.add_run(question), 9.7, True, TEAL)
        a = doc.add_paragraph()
        a.paragraph_format.left_indent = Inches(0.16)
        a.paragraph_format.space_after = Pt(3)
        set_font(a.add_run(answer), 9.5)

    final_note = doc.add_paragraph()
    final_note.paragraph_format.space_before = Pt(4)
    set_font(final_note.add_run("不会的问题  "), 9.5, True, TEAL)
    set_font(
        final_note.add_run("不要猜。先说现有数据能支持到哪一步，再指出尚未验证的条件，以及需要什么额外证据。"),
        9.5,
    )

    doc.core_properties.title = "多重检验校正课堂展示讲稿"
    doc.core_properties.subject = "生物医学数据科学第六讲课堂展示口语讲稿"
    doc.core_properties.author = " "
    doc.save(OUT)
    print(OUT.resolve())


if __name__ == "__main__":
    make_document()
