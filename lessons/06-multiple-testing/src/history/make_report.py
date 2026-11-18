import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(".")
OUT_DIR = ROOT / "outputs"
DATA = json.loads((ROOT / "work/analysis.json").read_text(encoding="utf-8"))
REPORT_PATH = OUT_DIR / "第6讲_多重检验校正_作业报告.docx"

INK = "20292D"
MUTED = "526166"
TEAL = "176D66"
PALE_TEAL = "E8F2EF"
PALE_GRAY = "F2F5F4"
GRID = "D5DEDC"
GOLD = "D29A25"
FONT = "Microsoft YaHei"


def set_run_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if italic is not None:
        run.italic = italic
    return run


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=90, start=110, bottom=90, end=110):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        element = borders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), GRID)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def mark_header_row(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_table(doc, headers, rows, widths, font_size=8.8):
    table = doc.add_table(rows=1, cols=len(headers))
    table.autofit = False
    table.alignment = 1
    set_table_borders(table)
    mark_header_row(table.rows[0])
    for index, text in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.width = Inches(widths[index])
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, TEAL)
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(str(text)), font_size, True, "FFFFFF")
    for row_no, row_data in enumerate(rows):
        row = table.add_row()
        prevent_row_split(row)
        for col_no, value in enumerate(row_data):
            cell = row.cells[col_no]
            cell.width = Inches(widths[col_no])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if row_no % 2 == 1:
                set_cell_shading(cell, PALE_GRAY)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.08
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_no == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_run_font(p.add_run(str(value)), font_size, col_no == 0, INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_body(doc, text, bold_lead=None, after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.18
    if bold_lead and text.startswith(bold_lead):
        set_run_font(p.add_run(bold_lead), 10.2, True, INK)
        set_run_font(p.add_run(text[len(bold_lead):]), 10.2, False, INK)
    else:
        set_run_font(p.add_run(text), 10.2, False, INK)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    set_run_font(p.add_run(text), 15 if level == 1 else 12, True, INK)
    return p


def add_label_value(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    set_run_font(p.add_run(label + "  "), 10.2, True, TEAL)
    set_run_font(p.add_run(value), 10.2, False, INK)
    return p


def add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, 8, False, MUTED)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run._r.addnext(fld)


def apply_doc_styles(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)
    section.header_distance = Inches(0.3)
    section.footer_distance = Inches(0.35)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.font.size = Pt(10.2)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.18

    for style_name, size in (("Heading 1", 15), ("Heading 2", 12)):
        style = doc.styles[style_name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(INK)
        style.paragraph_format.keep_with_next = True

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.paragraph_format.space_after = Pt(0)
    set_run_font(header.add_run("生物医学数据科学  |  第 6 讲"), 8, False, MUTED)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(0)
    set_run_font(footer.add_run("多重检验校正  ·  "), 8, False, MUTED)
    add_page_field(footer)


def make_report():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()
    apply_doc_styles(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(10)
    title.paragraph_format.space_after = Pt(3)
    set_run_font(title.add_run("多重检验校正"), 23, True, INK)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(7)
    set_run_font(subtitle.add_run("显著的结果，值得相信吗？"), 13, True, TEAL)
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.paragraph_format.space_after = Pt(12)
    set_run_font(meta.add_run("生物医学数据科学 · 第 6 讲　　汇报人：________"), 9.5, False, MUTED)

    callout = doc.add_paragraph()
    callout.paragraph_format.left_indent = Inches(0.12)
    callout.paragraph_format.right_indent = Inches(0.12)
    callout.paragraph_format.space_before = Pt(4)
    callout.paragraph_format.space_after = Pt(10)
    callout.paragraph_format.line_spacing = 1.2
    p_pr = callout._p.get_or_add_pPr()
    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), PALE_TEAL)
    p_pr.append(shade)
    border = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "20")
    left.set(qn("w:space"), "7")
    left.set(qn("w:color"), TEAL)
    border.append(left)
    p_pr.append(border)
    set_run_font(callout.add_run("核心判断　"), 10.3, True, TEAL)
    set_run_font(
        callout.add_run(
            "多重校正改变的是一组发现的错误控制方式，不会把显著结果自动变成真实基因或因果证据。"
            "在本次 AML 与 ALL 比较中，7,129 个探针经 BH 校正后有 535 个达到阈值；"
            "预先锁定的 10 个候选中，7 个在同一研究的留出样本中通过候选家族 Bonferroni 规则。"
            "这些结果支持继续验证，不足以证明机制、临床价值或独立外部可重复性。"
        ),
        10.3,
        False,
        INK,
    )

    add_heading(doc, "1 研究问题与分析协议")
    add_body(
        doc,
        "驱动问题：同时检验大量表达特征时，如何控制假阳性，并判断哪些差异值得进入下一步验证？"
        "工作单中的“1,000 个基因、53 个 raw p<0.05”是用于讨论的假设案例，不是本次 CSV 的计算结果。"
    )
    protocol = [
        ("研究对象", "Golub 白血病微阵列表达数据；ALL 与 AML 两种白血病亚型，无健康对照"),
        ("数据规模", "72 份样本、7,129 个探针；train 发现 38 份（ALL 27、AML 11），test 留出 34 份（ALL 20、AML 14）"),
        ("检验家族", "发现阶段使用 train 的全部 7,129 个探针；每个探针各做一个双侧检验"),
        ("单项方法", "Welch t 检验；H0 为两组总体均值相等；效应方向定义为 AML−ALL"),
        ("错误控制", "同一原始 p 值家族分别用 raw 0.05、Bonferroni adjusted p≤0.05、BH adjusted p≤0.05 比较"),
        ("验证规则", "在 BH 显著探针中按 raw p、探针名排序锁定前 10；之后只在 test 检验锁定名单，候选家族 Bonferroni≤0.05 且方向相同"),
    ]
    add_table(doc, ["项目", "预先固定的设置"], protocol, [1.25, 5.45], 8.7)
    add_body(
        doc,
        "保留数据的原始表达尺度。文件包含负值，因此不取 log、不把非正数改成缺失；均值差是表达强度差，不称为 fold change。"
        "sample_id、original_split、diagnosis 是标识或分组字段，不属于探针检验家族。"
    )

    add_heading(doc, "2 核心概念与解释边界")
    concepts = [
        ("p-value", "H0 与模型成立时，观察到当前或更极端统计量的概率", "不是 H0 为真的概率，也不表示效应大小"),
        ("检验家族 m", "为同一研究主张共同考虑的一组检验", "不能看完 p 值后只把报告项算入家族"),
        ("FWER", "P(V>0)，整个家族至少出现一次假发现的概率", "不等于平均假发现数量"),
        ("FDP / FDR", "单次 FDP=V/max(R,1)；FDR=重复研究中 FDP 的期望", "FDR≤0.05 不保证本次 FDP≤0.05，也不代表单项 95% 为真"),
        ("Bonferroni", "p_adj=min(mp,1)，主要控制 FWER", "检验有效时不要求相互独立；较严格，可能降低检出力"),
        ("BH adjusted p", "排序后按秩校正并作单调化，用于 BH 的发现规则", "其 FDR 保证依赖有效检验与独立或适当正依赖；不是自动计算 Storey q-value"),
        ("效应、CI、power", "分别描述差异大小、估计不确定性和检出真信号能力", "逐项 CI 不是家族同时区间；不显著不等于无效应"),
    ]
    add_table(doc, ["概念", "讲对的解释", "不能据此声称"], concepts, [1.25, 2.65, 2.8], 8.35)
    add_body(
        doc,
        "当 1,000 个零假设都成立且每项检验有效时，期望假发现数为 1,000×0.05=50；"
        "这个期望不要求检验相互独立。只有在独立且每项恰好以 0.05 错误拒绝的条件下，"
        "FWER=1−(1−0.05)^1000，接近 1。"
    )

    add_heading(doc, "3 已知真值模拟")
    add_body(
        doc,
        "模拟使用独立正态变量、两组样本量 27 与 11、1,000 个变量、固定随机种子 42，"
        "每种设置重复 1,000 次。已知真值使我们可以数出 V、FDP 与 power；这些模拟变量不是白血病 CSV 中的基因或患者。"
        "模拟用等方差 t 检验，因为生成机制确知正态且等方差；真实数据另用 Welch 检验。"
    )
    add_heading(doc, "3.1 全部零效应", 2)
    null = DATA["simulations"]["all_null"]
    null_rows = []
    for method, label in (("raw", "原始 p"), ("bonferroni", "Bonferroni"), ("bh", "BH")):
        s = null["summary"][method]
        null_rows.append(
            (
                label,
                str(null["first_run_R"][method]),
                f'{s["mean_R"]:.3f} ({s["sd_R"]:.3f})',
                f'{int(s["p05_R"])}–{int(s["p95_R"])}',
                f'{s["FWER"]:.3f}',
                f'{s["FDR"]:.3f}',
            )
        )
    add_table(
        doc,
        ["方法", "第 1 次 R", "平均 R（SD）", "R 的 5%–95%分位", "FWER", "FDR"],
        null_rows,
        [1.0, 0.78, 1.18, 1.25, 0.8, 0.8],
        8.2,
    )
    add_body(
        doc,
        "第一次全零模拟中，raw 规则发现 42 项，Bonferroni 与 BH 都是 0；单次结果不能说明长期表现。"
        "1,000 次汇总里 raw 平均发现约 49.944 项，FWER 与 FDR 估计均为 1.000。"
        "全零时任何发现都是假发现，因此一次实验只要有发现，FDP=1；没有发现时 FDP=0。"
        "校正方法的 FWER/FDR 估计约 0.05，具体有限重复比例会有 Monte Carlo 波动。"
    )
    add_heading(doc, "3.2 部分真信号与检出力", 2)
    sim_rows = []
    for delta_key, delta_label in (("signal_1.0", "δ=1.0"), ("signal_0.5", "δ=0.5")):
        sim = DATA["simulations"][delta_key]["summary"]
        for method, label in (("raw", "raw"), ("bonferroni", "Bonferroni"), ("bh", "BH")):
            s = sim[method]
            sim_rows.append(
                (
                    delta_label,
                    label,
                    f'{s["mean_R"]:.3f}',
                    f'{s["mean_V"]:.3f}',
                    f'{s["FWER"]:.3f}',
                    f'{s["FDR"]:.3f}',
                    f'{100*s["mean_power"]:.2f}%',
                )
            )
    add_table(
        doc,
        ["信号偏移", "方法", "平均 R", "平均 V", "FWER", "FDR", "平均 power"],
        sim_rows,
        [0.9, 1.05, 0.92, 0.92, 0.78, 0.78, 0.95],
        8.1,
    )
    add_body(
        doc,
        "δ=1.0 时 BH 的 FWER 估计为 0.754，但 FDR 约 0.046，说明 BH 目标是发现集合的错误比例期望，"
        "不是“至少一次错误”的概率。Bonferroni 更严格，平均 power 约 5.96%，BH 约 28.65%。"
        "信号减弱至 δ=0.5 后，BH 平均 power 约 0.34%，说明校正后检出真信号更困难；"
        "这不代表模拟中已知的真信号消失。raw 方法保留更多真信号的同时也承受更高错误比例。"
    )

    add_heading(doc, "4 真实表达数据结果")
    add_body(
        doc,
        f'发现阶段在 {DATA["group_counts"]["train"]["ALL"]} 份 ALL 与 '
        f'{DATA["group_counts"]["train"]["AML"]} 份 AML 样本中，对完整的 {DATA["family_size"]:,} 个探针家族逐项检验。'
        "由于真实数据没有单项真值标签，不能从这些发现数直接算出假发现数或真实 power。"
    )
    discovery_rows = [
        ("raw p≤0.05", f'{DATA["counts"]["raw"]:,}', "未校正筛查数"),
        ("Bonferroni adjusted p≤0.05", f'{DATA["counts"]["bonferroni"]:,}', "raw p 阈值 7.01×10⁻⁶"),
        ("BH adjusted p≤0.05", f'{DATA["counts"]["bh"]:,}', "第 535 个排序 p=0.003747"),
    ]
    add_table(doc, ["规则", "发现数", "解释"], discovery_rows, [2.35, 1.0, 3.35], 8.8)
    add_body(
        doc,
        "三种规则得到不同发现集合是控制目标不同的结果，不代表某种算法算错。Bonferroni 通常用于假阳性代价高的确认性检验；"
        "BH 常用于大规模探索性筛选，在假设条件合适时控制 FDR 并保留更多候选。"
    )
    add_heading(doc, "4.1 事后缩小家族会低估选择机会", 2)
    add_body(
        doc,
        "如果先从 7,129 个探针里按 raw p 选前 20 个，再只对这 20 个校正，调整结果就不再对应原来的搜索过程。"
        "首位探针 U22376_cds2_s_at 的完整家族 BH adjusted p 为 5.98×10⁻⁵；"
        "错误地按 m=20 计算会变成 1.68×10⁻⁷。较小的数值没有提供原 7,129 项家族的控制保证。"
        "只有在看这批数据之前就根据外部依据指定 20 个探针，研究问题也限于这 20 项时，m=20 才可能是合理家族。"
    )
    add_heading(doc, "4.2 效应大小与候选锁定", 2)
    detail = DATA["top_probe_detail"]
    add_body(
        doc,
        f'按预定规则锁定的首位候选为 {detail["probe"]}。发现组中 AML n=11、ALL n=27，'
        f'均值分别为 {detail["mean_AML"]:,.2f} 与 {detail["mean_ALL"]:,.2f}，'
        f'均值差 AML−ALL={detail["delta"]:,.2f}，95% 逐项 Welch CI '
        f'[{detail["ci95"][0]:,.2f}, {detail["ci95"][1]:,.2f}]，描述性 Cohen d={detail["cohen_d"]:.2f}。'
        f'其 BH adjusted p={detail["bh"]:.2e}。负差值表示该探针在 AML 组的测量值较低。'
    )
    add_body(
        doc,
        "这里的表达值是处理后的微阵列强度，均值差不能解释为 fold change。Cohen d 用合并组内标准差作标准化描述，"
        "不改变 Welch 检验允许两组方差不同的设定。逐项 CI 没有控制整个家族的同时覆盖率；"
        "候选按显著性筛出，发现组效应还可能有选择偏倚。"
    )

    add_heading(doc, "5 留出验证与结论")
    add_body(
        doc,
        "在查看 test 结果前，按 train 中 BH 显著、raw p 升序并以探针名破同分的规则锁定 10 项。"
        "验证家族大小为 10；判据是 test 内 Bonferroni adjusted p≤0.05，且差值方向与 train 相同。"
    )
    candidate_rows = []
    for i, row in enumerate(DATA["candidates"], 1):
        candidate_rows.append(
            (
                str(i),
                row["probe"],
                f'{row["delta"]:+,.0f}',
                f'{row["test_delta"]:+,.0f}',
                f'{row["test_bonferroni"]:.3g}',
                "通过" if row["replicated"] else "未达阈值",
            )
        )
    add_table(
        doc,
        ["序", "探针", "train Δ", "test Δ", "test Bonf p", "预定判据"],
        candidate_rows,
        [0.38, 1.68, 1.03, 1.03, 1.12, 1.05],
        8.0,
    )
    add_body(
        doc,
        "10 个候选的方向都一致，其中 7 个通过预定显著性判据；另外 3 个没有达到校正阈值。"
        "未达阈值不等于证明无差异，也不能事后替换候选。test 是同一研究中历史留出样本，不是新招募的独立外部队列。"
    )
    add_heading(doc, "主张—证据—推理—边界", 2)
    add_label_value(
        doc,
        "主张",
        "这份数据提供 AML 与 ALL 之间若干可继续验证的探针差异。",
    )
    add_label_value(
        doc,
        "证据",
        "train 的 7,129 项完整家族中，535 项 BH adjusted p≤0.05；预先锁定的 10 项中，7 项按候选家族 Bonferroni 规则通过且方向一致。",
    )
    add_label_value(
        doc,
        "推理",
        "校正后结果比 raw p 筛查更适合形成探索性候选；留出结果为这些候选提供了同研究内的重复证据。",
    )
    add_label_value(
        doc,
        "边界",
        "不能据此声称 535 项都是真实独立基因、差异由疾病亚型因果造成、候选具有临床价值或能在外部人群复现。",
    )

    add_heading(doc, "6 方法选择与可能追问")
    choices = [
        ("探索性组学筛选", "预先定义完整检验家族；常用 BH 控制 FDR，以较低检出力代价保留候选", "候选仍需机制评估与独立验证"),
        ("确认性临床终点", "预先指定主要终点；当任一假阳性的代价很高时，可优先控制 FWER，如 Bonferroni", "更严格的阈值会降低检出力"),
    ]
    add_table(doc, ["情境", "错误控制选择", "代价或后续"], choices, [1.55, 3.05, 2.1], 8.5)
    questions = [
        ("问：BH 显著是否意味着每个探针有 95% 概率为真？",
         "答：不是。FDR 是重复分析中发现集合错误比例的期望，不是单项后验概率。"),
        ("问：为什么 Bonferroni 只找到 26 项，BH 找到 535 项？",
         "答：目标不同。Bonferroni 控制整个家族至少一次假阳性的概率，通常更严格；BH 控制发现集合的 FDR，探索性筛选中通常检出力更高。"),
        ("问：test 有 7 项通过，能否称为外部验证？",
         "答：不能。test 是同一研究预留的历史样本；还需独立招募或独立队列，并预先固定候选、测量和判据。"),
        ("问：为什么不用 7,129 个 p 值直接判断真基因数？",
         "答：真实数据没有真值标签；探针还可能相关、一个基因可有多个探针，显著数不等于真阳性数或独立基因数。"),
    ]
    for q, a in questions:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        set_run_font(p.add_run(q), 9.7, True, TEAL)
        p2 = doc.add_paragraph()
        p2.paragraph_format.left_indent = Inches(0.16)
        p2.paragraph_format.space_after = Pt(5)
        p2.paragraph_format.line_spacing = 1.15
        set_run_font(p2.add_run(a), 9.7, False, INK)

    add_heading(doc, "AI 使用与核验说明")
    add_body(
        doc,
        "AI 用于整理报告和汇报结构、辅助编写统计核验程序及文字校对。程序直接读取所提供的 leukemia.csv，"
        "检查数据形状、组别与 SHA-256，并重算真实数据的 Welch 检验、Bonferroni、BH、候选效应和留出判据；"
        "模拟部分按固定种子、生成机制和重复次数重跑，汇总与随包教师参考表一致。"
    )
    add_body(
        doc,
        "提交/上台前的个人核验：请确认自己能解释 BH 与 Bonferroni 的计算目标、说出首位探针的均值差和 CI，"
        "并说明为什么 7/10 不能叫外部验证。"
    )

    add_heading(doc, "参考资料")
    references = [
        "Golub TR, et al. Molecular Classification of Cancer: Class Discovery and Class Prediction by Gene Expression Monitoring. Science. 1999;286(5439):531–537. doi:10.1126/science.286.5439.531.",
        "Benjamini Y, Hochberg Y. Controlling the False Discovery Rate: A Practical and Powerful Approach to Multiple Testing. Journal of the Royal Statistical Society: Series B. 1995;57(1):289–300.",
        "Storey JD. A Direct Approach to False Discovery Rates. Journal of the Royal Statistical Society: Series B. 2002;64(3):479–498.",
        "OpenML. leukemia data set, ID 1104, version 1. 本报告使用课程压缩包提供的 CSV 与元数据快照。",
        "课程材料：第 6 讲 multiple-test.md、experiment-design.md、data/README.md、teacher-reference.md 与 student-worksheet.md。",
    ]
    for ref in references:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.18)
        p.paragraph_format.first_line_indent = Inches(-0.18)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.12
        set_run_font(p.add_run(ref), 8.6, False, MUTED)

    doc.core_properties.title = "多重检验校正：显著的结果，值得相信吗？"
    doc.core_properties.subject = "生物医学数据科学第六讲 PBL 作业报告"
    doc.core_properties.author = " "
    doc.save(REPORT_PATH)
    print(REPORT_PATH.resolve())


if __name__ == "__main__":
    make_report()
