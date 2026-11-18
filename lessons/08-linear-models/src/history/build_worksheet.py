from pathlib import Path
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(r"C:\Users\A\AppData\Local\Temp\linear-models-review-20261102\8. Linear-models")
ASSETS = ROOT / "assets" / "pbl-linear-models-lda-logistic-regression"
OUT = Path(sys.argv[1])

INK = "202A30"
TEAL = "165B61"
PALE = "EEF4F4"
PALE_BLUE = "F3F6F8"
GRID = "D5DDE0"
MUTED = "52616A"
FONT = "Aptos"
EAST_ASIA_FONT = "Microsoft YaHei"


def set_run_font(run, size=10.5, bold=None, color=INK, name=FONT, italic=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), EAST_ASIA_FONT)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)


def set_style_font(style, size, bold=False, color=INK):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor.from_string(color)
    rpr = style._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), EAST_ASIA_FONT)
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_cell_margins(cell, top=100, start=110, bottom=100, end=110):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "5")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), GRID)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def add_page_number(paragraph):
    run = paragraph.add_run()
    set_run_font(run, size=8, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instr, separate, text, end):
        run._r.append(element)


def configure_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.68)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.72)
    section.right_margin = Inches(0.72)
    section.header_distance = Inches(0.30)
    section.footer_distance = Inches(0.30)

    set_style_font(doc.styles["Normal"], 10.5)
    doc.styles["Normal"].paragraph_format.space_after = Pt(6)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.16
    doc.styles["Normal"].paragraph_format.widow_control = True
    set_style_font(doc.styles["Title"], 23, True, "000000")
    doc.styles["Title"].paragraph_format.space_after = Pt(4)
    title_ppr = doc.styles["Title"]._element.pPr
    if title_ppr is not None:
        title_border = title_ppr.find(qn("w:pBdr"))
        if title_border is not None:
            title_ppr.remove(title_border)
    set_style_font(doc.styles["Subtitle"], 11, False, MUTED)
    doc.styles["Subtitle"].paragraph_format.space_after = Pt(10)
    set_style_font(doc.styles["Heading 1"], 16, True, "000000")
    doc.styles["Heading 1"].paragraph_format.space_before = Pt(15)
    doc.styles["Heading 1"].paragraph_format.space_after = Pt(7)
    doc.styles["Heading 1"].paragraph_format.keep_with_next = True
    set_style_font(doc.styles["Heading 2"], 12.5, True, "000000")
    doc.styles["Heading 2"].paragraph_format.space_before = Pt(9)
    doc.styles["Heading 2"].paragraph_format.space_after = Pt(4)
    doc.styles["Heading 2"].paragraph_format.keep_with_next = True
    set_style_font(doc.styles["Heading 3"], 11, True, "000000")
    doc.styles["Heading 3"].paragraph_format.space_before = Pt(7)
    doc.styles["Heading 3"].paragraph_format.space_after = Pt(3)
    doc.styles["Heading 3"].paragraph_format.keep_with_next = True
    for style_name in ("List Bullet", "List Number"):
        set_style_font(doc.styles[style_name], 10.5)
        doc.styles[style_name].paragraph_format.space_after = Pt(3)
        doc.styles[style_name].paragraph_format.line_spacing = 1.12

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run("第 8 讲  线性模型与概率分类")
    set_run_font(run, size=8, color=MUTED)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("学生工作单完成版  |  ")
    set_run_font(run, size=8, color=MUTED)
    add_page_number(footer)
    return doc


doc = configure_document()


def add_para(text="", bold_lead=None, style=None, align=None, after=6, keep=False):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.widow_control = True
    if keep:
        p.paragraph_format.keep_with_next = True
    if bold_lead and text.startswith(bold_lead):
        lead = p.add_run(bold_lead)
        set_run_font(lead, bold=True)
        tail = p.add_run(text[len(bold_lead):])
        set_run_font(tail)
    else:
        run = p.add_run(text)
        set_run_font(run)
    return p


def add_bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.14)
    run = p.add_run(text)
    set_run_font(run)
    return p


def add_heading(text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    run = p.add_run(text)
    set_run_font(run, size={1: 16, 2: 12.5, 3: 11}[level], bold=True, color="000000")
    return p


def add_formula(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(7)
    run = p.add_run(text)
    set_run_font(run, size=10.5, bold=True, name="Cambria")
    return p


def add_caption(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    set_run_font(run, size=8.5, color=MUTED)
    return p


def add_table(headers, rows, widths=None, aligns=None, font_size=9.1):
    table = doc.add_table(rows=1, cols=len(headers))
    table.autofit = False
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if widths is None:
        widths = [6.83 / len(headers)] * len(headers)
    for i, width in enumerate(widths):
        table.columns[i].width = Inches(width)
    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = Inches(widths[i])
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, TEAL)
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(str(text))
        set_run_font(r, size=font_size, bold=True, color="FFFFFF")
    set_repeat_table_header(table.rows[0])
    prevent_row_split(table.rows[0])

    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for col_index, value in enumerate(values):
            cell = cells[col_index]
            cell.width = Inches(widths[col_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if row_index % 2 == 1:
                set_cell_shading(cell, PALE_BLUE)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.08
            if aligns and col_index < len(aligns):
                p.alignment = aligns[col_index]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(value))
            set_run_font(r, size=font_size, color=INK)
        prevent_row_split(table.rows[-1])
    set_table_borders(table)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(3)
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.line_spacing = 0.5
    return table


# Opening
p = doc.add_paragraph(style="Title")
r = p.add_run("第 8 讲 线性模型与概率分类工作单")
set_run_font(r, size=23, bold=True, color="000000")
p = doc.add_paragraph(style="Subtitle")
r = p.add_run("Logistic Regression 与 LDA 的假设、概率和决策比较")
set_run_font(r, size=11, color=MUTED)
add_para("姓名：南铂山    学号：24520261155028    日期：2026-11-02", after=8)
add_para(
    "本工作单完成建模对象、参数量、无泄漏流程、错误代价、系数稳定性和 CERB 结论。按随包结果，LDA 在本次 5×FN+FP 规则下的代价略低，但不能据此宣称它普遍优于 Logistic。两项资料限制会影响结论强度：压缩包包含两份不一致的数据文件；完整版 notebook 在 Packet D 前已使用测试集标签做诊断和压力测试，因此所谓“锁定测试”并非真正未触碰的盲测。",
    bold_lead="本工作单完成",
    after=8,
)

# Task 1
add_heading("1 模型对象 参数与目标")
add_para("Logistic regression 和 LDA 都能给出线性分类边界和概率，但它们从不同方向建模。")
add_table(
    ["问题", "Logistic regression", "LDA"],
    [
        ["直接建模的对象", "条件概率 P(Y=1|X=x)。", "类条件分布 P(X|Y=k) 与类别先验 P(Y=k)。"],
        ["主要参数", "截距 β₀ 与特征系数 β。", "每类均值 μₖ、共享协方差 Σ、类别先验 πₖ。"],
        ["依据什么估计", "带标签的 (X,Y) 中，特征与结局的条件关系。", "每个类别内 X 的分布、两类共同的协方差结构和类别比例。"],
        ["主要目标", "最大化条件似然，等价于最小化二分类交叉熵。", "在高斯类条件模型下估计类条件似然与先验，再用 Bayes 定理求 posterior。"],
        ["概率从哪里来", "把线性 log-odds z 送入 sigmoid。", "把高斯密度乘先验，并对所有类别归一化。"],
    ],
    widths=[1.28, 2.78, 2.77],
)
add_para("两条概率流程", bold_lead="两条概率流程", after=2)
add_formula("Logistic：x → z=β₀+xᵀβ → sigmoid(z) → P(Y=1|x) → 阈值 t → 类别")
add_formula("LDA：x 与 (μₖ,Σ,πₖ) → p(x|Y=k)πₖ → Bayes 归一化 → P(Y=k|x) → 类别")
fig_path = ASSETS / "concept-logistic-lda-intuition.png"
if fig_path.exists():
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    picture = p.add_run().add_picture(str(fig_path), width=Inches(6.72))
    picture._inline.docPr.set("title", "Logistic sigmoid 与 LDA 线性边界示意")
    picture._inline.docPr.set("descr", "三幅一维示意图依次显示 sigmoid 概率曲线、LDA 两类高斯密度，以及共享协方差时的线性判别分数差。")
    add_caption("图 1  sigmoid 与 LDA 共享协方差的直观示意。该图是一维概念图，不是本数据的拟合结果。")
add_para(
    "两种方法都可以输出 posterior，也都可能产生线性边界。Logistic 的线性部分直接表达条件 log-odds；LDA 的线性判别分数来自高斯密度与先验。若特征与结局的条件 log-odds 关系近似线性，且重点是直接估计 P(Y|X)，Logistic 是自然起点；若类条件高斯和共享协方差近似可信，且维度相对样本量可控，LDA 可借助这些结构假设提高估计效率。不能只按某一个 AUC 选方法。",
)

# Task 2
add_heading("2 线性评分与决策边界")
add_formula("z = β₀ + xᵀβ")
add_para("x 是 d 维输入向量；β 是与各输入对应的权重；β₀ 是截距；z 是尚未压缩到 0–1 范围的线性评分。给定 z=c 时，满足 β₀+xᵀβ=c 的点构成一个超平面；二维时是一条直线。")
add_bullet("Logistic 的 z 是 sigmoid 前的 log-odds 评分。类别阈值 t 对应 z=log[t/(1−t)]。当 t=0.5 时，z=0。")
add_bullet("两类 LDA 比较判别函数 δ₁(x) 与 δ₀(x)。共享 Σ 时，二者差值可写成 wᵀx+b；边界由 δ₁(x)−δ₀(x)=0 给出。")
add_bullet("不同系数的大小只有在单位和编码一致时才可比较。本模型把六个数值输入标准化到开发集标准差单位，三个二元变量保持 0/1，四个名义变量用参照水平编码；相关变量同时入模时，系数还会受共线性影响。")
add_bullet("只调整概率阈值会改变决策边界的位置和类别标签，不会重估 β、μ 或 Σ，不会改变原始概率，也不会改变基于概率排序计算的 AUC。")

# Task 3
add_heading("3 Logistic Regression 的目标与参数量")
add_formula("p = 1/(1+e^(−z))     log[p/(1−p)] = z")
add_para("sigmoid 把任意实数 z 映射到 0 和 1 之间；logit 则把概率换回 log-odds。拟合时，模型为每个样本产生 pᵢ，并寻找 β₀、β，使真实标签获得较高概率。")
add_formula("L(β₀,β) = −Σᵢ [yᵢ log(pᵢ) + (1−yᵢ)log(1−pᵢ)]")
add_para("Notebook 的 LogisticRegression 使用 C=∞，没有施加有限强度的 L2 惩罚；拟合目标是降低上述条件交叉熵。训练数据标签进入损失，预测新个体时只需输入 X。")
add_table(
    ["编码块", "维度计算", "参数数"],
    [
        ["标准化数值特征", "age、resting_bp、serum_cholesterol、max_heart_rate、oldpeak、major_vessels", "6"],
        ["二元特征", "sex、fasting_blood_sugar、exercise_angina，各 1 列", "3"],
        ["chest_pain", "4 个水平，drop-first，4−1", "3"],
        ["resting_ecg", "3 个水平，drop-first，3−1", "2"],
        ["st_slope", "3 个水平，drop-first，3−1", "2"],
        ["thal", "3 个水平，drop-first，3−1", "2"],
        ["合计", "编码后 q=6+3+3+2+2+2=18；另加截距", "19"],
    ],
    widths=[1.52, 4.10, 1.21],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
)
add_para("四个分类变量分别以 chest_pain=1、resting_ecg=0、st_slope=1、thal=3 为参照水平。参照水平改变时，虚拟变量系数的含义随之改变；若对每个水平都编码且不设参照，可能出现完全共线。", bold_lead="参照水平")
add_para(
    "例：major_vessels 的开发集系数 β≈1.437，OR=e^β≈4.21。该变量在本流程中作为数值变量标准化，因此解释单位是增加 1 个开发集标准差，而不是增加 1 条血管；在其他已编码变量保持不变的模型条件下，估计 odds 约乘以 4.21。OR 不是概率倍数，也不是因果效应。Bootstrap 的系数 95% 区间约为 1.00–3.17，对应 OR 区间约 2.71–23.9，显示量级仍不精确。",
    bold_lead="条件性 OR",
)

# Task 4
add_heading("4 LDA 的假设与参数量")
add_formula("X|Y=k ~ N(μₖ,Σ)")
add_para("μₖ 描述第 k 类在各维特征上的平均位置；Σ 描述类内各特征的方差与协方差；πₖ=P(Y=k) 是类别先验。LDA 先以高斯密度估计 P(X|Y=k)，再计算 P(Y=k|X=x) ∝ p(x|Y=k)πₖ，最后按类别归一化。")
add_table(
    ["参数块", "两类公式", "本课 d=18"],
    [
        ["两类均值", "K×d", "2×18=36"],
        ["共享协方差", "d(d+1)/2", "18×19/2=171"],
        ["独立先验", "K−1", "1"],
        ["LDA 合计", "Kd+d(d+1)/2+(K−1)", "208"],
        ["类特异协方差", "K×d(d+1)/2", "2×171=342"],
        ["QDA 合计", "Kd+K d(d+1)/2+(K−1)", "379"],
    ],
    widths=[1.55, 3.45, 1.83],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
)
add_para("比较两类高斯对数密度时，共享 Σ 使两类共同的 −½xᵀΣ⁻¹x 项相消，剩下线性项和常数项，因此边界是线性的。若每类有自己的 Σₖ，二次项不能相消；边界通常变成二次曲面，参数量也增加，协方差估计的方差随之上升。")
add_para("本课 d=18，而每个开发集训练折约有 172–173 条记录。208 是均值、协方差和先验的名义参数计数，不等于 208 个互相独立的回归系数；但 pooled covariance 含 171 个元素，样本量不大时其估计可能波动。若 d 继续增加，协方差参数按 d² 增长，宜评估收缩 LDA 或降维，而非只看总样本数。")
add_para("LDA 可以输出 posterior probability。它与 Logistic 的差异不在“能否输出概率”，而在概率来源：LDA 依赖类条件分布、共享协方差和先验；Logistic 直接估计条件 log-odds。两种概率都需要单独检查校准。此处的 LDA 概率路线与 scikit-learn User Guide 的线性判别分析定义一致。")

# Task 5
add_heading("5 数据审计与无泄漏流程")
add_table(
    ["审计项", "Statlog 结果文件", "记录"],
    [
        ["分析文件", "assets/.../statlog-heart.csv", "270 行；13 个输入列及 1 个结局列。"],
        ["结局", "heart_disease_present", "阳性 120，阴性 150；阳性比例 44.4%。"],
        ["缺失", "statlog-heart.csv", "0 个缺失单元格；无需缺失值插补。"],
        ["划分", "80% development / 20% test", "开发集 216（96 阳性、120 阴性）；测试集 54（24 阳性、30 阴性）。"],
        ["随机协议", "stratified, random_state=2026", "按结局比例分层，保留固定随机种子。"],
    ],
    widths=[1.25, 2.30, 3.28],
)
add_para("来源：UCI Machine Learning Repository, Statlog (Heart), DOI 10.24432/C57303，CC BY 4.0。")
add_heading("资料文件的一致性", 2)
add_para(
    "数据源差异：包内的 data/heart.csv 不是上述 Statlog 分析文件：它有 270×18 个字段、557 个缺失单元格，并包含 sample_id、record_id、site 和 num；其中 `ca` 与 `thal` 缺失尤其多。学生工作单和 README 却把该文件描述为已无缺失，短版 notebook 也指向 data/heart.csv。相反，完整版 notebook 和 packet-B 至 packet-D 的结果表都使用 assets/.../statlog-heart.csv，后者才与 270 条、13 个原始特征、零缺失的分析设计一致。因此，本工作单的数值按完整版 notebook 及其 Statlog 输出表填写；若运行短版 notebook 指向另一文件，结果不能与这些表直接混用。",
    bold_lead="数据源差异：",
)
add_para(
    "排除字段：不能作为输入的字段包括结局 heart_disease_present、原始结局 num、sample_id/record_id 等标识符，以及 site。结局和 num 会直接泄漏答案；标识符没有可解释的预测含义；site 可能让模型记住中心构成，妨碍跨中心迁移。本课完整 Statlog 文件本身只保留 13 项预测变量和重编码结局。",
    bold_lead="排除字段：",
)
add_para("无泄漏顺序：先分层拆出开发集和测试集；对每个开发集训练折，仅在训练折拟合 StandardScaler 与 OneHotEncoder；将已拟合的转换器应用于该折验证数据；在训练折估计模型，再计算验证预测和指标。若先用全体样本计算均值、标准差或分类水平再交叉验证，验证折会影响预处理参数，验证结果将偏乐观。")
add_para(
    "协议偏离：进一步核对发现，完整版 notebook 的 Packet C 又用全体 X、y 计算类内偏态和协方差；并在 Packet D 前用 X_test、y_test 比较加入近重复年龄变量前后的 AUC。这让测试集参与了诊断和压力测试。即使 Packet D 没有再调阈值，这份测试集也不再是严格未触碰的最终检验；正确做法是把上述诊断限制在 development 内，并另留一个未查看的外部或新划分测试集。",
    bold_lead="协议偏离：",
)

# Task 6
add_heading("6 比较假设后果与预测证据")
add_para("先区分建模路线和参数规模，再看 development 上 5 折×5 次重复分层交叉验证。Logistic 与 LDA 的预处理均在每个训练折内拟合。")
add_table(
    ["模型", "概率路线与主要假设", "名义参数量"],
    [
        ["Prevalence baseline", "不使用特征；只估计类别先验。", "1"],
        ["Gaussian NB", "类条件逐特征高斯；给定类别后条件独立。", "约 73"],
        ["LDA", "类条件多元高斯；各类别共享协方差。", "208"],
        ["Logistic regression", "直接建模条件概率；log-odds 对输入线性。", "19"],
    ],
    widths=[1.55, 4.10, 1.18],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    font_size=8.8,
)
add_para("Gaussian NB 的约 73 个参数由 36 个类均值、36 个类内方差和 1 个独立先验组成。")
add_para("下表数值为 25 个验证折的均值±折间标准差；Sensitivity 和 Specificity 用 0.5 阈值。折间标准差是描述性波动，不是置信区间；重复折彼此重叠，不能据此作显著性判断。Brier 和 log loss 越低越好。")
add_table(
    ["模型", "ROC-AUC", "PR-AUC", "Sens. @0.5", "Spec. @0.5", "Brier", "Log loss"],
    [
        ["Prevalence baseline", "0.500±0.000", "0.444±0.005", "0.000±0.000", "1.000±0.000", "0.247±0.001", "0.687±0.001"],
        ["Gaussian NB", "0.861±0.038", "0.823±0.071", "0.743±0.177", "0.722±0.210", "0.233±0.078", "1.658±0.924"],
        ["LDA", "0.910±0.030", "0.906±0.035", "0.777±0.083", "0.877±0.057", "0.125±0.027", "0.409±0.099"],
        ["Logistic regression", "0.903±0.033", "0.896±0.037", "0.796±0.062", "0.862±0.076", "0.128±0.027", "0.429±0.108"],
    ],
    widths=[1.28, 0.91, 0.91, 1.03, 1.03, 0.80, 0.87],
    aligns=[
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
    ],
    font_size=7.8,
)
add_para("LDA 与 Logistic 的交叉验证 ROC-AUC 只差约 0.008，PR-AUC 也接近。LDA 的 Brier 和 log loss 略低，Logistic 的 sensitivity 均值略高；没有一个模型在所有指标上占优。LDA 平均 AUC 略高不足以证明更适合。还需看校准、错误代价、模型假设与参数稳定性，以及真正未触碰数据上的表现。若 Logistic 的 AUC 略高，也不意味着 LDA 毫无价值：在共享协方差假设接近成立时，生成式结构可能有效率优势；但要由目标数据验证，而非凭算法标签推断。")
add_para("本任务中 baseline 在 0.5 阈值下 sensitivity=0，是因为它总把样本判为阴性；其 specificity=1 并不代表实用。AUC 描述排序，Brier 与 log loss 衡量概率误差，阈值下的 FN/FP 反映行动后果，三者不能互相替代。")

# Task 7
add_heading("7 阈值与 5 比 1 错误代价")
add_para("阈值在 development 的 5 折 OOF 概率上选择，代价为 C=5×FN+FP；审计网格为 0.05 至 0.95、步长 0.005。按 packet-C-threshold-audit.csv 导出的选择结果，Logistic 阈值为 0.14，LDA 为 0.11。")
add_table(
    ["模型", "阈值", "开发集 OOF 混淆矩阵 TN/FP/FN/TP", "Sensitivity", "Specificity", "Precision", "代价"],
    [
        ["Logistic", "0.14", "76 / 44 / 6 / 90", "0.938", "0.633", "0.672", "74"],
        ["LDA", "0.11", "80 / 40 / 7 / 89", "0.927", "0.667", "0.690", "75"],
    ],
    widths=[0.91, 0.62, 2.16, 0.84, 0.84, 0.78, 0.68],
    aligns=[
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
    ],
    font_size=8.3,
)
add_para("开发 OOF 上 Logistic 的代价为 44+5×6=74，LDA 为 40+5×7=75。由此可见，选阈值不只是“越低越好”：降低阈值通常减少 FN，却会增加 FP；具体最小代价取决于概率排序和样本。若概率已校准且只有 FP=1、FN=5 两种误判成本，理论 Bayes 阈值为 1/(1+5)=0.167；本题的经验 OOF 阈值是有限样本、给定网格下的最优值，不应解释成临床通用界线。")
add_heading("随包测试输出", 2)
add_para("以下是资产中已计算的测试输出。因为 Task 5 所述的测试标签提前使用，这些数字只能作为该次划分的内部描述，不是盲测确认。")
add_table(
    ["模型", "阈值", "测试混淆矩阵 TN/FP/FN/TP", "Sensitivity", "Specificity", "Precision", "代价"],
    [
        ["Logistic", "0.14", "15 / 15 / 1 / 23", "0.958", "0.500", "0.605", "20"],
        ["LDA", "0.11", "18 / 12 / 1 / 23", "0.958", "0.600", "0.657", "17"],
    ],
    widths=[0.91, 0.62, 2.16, 0.84, 0.84, 0.78, 0.68],
    aligns=[
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
    ],
    font_size=8.3,
)
add_para("测试集有 24 个阳性和 30 个阴性。两模型都检出 23/24 个阳性；LDA 少 3 个误报，因此在已给定的测试输出中代价为 17，低于 Logistic 的 20。模型参数和概率没有因阈值改变；只有将概率转成阳性/阴性的规则改变。")
add_para("锁定测试被反复用于选择特征、模型或阈值后，就成为开发资料。此时再报告其性能会低估选择造成的乐观偏差，不能把它称为最终独立评价。")

# Task 8
add_heading("8 系数稳定性与 LDA 假设审计")
add_para("LDA 的理论假设只近似成立。连续特征的类内偏态审计中，major_vessels 的最大绝对偏度约 2.483，serum_cholesterol 约 1.783，oldpeak 约 1.556；两类连续变量协方差矩阵的相对 Frobenius 差约 0.559。此外，数据还含二元和 one-hot 离散变量，本来就不是连续多元高斯。")
add_para("这些诊断提示假设压力，不是机械的弃用规则。模型在该拆分仍显示良好排序和概率误差，但因为诊断用到了全数据，不能把它们当作只在开发集完成的独立假设检验。后续应在开发集做类别分布、协方差、异常值与校准检查，并比较 shrinkage LDA、把 major_vessels 按类别编码等敏感性方案。")
add_table(
    ["Logistic 特征", "β", "OR=e^β", "Bootstrap β 95% 区间", "符号稳定率"],
    [
        ["major_vessels（每 1 SD）", "1.437", "4.209", "0.996 至 3.173", "1.000"],
        ["age（每 1 SD）", "−0.145", "0.865", "−0.903 至 0.445", "0.700"],
        ["chest_pain=2 vs 1", "1.593", "4.917", "−1.544 至 4.906", "0.863"],
        ["chest_pain=4 vs 1", "2.578", "13.167", "1.023 至 6.175", "1.000"],
    ],
    widths=[1.77, 0.63, 0.78, 2.07, 1.08],
    aligns=[
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
    ],
    font_size=8.5,
)
add_para("系数表来自开发集的 300 次非参数 bootstrap。解释例子：在当前编码与 Logistic 模型中，major_vessels 每增加 1 个开发集标准差，在其他入模变量固定时，与心脏病结局 odds 约乘以 4.21 相关。major_vessels 实为 0–3 的计数，本模型将它当作标准化数值变量，因此隐含线性 log-odds 趋势；该 OR 不是“概率增加 4.21 倍”，也不能支持因果解释。")
add_para("加入人为构造的近重复年龄变量后，age 系数从 −0.145 变为 18.389，近重复项系数为 −18.568；随包 AUC 从 0.908 变为 0.917。高度相关变量可以用很大的相反权重生成相近的总评分，因此单个系数可能剧烈变化而预测排序变化较小。近重复项只是压力测试，不是原始变量；由于压力测试用到测试标签，AUC 差异仅作探索性说明。")
add_para("预测稳定不等于系数稳定。共线性会让信息在相关变量之间重新分配，变量层面的解释比总体预测更敏感。报告时应同时查看 bootstrap 区间、方向稳定率、校准和阈值指标，并在独立样本复现。")

# Task 9
add_heading("9 冻结协议与 CERB 结论")
add_para("本次冻结的候选方案为 LDA、训练折内标准化与 one-hot、development OOF 选择的阈值 0.11、错误代价 5×FN+FP。评价区分时报告 ROC-AUC 与 PR-AUC；评价概率时报告 Brier 与 log loss；评价行动时报告混淆矩阵、sensitivity、specificity、precision 和加权代价。")
add_table(
    ["模型", "阈值", "ROC-AUC", "PR-AUC", "Accuracy", "Balanced acc.", "Brier", "Log loss"],
    [
        ["LDA", "0.11", "0.906", "0.896", "0.759", "0.779", "0.116", "0.438"],
        ["Logistic", "0.14", "0.908", "0.896", "0.704", "0.729", "0.111", "0.429"],
    ],
    widths=[0.95, 0.58, 0.78, 0.78, 0.82, 1.02, 0.84, 1.06],
    aligns=[
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
    ],
    font_size=8.1,
)
add_para("点估计取自 packet-d-locked-test.csv；对应混淆矩阵与错误代价见任务 7。LDA 的 AUC 并非最高，但在该固定代价下少 3 个加权错误单位；Logistic 的 Brier 与 log loss 略低。")
add_table(
    ["指标", "LDA Bootstrap 中位数与 95% 区间", "Logistic Bootstrap 中位数与 95% 区间"],
    [
        ["ROC-AUC", "0.910（0.806–0.978）", "0.911（0.806–0.979）"],
        ["PR-AUC", "0.902（0.799–0.974）", "0.903（0.796–0.977）"],
        ["Brier", "0.115（0.053–0.195）", "0.111（0.054–0.184）"],
        ["Sensitivity", "0.958（0.875–1.000）", "0.958（0.875–1.000）"],
        ["Specificity", "0.600（0.433–0.767）", "0.500（0.333–0.667）"],
    ],
    widths=[1.38, 2.72, 2.73],
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
)
add_para("区间来自对已拟合模型的测试集预测做 1,000 次分层 bootstrap；它没有把模型重新拟合的不确定性、数据来源偏差或测试集已被提前查看的不确定性纳入。两模型的区间明显重叠。")
add_heading("CERB")
add_para(
    "Claim：若问题严格限定为本次教学数据和 5×FN+FP 的内部比较，我暂选 LDA，并冻结阈值 0.11。这个选择不是“LDA 总体更好”，而是一个依赖当前数据划分、代价规则与编码的候选决定。",
    bold_lead="Claim：",
)
add_para(
    "Evidence：development 重复交叉验证的 ROC-AUC 为 LDA 0.910±0.030、Logistic 0.903±0.033，差值约 0.008；测试输出中两者 AUC 分别为 0.906 和 0.908。按冻结阈值，二者 sensitivity 都为 0.958；LDA specificity 为 0.600、代价为 17，Logistic specificity 为 0.500、代价为 20。Bootstrap 区间宽且重叠；Logistic 的 Brier 略低（0.111 对 0.116）。",
    bold_lead="Evidence：",
)
add_para(
    "Reasoning：漏诊成本设为误报的五倍，因此先关注 FN，再比较相同 FN 下的 FP。当前测试输出两模型各漏掉 1 例，LDA 少误报 3 例，所以按给定公式少 3 个代价单位。与此同时，AUC、PR-AUC 和校准误差都很接近，没有证据表明某模型在所有方面占优；LDA 的生成分布假设也受到离散变量、偏态与协方差差异挑战。",
    bold_lead="Reasoning：",
)
add_para(
    "Boundary：数据是小型教学样本，测试集只有 54 条记录，且测试标签已在 Packet C 的假设审计和年龄压力测试中使用。因此这些数字不能称为独立盲测证据，不能声称 LDA 普遍优于 Logistic，不能推断临床获益或跨医院有效性，也不能把 major_vessels 的 OR 写成干预因果效应。",
    bold_lead="Boundary：",
)
add_para(
    "Next step：先统一数据源，使课程 worksheet、README、短版 notebook 与完整版 notebook 指向同一文件；把所有假设审计和变量压力测试限于 development；在未接触的独立测试集或外部中心上按预先冻结的流程评价校准、阈值代价和亚组表现。若模型将用于临床决策，还需按目标人群重新验证患病率、资源约束和错误成本。",
    bold_lead="Next step：",
)

# Homework / exit ticket
add_heading("作业与退出卡")
add_heading("1 两种模型有什么不同", 2)
add_para("Logistic 直接建模 P(Y|X)，通过最小化交叉熵估计 log-odds 系数；LDA 建模 P(X|Y=k) 和类别先验，再经 Bayes 定理得到 P(Y|X)。前者的主要函数形式假设是 log-odds 对输入近似线性；后者要求类条件分布近似多元高斯且类别共享协方差。")
add_heading("2 哪类参数会快速增加", 2)
add_para("LDA 的共享协方差需要 d(d+1)/2 个参数，随维度按平方增长。本课 d=18 时有 171 个协方差参数；维度继续上升而样本量不变，会增加协方差估计噪声。")
add_heading("3 为什么 OR=2 不代表患病概率翻倍", 2)
add_para("OR 比较 odds，不是 probability。若起始概率 p=0.10，odds=1/9；odds 翻倍后为 2/9，对应概率 2/11≈0.182，而不是 0.20。起始概率不同，odds 翻倍对应的概率变化也不同。")
add_heading("4 为什么 AUC 校准与固定阈值错误不能互换", 2)
add_para("AUC 看阳性是否整体排在阴性之前；Brier 和 log loss 看概率离真实标签有多远；阈值后的 FN、FP、sensitivity、specificity 与 precision 看决策结果。模型可有高 AUC 却概率失准，也可在一个阈值下漏诊多，故需根据任务并列报告。")
add_heading("5 按假设和目标选择方法", 2)
add_para("先明确目标是排序、风险概率还是带代价的行动，再检查数据类型、样本量、维度、校准和错误代价。若条件 log-odds 近似线性且希望直接建模 P(Y|X)，优先评估 Logistic；若类条件高斯和共享协方差近似合理、协方差可稳定估计，可评估 LDA。用 development 选择阈值并冻结，最终只在未参与选择的数据上评估；不以单一分数宣布赢家。")

OUT.parent.mkdir(parents=True, exist_ok=True)
doc.core_properties.title = "第 8 讲 线性模型与概率分类工作单"
doc.core_properties.subject = "LDA 与 Logistic Regression 的假设、概率、参数与决策"
doc.core_properties.author = "南铂山"
doc.save(OUT)
print(f"Wrote {OUT}")
