import csv
import io
import json
import math
import os
import statistics
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph


ROOT = Path(__file__).resolve().parent
CONTENT_PATH = ROOT / "presentation-content.json"
DATA_ZIP = Path(r"C:\Users\A\Downloads\hypothesis.zip")
SOURCE_WORKSHEET = Path(
    r"C:\Users\A\xwechat_files\wxid_kbcgskbdfxu922_2eec\msg\file\2026-10\pbl-hypothesis-student-worksheet.docx"
)
OUTPUT_DIR = ROOT / "deliverables"
WORKSHEET_OUT = Path(r"C:\Users\A\Downloads\pbl-hypothesis-student-worksheet-completed.docx")
SCRIPT_OUT = Path(r"C:\Users\A\Downloads\统计检验与假设设计_逐页演讲稿.docx")


def regularized_beta(a, b, x):
    """Regularized incomplete beta, used for a Student t tail probability."""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0

    def fraction(aa, bb, xx):
        max_iter = 300
        eps = 3e-14
        tiny = 1e-300
        qab = aa + bb
        qap = aa + 1.0
        qam = aa - 1.0
        c = 1.0
        d = 1.0 - qab * xx / qap
        if abs(d) < tiny:
            d = tiny
        d = 1.0 / d
        h = d
        for m in range(1, max_iter + 1):
            m2 = 2 * m
            term = m * (bb - m) * xx / ((qam + m2) * (aa + m2))
            d = 1.0 + term * d
            if abs(d) < tiny:
                d = tiny
            c = 1.0 + term / c
            if abs(c) < tiny:
                c = tiny
            d = 1.0 / d
            h *= d * c
            term = -(aa + m) * (qab + m) * xx / ((aa + m2) * (qap + m2))
            d = 1.0 + term * d
            if abs(d) < tiny:
                d = tiny
            c = 1.0 + term / c
            if abs(c) < tiny:
                c = tiny
            d = 1.0 / d
            delta = d * c
            h *= delta
            if abs(delta - 1.0) < eps:
                break
        return h

    log_bt = (
        math.lgamma(a + b)
        - math.lgamma(a)
        - math.lgamma(b)
        + a * math.log(x)
        + b * math.log1p(-x)
    )
    bt = math.exp(log_bt)
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * fraction(a, b, x) / a
    return 1.0 - bt * fraction(b, a, 1.0 - x) / b


def t_two_sided_p(t_value, degrees_freedom):
    x = degrees_freedom / (degrees_freedom + t_value * t_value)
    return regularized_beta(degrees_freedom / 2.0, 0.5, x)


def t_critical_975(degrees_freedom):
    low, high = 0.0, 8.0
    for _ in range(100):
        mid = (low + high) / 2
        if t_two_sided_p(mid, degrees_freedom) > 0.05:
            low = mid
        else:
            high = mid
    return (low + high) / 2


def read_primary_data():
    with zipfile.ZipFile(DATA_ZIP) as archive:
        text = archive.read("data/heart-disease-hypothesis.csv").decode("utf-8-sig")
    rows = list(csv.DictReader(io.StringIO(text)))
    groups = {0: [], 1: []}
    missing = {0: 0, 1: 0}
    sites = set()
    raw_counts = {0: 0, 1: 0}
    for row in rows:
        group = int(float(row["disease"]))
        raw_counts[group] += 1
        sites.add(row["site"])
        value = (row.get("thalach") or "").strip()
        if value:
            groups[group].append(float(value))
        else:
            missing[group] += 1

    means = {g: statistics.mean(values) for g, values in groups.items()}
    variances = {g: statistics.variance(values) for g, values in groups.items()}
    standard_deviations = {g: math.sqrt(variances[g]) for g in groups}
    standard_error = math.sqrt(
        variances[1] / len(groups[1]) + variances[0] / len(groups[0])
    )
    difference = means[1] - means[0]
    t_value = difference / standard_error
    df = (
        (variances[1] / len(groups[1]) + variances[0] / len(groups[0])) ** 2
        / (
            (variances[1] / len(groups[1])) ** 2 / (len(groups[1]) - 1)
            + (variances[0] / len(groups[0])) ** 2 / (len(groups[0]) - 1)
        )
    )
    critical = t_critical_975(df)
    p_value = t_two_sided_p(t_value, df)
    ci = (
        difference - critical * standard_error,
        difference + critical * standard_error,
    )

    assert len(rows) == 920
    assert raw_counts == {0: 411, 1: 509}
    assert len(sites) == 4
    assert len(groups[0]) == 391 and len(groups[1]) == 474
    assert missing == {0: 20, 1: 35}
    assert abs(means[1] - 128.2616) < 0.001
    assert abs(means[0] - 148.8005) < 0.001
    assert abs(t_value + 12.633) < 0.001
    assert abs(p_value - 1.30e-33) / 1.30e-33 < 0.03
    assert abs(ci[0] + 23.73) < 0.01 and abs(ci[1] + 17.35) < 0.01

    return {
        "rows": rows,
        "groups": groups,
        "raw_counts": raw_counts,
        "missing": missing,
        "means": means,
        "sds": standard_deviations,
        "difference": difference,
        "t": t_value,
        "df": df,
        "p": p_value,
        "ci": ci,
    }


def east_asia_font(run, name="Microsoft YaHei"):
    run.font.name = name
    r_pr = run._element.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), name)


def style_run(run, size=10.5, color="202020", bold=False, italic=False):
    east_asia_font(run)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.bold = bold
    run.italic = italic


def paragraph_after(anchor, label, body, document, size=10.5):
    node = OxmlElement("w:p")
    anchor._p.addnext(node)
    paragraph = Paragraph(node, anchor._parent)
    paragraph.style = document.styles["Normal"]
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.12
    paragraph.paragraph_format.keep_together = True
    head = paragraph.add_run(label)
    style_run(head, size=size, color="1F4E79", bold=True)
    text = paragraph.add_run(body)
    style_run(text, size=size)
    return paragraph


def patch_only_document_xml(original_path, modified_path, output_path):
    target = "word/document.xml"
    with zipfile.ZipFile(original_path) as original, zipfile.ZipFile(modified_path) as modified:
        modified_xml = modified.read(target)
        with zipfile.ZipFile(output_path, "w") as result:
            for info in original.infolist():
                data = modified_xml if info.filename == target else original.read(info.filename)
                result.writestr(info, data)

    with zipfile.ZipFile(original_path) as original, zipfile.ZipFile(output_path) as result:
        for name in original.namelist():
            if name != target:
                assert original.read(name) == result.read(name), f"Unexpected package change: {name}"


def make_completed_worksheet(stats):
    document = Document(SOURCE_WORKSHEET)
    paragraphs = {p.text.strip(): p for p in document.paragraphs if p.text.strip()}
    prompt_1 = paragraphs[
        "填写：研究对象、比较对象、主要终点、分析单位、H0、H1、α，以及双侧或单侧的理由。"
    ]
    prompt_2 = paragraphs[
        "各写一个 Type I 和 Type II error。说明样本量、真实效应和测量噪声如何影响 power，并解释未拒绝 H0 的边界。"
    ]
    prompt_4 = paragraphs[
        "预先指定 thalach 为主要终点。报告每组 n、描述统计、效应量、95% CI、p-value、H0 决策，并说明统计关联、实际意义和因果边界。"
    ]
    prompt_5 = paragraphs[
        "指出至少两个混杂因素，说明控制或敏感性分析；解释只挑显著指标的风险。用“主张—证据—推理—边界”写四句话。"
    ]

    paragraph_after(
        prompt_1,
        "作答｜",
        "研究对象为 disease 与 thalach 均完整的四中心受试者；比较 disease=1（num>0）与 disease=0（num=0）。主要终点为 thalach（最大运动心率，次/分），分析单位为每位受试者。定义 Δ=μ1−μ0，H0：Δ=0；H1：Δ≠0；α=0.05。题目问是否有关或不同，没有预先指定方向，因此采用双侧检验。",
        document,
    )
    paragraph_after(
        prompt_2,
        "作答｜",
        "Type I error：真实均值无差异，却拒绝 H0，误报有关联。Type II error：真实均值不同，却未拒绝 H0，漏掉差异。其他条件相同时，样本量增加或真实效应变大通常提高 power；测量噪声增大通常降低 power。未拒绝 H0 只表示证据不足，不等于证明两组相同；应结合置信区间判断数据仍允许多大的差异。",
        document,
    )

    table = document.tables[0]
    method_rows = [
        [
            "Welch 独立样本 t 检验",
            "连续结局；两组独立；检查分布与离群值；不要求方差相等。",
            "Mann–Whitney U（秩/分布差异）",
        ],
        [
            "配对 t 检验",
            "同一受试者；分析个体内差值；差值近似正态。",
            "Wilcoxon 符号秩检验",
        ],
        [
            "χ² 独立性检验",
            "观察独立；期望频数足够。",
            "Fisher 精确检验（稀疏表）",
        ],
        [
            "Pearson 相关 / 线性回归",
            "近似线性；观察独立；检查异常点。",
            "Spearman 等级相关",
        ],
        [
            "Logistic 回归",
            "结局二分类；检查 logit 线性与共线性；分开训练/测试。",
            "LDA 或正则化 Logistic",
        ],
    ]
    for row_index, values in enumerate(method_rows, start=1):
        for column_index, value in enumerate(values, start=1):
            cell = table.cell(row_index, column_index)
            cell.text = value
            cell.vertical_alignment = 1
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(1)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    style_run(run, size=9.2, color="202020")

    p_sample = paragraph_after(
        prompt_4,
        "样本与缺失｜",
        "原始 N=920（疾病证据组 509、无疾病证据组 411）。thalach 缺失 55 人（两组分别 35、20 人）；完整病例 n=865（疾病证据组 474、无疾病证据组 391）。",
        document,
        size=10.2,
    )
    p_desc = paragraph_after(
        p_sample,
        "描述统计｜",
        "疾病证据组：均值 128.26、SD 24.02、中位数 128；无疾病证据组：均值 148.80、SD 23.61、中位数 151，单位均为次/分。",
        document,
        size=10.2,
    )
    p_test = paragraph_after(
        p_desc,
        "检验结果｜",
        "按 Δ=μ1−μ0，均值差为 −20.54 次/分；Welch t=−12.633，df≈837.18，95% CI [−23.73, −17.35]，双侧 p≈1.30×10⁻³³。α=0.05 时拒绝 H0。p 值是在 H0 与检验假设成立时，观察到当前或更极端结果的概率，不是 H0 为真的概率。",
        document,
        size=10.2,
    )
    paragraph_after(
        p_test,
        "解释与边界｜",
        "区间未包含 0，支持两组总体均值不同，样本均值相差约 20.5 次/分。但资料未给出临床重要差异阈值，不能只凭小 p 值认定临床意义。该数据是未调整的观察性跨中心资料，只支持关联，不证明疾病导致心率改变、治疗有效或具备个体诊断能力。",
        document,
        size=10.2,
    )

    p_conf = paragraph_after(
        prompt_5,
        "混杂与稳健性｜",
        "年龄和性别可能同时关联 disease 构成与 thalach；site 可能反映转诊人群、病例构成或测量流程差异。后续可比较未调整结果与 thalach ~ disease + age + sex + C(site) 的调整结果，并按中心检查方向；报告 thalach 缺失比例及有无缺失者的可观测特征。用药、运动能力未提供，不能声称已控制。",
        document,
        size=10.2,
    )
    p_multi = paragraph_after(
        p_conf,
        "多重比较｜",
        "若 20 个零假设都为真且 20 次检验相互独立，至少一次假阳性的概率为 1−0.95²⁰≈64.2%；指标相关时该数值不再精确。应预先指定 thalach 为主要终点，其余标为探索性；确认性多重检验可用 Holm/Bonferroni，探索性分析可控制 FDR。",
        document,
        size=10.2,
    )
    p_claim = paragraph_after(
        p_multi,
        "主张｜",
        "在 thalach 完整的四中心样本中，疾病证据组平均最大运动心率低于无疾病证据组。",
        document,
        size=10.2,
    )
    p_evidence = paragraph_after(
        p_claim,
        "证据｜",
        "两组 n=474 和 n=391；均值差 −20.54 次/分，95% CI [−23.73, −17.35]，p≈1.30×10⁻³³。",
        document,
        size=10.2,
    )
    p_reasoning = paragraph_after(
        p_evidence,
        "推理｜",
        "置信区间不包含 0，且差值为负；在 α=0.05 下拒绝两组均值相同的 H0。",
        document,
        size=10.2,
    )
    paragraph_after(
        p_reasoning,
        "边界｜",
        "这是未调整的观察性完整病例比较；混杂、缺失与中心差异仍可能影响估计，不能据此推断因果或个体诊断能力。",
        document,
        size=10.2,
    )

    with tempfile.TemporaryDirectory(dir=ROOT) as temp_dir:
        modified = Path(temp_dir) / "modified.docx"
        document.save(modified)
        patch_only_document_xml(SOURCE_WORKSHEET, modified, WORKSHEET_OUT)


def set_document_styles(document):
    normal = document.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(32, 32, 32)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.12

    title_style = document.styles["Title"]
    title_style.font.name = "Microsoft YaHei"
    title_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    title_style.font.size = Pt(21)
    title_style.font.bold = True
    title_style.font.color.rgb = RGBColor(0, 0, 0)
    title_style.font.underline = False
    title_ppr = title_style._element.get_or_add_pPr()
    for border in list(title_ppr.findall(qn("w:pBdr"))):
        title_ppr.remove(border)

    for style_name, size in (("Heading 1", 14), ("Heading 2", 12)):
        style = document.styles[style_name]
        style.font.name = "Microsoft YaHei"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(31, 78, 121)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(5)
        style.paragraph_format.keep_with_next = True


def add_labeled_paragraph(document, label, text, size=10.5):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.12
    first = paragraph.add_run(label)
    style_run(first, size=size, color="1F4E79", bold=True)
    rest = paragraph.add_run(text)
    style_run(rest, size=size)
    return paragraph


def make_speaker_script():
    with CONTENT_PATH.open(encoding="utf-8") as stream:
        content = json.load(stream)

    document = Document()
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.82)
    section.right_margin = Inches(0.82)
    set_document_styles(document)

    document.core_properties.title = content["title"]
    document.core_properties.subject = content["subtitle"]

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title_ppr = title._p.get_or_add_pPr()
    for border in list(title_ppr.findall(qn("w:pBdr"))):
        title_ppr.remove(border)
    title_run = title.add_run(content["title"])
    style_run(title_run, size=21, color="000000", bold=True)
    add_labeled_paragraph(document, "课程主题｜", content["subtitle"])
    add_labeled_paragraph(
        document,
        "配套文件｜",
        "统计检验与假设设计演讲PPT（11 页，含演讲者备注）",
    )
    add_labeled_paragraph(
        document,
        "使用方式｜",
        "按页讲述即可；每页先看本页作用，再用讲稿说明概念、数据结果和结论边界。讲稿不是工作纸的逐格朗读。",
    )

    document.add_heading("一 核心数据与最终判断", level=1)
    add_labeled_paragraph(
        document,
        "核心结果｜",
        "在 thalach 完整的 865 名受试者中，疾病证据组（n=474）的平均最大运动心率为 128.26 次/分，无疾病证据组（n=391）为 148.80 次/分；均值差为 −20.54 次/分，95% CI [−23.73, −17.35]，双侧 Welch 检验 p≈1.30×10⁻³³。",
    )
    add_labeled_paragraph(
        document,
        "最终判断｜",
        "数据支持两组总体均值存在统计差异，但这是未调整的观察性完整病例比较。它不能确立因果，也不能仅凭 p 值断言临床重要性或个体诊断能力。",
    )

    document.add_heading("二 逐页讲稿", level=1)
    purposes = [
        "说明汇报主题和分析主线。",
        "明确研究对象、比较组、主要终点和分析单位。",
        "给出可检验的 H0/H1、显著性水平与双侧检验理由。",
        "解释两类错误、power 及未拒绝 H0 的含义。",
        "根据变量类型和研究设计选择检验。",
        "区分连续变量关联和二分类预测问题。",
        "说明原始样本、缺失处理和实际分析人数。",
        "报告主要效应估计、置信区间、p 值与 H0 决策。",
        "区分统计显著性、临床重要性和因果结论。",
        "说明混杂、多重比较和后续稳健性检查。",
        "用主张、证据、推理和边界作结。"
    ]
    for index, slide in enumerate(content["slides"], start=1):
        document.add_heading(f"第 {index:02d} 页  {slide['title']}", level=2)
        purpose = add_labeled_paragraph(document, "本页作用｜", purposes[index - 1])
        purpose.paragraph_format.keep_with_next = True
        notes = add_labeled_paragraph(document, "讲稿｜", slide["notes"])
        notes.paragraph_format.keep_together = True

    document.add_heading("三 课堂追问准备", level=1)
    for question, answer in content["questions"]:
        add_labeled_paragraph(document, question + " ", answer)

    document.add_heading("四 结束语", level=1)
    closing = document.add_paragraph()
    closing.paragraph_format.space_after = Pt(8)
    run = closing.add_run(content["closing"])
    style_run(run, size=11, color="202020")

    document.add_heading("五 AI 使用与核对说明", level=1)
    add_labeled_paragraph(document, "说明｜", content["aiDisclosure"])

    document.add_heading("六 视觉素材说明", level=1)
    add_labeled_paragraph(document, "来源｜", content["visualCredit"])

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if SCRIPT_OUT.exists():
        raise FileExistsError(f"Refusing to overwrite {SCRIPT_OUT}")
    document.save(SCRIPT_OUT)


def main():
    if not SOURCE_WORKSHEET.exists() or not DATA_ZIP.exists():
        raise FileNotFoundError("A supplied source file is missing.")
    if WORKSHEET_OUT.exists():
        raise FileExistsError(f"Refusing to overwrite {WORKSHEET_OUT}")
    stats = read_primary_data()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    make_completed_worksheet(stats)
    make_speaker_script()
    print(f"Completed worksheet: {WORKSHEET_OUT}")
    print(f"Speaker script: {SCRIPT_OUT}")
    print(
        f"Welch t={stats['t']:.6f}, df={stats['df']:.4f}, "
        f"p={stats['p']:.6g}, 95% CI=({stats['ci'][0]:.4f}, {stats['ci'][1]:.4f})"
    )


if __name__ == "__main__":
    main()
