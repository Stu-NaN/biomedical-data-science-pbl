import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
CONTENT_PATH = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "presenter-content-v2.json"
OUTPUT_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "deliverables" / "统计检验与假设设计_口语详稿_v2.docx"

INK = "202A30"
MUTED = "5C676D"
TEAL = "007E80"
FONT = "Microsoft YaHei"


def set_run(run, size=10.5, color=INK, bold=False, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.bold = bold
    r_pr = run._element.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), FONT)


def labeled_paragraph(doc, label, body, keep=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.keep_together = keep
    head = p.add_run(label)
    set_run(head, color=TEAL, bold=True)
    text = p.add_run(body)
    set_run(text)
    return p


def add_body_paragraphs(doc, text):
    for block in text.split("\n\n"):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_together = True
        run = p.add_run(block.strip())
        set_run(run, size=10.5)


def add_code(doc, text):
    for line in text.splitlines():
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.22)
        p.paragraph_format.right_indent = Inches(0.12)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line if line else " ")
        set_run(run, size=8.7, color="37474F", font="Consolas")


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, separate, placeholder, end])
    set_run(run, size=8, color=MUTED)


def build():
    data = json.loads(CONTENT_PATH.read_text(encoding="utf-8"))
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.58)
    section.bottom_margin = Inches(0.62)
    section.left_margin = Inches(0.72)
    section.right_margin = Inches(0.72)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.12

    for name, size in (("Title", 22), ("Heading 1", 15), ("Heading 2", 12)):
        style = doc.styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string("000000")
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(5)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(header.add_run("统计检验与假设设计  逐页口语详稿"), size=8, color=MUTED)
    add_page_number(section.footer.paragraphs[0])

    title = doc.add_paragraph(style="Title")
    title.paragraph_format.space_after = Pt(3)
    set_run(title.add_run("统计检验与假设设计口语详稿（v2）"), size=22, color="000000", bold=True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    set_run(p.add_run(data["subtitle"]), size=12, color=MUTED)

    doc.add_heading("怎么使用这份讲稿", level=1)
    labeled_paragraph(
        doc,
        "演讲顺序｜",
        "每页先用“口语稿”自然开场，再结合屏幕上的数字讲“本页要点”。停在“观众互动”处留几秒，让听众先回答；之后再用提示把答案接回来。代码页先读代码，再用人话解释输入、处理和输出。",
    )
    labeled_paragraph(
        doc,
        "分析边界｜",
        "本次实际完成了分组描述统计、thalach 缺失计数、未调整的 Welch 独立样本均值比较、均值差置信区间和结果图。没有运行年龄/性别/中心调整模型、疾病预测模型或因果分析。文中后续模型均明确标注为建议。",
    )
    labeled_paragraph(
        doc,
        "核心数据｜",
        "原始 920 人；thalach 缺失 55 人；完整病例 865 人。疾病证据组 n=474，均值 128.26 次/分；无疾病证据组 n=391，均值 148.80 次/分。均值差（疾病组减无疾病证据组）为 −20.54 次/分，95% CI [−23.73, −17.35]，双侧 Welch p≈1.30×10⁻³³。",
    )
    labeled_paragraph(
        doc,
        "代码说明｜",
        "分析数值由 Python 标准库计算。代码片段展示关键逻辑；p 值通过脚本中的 t 分布尾概率函数求出，该函数使用正则化不完全 Beta 函数计算。调整模型片段只是下一步建议，不是本次运行结果。",
    )
    doc.add_page_break()

    for slide in data["slides"]:
        doc.add_heading("第{0:02d}页  {1}".format(slide["index"], slide["title"]), level=1)
        labeled_paragraph(doc, "本页要讲清楚｜", slide["goal"])
        labeled_paragraph(doc, "一句话记住｜", slide["key"])

        doc.add_heading("可以直接这样讲", level=2)
        add_body_paragraphs(doc, slide["talk"])

        doc.add_heading("本页分析步骤", level=2)
        for item in slide["process"]:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Inches(0.22)
            p.paragraph_format.first_line_indent = Inches(-0.12)
            p.paragraph_format.space_after = Pt(2)
            set_run(p.add_run(item), size=9.8)

        doc.add_heading("代码或计算逻辑", level=2)
        add_code(doc, slide["code"])
        labeled_paragraph(doc, "代码翻译｜", slide["codeExplanation"])

        doc.add_heading("观众参与与衔接", level=2)
        labeled_paragraph(doc, "提问｜", slide["ask"])
        labeled_paragraph(doc, "接话提示｜", slide["guide"])
        labeled_paragraph(doc, "过渡｜", slide["transition"])
        labeled_paragraph(doc, "避免误讲｜", slide["pitfall"])

    doc.add_heading("常见追问参考", level=1)
    for item in data["questions"]:
        labeled_paragraph(doc, item["question"] + " ", item["answer"], keep=True)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    if OUTPUT_PATH.exists():
        raise FileExistsError("Refusing to overwrite: {}".format(OUTPUT_PATH))
    doc.core_properties.title = "统计检验与假设设计口语详稿（v2）"
    doc.core_properties.subject = data["subtitle"]
    doc.save(OUTPUT_PATH)
    print(OUTPUT_PATH)


if __name__ == "__main__":
    build()
