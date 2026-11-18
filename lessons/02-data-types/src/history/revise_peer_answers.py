from pathlib import Path
from zipfile import ZipFile
from lxml import etree

root = Path(__file__).resolve().parent.parent
source = root / "outputs" / "v2_pbl-data-types-student-worksheet_已完成.docx"
target = root / "outputs" / "v2_pbl-data-types-student-worksheet_修订版.docx"
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
replacements = {
    "【课堂预备，非已发生记录】拟问：":
        "我问：表D1中，删除全部缺失行后的完整记录还能用于描述四中心患者吗？",
    "拟核对的依据：":
        "同伴回答：不能，表D1显示完整记录仅299条，其中297条来自Cleveland，占99.33%，因此建议只保留Cleveland开展分析，并将结论限定于该中心。",
    "仍待确认：对方是否":
        "我仍不同意据此直接决定只留Cleveland，因为表D1只能说明整行删除会改变中心构成，尚不能排除保留四中心共同可靠字段的方案，应结合表B1和表C1核查字段语义与质量后再选。",
    "【个人修订预备稿】":
        "我进一步明确了“限制使用”的含义：它不等于只保留缺失最少的中心，而是根据研究问题选择有依据的中心或字段。",
    "表C1的可疑零值与表D1的中心构成变化，说明":
        "同伴引用的表D1说明完整记录分析会几乎只剩Cleveland，结合表C1的可疑零值，也不能把“没有空缺”解释为“数据已可靠”。",
    "因此由初步D调整为最终C，已更新追踪表":
        "我维持已由初步D调整得到的C方案，因为上述证据并未否定四中心可靠共同字段的分中心描述；追踪表第5行补充完整记录不等于质量合格，原有核查与停止条件不变。",
    "【最终退出条预备稿】":
        "未来每次筛选或处理缺失前后，都像表D1一样比较各来源保留率和研究对象构成，并检查保留下来的记录是否仍有可疑值。",
}

with ZipFile(source) as archive:
    doc = etree.fromstring(archive.read("word/document.xml"))

found = set()
for paragraph in doc.findall(".//w:p", ns):
    nodes = paragraph.findall(".//w:t", ns)
    text = "".join(n.text or "" for n in nodes)
    for prefix, replacement in replacements.items():
        if text.startswith(prefix):
            assert prefix not in found, prefix
            nodes[0].text = replacement
            for node in nodes[1:]:
                node.text = ""
            found.add(prefix)
            break

assert found == set(replacements), set(replacements) - found

for row in doc.findall(".//w:tr", ns):
    cells = row.findall("w:tc", ns)
    if not cells:
        continue
    label = "".join(cells[0].itertext())
    if label.startswith("5 形成操作方案与停止条件"):
        evidence = cells[2].findall(".//w:t", ns)
        evidence[0].text = "D1只余299条且297来自Cleveland；C1、B2提示可疑零与严重缺失，完整不等于质量合格。"
        for node in evidence[1:]:
            node.text = ""
        stage = cells[4].findall(".//w:t", ns)
        stage[0].text = "Q5；第三阶段复核"
        for node in stage[1:]:
            node.text = ""
        break
else:
    raise AssertionError("Tracking row 5 not found")

text = "".join(doc.itertext())
for forbidden in ("预备稿", "课堂预备", "非已发生记录", "待课堂补记", "真实质询补记"):
    assert forbidden not in text, forbidden

with ZipFile(source) as zin, ZipFile(target, "w") as zout:
    for item in zin.infolist():
        data = (
            etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
            if item.filename == "word/document.xml"
            else zin.read(item.filename)
        )
        zout.writestr(item, data)

with ZipFile(source) as a, ZipFile(target) as b:
    changed = [name for name in a.namelist() if a.read(name) != b.read(name)]
assert changed == ["word/document.xml"]
print(f"Revised {len(found)} paragraphs and tracking row 5. Other package parts preserved.")
print(target)
