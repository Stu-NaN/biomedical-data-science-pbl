from pathlib import Path
import hashlib
import json
import re
import sys
from zipfile import ZipFile
from lxml import etree
from docx import Document
from pypdf import PdfReader

root = Path(__file__).parent
output = Path(r"C:\Users\A\Desktop\第9讲作业_SVM与Margin_Claude协作修订版.docx")
render = root / sys.argv[1]
metadata = json.loads((root / "source_metadata.json").read_text(encoding="utf-8"))
results = json.loads((root / "verified_results.json").read_text(encoding="utf-8"))
assert hashlib.sha256(Path(metadata["source"]).read_bytes()).hexdigest() == metadata["sha256"]
assert (root / "claude_desktop_review.md").stat().st_size > 1000

doc = Document(output)
assert len(doc.tables) == 14
assert len(doc.inline_shapes) == 8
captions = [p.text for p in doc.paragraphs if p.style.name == "Caption"]
assert [int(re.match(r"图 (\d+)", c)[1]) for c in captions if c.startswith("图 ")] == list(range(1, 9))
assert [int(re.match(r"表 (\d+)", c)[1]) for c in captions if c.startswith("表 ")] == list(range(1, 15))
assert doc.tables[1].cell(1, 3).text == "V061（AB）"
assert doc.tables[1].cell(3, 3).text == "V217（NO）"

with ZipFile(output) as archive:
    xml = etree.fromstring(archive.read("word/document.xml"))
    namespace = {
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    }
    equation_count = len(xml.xpath("//m:oMath", namespaces=namespace))
    assert equation_count == 5
    assert len(xml.xpath("//wp:docPr[@descr]", namespaces=namespace)) == 8
    assert len([n for n in archive.namelist() if n.startswith("word/media/")]) == 8

assert results["geometry_diagnostics"]["cropped_training_count"] == 7
assert [r["n"] for r in results["score_bins_linear"]] == [9, 25, 4, 24]
assert [r["abnormal"] for r in results["score_bins_linear"]] == [4, 12, 3, 23]
assert results["small_gamma_checks"][0]["prediction_disagreements"] == 0
assert results["small_gamma_checks"][2]["prediction_disagreements"] == 2
assert len(results["newly_correct_dh"]) == 8
assert results["metrics"]["linear"]["TN"] == 18
assert results["metrics"]["linear"]["FN"] == 16
assert results["metrics"]["rbf"]["TN"] == 16
assert results["metrics"]["rbf"]["FN"] == 8

pdf = PdfReader(next(render.glob("*.pdf")))
assert len(list(render.glob("page-*.png"))) == len(pdf.pages)
page_texts = [page.extract_text() for page in pdf.pages]
heading_page = next(i for i, text in enumerate(page_texts) if "对 SVM 的综合理解" in text)
ending_page = next(i for i, text in enumerate(page_texts) if "这次能从数据里直接说明的" in text)
assert heading_page == ending_page
report = {
    "output": str(output),
    "original_unchanged": True,
    "claude_review_received": True,
    "pages": len(pdf.pages),
    "tables": len(doc.tables),
    "figures": len(doc.inline_shapes),
    "native_equations": equation_count,
    "all_captions_unique_and_ordered": True,
    "summary_within_one_page": True,
    "independent_experiment_checks": "passed",
}
(root / "final_checks.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(report, ensure_ascii=False, indent=2))
