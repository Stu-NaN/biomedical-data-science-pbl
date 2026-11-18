"""Validate course structure, imported hashes, document integrity, and secret exclusions."""

from __future__ import annotations

import ast
import csv
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
SKIP = {".git", "__pycache__", ".venv", "env", "node_modules", ".local"}
TEXT = {".py", ".mjs", ".ps1", ".cmd", ".json", ".jsonl", ".md", ".txt", ".yml", ".yaml"}
SECRET = re.compile(r"\b(?:gh[pousr]_[A-Za-z0-9]{25,}|github_pat_[A-Za-z0-9_]{25,}|sk-[A-Za-z0-9_-]{30,})\b")


def files():
    return sorted(p for p in ROOT.rglob("*") if p.is_file() and not set(p.relative_to(ROOT).parts) & SKIP)


def main():
    errors, warnings = [], []
    paths = files()
    imported = json.loads((ROOT / "provenance/imported-files.json").read_text(encoding="utf-8"))
    for record in imported:
        path = ROOT / record["path"]
        if not path.is_file():
            errors.append(f"Missing imported file: {record['path']}")
        elif hashlib.sha256(path.read_bytes()).hexdigest() != record["sha256"]:
            errors.append(f"Changed imported file: {record['path']}")
    supplied = json.loads((ROOT / "provenance/supplied-files.json").read_text(encoding="utf-8"))
    formal = [r for r in supplied if r.get("role") == "user-confirmed-homework"]
    for record in formal:
        matches = [p for p in paths if p.name == record["name"] and p.parent.name == "homework"]
        if len(matches) != 1 or hashlib.sha256(matches[0].read_bytes()).hexdigest() != record["sha256"]:
            errors.append(f"Formal homework not preserved: {record['name']}")
    if len(formal) != 8:
        errors.append(f"Expected eight user-confirmed homework files, found {len(formal)}")
    with (ROOT / "provenance/file-index.csv").open(encoding="utf-8-sig", newline="") as stream:
        indexed = list(csv.DictReader(stream))
    excluded = {"provenance/file-index.csv", "provenance/validation.json"}
    expected = {p.relative_to(ROOT).as_posix() for p in paths} - excluded
    indexed_paths = {record["path"] for record in indexed}
    if len(indexed_paths) != len(indexed):
        errors.append("Duplicate paths in file-index.csv")
    for missing in sorted(expected - indexed_paths):
        errors.append(f"File missing from index: {missing}")
    for stale in sorted(indexed_paths - expected):
        errors.append(f"Index references a missing file: {stale}")
    for record in indexed:
        path = ROOT / record["path"]
        if path.is_file() and (
            hashlib.sha256(path.read_bytes()).hexdigest() != record["sha256"]
            or path.stat().st_size != int(record["bytes"])
        ):
            errors.append(f"Index checksum mismatch: {record['path']}")
        for field in ["related_code", "related_history", "related_ai_dialogue"]:
            if not (ROOT / record[field]).is_file():
                errors.append(f"Broken {field} association: {record['path']}")
    docs, py = 0, 0
    for path in paths:
        rel = path.relative_to(ROOT).as_posix()
        if path.stat().st_size >= 100 * 1024**2:
            errors.append(f"File exceeds GitHub regular-file limit: {rel}")
        if path.name in {"auth.json", "hosts.yml", ".env"}:
            errors.append(f"Forbidden credentials filename: {rel}")
        if "teacher" in path.parts and "data" in rel:
            errors.append(f"Teacher-only data included: {rel}")
        if path.suffix in TEXT:
            text = path.read_text(encoding="utf-8-sig")
            if SECRET.search(text):
                errors.append(f"Potential credential in: {rel}")
            if path.suffix == ".py":
                try:
                    ast.parse(text, filename=rel)
                    py += 1
                except SyntaxError as exc:
                    errors.append(f"Python syntax error: {rel}: {exc}")
        if path.suffix in {".pptx", ".docx"}:
            try:
                with zipfile.ZipFile(path) as archive:
                    if archive.testzip():
                        raise ValueError("CRC mismatch")
                    main_part = "word/document.xml" if path.suffix == ".docx" else "ppt/presentation.xml"
                    ElementTree.fromstring(archive.read(main_part))
                docs += 1
            except Exception as exc:
                errors.append(f"Invalid Office file: {rel}: {exc}")
    lessons = sorted(p for p in (ROOT / "lessons").iterdir() if p.is_dir())
    for base in lessons:
        for required in ["README.md", "reflection.md", "src/analyze.py", "data/README.md",
                         "ai-dialogues/conversation.md", "outputs/reproduced/results.json"]:
            if not (base / required).is_file():
                errors.append(f"Missing lesson component: {base.name}/{required}")
        result = json.loads((base / "outputs/reproduced/results.json").read_text(encoding="utf-8"))
        if result.get("locked_test_label_files_opened") is not False:
            errors.append(f"Missing locked-label access declaration: {base.name}")
        if result.get("held_out_evaluation_performed") is not False:
            errors.append(f"Unexpected held-out evaluation claim: {base.name}")
    report = {
        "status": "passed" if not errors else "failed", "lesson_count": len(lessons),
        "file_count": len(paths), "import_records_checked": len(imported), "formal_homework_checked": len(formal),
        "indexed_files_checked": len(indexed),
        "office_packages_checked": docs, "python_files_parsed": py,
        "total_bytes": sum(p.stat().st_size for p in paths), "errors": errors, "warnings": warnings,
        "scope": "Structural, checksum and syntax checks; not visual re-review of unchanged historical documents.",
    }
    destination = ROOT / "provenance/validation.json"
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
