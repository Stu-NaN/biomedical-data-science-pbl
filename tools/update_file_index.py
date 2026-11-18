"""Refresh navigation and checksums after intentional archive changes."""

from __future__ import annotations

import csv
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKIP = {".git", "__pycache__", ".venv", "env", "node_modules", ".local"}
EXCLUDED = {"provenance/file-index.csv", "provenance/validation.json"}
FIELDS = [
    "path", "lesson", "source", "category", "related_code", "related_history",
    "related_ai_dialogue", "association", "bytes", "sha256",
]


def archive_files():
    return sorted(
        path for path in ROOT.rglob("*")
        if path.is_file()
        and not set(path.relative_to(ROOT).parts) & SKIP
        and path.relative_to(ROOT).as_posix() not in EXCLUDED
    )


def main():
    records = json.loads((ROOT / "provenance/imported-files.json").read_text(encoding="utf-8"))
    imported = {record["path"]: record for record in records}
    rows = []
    for path in archive_files():
        relative = path.relative_to(ROOT)
        lesson = relative.parts[1] if relative.parts[0] == "lessons" else ""
        entry = imported.get(relative.as_posix(), {})
        rows.append({
            "path": relative.as_posix(),
            "lesson": lesson,
            "source": entry.get("source", "generated-during-archive-assembly"),
            "category": entry.get("category", "new-archive-file"),
            "related_code": f"lessons/{lesson}/src/analyze.py" if lesson else "tools/run_lesson.py",
            "related_history": f"lessons/{lesson}/src/history/README.md" if lesson else "provenance/assembly-notes.md",
            "related_ai_dialogue": f"lessons/{lesson}/ai-dialogues/conversation.md" if lesson else "provenance/assembly-notes.md",
            "association": "lesson-level navigation, not proof of exact file generation",
            "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        })
    with (ROOT / "provenance/file-index.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Indexed {len(rows)} files")


if __name__ == "__main__":
    main()
