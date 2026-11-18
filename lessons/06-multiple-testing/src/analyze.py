"""Portable archive check for 06-multiple-testing; run from any working directory."""
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
if __name__ == "__main__":
    sys.argv = [str(ROOT / "tools/run_lesson.py"), "06-multiple-testing"]
    runpy.run_path(str(ROOT / "tools/run_lesson.py"), run_name="__main__")
