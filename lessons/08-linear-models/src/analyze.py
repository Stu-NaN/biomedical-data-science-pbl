"""Portable archive check for 08-linear-models; run from any working directory."""
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
if __name__ == "__main__":
    sys.argv = [str(ROOT / "tools/run_lesson.py"), "08-linear-models"]
    runpy.run_path(str(ROOT / "tools/run_lesson.py"), run_name="__main__")
