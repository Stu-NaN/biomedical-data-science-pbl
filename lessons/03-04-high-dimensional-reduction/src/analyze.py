"""Portable archive check for 03-04-high-dimensional-reduction; run from any working directory."""
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
if __name__ == "__main__":
    sys.argv = [str(ROOT / "tools/run_lesson.py"), "03-04-high-dimensional-reduction"]
    runpy.run_path(str(ROOT / "tools/run_lesson.py"), run_name="__main__")
