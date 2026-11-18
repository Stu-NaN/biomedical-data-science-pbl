import hashlib
import importlib.util
import json
import unittest
from pathlib import Path

import numpy as np
from scipy.stats import false_discovery_control

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("run_lesson", ROOT / "tools/run_lesson.py")
RUNNER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(RUNNER)


class ArchiveTests(unittest.TestCase):
    def test_bh_matches_scipy(self):
        p = np.random.default_rng(19).uniform(size=100)
        np.testing.assert_allclose(RUNNER.bh(p), false_discovery_control(p, method="bh"))

    def test_bh_ties_and_boundaries(self):
        p = np.array([0, .01, .01, .8, 1.])
        np.testing.assert_allclose(RUNNER.bh(p), false_discovery_control(p, method="bh"))
        self.assertTrue(np.all(np.diff(RUNNER.bh(p)[np.argsort(p)]) >= 0))

    def test_supplied_homework_preserved(self):
        records = json.loads((ROOT / "provenance/supplied-files.json").read_text(encoding="utf-8"))
        formal = [r for r in records if r.get("role") == "user-confirmed-homework"]
        self.assertEqual(len(formal), 8)
        for entry in formal:
            paths = [p for p in (ROOT / "lessons").rglob(entry["name"]) if p.parent.name == "homework"]
            self.assertEqual(len(paths), 1)
            self.assertEqual(hashlib.sha256(paths[0].read_bytes()).hexdigest(), entry["sha256"])

    def test_checks_do_not_claim_heldout_evaluation(self):
        for lesson in RUNNER.LESSONS:
            result = json.loads((ROOT / "lessons" / lesson / "outputs/reproduced/results.json").read_text(encoding="utf-8"))
            self.assertFalse(result["held_out_evaluation_performed"])
            self.assertFalse(result["locked_test_label_files_opened"])


if __name__ == "__main__":
    unittest.main()
