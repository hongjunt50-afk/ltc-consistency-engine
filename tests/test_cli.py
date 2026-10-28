"""命令行端到端冒烟：三条子命令都要能跑通并给出正确退出码。

存在理由：单元测试只调函数，CLI 的参数拼装曾经漏过（score_quality 少写关键字参数），
函数级测试全绿但命令行报 TypeError，所以单独盯一层。
"""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import ltc_rules as R  # noqa: E402


def run(*args):
    return subprocess.run([sys.executable, "ltc_rules.py", *map(str, args)],
                          cwd=ROOT, capture_output=True, text=True, encoding="utf-8")


class TestCli(unittest.TestCase):
    def test_audit_clean_exits_zero(self):
        p = run("audit", "tests/fixtures/clean_case.json")
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertEqual(json.loads(p.stdout)["findings"], [])

    def test_audit_dirty_exits_one(self):
        p = run("audit", "tests/fixtures/sample_case.json")
        self.assertEqual(p.returncode, 1)
        self.assertGreater(json.loads(p.stdout)["summary"]["hard_violations"], 0)

    def _tmp(self, payload):
        f = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False,
                                        encoding="utf-8")
        json.dump(payload, f, ensure_ascii=False)
        f.close()
        return Path(f.name)

    def test_inspection_full_marks(self):
        scores = {s["section"]: [i["score"] for i in s["items"]]
                  for s in R.INSPECTION["sections"]}
        p = run("inspection", self._tmp(scores))
        self.assertEqual(p.returncode, 0, p.stderr)
        out = json.loads(p.stdout)
        self.assertEqual(out["total_score"], 100)

    def test_inspection_invalid_exits_one(self):
        p = run("inspection", self._tmp({"机构管理": [27]}))
        self.assertEqual(p.returncode, 1)

    def test_quality_pass_and_fail(self):
        ok = {"scores": {"A": [5] * 20, "B": [3] * 16, "C": [3] * 5}}
        p = run("quality", self._tmp(ok))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertEqual(json.loads(p.stdout)["verdict"], "合格")

        bad = {"scores": {"A": [5] * 20, "B": [2] * 16, "C": [5] * 5}}
        p = run("quality", self._tmp(bad))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertEqual(json.loads(p.stdout)["verdict"], "不合格")

    def test_quality_hard_fail(self):
        payload = {"scores": {"A": [5] * 20, "B": [5] * 16, "C": [5] * 5},
                   "hard_fail_hits": ["出现打骂参保人行为"]}
        p = run("quality", self._tmp(payload))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertEqual(json.loads(p.stdout)["verdict"], "不合格")

    def test_unknown_command(self):
        p = run("bogus", "tests/fixtures/clean_case.json")
        self.assertEqual(p.returncode, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
