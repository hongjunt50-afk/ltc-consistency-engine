"""规则引擎自检：每条规则都要有命中用例和干净用例（防止规则永远不触发或乱触发）。"""
import copy
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import ltc_rules as R  # noqa: E402

FIXTURE = json.loads((ROOT / "tests" / "fixtures" / "sample_case.json").read_text(encoding="utf-8"))
CLEAN = json.loads((ROOT / "tests" / "fixtures" / "clean_case.json").read_text(encoding="utf-8"))


def clean_case() -> dict:
    """一条完全合规的最小用例：用于验证"不该报的别报"。"""
    return {
        "elders": [{
            "elder_id": "E1", "service_type": "居家护理",
            "address": "甲市乙区丙路 1 号",
            "plan": [{"name": "擦浴", "freq_per_month": 1}],
        }],
        "caregivers": [{"caregiver_id": "C1", "qualification": "长期照护师"}],
        "records": [{
            "record_no": "R1", "elder_id": "E1", "caregiver_id": "C1",
            "service_type": "居家护理", "service_address": "甲市乙区丙路 1 号",
            "service_date": "2026-10-05", "start_time": "08:00", "end_time": "09:00",
            "duration_min": 60, "items": ["擦浴"], "item_count": 1,
            "caregiver_sign": True, "elder_sign": True,
            "submit_time": "2026-10-05 20:00",
        }],
    }


def rules_of(findings):
    return [f.rule for f in findings]


class TestCatalog(unittest.TestCase):
    def test_counts_match_policy(self):
        """36 项 = 生活照护 20 + 医疗护理 16（医保发〔2025〕24号）。"""
        items = R.CATALOG["items"]
        self.assertEqual(len(items), 36)
        self.assertEqual(sum(1 for i in items if i["category"] == "生活照护类"), 20)
        self.assertEqual(sum(1 for i in items if i["category"] == "医疗护理类"), 16)

    def test_codes_and_seq_unique(self):
        items = R.CATALOG["items"]
        self.assertEqual(len({i["code"] for i in items}), 36)
        self.assertEqual(len({i["seq"] for i in items}), 36)

    def test_inspection_totals(self):
        """巡查表 8 分区、满分 100、分项累加与分区声明一致。"""
        self.assertEqual(R.INSPECTION["max_score"], 100)
        self.assertEqual(len(R.INSPECTION["sections"]), 8)
        self.assertEqual(sum(s["max_score"] for s in R.INSPECTION["sections"]), 100)
        for sec in R.INSPECTION["sections"]:
            self.assertEqual(sum(i["score"] for i in sec["items"]), sec["max_score"],
                             f"{sec['section']} 分项求和与声明不符")
        self.assertEqual(sum(len(s["items"]) for s in R.INSPECTION["sections"]), 30)

    def test_quality_thresholds(self):
        self.assertAlmostEqual(R.QUALITY["pass_rate"], 0.60)
        counts = {t["id"]: len(t["items"]) for t in R.QUALITY["tables"]}
        self.assertEqual(counts, {"A": 20, "B": 16, "C": 5})
        self.assertEqual(len(R.QUALITY["hard_fail"]), 2)


class TestConsistencyRules(unittest.TestCase):
    def test_clean_case_has_no_findings(self):
        self.assertEqual(R.check_consistency(clean_case()), [])

    def test_unknown_service(self):
        c = clean_case()
        c["records"][0]["items"] = ["洗脚"]
        c["records"][0]["item_count"] = 1
        self.assertIn("E_UNKNOWN_SERVICE", rules_of(R.check_consistency(c)))

    def test_bad_time_and_duration(self):
        c = clean_case()
        c["records"][0]["end_time"] = "07:30"
        self.assertIn("E_BAD_TIME", rules_of(R.check_consistency(c)))

        c = clean_case()
        c["records"][0]["duration_min"] = 45  # 起止推算是 60
        self.assertIn("E_DURATION_MISMATCH", rules_of(R.check_consistency(c)))

    def test_item_count(self):
        c = clean_case()
        c["records"][0]["items"] = ["擦浴", "洗发"]
        c["records"][0]["item_count"] = 5
        self.assertIn("I_ITEM_COUNT_MISMATCH", rules_of(R.check_consistency(c)))

    def test_non_certified_qualification(self):
        """非持证人员执行医疗护理类项目 = 硬违规（质量评价表的直接不合格情形）。"""
        c = clean_case()
        c["caregivers"][0]["qualification"] = "其他"
        c["records"][0]["items"] = ["吸痰"]
        c["elders"][0]["plan"] = [{"name": "吸痰"}]
        self.assertIn("E_QUAL_MISMATCH", rules_of(R.check_consistency(c)))

    def test_unfilled_qualification_is_hard_violation(self):
        c = clean_case()
        c["caregivers"][0]["qualification"] = ""
        c["records"][0]["items"] = ["吸痰"]
        c["elders"][0]["plan"] = [{"name": "吸痰"}]
        self.assertIn("E_QUAL_MISMATCH", rules_of(R.check_consistency(c)))

    def test_scope_review_for_nursing_assistant(self):
        """养老护理员/医疗护理员做医疗护理类项目属执业范围问题，报 W 不报 E（各地口径不一）。"""
        c = clean_case()
        c["caregivers"][0]["qualification"] = "养老护理员"
        c["records"][0]["items"] = ["吸痰"]
        c["elders"][0]["plan"] = [{"name": "吸痰"}]
        got = rules_of(R.check_consistency(c))
        self.assertIn("W_QUAL_SCOPE_REVIEW", got)
        self.assertNotIn("E_QUAL_MISMATCH", got)

    def test_nurse_can_do_medical(self):
        c = clean_case()
        c["caregivers"][0]["qualification"] = "护士"
        c["records"][0]["items"] = ["吸痰"]
        c["elders"][0]["plan"] = [{"name": "吸痰"}]
        got = rules_of(R.check_consistency(c))
        self.assertNotIn("E_QUAL_MISMATCH", got)
        self.assertNotIn("W_QUAL_SCOPE_REVIEW", got)

    def test_missing_signatures(self):
        c = clean_case()
        c["records"][0]["caregiver_sign"] = False
        c["records"][0]["elder_sign"] = False
        self.assertEqual(rules_of(R.check_consistency(c)).count("E_NO_SIGN"), 2)

    def test_abnormal_without_remedy(self):
        c = clean_case()
        c["records"][0]["abnormal_note"] = "诉头晕"
        self.assertEqual(rules_of(R.check_consistency(c)).count("E_ABNORMAL_NO_REMEDY"), 2)

        c["records"][0]["remedy"] = "通知家属并观察"
        c["records"][0]["remedy_time"] = "2026-10-05 09:10"
        self.assertNotIn("E_ABNORMAL_NO_REMEDY", rules_of(R.check_consistency(c)))

    def test_off_plan(self):
        c = clean_case()
        c["records"][0]["items"] = ["洗发"]  # 计划里只有擦浴
        c["records"][0]["item_count"] = 1
        self.assertIn("W_OFF_PLAN", rules_of(R.check_consistency(c)))

    def test_service_type_and_address(self):
        c = clean_case()
        c["records"][0]["service_type"] = "机构护理"
        self.assertIn("W_SERVICE_TYPE_MISMATCH", rules_of(R.check_consistency(c)))

        c = clean_case()
        c["records"][0]["service_address"] = "别的地方"
        self.assertIn("W_ADDRESS_MISMATCH", rules_of(R.check_consistency(c)))

    def test_late_entry(self):
        c = clean_case()
        c["records"][0]["submit_time"] = "2026-10-20 09:00"
        self.assertIn("W_LATE_ENTRY", rules_of(R.check_consistency(c)))

    def test_overlap_same_caregiver_and_elder(self):
        """同一护理员同一时段两条记录 = 代打卡/虚记服务的典型形态。"""
        c = clean_case()
        second = copy.deepcopy(c["records"][0])
        second["record_no"] = "R2"
        second["start_time"] = "08:30"
        second["end_time"] = "09:30"
        c["records"].append(second)
        got = rules_of(R.check_consistency(c))
        self.assertIn("E_CAREGIVER_OVERLAP", got)
        self.assertIn("E_ELDER_OVERLAP", got)

    def test_overlap_touching_boundary_ok(self):
        """首尾相接（09:00 结束、09:00 开始）不算重叠。"""
        c = clean_case()
        second = copy.deepcopy(c["records"][0])
        second.update(record_no="R2", start_time="09:00", end_time="10:00")
        c["records"].append(second)
        got = rules_of(R.check_consistency(c))
        self.assertNotIn("E_CAREGIVER_OVERLAP", got)
        self.assertNotIn("E_ELDER_OVERLAP", got)

    def test_excess_daily_threshold_is_configurable(self):
        c = clean_case()
        c["records"][0]["end_time"] = "18:00"  # 600 分钟
        c["records"][0]["duration_min"] = 600
        self.assertIn("W_EXCESS_DAILY", rules_of(R.check_consistency(c)))
        self.assertNotIn("W_EXCESS_DAILY",
                         rules_of(R.check_consistency(c, day_max_minutes=1000)))

    def test_frequency_against_plan(self):
        c = clean_case()
        c["elders"][0]["plan"] = [{"name": "擦浴", "freq_per_month": 8}]
        self.assertEqual(rules_of(R.check_consistency(c)).count("W_FREQ_SHORT"), 1)

        c["elders"][0]["plan"] = [{"name": "擦浴", "freq_per_month": 1}]
        self.assertNotIn("W_FREQ_SHORT", rules_of(R.check_consistency(c)))

        c["elders"][0]["plan"] = [{"name": "擦浴", "freq_per_month": 1}]
        c["records"].append({**copy.deepcopy(c["records"][0]), "record_no": "R2",
                             "service_date": "2026-10-06"})
        self.assertIn("W_FREQ_OVER", rules_of(R.check_consistency(c)))

    def test_fixture_end_to_end(self):
        result = R.audit(FIXTURE)
        self.assertEqual(result["records_checked"], 5)
        got = {f["rule"] for f in result["findings"]}
        for expect in ("E_QUAL_MISMATCH", "E_UNKNOWN_SERVICE", "E_BAD_TIME",
                       "E_ABNORMAL_NO_REMEDY", "E_CAREGIVER_OVERLAP", "W_LATE_ENTRY",
                       "I_ITEM_COUNT_MISMATCH"):
            self.assertIn(expect, got)

    def test_every_finding_has_clause(self):
        """每条结论都要能指向具体文书条款，否则无法向机构/监管解释。"""
        for f in R.audit(FIXTURE)["findings"]:
            self.assertTrue(f["clause"].strip(), f["rule"])
            self.assertIn("长期护理", f["clause"] + "长期护理")

    def test_clean_fixture_zero_findings(self):
        """合规对照组：三条记录全部合规，必须零命中（假阳性闸门）。"""
        res = R.audit(CLEAN)
        self.assertEqual(res["findings"], [], [f["rule"] for f in res["findings"]])
        self.assertEqual(res["summary"],
                         {"hard_violations": 0, "needs_review": 0, "data_quality": 0})


class TestInspectionScoring(unittest.TestCase):
    def full_scores(self):
        return {s["section"]: [i["score"] for i in s["items"]]
                for s in R.INSPECTION["sections"]}

    def test_full_marks_is_100(self):
        res = R.score_inspection(self.full_scores())
        self.assertTrue(res["valid"], res["errors"])
        self.assertEqual(res["total_score"], 100)
        self.assertEqual(res["applicable_full"], 100)

    def test_zero_is_zero(self):
        scores = {k: [0] * len(v) for k, v in self.full_scores().items()}
        res = R.score_inspection(scores)
        self.assertEqual(res["total_score"], 0)

    def test_out_of_range_rejected(self):
        scores = self.full_scores()
        scores["机构管理"][0] = 99
        res = R.score_inspection(scores)
        self.assertFalse(res["valid"])
        self.assertTrue(any("超出" in e for e in res["errors"]))

    def test_wrong_item_count_rejected(self):
        scores = self.full_scores()
        scores["档案管理"] = [3]
        res = R.score_inspection(scores)
        self.assertFalse(res["valid"])

    def test_not_applicable_excluded_from_denominator(self):
        scores = self.full_scores()
        scores["综合评价"] = [None]  # 该项为加分性质，可判不适用
        res = R.score_inspection(scores)
        self.assertTrue(res["valid"])
        self.assertEqual(res["total_score"], 90)
        self.assertEqual(res["applicable_full"], 90)


class TestQualityScoring(unittest.TestCase):
    def scores(self, a=5, b=5, c=5):
        return {tid: [v] * n for tid, n, v in
                (("A", 20, a), ("B", 16, b), ("C", 5, c))}

    def test_all_full_is_pass(self):
        res = R.score_quality(self.scores())
        self.assertEqual(res["verdict"], "合格")
        self.assertTrue(all(t["pass"] for t in res["tables"]))

    def test_exactly_60_percent_passes(self):
        """60% 是合格线，恰好 60% 应判合格（>= 而非 >）。"""
        res = R.score_quality(self.scores(a=3, b=3, c=3))  # 3/5 = 60%
        self.assertEqual(res["verdict"], "合格")
        for t in res["tables"]:
            self.assertAlmostEqual(t["rate"], 0.60)

    def test_below_60_fails(self):
        res = R.score_quality(self.scores(a=3, b=2, c=3))  # B 类 40%
        self.assertEqual(res["verdict"], "不合格")
        b = next(t for t in res["tables"] if t["id"] == "B")
        self.assertFalse(b["pass"])

    def test_hard_fail_override(self):
        """即使三类满分，命中「打骂参保人」也直接判不合格。"""
        res = R.score_quality(self.scores(), hard_fail_hits=["出现打骂参保人行为"])
        self.assertEqual(res["verdict"], "不合格")
        self.assertEqual(len(res["hard_fail_hits"]), 1)

    def test_out_of_range_rejected(self):
        s = self.scores()
        s["A"][0] = 9
        res = R.score_quality(s)
        self.assertFalse(res["valid"])

    def test_wrong_count_rejected(self):
        s = self.scores()
        s["B"] = [5] * 10
        res = R.score_quality(s)
        self.assertFalse(res["valid"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
