"""Excel 导入适配器测试：别名匹配、格式还原、关联键、错误报告、往返一致。"""
import json
import sys
import tempfile
import unittest
from datetime import datetime, time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import ltc_excel as X  # noqa: E402
import ltc_rules as R  # noqa: E402

try:
    from openpyxl import Workbook
except ImportError:  # pragma: no cover
    Workbook = None


def build(rows_by_sheet: dict[str, list[list]]):
    """按 {表名: 行列表} 造一个临时工作簿，返回路径（调用方负责清理）。"""
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in rows_by_sheet.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    f = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    f.close()
    wb.save(f.name)
    return Path(f.name)


RECORD_HEADER = ["服务编号", "参保人姓名", "护理人员姓名", "服务日期", "开始时间", "结束时间",
                 "时长（分钟）", "服务项目", "服务项目数合计", "护理人员签字",
                 "参保人员（家属）签名"]


class TestCellNormalisation(unittest.TestCase):
    def test_text_strips_and_drops_float_artifacts(self):
        self.assertEqual(X._text("  张三 "), "张三")
        self.assertEqual(X._text(None), "")
        self.assertEqual(X._text(110101199001011234.0), "110101199001011232")
        self.assertEqual(X._text(True), "是")

    def test_float_identity_precision_loss_is_reported(self):
        """Excel 把 18 位身份证当数字存会丢末两位，这是真风险，导入时要提示而不能装看不见。"""
        p = build({
            "参保人": [["参保人姓名", "身份证号"],
                    ["张三", 110101199001011234.0],
                    ["李四", "110101199001011234"]],
        })
        try:
            case, report = X.load_case(p)
        finally:
            p.unlink(missing_ok=True)
        ids = {e["name"]: e["id_card"] for e in case["elders"]}
        self.assertEqual(ids["李四"], "110101199001011234")   # 文本格式不丢
        self.assertNotEqual(ids["张三"], "110101199001011234")  # 数字格式已丢
        self.assertTrue(any("精度" in i["message"] for i in report.issues), report.as_dict())

    def test_date_variants(self):
        for raw in ("2026-10-05", "2026/10/5", "2026.10.05", "2026年10月5日", "20261005",
                    datetime(2026, 10, 5, 8, 30)):
            self.assertEqual(X._date(raw), "2026-10-05", raw)

    def test_date_unparseable_passes_through(self):
        """不认识的日期不静默变空，原样交给规则引擎报数据质量。"""
        self.assertEqual(X._date("十月五日"), "十月五日")

    def test_time_variants(self):
        self.assertEqual(X._time("8:30"), "08:30")
        self.assertEqual(X._time("08：30"), "08:30")
        self.assertEqual(X._time("08:30:00"), "08:30")
        self.assertEqual(X._time(time(9, 5)), "09:05")
        self.assertEqual(X._time(datetime(2026, 10, 5, 14, 0)), "14:00")

    def test_excel_fractional_time(self):
        """openpyxl 读纯时间格会得到 0~1 的小数，8:00 = 1/3。"""
        self.assertEqual(X._time(8 / 24), "08:00")
        self.assertEqual(X._time(14.5 / 24), "14:30")

    def test_sign_variants(self):
        for yes in ("是", "√", "✓", "有", "Y", "true", "1", "已签"):
            self.assertTrue(X._sign(yes), yes)
        for no in ("否", "×", "无", "N", "0", "", None, "未签"):
            self.assertFalse(X._sign(no), no)
        self.assertTrue(X._sign("张三"))  # 手写姓名视作已签

    def test_items_split_and_code_lookup(self):
        got = X._items("擦浴、洗发；SH01, 不存在的项目")
        self.assertEqual(got, ["擦浴", "洗发", "协助进食/水及指导（含管饲饮食）", "不存在的项目"])

    def test_items_single(self):
        self.assertEqual(X._items("擦浴"), ["擦浴"])
        self.assertEqual(X._items(""), [])
        self.assertEqual(X._items(None), [])

    def test_int(self):
        self.assertEqual(X._int("60分钟"), 60)
        self.assertEqual(X._int(""), None)
        self.assertEqual(X._int(30.0), 30)


class TestSheetAndHeaderMatching(unittest.TestCase):
    def test_sheet_alias(self):
        self.assertEqual(X._classify_sheet("服务记录单"), "records")
        self.assertEqual(X._classify_sheet("护理记录"), "records")
        self.assertEqual(X._classify_sheet("护理服务计划"), "plan")
        self.assertEqual(X._classify_sheet("护理计划"), "plan")
        self.assertEqual(X._classify_sheet("护理人员台账"), "caregivers")
        self.assertEqual(X._classify_sheet("参保人信息"), "elders")
        self.assertIsNone(X._classify_sheet("Sheet1"))

    def test_longest_alias_wins(self):
        """「护理服务计划」不能被「护理人员」之类的短别名抢走。"""
        self.assertEqual(X._classify_sheet("护理服务计划表"), "plan")

    def test_cross_category_alias_prefers_longer(self):
        """跨类别歧义表名取更长的别名：「参保人服务记录」应识别为服务记录而非参保人表。"""
        self.assertEqual(X._classify_sheet("参保人服务记录"), "records")
        self.assertEqual(X._classify_sheet("参保人信息"), "elders")

    def test_header_normalisation(self):
        self.assertEqual(X._norm_header("时长 (分钟)"), "时长分钟")
        self.assertEqual(X._norm_header("  服务日期 "), "服务日期")

    def test_field_aliases_resolve(self):
        cols = X._build_header_index(["老人姓名", "身份证", "护理方式", "评估等级", "家庭住址"])
        self.assertEqual(set(cols), {"elder_name", "id_card", "service_type",
                                     "disability_level", "address"})

    def test_generic_name_does_not_steal_specific_column(self):
        """表里同时有「参保人姓名」和「护理人员姓名」时，两边必须各归各位。"""
        cols = X._build_header_index(["服务编号", "参保人姓名", "护理人员姓名", "服务日期"])
        self.assertEqual(cols["elder_name"], 1)
        self.assertEqual(cols["caregiver_name"], 2)

    def test_generic_name_falls_back(self):
        """只有一列「姓名」时，谁列在前谁先匹配（老人在前是主表惯例）。"""
        cols = X._build_header_index(["姓名", "资质"])
        self.assertEqual(cols["elder_name"], 0)
        self.assertEqual(cols["qualification"], 1)


@unittest.skipIf(Workbook is None, "需要 openpyxl")
class TestLoadCase(unittest.TestCase):
    def setUp(self):
        self.tmp = []

    def tearDown(self):
        for p in self.tmp:
            Path(p).unlink(missing_ok=True)

    def _build(self, sheets):
        p = build(sheets)
        self.tmp.append(p)
        return p

    def test_minimal_workbook(self):
        p = self._build({
            "参保人信息": [["参保人姓名", "护理方式", "居住地址"],
                       ["张三", "居家护理", "甲市乙区丙路 1 号"]],
            "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录单": [RECORD_HEADER,
                       ["R1", "张三", "李某", "2026/10/5", "8:00", "9:00", 60,
                        "擦浴", 1, "是", "是"]],
        })
        case, report = X.load_case(p)
        self.assertEqual(len(case["elders"]), 1)
        self.assertEqual(len(case["caregivers"]), 1)
        self.assertEqual(len(case["records"]), 1)
        r = case["records"][0]
        self.assertEqual(r["service_date"], "2026-10-05")
        self.assertEqual(r["start_time"], "08:00")
        self.assertEqual(r["items"], ["擦浴"])
        self.assertTrue(r["caregiver_sign"])
        self.assertEqual([i for i in report.issues if i["level"] == "E"], [])

    def test_alternate_headers_and_order(self):
        """表头换写法、工作表顺序打乱，都应能导入。"""
        p = self._build({
            "服务明细": [["日期", "姓名", "服务项目", "开始时间", "结束时间", "护理员姓名",
                      "身份确认", "护理员确认"],
                     ["2026-10-06", "王五", "洗发", "10:00", "11:00", "赵某", "√", "√"]],
            "人员台账": [["人员姓名", "职业资格"], ["赵某", "长期照护师"]],
            "服务对象": [["姓名", "护理方式"], ["王五", "机构护理"]],
        })
        case, report = X.load_case(p)
        self.assertEqual(len(case["records"]), 1, report.as_dict())
        self.assertEqual(case["records"][0]["items"], ["洗发"])

    def test_unknown_elder_is_reported_and_skipped(self):
        p = self._build({
            "参保人": [["参保人姓名"], ["张三"]],
            "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录": [RECORD_HEADER,
                     ["R1", "查无此人", "李某", "2026-10-05", "08:00", "09:00", 60, "擦浴", 1, "是", "是"]],
        })
        case, report = X.load_case(p)
        self.assertEqual(case["records"], [])
        self.assertEqual(len([i for i in report.issues if i["level"] == "E"]), 1)

    def test_caregiver_not_in_roster_backfilled_with_empty_qualification(self):
        """记录里出现的人员必须进人员表，否则资质校验会被静默跳过。"""
        p = self._build({
            "参保人": [["参保人姓名"], ["张三"]],
            "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录": [RECORD_HEADER,
                     ["R1", "张三", "陌生人", "2026-10-05", "08:00", "09:00", 60, "吸痰", 1, "是", "是"]],
        })
        case, report = X.load_case(p)
        quals = {c["caregiver_id"]: c["qualification"] for c in case["caregivers"]}
        self.assertEqual(quals["陌生人"], "")
        self.assertTrue(any(i["level"] == "W" for i in report.issues))
        # 空资质 + 医疗护理项目 → 应为硬违规，说明补录让校验生效了
        self.assertIn("E_QUAL_MISMATCH", [f.rule for f in R.check_consistency(case)])

    def test_plan_attaches_to_elder(self):
        p = self._build({
            "参保人": [["参保人姓名", "护理方式"], ["张三", "居家护理"]],
            "护理服务计划": [["参保人姓名", "服务项目", "月频次"],
                        ["张三", "擦浴", 8], ["张三", "洗发", 4]],
        })
        case, _ = X.load_case(p)
        plan = case["elders"][0]["plan"]
        self.assertEqual([(x["name"], x["freq_per_month"]) for x in plan],
                         [("擦浴", 8), ("洗发", 4)])

    def test_plan_for_unknown_elder_is_skipped(self):
        p = self._build({
            "参保人": [["参保人姓名"], ["张三"]],
            "护理服务计划": [["参保人姓名", "服务项目", "月频次"], ["李四", "擦浴", 8]],
        })
        case, report = X.load_case(p)
        self.assertEqual(case["elders"][0]["plan"], [])
        self.assertTrue(any("不在参保人表中" in i["message"] for i in report.issues))

    def test_missing_sheet_is_fatal(self):
        p = self._build({"参保人": [["参保人姓名"], ["张三"]]})
        _, report = X.load_case(p)
        self.assertEqual(len([i for i in report.issues if i["level"] == "E"]), 2)  # 缺 caregivers + records

    def test_missing_required_column_is_fatal(self):
        p = self._build({
            "参保人": [["参保人姓名"], ["张三"]],
            "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录": [["服务编号", "参保人姓名", "服务项目"], ["R1", "张三", "擦浴"]],
        })
        _, report = X.load_case(p)
        missing = [i["field"] for i in report.issues if i["level"] == "E" and i["field"]]
        for field in ("service_date", "start_time"):
            self.assertIn(field, missing)

    def test_item_count_autofill(self):
        """没填项目数合计时按实际项目数补，避免每行都报数据质量。"""
        p = self._build({
            "参保人": [["参保人姓名"], ["张三"]],
            "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录": [["服务编号", "参保人姓名", "护理人员姓名", "服务日期", "开始时间",
                      "结束时间", "服务项目", "护理人员签字", "参保人员（家属）签名"],
                     ["R1", "张三", "李某", "2026-10-05", "08:00", "09:00", "擦浴、洗发", "是", "是"]],
        })
        case, _ = X.load_case(p)
        self.assertEqual(case["records"][0]["item_count"], 2)
        self.assertEqual(R.check_consistency(case), [])

    def test_header_can_be_below_title_row(self):
        """机构常见做法：第一行是大标题，第二行才是表头。"""
        p = self._build({
            "参保人": [["某机构参保人台账"], ["参保人姓名", "护理方式"], ["张三", "居家护理"]],
            "护理人员": [["护理人员台账"], ["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录": [["长期护理保险服务记录单"], RECORD_HEADER,
                     ["R1", "张三", "李某", "2026-10-05", "08:00", "09:00", 60, "擦浴", 1, "是", "是"]],
        })
        case, report = X.load_case(p)
        self.assertEqual(len(case["records"]), 1, report.as_dict())
        self.assertEqual(int(report.stats["elders"]), 1)

    def test_blank_rows_ignored(self):
        p = self._build({
            "参保人": [["参保人姓名"], ["张三"], [None], [""]],
            "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
            "服务记录": [RECORD_HEADER,
                     ["R1", "张三", "李某", "2026-10-05", "08:00", "09:00", 60, "擦浴", 1, "是", "是"],
                     [None] * 11],
        })
        case, _ = X.load_case(p)
        self.assertEqual(len(case["records"]), 1)


@unittest.skipIf(Workbook is None, "需要 openpyxl")
class TestRoundTrip(unittest.TestCase):
    """模板 → 导入 → case，必须与源 case 在规则引擎看来等价。"""

    def _roundtrip(self, fixture: str):
        src = json.loads((ROOT / "tests" / "fixtures" / fixture).read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as d:
            xlsx = Path(d) / "t.xlsx"
            X.write_template(xlsx, ROOT / "tests" / "fixtures" / fixture)
            case, report = X.load_case(xlsx)
        self.assertEqual([i for i in report.issues if i["level"] == "E"], [], report.as_dict())
        return src, case

    def _signature(self, findings):
        """只比规则名和记录号：导入会补 item_count、按姓名做键，细节字段名会变。"""
        return sorted((f.rule, f.record_no) for f in findings)

    def test_sample_fixture_equivalent(self):
        src, case = self._roundtrip("sample_case.json")
        self.assertEqual(len(case["records"]), len(src["records"]))
        self.assertEqual(self._signature(R.check_consistency(case)),
                         self._signature(R.check_consistency(src)))

    def test_clean_fixture_equivalent(self):
        src, case = self._roundtrip("clean_case.json")
        self.assertEqual(R.check_consistency(case), [])
        self.assertEqual(self._signature(R.check_consistency(case)),
                         self._signature(R.check_consistency(src)))

    def test_template_has_guide_sheet_with_catalog(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as d:
            xlsx = Path(d) / "t.xlsx"
            X.write_template(xlsx)
            wb = load_workbook(xlsx, read_only=True)
            self.assertIn("填写说明", wb.sheetnames)
            rows = [list(r) for r in wb["填写说明"].iter_rows(values_only=True)]
            codes = [r[0] for r in rows if r[0] and str(r[0]).startswith(("SH", "YL"))]
            self.assertEqual(len(codes), 36)
            wb.close()


@unittest.skipIf(Workbook is None, "需要 openpyxl")
class TestReviewPipeline(unittest.TestCase):
    """一条命令的完整链路：Excel → case → 核验 → 异常清单 xlsx。"""

    def test_review_writes_case_and_report(self):
        import subprocess
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            xlsx = d / "in.xlsx"
            X.write_template(xlsx, ROOT / "tests" / "fixtures" / "sample_case.json")
            p = subprocess.run([sys.executable, "ltc_excel.py", "review", str(xlsx)],
                               cwd=ROOT, capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(p.returncode, 1, p.stderr)  # sample 有硬违规
            out = json.loads(p.stdout)
            self.assertTrue(Path(out["case"]).exists())
            self.assertTrue(Path(out["report"]).exists())
            self.assertGreater(out["summary"]["hard_violations"], 0)

            from openpyxl import load_workbook
            wb = load_workbook(out["report"])
            self.assertEqual(wb.sheetnames[:2], ["异常清单", "汇总"])
            rows = list(wb["异常清单"].iter_rows(values_only=True))
            self.assertEqual(rows[0][:3], ("严重度", "规则", "对象"))
            # 每条异常都要有文书依据，否则机构无法整改
            for r in rows[1:]:
                self.assertTrue(r[4], r)
            # 对象列不允许空
            for r in rows[1:]:
                self.assertTrue(r[2], r)
            wb.close()

    def test_review_clean_exits_zero(self):
        import subprocess
        with tempfile.TemporaryDirectory() as d:
            xlsx = Path(d) / "in.xlsx"
            X.write_template(xlsx, ROOT / "tests" / "fixtures" / "clean_case.json")
            p = subprocess.run([sys.executable, "ltc_excel.py", "review", str(xlsx)],
                               cwd=ROOT, capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
            self.assertEqual(json.loads(p.stdout)["summary"]["hard_violations"], 0)

    def test_strict_mode_fails_on_import_warnings(self):
        import subprocess
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            xlsx = d / "in.xlsx"
            wb = Workbook()
            wb.remove(wb.active)
            for title, rows in {
                "参保人": [["参保人姓名", "身份证号"], ["张三", 110101199001011234.0]],
                "护理人员": [["护理人员姓名", "资质"], ["李某", "护士"]],
                "服务记录": [["服务编号", "参保人姓名", "护理人员姓名", "服务日期",
                          "开始时间", "结束时间", "服务项目", "护理人员签字", "家属签字"],
                         ["R1", "张三", "李某", "2026-10-05", "08:00", "09:00", "擦浴", "是", "是"]],
            }.items():
                ws = wb.create_sheet(title)
                for r in rows:
                    ws.append(r)
            wb.save(xlsx)

            base = [sys.executable, "ltc_excel.py", "review", str(xlsx)]
            p1 = subprocess.run(base, cwd=ROOT, capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(p1.returncode, 0, p1.stdout)   # 只有 W，不算失败
            p2 = subprocess.run(base + ["--strict"], cwd=ROOT, capture_output=True,
                                text=True, encoding="utf-8")
            self.assertEqual(p2.returncode, 1, p2.stdout)   # strict 下 W 也失败


if __name__ == "__main__":
    unittest.main(verbosity=2)
