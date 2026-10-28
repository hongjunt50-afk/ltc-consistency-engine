"""长护险服务过程一致性核验与评分：把 2026 版官方文书编译成可执行规则。

覆盖三块，来源均为《长期护理保险服务管理文书（2026 年版）》（国家医保局办公室印发）：

1. 服务过程一致性核验（服务记录单 vs 护理服务计划 vs 36 项目录）
   对应文书：表 3 长期护理保险服务记录单、表 2 护理服务计划表
2. 巡查记录表评分（8 分区 30 项，满分 100）
   对应文书：表 3 定点长护服务机构巡查记录表
3. 质量评价表算分（A 生活照护 / B 医疗护理 / C 其他，每项 0–5 分，三类得分率均 ≥60% 合格）
   对应文书：表 2 长期护理服务质量评价表

文书原文由 tools/extract_ruleset.py 从 PDF 抽成 data/derived/*.json，本模块只读 JSON，不内置条文。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
DERIVED = ROOT / "data" / "derived"


def _load(name: str) -> dict:
    path = DERIVED / name
    if not path.exists():
        raise FileNotFoundError(f"缺少 {path}，先跑 python tools/extract_ruleset.py")
    return json.loads(path.read_text(encoding="utf-8"))


CATALOG = _load("services_catalog.json")
INSPECTION = _load("inspection_checklist.json")
QUALITY = _load("quality_review.json")
RECORD_SCHEMA = _load("service_record_schema.json")

BY_NAME = {i["name"]: i for i in CATALOG["items"]}
BY_CODE = {i["code"]: i for i in CATALOG["items"]}
MEDICAL_CODES = {i["code"] for i in CATALOG["items"] if i["category"] == "医疗护理类"}

# 《服务记录单》"护理人员资质"栏的六种取值。前五类都是持证人员；
# 质量评价表只在「非持证人员提供医疗护理」时才直接判不合格，所以不能按"护士/医生专属"来判。
CERTIFIED_QUALIFICATIONS = {"长期照护师", "护士", "医生", "医疗护理员", "养老护理员"}
# 这五类中，执业范围偏生活照护、做医疗护理类项目需人工复核（各地口径不一）
SCOPE_REVIEW_QUALIFICATIONS = {"养老护理员", "医疗护理员"}


# ---------------------------------------------------------------- 通用

@dataclass
class Finding:
    """一条可解释的核验结果：命中哪条规则、哪个字段冲突、依据哪份文书。"""
    rule: str
    severity: str  # E 硬违规 / W 需复核 / I 数据质量
    message: str
    clause: str
    record_no: str | None = None
    evidence: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {
            "rule": self.rule, "severity": self.severity, "message": self.message,
            "clause": self.clause, "record_no": self.record_no, "evidence": self.evidence,
        }


def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def _parse_time(s: str) -> time:
    return datetime.strptime(s, "%H:%M").time()


def _minutes(start: str, end: str) -> int:
    a, b = _parse_time(start), _parse_time(end)
    return (b.hour * 60 + b.minute) - (a.hour * 60 + a.minute)


# ---------------------------------------------------------------- 1. 一致性核验

def check_consistency(case: dict, *, day_max_minutes: int = 480,
                      late_entry_days: int = 3, duration_tolerance: int = 1) -> list[Finding]:
    """核验一批服务记录与护理计划、服务目录、人员资质、签字的自洽性。

    case 结构见 tests/fixtures/sample_case.json。所有阈值显式入参，不改代码即可按地区政策调。
    """
    out: list[Finding] = []
    elders = {e["elder_id"]: e for e in case.get("elders", [])}
    caregivers = {c["caregiver_id"]: c for c in case.get("caregivers", [])}
    records = case.get("records", [])

    for r in records:
        rid = r.get("record_no")
        elder = elders.get(r.get("elder_id"))
        caregiver = caregivers.get(r.get("caregiver_id"))

        # E_UNKNOWN_SERVICE：目录实行准入法管理，未列入不予支付
        for name in r.get("items", []):
            if name not in BY_NAME and not any(i["code"] == name for i in CATALOG["items"]):
                out.append(Finding(
                    "E_UNKNOWN_SERVICE", "E", f"服务项目「{name}」不在国家 36 项目录内", 
                    "国家长期护理保险服务项目目录（试行）医保发〔2025〕24号 凡例·三（准入法管理）",
                    rid, {"item": name}))

        # E_BAD_TIME / E_DURATION_MISMATCH
        if r.get("start_time") and r.get("end_time"):
            diff = _minutes(r["start_time"], r["end_time"])
            if diff <= 0:
                out.append(Finding(
                    "E_BAD_TIME", "E",
                    f"服务结束时间 {r['end_time']} 不晚于开始时间 {r['start_time']}",
                    "长期护理保险服务记录单·服务记录栏（服务开始时间/服务结束时间）",
                    rid, {"start_time": r["start_time"], "end_time": r["end_time"]}))
            elif r.get("duration_min") is not None and abs(diff - r["duration_min"]) > duration_tolerance:
                out.append(Finding(
                    "E_DURATION_MISMATCH", "E",
                    f"填报时长 {r['duration_min']} 分钟与起止时间推算的 {diff} 分钟不符",
                    "长期护理保险服务记录单·服务记录栏（时长应按起止时间填写）",
                    rid, {"declared": r["duration_min"], "computed": diff,
                          "start_time": r["start_time"], "end_time": r["end_time"]}))

        # I_ITEM_COUNT_MISMATCH
        if r.get("item_count") is not None and r.get("items") is not None:
            if r["item_count"] != len(r["items"]):
                out.append(Finding(
                    "I_ITEM_COUNT_MISMATCH", "I",
                    f"服务项目数合计填 {r['item_count']}，实际列了 {len(r['items'])} 项",
                    "长期护理保险服务记录单·服务记录栏（服务项目数合计）",
                    rid, {"declared": r["item_count"], "actual": len(r["items"])}))

        # E_QUAL_MISMATCH：医疗护理类项目须由护士/医生执行
        if caregiver:
            qual = (caregiver.get("qualification") or "").strip()
            med_items = [n for n in r.get("items", [])
                         if n in BY_NAME and BY_NAME[n]["category"] == "医疗护理类"]
            if med_items and qual not in CERTIFIED_QUALIFICATIONS:
                # 质量评价表只把「非持证人员提供医疗护理」列为直接不合格，其余为无资质执行
                out.append(Finding(
                    "E_QUAL_MISMATCH", "E",
                    f"资质「{qual or '未填'}」不属于持证护理人员，却执行了医疗护理类项目 {med_items}",
                    "长期护理保险服务记录单·护理人员资质；质量评价表·其他类第 1 项（非持证人员提供医疗护理直接判不合格）",
                    rid, {"qualification": qual, "items": med_items,
                          "certified": sorted(CERTIFIED_QUALIFICATIONS)}))
            elif med_items and qual in SCOPE_REVIEW_QUALIFICATIONS:
                out.append(Finding(
                    "W_QUAL_SCOPE_REVIEW", "W",
                    f"资质「{qual}」执行医疗护理类项目 {med_items}，需核对该地区是否允许其执业范围",
                    "长期护理保险服务记录单·护理人员资质；巡查记录表·机构管理（人员类型与服务范围相匹配）",
                    rid, {"qualification": qual, "items": med_items}))

        # E_NO_SIGN：一服务一记录，须双方签字
        if not r.get("caregiver_sign"):
            out.append(Finding("E_NO_SIGN", "E", "缺护理人员签字",
                               "长期护理保险服务记录单·签字确认栏（护理人员签字）", rid, {}))
        if not r.get("elder_sign"):
            out.append(Finding("E_NO_SIGN", "E", "缺参保人员（家属）签名",
                               "长期护理保险服务记录单·签字确认栏（参保人员（家属）签名）", rid, {}))

        # E_ABNORMAL_NO_REMEDY：写了异常必须写处理措施与处理时间
        if r.get("abnormal_note"):
            for k, label in (("remedy", "处理措施"), ("remedy_time", "处理时间")):
                if not r.get(k):
                    out.append(Finding(
                        "E_ABNORMAL_NO_REMEDY", "E",
                        f"记录了异常情况（{r['abnormal_note'][:20]}…）但缺{label}",
                        "长期护理保险服务记录单·异常情况说明/处理措施/处理时间",
                        rid, {"abnormal_note": r["abnormal_note"], "missing": k}))

        # W_OFF_PLAN：服务项目不在该老人护理计划内
        planned = {p.get("name") or p.get("code") for p in (elder or {}).get("plan", [])}
        planned.discard(None)
        if planned:
            for name in r.get("items", []):
                if name in BY_NAME and name not in planned:
                    out.append(Finding(
                        "W_OFF_PLAN", "W", f"服务项目「{name}」不在该参保人的护理服务计划内",
                        "护理服务计划表·服务项目栏；巡查记录表·服务管理（按护理服务计划提供服务）",
                        rid, {"item": name, "planned_count": len(planned)}))

        # W_SERVICE_TYPE_MISMATCH
        if elder and elder.get("service_type") and r.get("service_type") \
                and r["service_type"] != elder["service_type"]:
            out.append(Finding(
                "W_SERVICE_TYPE_MISMATCH", "W",
                f"记录服务类型「{r['service_type']}」与备案护理方式「{elder['service_type']}」不一致",
                "长期护理保险服务记录单·服务类型；失能人员档案信息表·护理方式",
                rid, {"record": r["service_type"], "enrolled": elder["service_type"]}))

        # W_ADDRESS_MISMATCH：居家护理须在备案地址
        if r.get("service_type") == "居家护理" and elder and elder.get("address") \
                and r.get("service_address") and r["service_address"] != elder["address"]:
            out.append(Finding(
                "W_ADDRESS_MISMATCH", "W",
                f"服务地址「{r['service_address']}」与备案居住地址「{elder['address']}」不一致",
                "长期护理保险服务记录单·服务地址；失能人员档案信息表·居住地址",
                rid, {"record": r["service_address"], "enrolled": elder["address"]}))

        # W_LATE_ENTRY：事后补录
        if r.get("submit_time") and r.get("service_date"):
            lag = (_parse_date(r["submit_time"][:10]) - _parse_date(r["service_date"])).days
            if lag > late_entry_days:
                out.append(Finding(
                    "W_LATE_ENTRY", "W",
                    f"服务日期 {r['service_date']}，录入日期 {r['submit_time'][:10]}，滞后 {lag} 天",
                    "长期护理保险服务记录单（按次如实填写）；长期护理服务月度汇总表（按月归集）",
                    rid, {"service_date": r["service_date"], "submit_time": r["submit_time"],
                          "lag_days": lag, "tolerance_days": late_entry_days}))

    # E_CAREGIVER_OVERLAP / E_ELDER_OVERLAP / W_EXCESS_DAILY：跨记录校验
    by_caregiver: dict[tuple, list] = {}
    by_elder: dict[tuple, list] = {}
    by_day: dict[tuple, int] = {}
    for r in records:
        if not (r.get("service_date") and r.get("start_time") and r.get("end_time")):
            continue
        span = (_parse_time(r["start_time"]), _parse_time(r["end_time"]))
        if span[1] <= span[0]:
            continue  # 已由 E_BAD_TIME 报过
        by_caregiver.setdefault((r.get("caregiver_id"), r["service_date"]), []).append((span, r))
        by_elder.setdefault((r.get("elder_id"), r["service_date"]), []).append((span, r))
        by_day[(r.get("elder_id"), r["service_date"])] = \
            by_day.get((r.get("elder_id"), r["service_date"]), 0) + _minutes(r["start_time"], r["end_time"])

    for (cid, day), rows in by_caregiver.items():
        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                (s1, e1), r1 = rows[i]
                (s2, e2), r2 = rows[j]
                if s1 < e2 and s2 < e1:
                    out.append(Finding(
                        "E_CAREGIVER_OVERLAP", "E",
                        f"护理人员 {cid} 在 {day} 的 {r1['record_no']} 与 {r2['record_no']} 时段重叠",
                        "长期护理保险服务记录单·服务开始/结束时间；巡查记录表·费用结算（不得重复收取）",
                        r1["record_no"], {"other_record": r2["record_no"],
                                          "a": [r1["start_time"], r1["end_time"]],
                                          "b": [r2["start_time"], r2["end_time"]]}))

    for (eid, day), rows in by_elder.items():
        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                (s1, e1), r1 = rows[i]
                (s2, e2), r2 = rows[j]
                if s1 < e2 and s2 < e1:
                    out.append(Finding(
                        "E_ELDER_OVERLAP", "E",
                        f"参保人 {eid} 在 {day} 同时段有两条服务记录（{r1['record_no']} / {r2['record_no']}）",
                        "长期护理保险服务记录单·一服务一记录；巡查记录表·费用结算（不得重复收取）",
                        r2["record_no"], {"other_record": r1["record_no"]}))

    for (eid, day), total in by_day.items():
        if total > day_max_minutes:
            out.append(Finding(
                "W_EXCESS_DAILY", "W",
                f"参保人 {eid} 在 {day} 累计服务 {total} 分钟，超过单日阈值 {day_max_minutes} 分钟",
                "国家长期护理保险服务项目目录·四（各地应测算确定项目服务时长与享受频次）",
                None, {"elder_id": eid, "date": day, "total_minutes": total,
                       "threshold": day_max_minutes}))

    # W_FREQ_SHORT / W_FREQ_OVER：按月比对计划频次
    for elder in elders.values():
        plan = elder.get("plan") or []
        if not plan:
            continue
        month_counts: dict[str, dict[str, int]] = {}
        for r in records:
            if r.get("elder_id") != elder["elder_id"] or not r.get("service_date"):
                continue
            ym = r["service_date"][:7]
            for name in r.get("items", []):
                month_counts.setdefault(ym, {}).setdefault(name, 0)
                month_counts[ym][name] += 1
        for ym, counter in month_counts.items():
            for p in plan:
                name = p.get("name") or p.get("code")
                want = p.get("freq_per_month")
                if want is None:
                    continue
                got = counter.get(name, 0)
                if got < want:
                    out.append(Finding(
                        "W_FREQ_SHORT", "W",
                        f"{elder['elder_id']} 在 {ym} 的「{name}」执行 {got} 次，低于计划 {want} 次",
                        "护理服务计划表·建议频次/确定频次；长期护理服务月度汇总表",
                        None, {"elder_id": elder["elder_id"], "month": ym,
                               "item": name, "planned": want, "actual": got}))
                elif got > want:
                    out.append(Finding(
                        "W_FREQ_OVER", "W",
                        f"{elder['elder_id']} 在 {ym} 的「{name}」执行 {got} 次，超出计划 {want} 次",
                        "护理服务计划表·确定频次；巡查记录表·费用结算（不得重复、分解收取）",
                        None, {"elder_id": elder["elder_id"], "month": ym,
                               "item": name, "planned": want, "actual": got}))

    return out


# ---------------------------------------------------------------- 2. 巡查表评分

def score_inspection(scores: dict[str, list[int | None]]) -> dict:
    """按《定点长护服务机构巡查记录表》算分。

    scores 形如 {"机构管理": [3,3,None,...], ...}，None 表示该检查项未适用（不计入分母）。
    返回分区得分、总分与逐项校验错误。
    """
    errors: list[str] = []
    sections_out = []
    total, total_applicable = 0, 0

    for sec in INSPECTION["sections"]:
        name = sec["section"]
        got = scores.get(name)
        if got is None:
            errors.append(f"缺少分区「{name}」的评分")
            continue
        if len(got) != len(sec["items"]):
            errors.append(f"分区「{name}」应有 {len(sec['items'])} 项评分，给了 {len(got)} 项")
            continue
        sec_score, sec_full = 0, 0
        for idx, (item, val) in enumerate(zip(sec["items"], got)):
            if val is None:
                continue
            if not isinstance(val, (int, float)) or val < 0 or val > item["score"]:
                errors.append(f"「{name}」第 {idx + 1} 项得分 {val} 超出 0–{item['score']} 范围")
                continue
            sec_score += val
            sec_full += item["score"]
        sections_out.append({"section": name, "score": sec_score,
                             "max_score": sec["max_score"], "applicable_full": sec_full})
        total += sec_score
        total_applicable += sec_full

    return {
        "form": INSPECTION["title"],
        "sections": sections_out,
        "total_score": total,
        "max_score": INSPECTION["max_score"],
        "applicable_full": total_applicable,
        "errors": errors,
        "valid": not errors,
    }


# ---------------------------------------------------------------- 3. 质量评价算分

def score_quality(scores: dict[str, list[int | None]], *, hard_fail_hits: list[str] | None = None) -> dict:
    """按《长期护理服务质量评价表》算分：A/B/C 三类得分率均 ≥60% 才合格。

    scores 形如 {"A": [5,4,...], "B": [...], "C": [...]}；None 视为该项不适用。
    hard_fail_hits 传命中的直接不合格条件原文（如「出现打骂参保人」）。
    """
    errors: list[str] = []
    tables_out = []
    all_pass = True

    for table in QUALITY["tables"]:
        tid = table["id"]
        got = scores.get(tid)
        if got is None:
            if table["items"]:  # A/B 类必须评；C 类在只做生活照护时可能整体不适用
                errors.append(f"缺少表 {tid}（{table['name']}）的评分")
                all_pass = False
            continue
        if len(got) != len(table["items"]):
            errors.append(f"表 {tid} 应有 {len(table['items'])} 项评分，给了 {len(got)} 项")
            continue
        vals = [v for v in got if v is not None]
        for k, v in enumerate(vals):
            if not isinstance(v, (int, float)) or v < 0 or v > QUALITY["score_per_item"]:
                errors.append(f"表 {tid} 第 {k + 1} 个有效项得分 {v} 超出 0–{QUALITY['score_per_item']} 范围")
        if not vals:
            tables_out.append({"id": tid, "name": table["name"], "items": 0,
                               "rate": None, "pass": None})
            continue
        rate = sum(vals) / (len(vals) * QUALITY["score_per_item"])
        passed = rate >= QUALITY["pass_rate"]
        all_pass &= passed
        tables_out.append({"id": tid, "name": table["name"], "items": len(vals),
                           "score": sum(vals), "full": len(vals) * QUALITY["score_per_item"],
                           "rate": round(rate, 4), "pass": passed})

    hits = [h for h in (hard_fail_hits or []) if any(h in f["condition"] or f["condition"] in h
                                                     for f in QUALITY["hard_fail"])]
    verdict = "不合格" if (hits or not all_pass) else "合格"
    return {
        "form": QUALITY["title"],
        "tables": tables_out,
        "pass_line": QUALITY["pass_rate"],
        "hard_fail_hits": hits,
        "verdict": verdict,
        "errors": errors,
        "valid": not errors,
    }


# ---------------------------------------------------------------- 汇总报告

def audit(case: dict, **kwargs) -> dict:
    """一致性核验汇总：按严重度分级，给出可解释证据清单。"""
    findings = check_consistency(case, **kwargs)
    by_sev = {"E": [], "W": [], "I": []}
    for f in findings:
        by_sev[f.severity].append(f.as_dict())
    counts = {k: len(v) for k, v in by_sev.items()}
    return {
        "records_checked": len(case.get("records", [])),
        "summary": {
            "hard_violations": counts["E"],
            "needs_review": counts["W"],
            "data_quality": counts["I"],
        },
        "findings": by_sev["E"] + by_sev["W"] + by_sev["I"],
    }


def main(argv: list[str] | None = None) -> int:
    """命令行：audit <case.json> | inspection <scores.json> | quality <scores.json>"""
    import sys
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print(__doc__)
        print("用法: python ltc_rules.py {audit|inspection|quality} <input.json>")
        return 2
    cmd, path = argv[0], Path(argv[1])
    payload = json.loads(path.read_text(encoding="utf-8"))

    if cmd == "audit":
        result = audit(payload)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["summary"]["hard_violations"] == 0 else 1
    if cmd == "inspection":
        result = score_inspection(payload)
    elif cmd == "quality":
        result = score_quality(payload.get("scores", payload),
                               hard_fail_hits=payload.get("hard_fail_hits"))
    else:
        print(f"未知命令: {cmd}")
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
