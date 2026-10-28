"""长护险服务记录 Excel 导入适配器：把机构实际在用的表格导入成规则引擎的 case 结构。

设计取舍：
- 不假设机构的表头写法统一，用别名表（每字段一组常见写法）做列匹配；
- 一个工作簿里按"表单名"分表，表名和表头都做模糊匹配，顺序不限；
- 校验失败不静默：逐行给出「哪张表第几行哪个字段」，导入报告与数据一起返回。

唯一权威输出是 case dict（结构见 tests/fixtures/*.json），规则引擎不关心数据从哪来。
"""
from __future__ import annotations

import json
import re
import sys
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
CATALOG = json.loads((ROOT / "data" / "derived" / "services_catalog.json").read_text(encoding="utf-8"))
NAME_BY_CODE = {i["code"]: i["name"] for i in CATALOG["items"]}
NAME_BY_NAME = {i["name"]: i["name"] for i in CATALOG["items"]}

# 表单类型 → (表名的可能写法, 匹配优先级)。表名歧义时按 SHEET_PRIORITY 高的赢：
# 服务记录表的名字里常带「参保人」「护理人员」，所以记录优先；护理计划常带「护理人员」次之。
SHEET_ALIASES = {
    "records": ["服务记录", "服务记录单", "护理记录", "记录单", "服务明细"],
    "plan": ["护理服务计划", "护理计划", "服务计划", "计 划"],
    "caregivers": ["护理人员", "照护人员", "人员信息", "人员台账"],
    "elders": ["参保人", "服务对象", "失能人员", "老人信息", "档案信息"],
}

# 字段别名。键是 case 结构里的字段名，值是可接受的表头写法（去空格后精确匹配）
FIELD_ALIASES: dict[str, list[str]] = {
    # 参保人
    "elder_name": ["参保人姓名", "姓名", "老人姓名", "服务对象姓名", "参保人员姓名"],
    "id_card": ["身份证号", "身份证号码", "身份证", "证件号码"],
    "service_type": ["护理方式", "服务类型", "服务方式", "护理类型"],
    "disability_level": ["失能等级", "失能等级评估结果", "评估等级", "护理等级"],
    "address": ["居住地址", "现住地址", "服务地址", "家庭住址", "地址"],
    # 计划
    "plan_item": ["服务项目", "项目名称", "护理项目"],
    "freq_per_month": ["月频次", "每月频次", "频次", "确定频次", "每月次数"],
    # 人员
    "caregiver_name": ["护理人员姓名", "护理员姓名", "照护人员姓名", "人员姓名", "姓名"],
    "qualification": ["护理人员资质", "资质", "人员资质", "职业资格", "职业"],
    # 记录
    "record_no": ["服务编号", "编号", "记录编号", "服务单号"],
    "service_date": ["服务日期", "日期", "服务时间", "护理日期"],
    "start_time": ["服务开始时间", "开始时间", "起始时间"],
    "end_time": ["服务结束时间", "结束时间", "终止时间"],
    "duration_min": ["时长（分钟）", "时长(分钟)", "时长", "分钟数", "服务时长"],
    "items": ["服务项目", "服务项目内容", "项目", "护理项目", "实际服务项目"],
    "item_count": ["服务项目数合计", "项目数", "服务项数", "项目数合计"],
    "caregiver_sign": ["护理人员签字", "护理员签字", "人员签字", "签字"],
    "elder_sign": ["参保人员（家属）签名", "参保人签字", "家属签字", "确认签字", "被服务方签字"],
    "abnormal_note": ["异常情况说明", "异常情况", "异常说明"],
    "remedy": ["处理措施", "处理", "处置措施"],
    "remedy_time": ["处理时间", "处置时间"],
    "submit_time": ["录入时间", "填报时间", "提交时间", "上传时间"],
}

REQUIRED = {
    "elders": ["elder_name"],
    "plan": ["elder_name", "plan_item"],
    "caregivers": ["caregiver_name", "qualification"],
    "records": ["elder_name", "caregiver_name", "service_date",
                "start_time", "end_time", "items"],
}

YES = {"是", "√", "✓", "有", "y", "yes", "true", "1", "已签", "已"}
NO = {"否", "×", "✗", "无", "n", "no", "false", "0", "未签", "未"}
ITEM_SEP = re.compile(r"[、；;，,\n]+")


# ---------------------------------------------------------------- 单元格归一

def _text(v: Any) -> str:
    """单元格 → 去空白字符串；数值型身份证/编号不转成科学计数法。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).replace("\u3000", " ").strip()


def _looks_precision_lost(v: Any) -> bool:
    """18 位身份证被 Excel 当数字存会丢末位（IEEE754 双精度只保证 15~17 位有效数字）。"""
    return isinstance(v, float) and abs(v) >= 1e15


def _date(v: Any) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = _text(v)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s  # 原样返回，让规则引擎报数据质量而不是在这里吞掉


def _time(v: Any) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, time):
        return v.strftime("%H:%M")
    if isinstance(v, (int, float)) and 0 <= float(v) < 1:
        # Excel 时间存的是 0~1 的小数
        total = round(float(v) * 24 * 60)
        return f"{total // 60:02d}:{total % 60:02d}"
    s = _text(v).replace("：", ":")
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def _sign(v: Any) -> bool:
    s = _text(v).lower()
    if s in YES:
        return True
    if s in NO or s == "":
        return False
    return True  # 有内容（比如手写姓名）即视为已签


def _items(v: Any) -> list[str]:
    """一格写多个项目：按、；;，,和换行拆；拆出来的编号尽量还原成项目名。"""
    s = _text(v)
    if not s:
        return []
    out = []
    for part in ITEM_SEP.split(s):
        p = part.strip()
        if not p:
            continue
        out.append(NAME_BY_CODE.get(p.upper(), NAME_BY_NAME.get(p, p)))
    return out


def _int(v: Any) -> int | None:
    s = _text(v)
    if not s:
        return None
    m = re.search(r"-?\d+", s)
    return int(m.group()) if m else None


# ---------------------------------------------------------------- 表头与列定位

def _norm_header(s: Any) -> str:
    return re.sub(r"[\s()（）]", "", _text(s))


def _build_header_index(header_row: list[Any]) -> dict[str, int]:
    """表头 → {字段名: 列下标}。别名按「具体 → 通用」排列，先命中者胜（如先「护理人员姓名」后「姓名」）。"""
    norm = [_norm_header(h) for h in header_row]
    index: dict[str, int] = {}
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            key = _norm_header(alias)
            if key in norm:
                index[field] = norm.index(key)
                break
    return index


def _classify_sheet(name: str) -> str | None:
    """表名 → 表单类型。命中越长越具体；同长度按 SHEET_ALIASES 的书写顺序定优先级。"""
    target = _norm_header(name)
    best, best_len = None, 0
    for kind, aliases in SHEET_ALIASES.items():
        for alias in aliases:
            a = _norm_header(alias)
            if a and a in target and len(a) > best_len:
                best, best_len = kind, len(a)
    return best


# ---------------------------------------------------------------- 主流程

class ImportReport:
    def __init__(self) -> None:
        self.issues: list[dict] = []
        self.stats: dict[str, int] = {}

    def add(self, sheet: str, row: int | None, field: str | None, msg: str, level: str = "W") -> None:
        self.issues.append({"sheet": sheet, "row": row, "field": field,
                            "message": msg, "level": level})

    def as_dict(self) -> dict:
        return {"stats": self.stats, "issue_count": len(self.issues), "issues": self.issues}


def _cell(row, idx):
    return row[idx] if idx is not None and idx < len(row) else None


def load_case(path: str | Path) -> tuple[dict, ImportReport]:
    """读取工作簿，返回 (case, 报告)。报告里有逐行问题，不抛异常。"""
    try:
        from openpyxl import load_workbook
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("需要 openpyxl：pip install openpyxl") from e

    wb = load_workbook(Path(path), data_only=True, read_only=True)
    report = ImportReport()
    sheets: dict[str, tuple[str, list]] = {}
    for ws in wb.worksheets:
        kind = _classify_sheet(ws.title)
        if kind and kind not in sheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            sheets[kind] = (ws.title, rows)

    for kind in ("elders", "caregivers", "records"):
        if kind not in sheets:
            report.add(kind, None, None, f"缺少工作表（表名需含「{SHEET_ALIASES[kind][0]}」）", "E")

    elders: dict[str, dict] = {}
    caregivers: dict[str, dict] = {}
    records: list[dict] = []

    # ---- 参保人
    if "elders" in sheets:
        title, rows = sheets["elders"]
        hdr_i, cols = _find_header(rows, report, title, "elders")
        if cols:
            for n, row in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
                name = _text(_cell(row, cols.get("elder_name")))
                if not name:
                    continue
                key = name
                if key in elders:
                    report.add(title, n, "elder_name", f"参保人「{name}」重复，保留首条", "W")
                    continue
                raw_id = _cell(row, cols.get("id_card"))
                if _looks_precision_lost(raw_id):
                    report.add(title, n, "id_card",
                               f"身份证号被 Excel 存成了数字，末几位精度已丢失（读到 {_text(raw_id)}），"
                               f"请把该列设为文本格式重新填写", "W")
                elders[key] = {
                    "elder_id": key,
                    "name": name,
                    "id_card": _text(raw_id),
                    "service_type": _text(_cell(row, cols.get("service_type"))),
                    "disability_level": _text(_cell(row, cols.get("disability_level"))),
                    "address": _text(_cell(row, cols.get("address"))),
                    "plan": [],
                }
            report.stats["elders"] = len(elders)

    # ---- 护理计划
    if "plan" in sheets:
        title, rows = sheets["plan"]
        hdr_i, cols = _find_header(rows, report, title, "plan")
        if cols:
            n_plan = 0
            for n, row in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
                name = _text(_cell(row, cols.get("elder_name")))
                item = _text(_cell(row, cols.get("plan_item")))
                if not name or not item:
                    continue
                if name not in elders:
                    report.add(title, n, "elder_name",
                               f"护理计划里的「{name}」不在参保人表中，已跳过该行", "W")
                    continue
                for it in _items(item):
                    elders[name]["plan"].append({
                        "name": it,
                        "freq_per_month": _int(_cell(row, cols.get("freq_per_month"))),
                    })
                    n_plan += 1
            report.stats["plan_items"] = n_plan

    # ---- 护理人员
    if "caregivers" in sheets:
        title, rows = sheets["caregivers"]
        hdr_i, cols = _find_header(rows, report, title, "caregivers")
        if cols:
            for n, row in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
                name = _text(_cell(row, cols.get("caregiver_name")))
                if not name:
                    continue
                if name in caregivers:
                    report.add(title, n, "caregiver_name", f"护理人员「{name}」重复，保留首条", "W")
                    continue
                qual = _text(_cell(row, cols.get("qualification")))
                caregivers[name] = {"caregiver_id": name, "name": name, "qualification": qual}
            report.stats["caregivers"] = len(caregivers)

    # ---- 服务记录
    if "records" in sheets:
        title, rows = sheets["records"]
        hdr_i, cols = _find_header(rows, report, title, "records")
        if cols:
            for n, row in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
                elder_key = _text(_cell(row, cols.get("elder_name")))
                cg_key = _text(_cell(row, cols.get("caregiver_name")))
                if not any(_text(v) for v in row):
                    continue  # 空行
                if not elder_key:
                    report.add(title, n, "elder_name", "缺参保人姓名，该行未导入", "E")
                    continue
                if elder_key not in elders:
                    report.add(title, n, "elder_name",
                               f"参保人「{elder_key}」不在参保人表中，该行未导入", "E")
                    continue
                if cg_key and cg_key not in caregivers:
                    # 记录里出现但人员表没有：补一条空资质的人员，让规则引擎报资质问题
                    caregivers[cg_key] = {"caregiver_id": cg_key, "name": cg_key, "qualification": ""}
                    report.add(title, n, "caregiver_name",
                               f"护理人员「{cg_key}」不在人员表中，已补录但资质为空", "W")
                rec = {
                    "record_no": _text(_cell(row, cols.get("record_no"))) or f"{title}!{n}",
                    "elder_id": elder_key,
                    "caregiver_id": cg_key or "",
                    "service_type": _text(_cell(row, cols.get("service_type")))
                                    or elders[elder_key]["service_type"],
                    "service_address": _text(_cell(row, cols.get("address")))
                                       or elders[elder_key]["address"],
                    "service_date": _date(_cell(row, cols.get("service_date"))),
                    "start_time": _time(_cell(row, cols.get("start_time"))),
                    "end_time": _time(_cell(row, cols.get("end_time"))),
                    "duration_min": _int(_cell(row, cols.get("duration_min"))),
                    "items": _items(_cell(row, cols.get("items"))),
                    "item_count": _int(_cell(row, cols.get("item_count"))),
                    "caregiver_sign": _sign(_cell(row, cols.get("caregiver_sign"))),
                    "elder_sign": _sign(_cell(row, cols.get("elder_sign"))),
                    "abnormal_note": _text(_cell(row, cols.get("abnormal_note"))),
                    "remedy": _text(_cell(row, cols.get("remedy"))),
                    "remedy_time": _text(_cell(row, cols.get("remedy_time"))),
                    "submit_time": _text(_cell(row, cols.get("submit_time"))),
                    "source": {"sheet": title, "row": n},
                }
                if not rec["items"]:
                    report.add(title, n, "items", "服务项目为空，该行仍导入（会报项目数不符）", "W")
                if rec["item_count"] is None:
                    rec["item_count"] = len(rec["items"])
                records.append(rec)
            report.stats["records"] = len(records)

    wb.close()
    case = {
        "case_id": Path(path).stem,
        "source_file": Path(path).name,
        "elders": list(elders.values()),
        "caregivers": list(caregivers.values()),
        "records": records,
    }
    return case, report


def _find_header(rows: list[list], report: ImportReport, title: str, kind: str) -> tuple[int, dict[str, int]]:
    """找表头行：前 10 行里匹配字段最多的那一行（至少命中 1 个必需字段）。

    不要求命中 ≥2 个字段——单列台账（只有「参保人姓名」一列）也是合法输入。
    """
    best_i, best_cols = -1, {}
    for i, row in enumerate(rows[:10]):
        cols = _build_header_index(row)
        if len(cols) > len(best_cols):
            best_i, best_cols = i, cols
    if not best_cols or not any(f in best_cols for f in REQUIRED[kind]):
        report.add(title, None, None,
                   f"找不到表头行（前 10 行里没匹配到必需字段「{FIELD_ALIASES[REQUIRED[kind][0]][0]}」）", "E")
        return -1, {}
    for f in REQUIRED[kind]:
        if f not in best_cols:
            report.add(title, best_i + 1, f,
                       f"缺必需列「{FIELD_ALIASES[f][0]}」（可接受的写法：{'/'.join(FIELD_ALIASES[f][:3])}）", "E")
    return best_i, best_cols


# ---------------------------------------------------------------- 模板导出

TEMPLATE_SPEC = {
    "参保人": ("elders", ["参保人姓名", "身份证号", "护理方式", "失能等级", "居住地址"]),
    "护理服务计划": ("plan", ["参保人姓名", "服务项目", "月频次"]),
    "护理人员": ("caregivers", ["护理人员姓名", "护理人员资质"]),
    "服务记录": ("records", ["服务编号", "参保人姓名", "护理人员姓名", "服务类型", "服务地址",
                          "服务日期", "开始时间", "结束时间", "时长（分钟）", "服务项目",
                          "服务项目数合计", "护理人员签字", "参保人员（家属）签名",
                          "异常情况说明", "处理措施", "处理时间", "录入时间"]),
}

GUIDE = [
    ["长护险服务记录导入模板 · 填写说明"],
    [""],
    ["1. 本模板按官方《长期护理保险服务管理文书（2026 年版）》的四张表组织，四个工作表名不要改。"],
    ["2. 服务项目一栏可写多个，用「、」或「；」分隔；也可直接写项目编号（如 SH01、YL03）。"],
    ["3. 服务日期写 2026-10-05 或 2026/10/5；开始/结束时间写 08:30。"],
    ["4. 签字栏写「是」或「√」表示已签，留空或「否」表示未签。"],
    ["5. 参保人姓名是各表之间的关联键，四张表里的写法必须一致。"],
    ["6. 导入后会自动跑一致性核验，异常清单按 E（硬违规）/ W（需复核）/ I（数据质量）分级。"],
    [""],
    ["36 项服务项目清单如下（可直接复制到表格里）："],
]


def write_template(path: str | Path, sample: str | Path | None = None) -> Path:
    """生成空白导入模板；给了 sample 案例 JSON 则把示例数据一并写进去。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    wb.remove(wb.active)

    guide = wb.create_sheet("填写说明")
    for row in GUIDE:
        guide.append(row)
    guide.append([""])
    guide.append(["编号", "类别", "项目名称"])
    for item in CATALOG["items"]:
        guide.append([item["code"], item["category"], item["name"]])
    guide.column_dimensions["A"].width = 12
    guide.column_dimensions["B"].width = 16
    guide.column_dimensions["C"].width = 46
    guide["A1"].font = Font(bold=True, size=13)

    data = None
    if sample:
        data = json.loads(Path(sample).read_text(encoding="utf-8"))

    for sheet_name, (_kind, headers) in TEMPLATE_SPEC.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, width in enumerate([16, 18, 16, 14, 30, 14, 12, 12, 14, 34, 14, 14, 20, 26, 22, 14, 20], start=1):
            if i <= len(headers):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
        if data:
            _fill_sample(ws, sheet_name, headers, data)

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def _fill_sample(ws, sheet_name: str, headers: list[str], data: dict) -> None:
    elders = {e["elder_id"]: e for e in data.get("elders", [])}
    caregivers = {c["caregiver_id"]: c for c in data.get("caregivers", [])}
    if sheet_name == "参保人":
        for e in data.get("elders", []):
            ws.append([e.get("name"), e.get("id_card", ""), e.get("service_type", ""),
                       e.get("disability_level", ""), e.get("address", "")])
    elif sheet_name == "护理服务计划":
        for e in data.get("elders", []):
            for p in e.get("plan", []):
                ws.append([e.get("name"), p.get("name"), p.get("freq_per_month")])
    elif sheet_name == "护理人员":
        for c in data.get("caregivers", []):
            ws.append([c.get("name"), c.get("qualification", "")])
    elif sheet_name == "服务记录":
        for r in data.get("records", []):
            ws.append([
                r.get("record_no"), elders.get(r.get("elder_id"), {}).get("name", r.get("elder_id")),
                caregivers.get(r.get("caregiver_id"), {}).get("name", r.get("caregiver_id")),
                r.get("service_type"), r.get("service_address"), r.get("service_date"),
                r.get("start_time"), r.get("end_time"), r.get("duration_min"),
                "、".join(r.get("items", [])), r.get("item_count"),
                "是" if r.get("caregiver_sign") else "否",
                "是" if r.get("elder_sign") else "否",
                r.get("abnormal_note", ""), r.get("remedy", ""), r.get("remedy_time", ""),
                r.get("submit_time", ""),
            ])


# ---------------------------------------------------------------- 异常清单导出

SEVERITY_LABEL = {"E": "硬违规", "W": "需复核", "I": "数据质量"}


def write_report(path: str | Path, result: dict, report: ImportReport | None = None) -> Path:
    """把核验结果写成 Excel 异常清单：机构拿到就能按行整改，每行带条款依据。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "异常清单"
    headers = ["严重度", "规则", "对象", "问题描述", "文书依据", "证据", "整改状态"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    colours = {"E": "FFC7CE", "W": "FFEB9C", "I": "DDEBF7"}
    for f in result.get("findings", []):
        ev = f.get("evidence", {})
        # 记录级问题看服务编号；按月/按人的问题（如频次）回落到参保人，别留空
        subject = f.get("record_no") or ev.get("elder_id") or ev.get("date") or ""
        ws.append([
            SEVERITY_LABEL.get(f["severity"], f["severity"]),
            f["rule"], subject,
            f["message"], f["clause"],
            json.dumps(ev, ensure_ascii=False),
            "",
        ])
        ws.cell(row=ws.max_row, column=1).fill = PatternFill(
            "solid", fgColor=colours.get(f["severity"], "FFFFFF"))

    # 长文本必须换行，否则规则名和描述会挤到相邻列里（实机验证过）
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for letter, width in (("A", 9), ("B", 26), ("C", 14), ("D", 46), ("E", 38), ("F", 34), ("G", 10)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    # 横向 + 缩到一页宽 + 每页重复表头，机构打印出来才像一份清单
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    summary = wb.create_sheet("汇总")
    summary.append(["核验对象", result.get("records_checked", 0), "条服务记录"])
    summary.append(["硬违规", result["summary"]["hard_violations"], "条"])
    summary.append(["需复核", result["summary"]["needs_review"], "条"])
    summary.append(["数据质量", result["summary"]["data_quality"], "条"])
    if report:
        summary.append([])
        summary.append(["导入统计"])
        for k, v in report.stats.items():
            summary.append([k, v])
        if report.issues:
            summary.append([])
            summary.append(["导入问题", len(report.issues), "条（明细见「导入问题」工作表）"])
    summary.column_dimensions["A"].width = 20
    summary.column_dimensions["B"].width = 12

    if report and report.issues:
        iss = wb.create_sheet("导入问题")
        iss.append(["级别", "工作表", "行号", "字段", "说明"])
        for c in iss[1]:
            c.font = Font(bold=True)
        for i in report.issues:
            iss.append([SEVERITY_LABEL.get(i["level"], i["level"]), i["sheet"],
                        i["row"] or "", i["field"] or "", i["message"]])
        for letter, width in (("A", 10), ("B", 22), ("C", 8), ("D", 18), ("E", 62)):
            iss.column_dimensions[letter].width = width

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


# ---------------------------------------------------------------- CLI

def main(argv: list[str] | None = None) -> int:
    """python ltc_excel.py template <out.xlsx> [--sample case.json]
       python ltc_excel.py import <in.xlsx> [--case out.json]
       python ltc_excel.py review <in.xlsx> [--case out.json] [--report out.xlsx] [--strict]
    """
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print(__doc__)
        print(main.__doc__)
        return 2
    cmd = argv[0]
    if cmd == "template":
        out = argv[1]
        sample = argv[argv.index("--sample") + 1] if "--sample" in argv else None
        path = write_template(out, sample)
        print(f"模板已生成: {path}")
        return 0
    if cmd == "import":
        src = argv[1]
        case, report = load_case(src)
        if "--case" in argv:
            dest = Path(argv[argv.index("--case") + 1])
            dest.write_text(json.dumps(case, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            print(f"case 已写出: {dest}")
        print(json.dumps(report.as_dict(), ensure_ascii=False, indent=2))
        return 0 if not any(i["level"] == "E" for i in report.issues) else 1
    if cmd == "review":
        # 一步到位：导入 → 核验 → 出异常清单。机构实际会用的就是这条。
        import ltc_rules as R

        src = argv[1]
        case, report = load_case(src)
        result = R.audit(case)
        stem = Path(src).stem
        case_out = Path(argv[argv.index("--case") + 1]) if "--case" in argv \
            else Path(src).with_name(f"{stem}_case.json")
        case_out.write_text(json.dumps(case, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        rep_out = Path(argv[argv.index("--report") + 1]) if "--report" in argv \
            else Path(src).with_name(f"{stem}_异常清单.xlsx")
        write_report(rep_out, result, report)
        print(json.dumps({
            "case": str(case_out), "report": str(rep_out),
            "records_checked": result["records_checked"],
            "summary": result["summary"],
            "import_issue_count": len(report.issues),
        }, ensure_ascii=False, indent=2))
        # --strict 时，导入告警也算失败（给要接入 CI 的机构用）
        if "--strict" in argv:
            return 0 if (result["summary"]["hard_violations"] == 0
                         and not report.issues) else 1
        return 0 if result["summary"]["hard_violations"] == 0 else 1
    print(f"未知命令: {cmd}")
    return 2


def demo() -> None:
    """自检：模板 → 回读，结构与统计必须自洽。"""
    import tempfile

    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "t.xlsx"
        write_template(p, ROOT / "tests" / "fixtures" / "clean_case.json")
        case, report = load_case(p)
        assert case["elders"], "参保人未导入"
        assert case["caregivers"], "护理人员未导入"
        assert case["records"], "服务记录未导入"
        assert all(r["items"] for r in case["records"]), f"项目未解析: {case['records'][0]}"
        assert not [i for i in report.issues if i["level"] == "E"], report.issues
        # 时段必须还原成 HH:MM
        for r in case["records"]:
            assert re.fullmatch(r"\d{2}:\d{2}", r["start_time"]), r["start_time"]
            assert re.fullmatch(r"\d{2}:\d{2}", r["end_time"]), r["end_time"]
            assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", r["service_date"]), r["service_date"]
    print("demo ok")


if __name__ == "__main__":
    if len(sys.argv) == 1:
        demo()
        raise SystemExit(0)
    raise SystemExit(main())
