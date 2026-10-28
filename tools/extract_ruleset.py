"""从官方 PDF《长期护理保险服务管理文书（2026 年版）》抽取结构化规则数据。

输入：data/source/nhsa-service-forms-2026.pdf（国家医保局办公室印发，来源 nhsa.gov.cn）
输出：data/derived/*.json

设计原则：可重复运行，每次运行都做自校验（数量、序号连续性、分区求和、三处副本一致），
校验失败直接非零退出，不产出半成品。
"""
import json
import re
import sys
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[1]
PDF = ROOT / "data" / "source" / "nhsa-service-forms-2026.pdf"
OUT = ROOT / "data" / "derived"

# PDF 内页码（1-based），经人工核对确定
PAGE_SUGGESTION = (12, 13)   # 表5 护理服务建议（36 项带编号）
PAGE_CARE_PLAN = (18, 19)    # 表2 护理服务计划表（带 □ 前缀）
PAGE_MONTHLY = (23, 24, 25)  # 表4 月度汇总表（数字独占行）
PAGE_RECORD_FORM = 21        # 表3 服务记录单
PAGE_INSPECTION = range(36, 40)  # 表3 巡查记录表
PAGE_QUALITY = range(30, 35)     # 表2 质量评价表（表1/表2/表3 三张）

LIFE, MED = 20, 16
FAILS = []


def fail(msg):
    FAILS.append(msg)
    print(f"  !! {msg}")


def doc():
    return pymupdf.open(PDF)


def lines_of(d, pages, drop_page_no=True):
    out = []
    for p in pages:
        ls = [x.strip() for x in d[p - 1].get_text().splitlines() if x.strip()]
        if drop_page_no and ls and re.fullmatch(r"\d{1,3}", ls[0]):
            ls = ls[1:]
        out.extend(ls)
    return out


STOP = ("合计项目数", "项目数合计", "月合计时长", "评估专家签字", "本人已充分知悉",
        "参保人或监护人", "长护专员签字", "定点长护机构", "护理人员签名", "护士签名", "注：")


def parse_numbered(lines, prefix_re):
    """`1.名称` / `□1.名称` 起项；续行并入名称（处理括号换行）。"""
    life, med, cur, buf = [], [], None, None

    def flush():
        nonlocal buf
        if buf:
            (cur if cur is not None else life).append(buf)
            buf = None

    for line in lines:
        if "生活照护类" in line:
            flush()
            cur = life
            continue
        if "医疗护理类" in line:
            flush()
            cur = med
            continue
        if cur is None:
            continue
        if line.isdigit():
            continue
        m = re.match(prefix_re, line)
        if m:
            flush()
            buf = {"no": int(m.group(1)), "name": m.group(2)}
        elif buf is not None:
            if any(s in line for s in STOP):
                flush()
                continue
            buf["name"] += line
    flush()
    return life, med


def parse_pairwise(lines):
    """月度汇总表：序号独占行 + 名称行。以"协助进食"为锚点，只接受连续序号。"""
    anchor = next(i for i, x in enumerate(lines) if x.startswith("协助进食"))
    i, pairs = anchor - 1, []
    while i + 1 < len(lines):
        num, name = lines[i], lines[i + 1]
        if not num.isdigit() or name.isdigit():
            i += 1
            continue
        if any(s in name for s in ("签名", "注：")):
            break
        if int(num) != len(pairs) + 1:
            i += 1
            continue
        pairs.append({"no": int(num), "name": name})
        i += 2
        while i < len(lines) and not lines[i].isdigit():
            if any(s in lines[i] for s in ("签名", "注：")):
                i = len(lines)
                break
            pairs[-1]["name"] += lines[i]
            i += 1
    life = [p for p in pairs if p["no"] <= LIFE]
    med = [{"no": p["no"] - LIFE, "name": p["name"]} for p in pairs if p["no"] > LIFE]
    return life, med


def build_catalog(d):
    """36 项服务目录。三处副本交叉校验，以服务建议表为准。"""
    print("[1/4] 服务项目目录")
    copies = {
        "服务建议表(表5)": parse_numbered(lines_of(d, PAGE_SUGGESTION), r"^(\d{1,2})[.．](.+)$"),
        "护理计划表(表2)": parse_numbered(lines_of(d, PAGE_CARE_PLAN), r"^□(\d{1,2})[.．](.+)$"),
        "月度汇总表(表4)": parse_pairwise(lines_of(d, PAGE_MONTHLY)),
    }

    base_life, base_med = copies["服务建议表(表5)"]
    for label, (life, med) in copies.items():
        ok_l = len(life) == LIFE and [x["no"] for x in life] == list(range(1, LIFE + 1))
        ok_m = len(med) == MED and [x["no"] for x in med] == list(range(1, MED + 1))
        print(f"  {label}: 生活 {len(life)}/{LIFE}{'✓' if ok_l else '✗'} 医疗 {len(med)}/{MED}{'✓' if ok_m else '✗'}")
        if not (ok_l and ok_m):
            fail(f"{label} 数量或序号不连续")

    for label, (life, med) in copies.items():
        if label.startswith("服务建议"):
            continue
        for tag, base, got in (("生活", base_life, life), ("医疗", base_med, med)):
            for b, g in zip(base, got):
                if b["name"] != g["name"]:
                    fail(f"{label} {tag}#{b['no']} 名称不一致: {b['name']!r} vs {g['name']!r}")

    catalog = []
    for item in base_life:
        catalog.append({"code": f"SH{item['no']:02d}", "seq": item["no"],
                        "category": "生活照护类", "name": item["name"],
                        "clause": "国家长期护理保险服务项目目录（试行）医保发〔2025〕24号 生活照护类"})
    for item in base_med:
        catalog.append({"code": f"YL{item['no']:02d}", "seq": LIFE + item["no"],
                        "category": "医疗护理类", "name": item["name"],
                        "clause": "国家长期护理保险服务项目目录（试行）医保发〔2025〕24号 医疗护理类"})
    print(f"  → 合计 {len(catalog)} 项（生活 {len(base_life)} + 医疗 {len(base_med)}）")
    return catalog


INSPECTION_SECTIONS = [
    ("机构管理", 27), ("服务管理", 20), ("费用结算", 10), ("信息管理", 15),
    ("档案管理", 6), ("风险防范", 4), ("违法违规行为查处", 8), ("综合评价", 10),
]


def build_inspection(d):
    """巡查记录表：8 分区 30 检查项，满分 100。"""
    print("[2/4] 巡查记录表")
    lines = lines_of(d, PAGE_INSPECTION)
    text = "\n".join(lines)

    declared = [int(x) for x in re.findall(r"（(\d+)分）", text)]
    print(f"  分区声明: {declared} 求和 {sum(declared)}")
    if sum(declared) != 100:
        fail(f"巡查表分区求和不等于 100，实为 {sum(declared)}")
    if declared != [t for _, t in INSPECTION_SECTIONS]:
        fail("巡查表分区顺序或分值与本脚本常量不符")

    # 按分区标题切段，段内按行分组：内容行累积 → 遇到纯数字分值行即结项
    positions = []
    for name, total in INSPECTION_SECTIONS:
        pat = r"\s*".join(map(re.escape, name))
        m = re.search(pat, text)
        if not m:
            fail(f"分区未找到: {name}")
            continue
        positions.append((m.start(), name, total))
    positions.sort()

    sections = []
    for k, (start, name, total) in enumerate(positions):
        end = positions[k + 1][0] if k + 1 < len(positions) else len(text)
        chunk = re.sub(r"（\d+分）", "", text[start:end], count=1)
        # 段落开头就是分区标题（可能跨行），去掉它，只留检查内容
        chunk = re.sub(r"^\s*" + r"\s*".join(map(re.escape, name)), "", chunk, count=1)
        items, buf = [], ""
        for line in chunk.splitlines():
            s = line.strip()
            if not s:
                continue
            if re.fullmatch(r"\d{1,2}", s):
                if buf:
                    items.append({"content": buf, "score": int(s)})
                    buf = ""
                    if sum(i["score"] for i in items) == total:
                        break
                continue
            if buf and sum(i["score"] for i in items) >= total:
                break
            buf += s
        got = sum(i["score"] for i in items)
        print(f"  {name}: {len(items)} 项 累加 {got}/{total} {'✓' if got == total else '✗'}")
        if got != total:
            fail(f"{name} 分项累加 {got} != 声明 {total}")
        sections.append({"section": name, "max_score": total, "items": items})

    total_items = sum(len(s["items"]) for s in sections)
    print(f"  → {len(sections)} 分区 / {total_items} 检查项 / 满分 {sum(s['max_score'] for s in sections)}")
    return {"title": "定点长护服务机构巡查记录表", "max_score": 100, "sections": sections}


def build_quality(d):
    """质量评价表：表1 生活照护 20 项 / 表2 医疗护理 16 项 / 表3 其他类 5 项，每项 0-5 分，得分率 ≥60% 合格。"""
    print("[3/4] 质量评价表")
    lines = lines_of(d, PAGE_QUALITY)
    text = "\n".join(lines)

    def parse_table(start_pat, seq_max):
        start = next((i for i, x in enumerate(lines) if re.search(start_pat, x)), None)
        if start is None:
            fail(f"未找到表起点: {start_pat}")
            return []
        items, i = [], start + 1
        while i < len(lines) and len(items) < seq_max:
            s = lines[i]
            if s.isdigit() and int(s) == len(items) + 1 and i + 1 < len(lines):
                name = lines[i + 1]
                detail = []
                i += 2
                while i < len(lines) and not (lines[i].isdigit() and lines[i].isdigit()
                                              and int(lines[i]) == len(items) + 2):
                    if re.match(r"^\(?[ABC]\)?小计", lines[i]) or lines[i].startswith("表"):
                        break
                    detail.append(lines[i])
                    i += 1
                items.append({"no": len(items) + 1, "name": name, "points": detail})
                continue
            i += 1
        return items

    life = parse_table(r"表1\s*[:：]\s*生活照护类", LIFE)
    med = parse_table(r"表2\s*[:：]\s*医疗护理类", MED)
    other = parse_table(r"表3\s*[:：]\s*其他类", 5)

    for label, got, want in (("生活照护", len(life), LIFE), ("医疗护理", len(med), MED), ("其他类", len(other), 5)):
        print(f"  {label}: {got}/{want} {'✓' if got == want else '✗'}")
        if got != want:
            fail(f"质量评价表 {label} 项数 {got} != {want}")

    if "60%" not in text:
        fail("未在质量评价表中找到 60% 合格线")

    # 直判不合格条件
    hard_fail = []
    for s in lines:
        if "打骂参保人" in s and "不合格" in s:
            hard_fail.append({"condition": "出现打骂参保人行为", "effect": "直接判定不合格"})
        if "非持证人员提供医疗护理" in s:
            hard_fail.append({"condition": "非持证人员提供医疗护理", "effect": "直接判定不合格"})

    print(f"  → 合格线 60%（A/B/C 三类各自得分率均需 ≥60%）; 直判不合格条件 {len(hard_fail)} 条")
    return {
        "title": "长期护理服务质量评价表",
        "score_per_item": 5,
        "pass_rate": 0.60,
        "pass_rule": "A、B、C 三类得分率均 ≥60% 为合格，有一类低于 60% 即不合格",
        "hard_fail": hard_fail,
        "tables": [
            {"id": "A", "name": "生活照护类", "items": life},
            {"id": "B", "name": "医疗护理类", "items": med},
            {"id": "C", "name": "其他类", "items": other},
        ],
    }


RECORD_FIELDS = [
    ("record_no", "服务编号", True, "表头"),
    ("org_name", "服务机构", True, "表头"),
    ("service_type", "服务类型", True, "表头：居家护理/机构护理/社区护理"),
    ("elder_name", "参保人姓名", True, "服务对象"),
    ("elder_id_card", "身份证号", True, "服务对象"),
    ("disability_level", "失能等级", True, "服务对象"),
    ("elder_phone", "联系电话", False, "服务对象"),
    ("service_address", "服务地址", True, "服务对象"),
    ("caregiver_name", "护理人员姓名", True, "服务主体"),
    ("caregiver_qualification", "护理人员资质", True, "服务主体：长期照护师/护士/医生/医疗护理员/养老护理员/其他"),
    ("bed_no", "床号", False, "机构护理类填写"),
    ("service_date", "服务日期", True, "服务记录栏"),
    ("start_time", "服务开始时间", True, "服务记录栏"),
    ("end_time", "服务结束时间", True, "服务记录栏"),
    ("duration_min", "时长（分钟）", True, "服务记录栏"),
    ("items", "服务项目内容", True, "服务记录栏"),
    ("item_count", "服务项目数合计", True, "服务记录栏"),
    ("abnormal_note", "异常情况说明", False, "异常与确认"),
    ("remedy", "处理措施", False, "异常与确认"),
    ("remedy_time", "处理时间", False, "异常与确认"),
    ("caregiver_sign", "护理人员签字", True, "签字确认栏"),
    ("elder_sign", "参保人员（家属）签名", True, "签字确认栏"),
]


def build_record_schema(d):
    """服务记录单字段表；同时校验 PDF 里确实存在这些字段名。"""
    print("[4/4] 服务记录单字段")
    text = d[PAGE_RECORD_FORM - 1].get_text()
    missing = [cn for _, cn, _, _ in RECORD_FIELDS if cn.replace("（", "(").replace("）", ")") not in
               text.replace("（", "(").replace("）", ")")]
    print(f"  字段 {len(RECORD_FIELDS)} 个，PDF 中未匹配 {len(missing)} 个")
    for m in missing:
        fail(f"服务记录单字段未在 PDF 中找到: {m}")
    return {
        "title": "长期护理保险服务记录单",
        "filling_rule": "一服务一记录，按次填写；记录服务开始时间、结束时间、实际服务项目与服务时长",
        "fields": [{"key": k, "label": cn, "required": req, "group": grp}
                   for k, cn, req, grp in RECORD_FIELDS],
    }


def main():
    if not PDF.exists():
        print(f"缺少源文件: {PDF}")
        return 2
    d = doc()
    print(f"源 PDF: {PDF.name}（{d.page_count} 页）\n")

    catalog = build_catalog(d)
    inspection = build_inspection(d)
    quality = build_quality(d)
    record = build_record_schema(d)

    if FAILS:
        print(f"\n校验失败 {len(FAILS)} 项，不写出产物：")
        for f in FAILS:
            print(f"  - {f}")
        return 1

    OUT.mkdir(parents=True, exist_ok=True)
    payloads = {
        "services_catalog.json": {
            "source": "国家长期护理保险服务项目目录（试行），医保发〔2025〕24号",
            "source_note": "官方目录附件为图片版，项目代码未在本次可解析的 PDF 中出现；code 字段为本文档自编的稳定标识（SH=生活照护, YL=医疗护理），seq 为官方序号",
            "count": len(catalog), "items": catalog,
        },
        "inspection_checklist.json": inspection,
        "quality_review.json": quality,
        "service_record_schema.json": record,
    }
    for name, payload in payloads.items():
        path = OUT / name
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"\n写出 {path.relative_to(ROOT)}（{path.stat().st_size:,} 字节）")

    print("\n全部校验通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
