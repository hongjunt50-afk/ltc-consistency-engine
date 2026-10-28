"""生成自包含的静态演示页：把示例台账跑一遍，结果渲染成 HTML。

用途：GitHub Pages 首页。访客不用装 Python，直接看到"输入 Excel → 输出异常清单"的效果。
"""
from __future__ import annotations

import html
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import ltc_excel as X  # noqa: E402
import ltc_rules as R  # noqa: E402

SEV = {"E": ("硬违规", "#d93025", "#fce8e6"),
       "W": ("需复核", "#b26a00", "#fff4e5"),
       "I": ("数据质量", "#1a73e8", "#e8f0fe")}

CSS = """
*{box-sizing:border-box}
body{margin:0;font:16px/1.65 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif;
     color:#1f2328;background:#f6f8fa}
.wrap{max-width:1080px;margin:0 auto;padding:0 24px 64px}
header{background:linear-gradient(135deg,#1f3a5f,#2c6e9b);color:#fff;padding:56px 0 48px;margin-bottom:32px}
header .wrap{padding-bottom:0}
h1{font-size:30px;margin:0 0 12px;font-weight:700}
.sub{opacity:.9;font-size:17px;margin:0}
.lead{background:#fff;border:1px solid #d0d7de;border-radius:10px;padding:24px 28px;margin-bottom:28px}
h2{font-size:20px;margin:36px 0 14px;padding-bottom:8px;border-bottom:1px solid #d0d7de}
.cards{display:flex;gap:16px;flex-wrap:wrap;margin:20px 0}
.card{flex:1 1 200px;background:#fff;border:1px solid #d0d7de;border-radius:10px;padding:18px 20px}
.card .n{font-size:32px;font-weight:700;line-height:1.1}
.card .l{font-size:13px;color:#57606a;margin-top:6px}
table{width:100%;border-collapse:collapse;background:#fff;font-size:14px;
      border:1px solid #d0d7de;border-radius:8px;overflow:hidden}
th{background:#eef2f6;text-align:left;padding:10px 12px;font-weight:600;border-bottom:1px solid #d0d7de}
td{padding:10px 12px;border-bottom:1px solid #eaeef2;vertical-align:top}
tr:last-child td{border-bottom:none}
.badge{display:inline-block;padding:2px 9px;border-radius:11px;font-size:12px;font-weight:600;white-space:nowrap}
.rule{font-family:ui-monospace,Consolas,monospace;font-size:12.5px;color:#57606a}
.clause{color:#57606a;font-size:13px}
code{background:#eff2f5;padding:1px 6px;border-radius:4px;font-size:13px}
pre{background:#0d1117;color:#e6edf3;padding:18px 20px;border-radius:10px;overflow-x:auto;font-size:13.5px;line-height:1.6}
footer{color:#57606a;font-size:13px;margin-top:48px;padding-top:20px;border-top:1px solid #d0d7de}
.flow{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:18px 0;font-size:14px}
.flow span{background:#fff;border:1px solid #d0d7de;border-radius:20px;padding:7px 15px}
.flow b{color:#8b949e}
"""


def cell(s: str) -> str:
    return html.escape(s, quote=False)


def badge(sev: str) -> str:
    label, fg, bg = SEV[sev]
    return f'<span class="badge" style="color:{fg};background:{bg}">{label}</span>'


def build_html() -> str:
    case, report = X.load_case(ROOT / "examples" / "示例台账.xlsx")
    result = R.audit(case)
    s = result["summary"]
    total = len(result["findings"])

    rows = []
    for f in result["findings"]:
        ev = f.get("evidence", {})
        subject = f.get("record_no") or ev.get("elder_id") or ev.get("date") or ""
        rows.append(
            "<tr>"
            f"<td>{badge(f['severity'])}</td>"
            f"<td class='rule'>{cell(f['rule'])}</td>"
            f"<td>{cell(str(subject))}</td>"
            f"<td>{cell(f['message'])}</td>"
            f"<td class='clause'>{cell(f['clause'])}</td>"
            "</tr>")

    catalog = R.CATALOG["items"]
    life = [i for i in catalog if i["category"] == "生活照护类"]
    med = [i for i in catalog if i["category"] == "医疗护理类"]

    insp = R.INSPECTION
    qual = R.QUALITY
    insp_rows = "".join(
        f"<tr><td>{cell(s['section'])}</td><td>{len(s['items'])}</td><td>{s['max_score']}</td></tr>"
        for s in insp["sections"])
    qual_rows = "".join(
        f"<tr><td>{t['id']}</td><td>{cell(t['name'])}</td><td>{len(t['items'])}</td>"
        f"<td>{len(t['items']) * qual['score_per_item']}</td></tr>"
        for t in qual["tables"])

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>长护险服务过程一致性核验引擎 · 演示</title>
<style>{CSS}</style>
</head>
<body>
<header><div class="wrap">
  <h1>长护险服务过程一致性核验引擎</h1>
  <p class="sub">把 2026 版官方文书编译成可执行、可解释、可核对的规则 · 每条结论都指向具体条款</p>
</div></header>

<div class="wrap">

<div class="lead">
  <strong>这个页面在演示什么。</strong>
  机构交来的服务记录、护理服务计划、人员资质放在一张 Excel 里，本工具逐条比对，
  把"服务项目不在国家 36 项目录""同一护理员同一时段两条记录""非持证人员做了医疗护理"
  这类问题挑出来，按硬违规 / 需复核 / 数据质量分级，并给出依据的文书条款。
  下面是仓库自带<b>合成示例台账</b>的真实运行结果（人物、机构、地址均为虚构）。
</div>

<div class="flow">
  <span>机构 Excel 台账</span><b>→</b><span>字段归一与别名匹配</span><b>→</b>
  <span>15 条一致性规则</span><b>→</b><span>异常清单 + 条款依据</span>
</div>

<div class="cards">
  <div class="card"><div class="n">{result['records_checked']}</div><div class="l">条服务记录</div></div>
  <div class="card"><div class="n" style="color:#d93025">{s['hard_violations']}</div><div class="l">硬违规（E）</div></div>
  <div class="card"><div class="n" style="color:#b26a00">{s['needs_review']}</div><div class="l">需复核（W）</div></div>
  <div class="card"><div class="n" style="color:#1a73e8">{s['data_quality']}</div><div class="l">数据质量（I）</div></div>
</div>

<h2>异常清单（{total} 条）</h2>
<table>
<thead><tr><th style="width:88px">严重度</th><th style="width:190px">规则</th>
<th style="width:92px">对象</th><th>问题描述</th><th style="width:32%">文书依据</th></tr></thead>
<tbody>
{''.join(rows)}
</tbody>
</table>

<h2>规则不是自己编的</h2>
<p>全部来自《长期护理保险服务管理文书（2026 年版）》与《国家长期护理保险服务项目目录（试行）》（医保发〔2025〕24 号）。
抽取脚本从官方 PDF 直接读，并在抽取时做四类自校验，任一失败就不产出文件：</p>
<ul>
  <li><b>数量</b>：36 项 = 生活照护 {len(life)} + 医疗护理 {len(med)}</li>
  <li><b>序号连续</b>：1..20 与 1..16 不许断号</li>
  <li><b>三处副本逐字一致</b>：同样 36 项在「护理服务建议表」「护理服务计划表」「月度汇总表」各出现一次，三种排版三种解析路径，得到的名称必须完全相同</li>
  <li><b>分区求和</b>：巡查记录表 {len(insp['sections'])} 个分区的声明分值必须相加为 {insp['max_score']}，且分区内各检查项累加等于该区声明总分</li>
</ul>

<h2>另外两块能力</h2>
<p><b>巡查记录表评分</b>：{len(insp['sections'])} 个分区、{sum(len(x['items']) for x in insp['sections'])} 个检查项、满分 {insp['max_score']}。</p>
<table>
<thead><tr><th style="width:220px">分区</th><th style="width:100px">检查项</th><th style="width:100px">分区分值</th></tr></thead>
<tbody>{insp_rows}</tbody>
</table>
<p><b>质量评价算分</b>：三类各自得分率均 ≥ {int(qual['pass_rate'] * 100)}% 才合格；命中「打骂参保人」「非持证人员提供医疗护理」直接判不合格。</p>
<table>
<thead><tr><th style="width:70px">表</th><th>评价类别</th><th style="width:100px">项数</th><th style="width:100px">满分</th></tr></thead>
<tbody>{qual_rows}</tbody>
</table>

<h2>自己跑一遍</h2>
<pre>git clone &lt;本仓库&gt; &amp;&amp; cd ltc-consistency-engine
pip install openpyxl

# 一条命令：导入 → 核验 → 出异常清单
python ltc_excel.py review examples/示例台账.xlsx

# 拿空白模板（含 36 项目录对照表）
python ltc_excel.py template 导入模板.xlsx

# 跑测试
python -m unittest discover -s tests</pre>

<footer>
  仓库内两份示例数据全部为合成数据，不含任何真实患者信息。
  本项目不做失能等级评定，不涉及疾病诊断。
</footer>

</div>
</body>
</html>
"""


def main() -> int:
    out = ROOT / "docs" / "index.html"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(build_html(), encoding="utf-8")
    print(f"演示页已生成: {out}（{out.stat().st_size:,} 字节）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
