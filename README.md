# 长护险服务过程一致性核验引擎 · LTC Consistency Engine

[![tests](https://github.com/hongjunt50-afk/ltc-consistency-engine/actions/workflows/tests.yml/badge.svg)](https://github.com/hongjunt50-afk/ltc-consistency-engine/actions/workflows/tests.yml)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

把《长期护理保险服务管理文书（2026 年版）》和《国家长期护理保险服务项目目录（试行）》编译成**可执行、可解释、可核对**的核验规则：输入机构的服务记录、护理计划、人员资质，输出异常清单，每条结论都指向具体文书条款。

> 定位：定点长护服务机构**自查**与经办机构**稽核辅助**工具，不涉及疾病诊断，不做评估等级判定。
>
> 在线看效果（纯静态页，不用装 Python）：<https://hongjunt50-afk.github.io/ltc-consistency-engine/>

## 为什么规则不是自己编的

规则全部来自官方公开文书，`tools/extract_ruleset.py` 从 PDF 直接抽取并在抽取时就做自校验（数量、序号连续性、分区分值累加、三处副本逐字一致），校验失败不产出文件。

源文件 `data/source/nhsa-service-forms-2026.pdf`（国家医保局办公室印发，820 KB / 50 页）不入版本库，重新获取：

```bash
curl -sL -o data/source/nhsa-service-forms-2026.pdf -A "Mozilla/5.0" \
  "https://www.nhsa.gov.cn/module/download/downfile.jsp?classid=0&filename=2f2918a59a0f46879de862e83ad5fbcc.pdf"
```

抽取产物 `data/derived/*.json` 已入库，因此 `ltc_rules.py` 和测试不依赖 PDF，克隆即可运行。

## 和已有项目的关系

不宣称首创，只说清楚自己填的是哪块：

| 环节 | 现状 | 本项目 |
|---|---|---|
| 失能等级评估 | GB/T 42195-2022 已有多个开源实现（如 `zzx749/shineng`），覆盖 26 个二级指标与五档判定 | **不做**，也不输出任何等级结论 |
| 机构侧业务系统 | 商业侧有成熟的机构管理系统在跑，开源侧也有若干工单/结算实现 | **不做**，本项目不碰业务流转 |
| 医保基金稽核 | 有面向结算数据的智能稽核项目，处理的是**基金支付侧** | 处理**服务过程侧**：记录、护理计划、人员资质三者是否自洽 |
| 服务过程质控 | GitHub 搜「长护险 质控」「护理 质控 稽核」「long-term care fraud detection」均为 0 结果 | 本项目要填的就是这块 |

## 三块能力

| 能力 | 输入 | 输出 | 政策依据 |
|---|---|---|---|
| 服务过程一致性核验 | 服务记录单 + 护理服务计划 + 人员资质 | 按 E（硬违规）/ W（需复核）/ I（数据质量）分级的异常清单，含条款引用与证据字段 | 服务记录单、护理服务计划表、36 项目录 |
| 巡查记录表评分 | 8 分区 30 项的实得分 | 分区得分、总分、适用分值与校验错误 | 定点长护服务机构巡查记录表（满分 100） |
| 质量评价表算分 | A/B/C 三类逐项 0–5 分 | 各类得分率、合格判定、直接不合格命中 | 长期护理服务质量评价表（三类得分率均 ≥60% 合格） |

## Excel 导入适配器

机构不会按你的 schema 交数据，会用自己攒的 Excel。`ltc_excel.py` 负责把"实际在用的表"翻译成引擎能吃的结构：

```bash
# 一条命令走完：导入 → 核验 → 出异常清单
python ltc_excel.py review 机构台账.xlsx
#   → 机构台账_case.json          （规整后的数据，可复核）
#   → 机构台账_异常清单.xlsx      （按行可整改，每条带文书依据）

# 先拿模板
python ltc_excel.py template 导入模板.xlsx            # 空白模板 + 36 项目录对照表
python ltc_excel.py template 示例.xlsx --sample tests/fixtures/sample_case.json  # 带示例数据
```

**能容错到什么程度**（都有测试锁着）：

| 情况 | 处理 |
|---|---|
| 表头写法不同 | 每字段一组别名，如参保人姓名 = 姓名/老人姓名/服务对象姓名；人员资质 = 资质/职业资格/职业 |
| 表名不同 | 「服务记录单」「护理记录」「服务明细」都识别；带「参保人」「护理人员」字样的表按最长别名判定归属 |
| 第一行是大标题 | 前 10 行里找匹配字段最多的一行作表头 |
| 一格写多个项目 | 按 `、；;，,` 拆分，也认 `SH01`/`YL03` 这类编号 |
| 日期写法多样 | `2026-10-05`、`2026/10/5`、`2026.10.05`、`2026年10月5日`、`20261005`、Excel 日期型 |
| 时间被存成小数 | openpyxl 读纯时间格得到 0~1 的小数，按 24 小时还原成 `HH:MM` |
| 签字栏写法 | `是/√/✓/有/已签` 为已签，`否/×/无/空` 为未签，手写姓名也算已签 |
| 记录里出现人员表没有的人 | 补录进人员表并把资质留空，让资质校验生效，同时报 W |
| 身份证被存成数字 | 18 位数字在 Excel 里会丢末几位，检测到就报 W 提示改文本格式，不装作没看见 |
| 没填项目数合计 | 按实际列出的项目数补，避免整表都报数据质量 |
| 缺表 / 缺必需列 | 报 E 并指明「哪张表、第几行、哪个字段、可写哪几种表头」 |

导入报告和数据一起返回，**不静默丢行**：每一处跳过（参保人对不上、人员不在名录、表头缺列）都带工作表名和行号。

## 30 秒看到效果

不用准备任何数据，仓库里带了一份合成的示例台账：

```bash
pip install openpyxl
python ltc_excel.py review examples/示例台账.xlsx
```

输出两个文件：`示例台账_case.json`（规整后的数据）和 `示例台账_异常清单.xlsx`（异常清单，按行可整改）。示例数据里埋了代打卡时段重叠、非持证人员做医疗护理、服务项目不在 36 项目录、异常无处理措施等问题，共命中 9 条硬违规 / 9 条需复核 / 1 条数据质量。

## 快速开始

```bash
git clone https://github.com/hongjunt50-afk/ltc-consistency-engine.git
cd ltc-consistency-engine

# 1. 从官方 PDF 抽取规则（自校验，失败不产出）
python tools/extract_ruleset.py

# 2. 一致性核验
python ltc_rules.py audit tests/fixtures/sample_case.json

# 3. 跑测试
python -m unittest discover -s tests -v
```

依赖：Python 3.10+。`ltc_rules.py` 与测试零第三方依赖（只读 JSON）；`tools/extract_ruleset.py` 需要 `pymupdf`；`ltc_excel.py` 需要 `openpyxl`。

```bash
pip install pymupdf openpyxl   # 仅在要用抽取和 Excel 导入时需要
```

## 一致性核验规则清单

| 规则 | 级别 | 判据 | 依据 |
|---|---|---|---|
| `E_UNKNOWN_SERVICE` | E | 服务项目不在国家 36 项目录内 | 目录凡例·三（准入法管理，未列入不予支付） |
| `E_BAD_TIME` | E | 结束时间不晚于开始时间 | 服务记录单·服务记录栏 |
| `E_DURATION_MISMATCH` | E | 填报时长与起止时间推算不符（默认容差 1 分钟） | 服务记录单·服务记录栏 |
| `E_QUAL_MISMATCH` | E | 非持证人员（资质不属于长期照护师/护士/医生/医疗护理员/养老护理员，或未填）执行医疗护理类项目 | 服务记录单·护理人员资质；质量评价表·其他类第 1 项（非持证人员提供医疗护理直接判不合格） |
| `E_NO_SIGN` | E | 缺护理人员签字或缺参保人（家属）签名 | 服务记录单·签字确认栏（一服务一记录） |
| `E_ABNORMAL_NO_REMEDY` | E | 记录了异常但缺处理措施或处理时间 | 服务记录单·异常情况说明/处理措施/处理时间 |
| `E_CAREGIVER_OVERLAP` | E | 同一护理员同日两段服务时段重叠 | 服务记录单；巡查表·费用结算（不得重复收取） |
| `E_ELDER_OVERLAP` | E | 同一参保人同时段两条服务记录 | 服务记录单·一服务一记录 |
| `W_QUAL_SCOPE_REVIEW` | W | 养老护理员/医疗护理员执行医疗护理类项目，需核对当地执业范围口径 | 服务记录单·护理人员资质；巡查表·机构管理（人员类型与服务范围相匹配） |
| `W_OFF_PLAN` | W | 服务项目不在该参保人护理计划内 | 护理服务计划表；巡查表·服务管理 |
| `W_SERVICE_TYPE_MISMATCH` | W | 记录服务类型与备案护理方式不符 | 服务记录单·服务类型；失能人员档案信息表 |
| `W_ADDRESS_MISMATCH` | W | 居家护理上门地址与备案居住地址不符 | 服务记录单·服务地址；档案信息表·居住地址 |
| `W_LATE_ENTRY` | W | 录入滞后服务日期超过阈值（默认 3 天） | 服务记录单（按次如实填写）；月度汇总表 |
| `W_EXCESS_DAILY` | W | 单日累计服务超阈值（默认 480 分钟） | 目录·四（各地应测算确定时长与频次） |
| `W_FREQ_SHORT` / `W_FREQ_OVER` | W | 月度执行频次低于/超出计划频次 | 护理服务计划表·确定频次；月度汇总表 |
| `I_ITEM_COUNT_MISMATCH` | I | 填报项目数与实际列出项数不符 | 服务记录单·服务项目数合计 |

阈值全部显式入参（`day_max_minutes` / `late_entry_days` / `duration_tolerance`），按地区政策调整不用改代码。

## 用例数据

`tests/fixtures/` 下两份**合成**数据（人物、机构、地址均虚构）：

- `sample_case.json`：5 条记录，植入多类问题，用于验证规则能命中
- `clean_case.json`：3 条完全合规记录，用于验证假阳性为零

## 仓库结构

```
ltc_rules.py                 规则引擎（零第三方依赖，只读 JSON）
ltc_excel.py                 Excel 导入适配器 + 异常清单导出（需 openpyxl）
tools/extract_ruleset.py     官方 PDF → 结构化规则 JSON（带自校验，需 pymupdf）
tools/render_demo_html.py    examples/示例台账.xlsx → docs/index.html（演示页）
data/source/                 官方 PDF 原件（不入库，见该目录 README）
data/derived/                抽取产物（36 项目录 / 巡查表 / 质量评价表 / 记录单字段）
examples/                    导入模板 / 示例台账 / 示例异常清单（全部合成数据）
tests/                       76 个测试（规则 35 + Excel 适配 34 + CLI 冒烟 7）+ 用例数据
docs/                        设计与边界说明 + 演示页
```

## 已知边界

- **项目代码不是官方码**：官方目录附件为图片版，本次可解析的 PDF 内未含项目代码。`SH01`/`YL01` 是本仓库自编的稳定标识，`seq` 才是官方序号。接真实系统时以各地落地版目录的官方代码为准。
- **项目内涵与服务要点未逐条入规则**：36 项的服务要点、质量评价表的评分要点只作为 JSON 留存供人查，未做成自动判定——那部分需要人对服务的现场观察，机器只能核字段自洽性。
- **资质规则按"持证 vs 非持证"分档，不按"护士/医生专属"**：文书并未规定医疗护理类项目只能由护士/医生执行，质量评价表只把"非持证人员提供医疗护理"列为直接不合格；而各地对养老护理员、医疗护理员的执业范围口径不一（浙江绍兴把"执业医师、护士、长期照护师、养老护理员、医疗护理员"并列为可提供长护服务的五类人员）。因此只有非持证人员执行才报 E，养老护理员/医疗护理员执行报 W 交人工核对。
- **不含失能等级评估**：评估部分已由国标与多个开源实现覆盖，本项目不做，也不输出任何评估等级结论。
