# 源文件说明

本目录存放官方政策文书的原件，**PDF 本身不入版本库**（`.gitignore` 已排除）。

## nhsa-service-forms-2026.pdf

《长期护理保险服务管理文书（2026 年版）》，国家医保局办公室印发，50 页，820 KB。

下载地址（公开文件）：

```
https://www.nhsa.gov.cn/module/download/downfile.jsp?classid=0&filename=2f2918a59a0f46879de862e83ad5fbcc.pdf
```

用途：`tools/extract_ruleset.py` 的输入。抽取产物已随仓库提供（`data/derived/`），
仅在需要复核条文或政策更新后重抽时才需要重新下载。

## 另一份依据文书

《国家长期护理保险服务项目目录（试行）》（医保发〔2025〕24 号），官方附件为图片版 PDF，
无法机器解析项目代码，故目录数据取自上述服务管理文书中的三处 36 项清单副本（三处交叉校验）。
