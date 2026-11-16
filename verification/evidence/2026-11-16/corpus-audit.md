# 夹具审计（2026-11-16）

用 `uv run --with pypdf --with openpyxl --with python-docx` 读取现有 demo 样例与新建 corpus，核对“文件实际内容”与“测试/文档声称”。

## 现有 demo 样例

| 夹具 | 实测 | 问题 |
|---|---|---|
| `review-comments.docx` | 4 段、1 评论、1 range、无 `commentsExtended`、1 ins + 1 del、1 脚注、1 尾注、1 header | 无法覆盖线程回复、多评论、表格/页眉评论 |
| `legal-contract.docx` | 125 段、1 表、1 header、1 sectPr | 与 `demo.docx` **md5 相同**（同一文件两个名字） |
| `charts-images.xlsx` | 2 sheet、1 图表、**0 图片**、1 合并区、**0 验证**、无冻结、A1:D8 | 名不符实；冻结/验证/图片路径零覆盖 |
| `large-grid.xlsx` | A1:BH551（551×60） | 不是大表 |
| `rotated-pages.pdf` | 2 页、`/Rotate` 均为 0 | 无旋转元数据；第二页只是横向 MediaBox |
| `scanned-invoice.pdf` | 3 页、每页 22 字符文本层 | 不是 image-only |
| `sample.pdf` / `large-contract.pdf` | 4 页 / 31 页 | “large” 不够大 |

## 新建 corpus（`test-data/corpus/`）

| 文件 | 特性 | 校验 |
|---|---|---|
| `docx-merged-table.docx` | vMerge×3、gridSpan、嵌套表、多段落单元格 | ✅ |
| `docx-mixed-orientation.docx` | 竖版 + 横版两个 section | ✅ |
| `docx-two-columns.docx` | `w:cols num=2` | ✅ |
| `docx-comments.docx` | 3 条评论（正文×2 + 表格×1）、1 脚注、header | ✅（LibreOffice 可转换） |
| `docx-large-table.docx` | 120 行表格 | ✅ |
| `xlsx-charts.xlsx` | bar/line/pie/area/radar/scatter 六种图表 | ✅ |
| `xlsx-conditional-formats.xlsx` | dataBar(num)、colorScale(percentile)、4 图标集、负值 | ✅ |
| `xlsx-validation.xlsx` | 列表 + 数值验证 | ✅ |
| `xlsx-frozen-hidden.xlsx` | 冻结 B3、隐藏行列、合并、1904 | ✅ |
| `xlsx-number-formats.xlsx` | 日期/货币/百分比/文本/`[Red]` | ✅ |
| `xlsx-large-20k.xlsx` | 2 万行公式 | ✅ |
| `pdf-rotated-90.pdf` | 4 页 `/Rotate 90` | ✅ |
| `pdf-image-only.pdf` | 1 页、文本层长度 0 | ✅ |
| `pdf-encrypted.pdf` | 密码 `secret` | ✅ |
| `pdf-300-pages.pdf` | 300 页 | ✅ |
| `pptx-deck-basic.pptx` | 5 页：占位符/项目符号/表格/图表/备注 | ✅ |

## 本轮修复

`docx-comments.docx` 最初包含**页眉评论**（python-docx `add_comment` 写在 header run 上），LibreOffice 直接拒绝加载（`source file could not be loaded`）。已改为“页眉文本 + 正文/表格评论”，并给注入的 `footnotes.xml` 补上 `[Content_Types].xml` override 与 `document.xml.rels` 关系。现在 LibreOffice 可正常转换（1 页），`verify_review_corpus.py` 通过。

结论：页眉评论需要 Word 真实保存的夹具，python-docx 生成的包不合法；后续归入真实文件来源。
