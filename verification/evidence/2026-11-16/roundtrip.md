# 编辑→导出→重开（V2，2026-11-16）

环境：本地 demo dist、Chromium 155、LibreOffice 26.8、python-docx/openpyxl。

## DOCX —— PASS

- 在 `DocxEditor` 起始文档输入 marker（`ROUNDTRIP-<ts>`）；
- 点击 `editor-export` 捕获下载；
- python-docx 重开：**marker 存在**；
- LibreOffice 转 PDF：**成功**（导出包合法）。

导出体积 4.1 KB；证据 `roundtrip/roundtrip.docx`、`roundtrip/roundtrip.pdf`。

## XLSX —— FAIL（P0）

- 上传 `test-data/corpus/xlsx-charts.xlsx`（6 图表），编辑 A1 后导出；
- 导出包：`xl/drawings/drawing1.xml` + rels 存在，rels 指向 `xl/charts/chart1..5.xml`；
- **5 个 chart part 全部缺失，`xl/media` 为空**；
- openpyxl 重开失败：`KeyError: There is no item named 'xl/charts/chart1.xml' in the archive`；
- 单元格编辑已写入 `sheet1.xml`（marker 可检索），说明丢失发生在导出装配阶段。

结论：带图表的 XLSX 导出后是**结构性损坏的包**，Excel/openpyxl 无法打开。已回填 #66，建议列为 P0。

证据：`roundtrip/roundtrip.xlsx`、`roundtrip/summary.json`。

## 深度回环（新增）

- **DOCX**：通过 demo 验收钩子插入表格 + 图片 + 格式化中间文字，导出后 python-docx 检查：2 张表、1 个 `word/media` part、1 个 `<w:drawing>`、2 个 bold run；LibreOffice 可打开。
- **XLSX**：使用无图表的 `xlsx-validation.xlsx`，选中 A4（名称框确认）加粗 + 红色填充，导出后 openpyxl 读回 `bold=True`、`fill=FFFF0000`。
- 结论：结构化内容与样式的导出链路在“无图表”路径下正常；图表路径仍是 #66 的 P0。证据：`roundtrip/deep.json`、`roundtrip/deep.docx`、`roundtrip/deep.xlsx`。
