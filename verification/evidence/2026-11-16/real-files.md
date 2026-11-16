# 真实文件验证（V1，2026-11-16）

来源：Apache POI / Apache PDFBox `test-data`（Apache-2.0），共 18 个文件，清单与 sha256 见 `test-data/real/manifest.json`。

## 结果

- **oracle（LibreOffice/pdfinfo）**：18/18 页数与收集时一致，**PASS**。
- **本地 viewer**：16/18 通过；2 个 XLSX 失败，根因同一个，已立 **#152**。

失败文件：

| 文件 | 现象 | 根因 |
|---|---|---|
| `poi-xlsx-45540_classic_Header.xlsx` | `INVALID_WORKBOOK`，`Cannot read properties of null (reading 'getElementsByTagName')` | 图表系列没有 `<c:tx>`，`readChartSeriesNameFromXml` 把 `null` 传给 `getFirstLocalDescendant` |
| `poi-xlsx-45544.xlsx` | 同上 | 同上 |

同一错误在主线程与 worker 路径都复现；CDP pause-on-exception 栈定位到 `getLocalDescendants → getFirstLocalDescendant → readChartSeriesNameFromXml`。

## 关键观测

| 文件 | 本地结果 |
|---|---|
| `poi-docx-60316.docx`（文本框 + SDT + 3 headers） | ready，1 页 |
| `poi-docx-58067.docx`（7 处修订） | ready，1 页 |
| `poi-pptx-45545_Comment.pptx`（11 页、11 notes、comments、媒体） | ready，11 缩略图（批注/备注未渲染，见 #100） |
| `poi-pptx-SmartArt.pptx` | ready，1 页（SmartArt 作为整体） |
| `poi-pptx-EmbeddedVideo.pptx` | ready，1 页（媒体交互未验证） |
| `poi-xlsx-123233_charts.xlsx` | ready，仅显示 1 个可见 sheet（其余 4 个 hidden，正确） |
| `poi-xlsx-46535.xlsx` | ready，2 个可见 sheet（3 个 hidden，正确），47 个打印页 |
| `poi-xlsx-47668.xlsx` | ready，1 个图片 |
| PDFBox rotations（6 页，90/180/270） | 6 个 page slot，旋转渲染 |
| PDFBox ligatures / 文本 | 1 页，文本层可搜索 |

## 结论

- 真实文件第一轮就暴露了一个 **P0 级解析崩溃**：只要图表系列没有标题，整个工作簿无法打开（#152）。
- 其余格式的真实文件首轮可加载；下一步 V2 用编辑→导出→重开验证内容是否真的保留，而不是“能打开”。
