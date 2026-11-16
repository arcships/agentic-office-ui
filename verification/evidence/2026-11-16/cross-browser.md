# 跨浏览器（V3，2026-11-16）

Chromium 155（系统）、Playwright Firefox 159、**系统 WebKitGTK 2.54.1（WebKitWebDriver）**。

| 用例 | Chromium | Firefox | 备注 |
|---|---|---|---|
| DOCX `demo.docx` | ready，6 页，首页 1525px | ready，6 页，首页 1526px | 1px 差异，分页一致 |
| XLSX `xlsx-charts.xlsx` | ready，4 canvas，6 图表 | 同 | 一致 |
| PDF `pdf-rotated-90.pdf` | 4 page slot | 4 | 一致 |
| PPTX `pptx-deck-basic.pptx` | ready，5 缩略图 | 5 | 一致 |

- 控制台/页面错误：两个浏览器均 **0**。
- 结论：两引擎行为一致；DOCX 的 6 页（而非 9 页）是共同缺陷，不是浏览器差异。

## WebKit（系统 WebKitGTK）

Playwright 自带的 WebKit 需要 Ubuntu 兼容库，但 Omarchy/Arch 自带 `WebKitWebDriver`，直接走 W3C WebDriver：

| 用例 | 结果 |
|---|---|
| DOCX `demo.docx` | ready，6 页，首页 1524px |
| XLSX `xlsx-charts.xlsx` | ready，4 canvas，6 图表 |
| PDF `pdf-rotated-90.pdf` | 4 slot / 4 图 |
| PPTX `pptx-deck-basic.pptx` | ready，5 缩略图 |

脚本：`tests/blackbox/webkit_webdriver.py`；结果 `cross-browser/webkit.json` + 4 张 `webkit-*.png`。WebKitGTK 2.54.1 与 Chromium/Firefox 结果一致。
