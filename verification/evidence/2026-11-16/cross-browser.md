# 跨浏览器（V3，2026-11-16）

Chromium 155（系统）与 Playwright Firefox 159；WebKit 因宿主库缺失未能启动。

| 用例 | Chromium | Firefox | 备注 |
|---|---|---|---|
| DOCX `demo.docx` | ready，6 页，首页 1525px | ready，6 页，首页 1526px | 1px 差异，分页一致 |
| XLSX `xlsx-charts.xlsx` | ready，4 canvas，6 图表 | 同 | 一致 |
| PDF `pdf-rotated-90.pdf` | 4 page slot | 4 | 一致 |
| PPTX `pptx-deck-basic.pptx` | ready，5 缩略图 | 5 | 一致 |

- 控制台/页面错误：两个浏览器均 **0**。
- 结论：两引擎行为一致；DOCX 的 6 页（而非 9 页）是共同缺陷，不是浏览器差异。

## WebKit

- 二进制已下载，但启动失败：缺少 `libicu74`、`libbacktrace.so.0`、`libxml2.so.2`、`libjxl.so.0.8`；宿主为 Arch（ICU 78），无 sudo 安装 Ubuntu 兼容库。
- 记录为 **SKIP（环境阻塞）**，不作为通过。后续可在 CI 的 Ubuntu runner 上补跑。
