# 渲染与行为观测（2026-11-16）

环境：Chromium 155、LibreOffice 26.8.1.1、上游 `react-docx@28a3bd9` / `react-xlsx@6277fe7`、本地 demo dist。

## 1. DOCX 页数（同一文件）

| 夹具 | 上游 | 本地 | LibreOffice | 结论 |
|---|---:|---:|---:|---|
| `demo.docx`（= `legal-contract.docx`） | 9 | 6 | 9 | 本地少 3 页，首页 1525px vs 纸张 1056px |
| `docx-comments.docx` | 1 | 1 | 1 | 一致 |
| `docx-merged-table.docx` | — | 1 | 1 | 一致 |

本地首页 `surface` 高度 1525px，纸张内容高 1056px；9/6 = 1.5、1525/1056 ≈ 1.44，说明本地把约一页半内容塞进一页。

## 2. 评论/修订

同一 `docx-comments.docx`：

- 上游：`Page 1 of 1`、32 words，**不显示批注卡**，默认最终文本（Review 面板按需打开）。
- 本地：1 页，**3 张 gutter 卡片**（2 修订 + 1 评论），内联插入/删除同时可见。
- 实测坐标（本地）：卡片 top 434/480/552，正文锚点 487/507/507；评论卡比锚点低 65px。
- 修订样式：插入 `rgb(220,38,38)`、删除 `rgb(185,28,28)`，两者 `text-decoration: none`。

## 3. XLSX

同一 `xlsx-charts.xlsx`：

- 上游：6 个图表全部渲染（SVG），但柱/折线**缺分类轴刻度**。
- 本地：6 个图表，分类轴与数值轴完整；图表位于 F2/N2 等锚点。
- 本地示例 `welcome.xlsx`：17 个 sheet tab 正常加载。

## 4. PDF

- `sample.pdf`：4 个 page slot、4 张 595×842 图；首页非白像素 13.6%。
- `large-contract.pdf`：31 页**全部预栅格化**（31 个 `<img>`，scrollHeight 26374）。
- fitWidth：slot 宽度 1191px（容器 ~1424），生效。

## 5. PPTX

`pptx-deck-basic.pptx`（5 页）：

- 浏览：`state=ready`、`1 / 5`、5 缩略图、13 个对象节点。
- 演示：`pptx-viewer--present` 生效，播放控件存在；点击舞台 slide 0 → 1，无 console error。
- LibreOffice 对照：5 页，一致。

## 6. 视口矩阵（本地）

1440×900 / 1280×720 / 768×1024 / 390×844 × 5 路由：无 console error、无页面级横向滚动。

## 7. 证据

- 截图：`screenshots/`（上游/本地成对）
- oracle：`office-oracle/summary.json`
- 双跑：`upstream-parity/summary.json`
