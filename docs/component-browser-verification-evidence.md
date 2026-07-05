# Vue Extend UI 组件浏览器验证证据

验证目标：按 `docs/component-browser-verification-plan.md` 覆盖 demo 全路由、全部组件、真实物料、交互状态、错误状态、响应式与最终命令门禁。

验证环境：内置浏览器，demo 地址 `http://localhost:5000`，项目根目录 `D:\code\ui\vue-extend`。

## 全局路由与响应式

覆盖路由：

- `/#/`
- `/#/docx-viewer`
- `/#/docx-editor`
- `/#/xlsx-viewer`
- `/#/pdf-viewer`
- `/#/components`

视口矩阵：

| 视口 | 结果 |
|---|---|
| 1440×900 | 6 个路由均有内容，document page overflow=false，未出现错误文本 |
| 1280×720 | 6 个路由均有内容，document page overflow=false，未出现错误文本 |
| 768×1024 | 6 个路由均有内容，document page overflow=false；DOCX Editor / XLSX Viewer 使用内部滚动容器承载宽内容 |
| 390×844 | 6 个路由均有内容，document page overflow=false；DOCX Editor / XLSX Viewer 使用内部滚动容器承载宽内容 |

浏览器错误检查：

- Page errors：无
- Console business errors：无；仅 Vite hot update 与 DimCode inspect 调试日志

## 真实物料校验

命令：

```bash
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts\verify_test_materials.py
```

结果：通过。

覆盖物料：

- DOCX：`demo.docx`、`legal-contract.docx`、`invoice-table.docx`、`report-with-image.docx`、`chinese-mixed.docx`、`corrupted.docx`
- PDF：`sample.pdf`、`scanned-invoice.pdf`、`rotated-pages.pdf`、`large-contract.pdf`、`corrupted.pdf`
- XLSX：`financial-model.xlsx`、`sales-table.xlsx`、`charts-images.xlsx`、`large-grid.xlsx`、`corrupted.xlsx`
- 图片：`invoice.png`、`contract-page.png`
- JSON：`field-citations.json`、`ocr-layout.json`、`manifest.json`

## Home

验证项：

- 首页标题与导航卡片可见
- DOCX Viewer / DOCX Editor / XLSX Viewer / PDF Viewer / Components 卡片均显示 `✅ Ready`
- API Compatibility Map 中 DOCX 与 XLSX API 项均显示 `✅ 1:1`
- 路由跳转可进入每个页面
- 响应式矩阵 page overflow=false

## DOCX Viewer

页面：`/#/docx-viewer`

浏览器样本矩阵：

| 文件 | 证据 |
|---|---|
| `demo.docx` | `Loaded: demo.docx`，出现 `MASTER SERVICES AGREEMENT` |
| `legal-contract.docx` | `Loaded: legal-contract.docx`，出现 `MASTER SERVICES AGREEMENT` |
| `invoice-table.docx` | `Loaded: invoice-table.docx`，出现 `INVOICE INV-2026-0705` / `Total Due` |
| `report-with-image.docx` | `Loaded: report-with-image.docx`，出现 `Quarterly Operations Report` / `Executive summary` |
| `chinese-mixed.docx` | `Loaded: chinese-mixed.docx`，出现中文混排内容 |
| `corrupted.docx` | `Loaded: corrupted.docx`，页面显示 DOCX 解析错误态 |

通过标准覆盖：空态、loading、成功态、真实文本、标题、表格、图片、中文混排、损坏文件错误反馈、内部 viewer 滚动。上传路径已验证：上传 `invoice-table.docx` 后显示 `uploaded-invoice-table.docx` 与发票表格内容。

## DOCX Editor

页面：`/#/docx-editor`

验证项：

- 初始文档显示 `Editable Verification Document`
- 页面包含 3 个 `contenteditable` 段落
- 第二段输入 `浏览器验证 mixed English 123` 后页面可见
- 输入后 Undo 按钮启用
- Undo 恢复原文 `Click this paragraph and type...`
- Redo 按钮启用并恢复 `浏览器验证 mixed English 123`
- Bold 后当前文本 `fontWeight = 700`
- Theme toggle 后页面背景变化
- `@extend-ai/vue-docx` typecheck/build 通过

## XLSX Viewer

页面：`/#/xlsx-viewer`

样本矩阵：

| 文件 | 证据 |
|---|---|
| `financial-model.xlsx` | 显示文件名；tabs 为 `Assumptions` / `P&L` / `Notes`；显示 `Starting Customers` |
| `sales-table.xlsx` | 显示 `Sales Data`；显示长文本 `Long note validating overflow...` |
| `charts-images.xlsx` | 显示 `Dashboard`；显示 `Revenue` |
| `large-grid.xlsx` | 显示 `Large Grid`；列名覆盖 `AA` 到 `BH` |
| `corrupted.xlsx` | 显示 `Failed to load workbook`；sheet tabs 清空 |

交互矩阵：

- 单元格点击显示 selection outline
- 双击 `Starting Customers` 输入框初始值为 `Starting Customers`
- Enter 提交 `Edited Customers` 后页面可见
- Undo 恢复原值，Redo 恢复编辑值
- Escape 取消 `Cancelled Edit`，原值保留
- Tab 提交 `Tab Edit` 并移动 selection 到下一格
- Zoom 100%→110%，列宽样式变为 `88px`
- Read only 模式阻止双击编辑
- Download 按钮存在且可用
- 上传路径已验证：上传 `sales-table.xlsx` 后显示 `uploaded-sales-table.xlsx` 与 `Sales Data`

## PDF Viewer

页面：`/#/pdf-viewer`

样本矩阵：

| 文件 | 证据 |
|---|---|
| `sample.pdf` | 4 thumbnails；Next 后 `2 / 4`；iframe hash `/samples/sample.pdf#page=2&zoom=125` |
| `scanned-invoice.pdf` | 3 thumbnails；Next 后 `2 / 3` |
| `rotated-pages.pdf` | 2 thumbnails；Next 后 `2 / 2` |
| `large-contract.pdf` | 31 thumbnails；Next 后 `2 / 31` |
| `corrupted.pdf` | 显示 `Unable to load PDF document.`；thumbnails 为 0 |

交互矩阵：

- Next/Previous 更新页码状态
- thumbnail 数量匹配页数
- zoom 更新 iframe hash 与 transform
- rotate 产生可观察 transform matrix
- search 输入后显示结果反馈
- download 失败有错误捕获路径
- 上传路径已验证：上传 `sample.pdf` 后显示 `uploaded-sample.pdf` 与 4 页状态

## Components

页面：`/#/components`

| 组件 | 证据 |
|---|---|
| SignaturePad | canvas CSS 尺寸 400×200；绘制后出现 `Signature captured (...)`；`Check empty` 显示 `Has signature`；Clear 后显示 `Empty` |
| FileUpload | PDF 拖拽显示 `Uploaded: matrix.pdf` |
| FileUpload accept | XLSX 拖入 PDF/DOCX 上传区显示 `not an accepted file type` |
| FileUpload maxSize | 3MB PDF 显示 `exceeds the 2.0MB size limit` |
| FileUpload disabled | disabled 上传区拖入文件后无接收状态 |
| FileThumbnail | PDF/DOCX/XLSX/PNG/unknown 文件名均可见；图片缩略图宽度为 48px |
| BoundingBoxCitations | invoice 字段可见；点击字段后 selected field 更新 |
| LayoutBlocks | 7 个 OCR blocks 可见；点击 row 后 selected block 更新；空态显示 `No layout blocks detected` |
| Spinner | spinner 元素可见 |
| Tooltip | hover/focus 后显示 tooltip 文案 |

## 命令门禁

最终执行并通过：

```bash
pnpm typecheck
pnpm build
```

`pnpm build` 仅输出 Vite 动态/静态 import chunking warning，构建状态为成功。

## 修复摘要

- DOCX Viewer：补齐多样本入口、可见错误态、JS DOCX fallback。
- DOCX Editor：补 starter model、contenteditable 写回 model、selection、format、响应式 undo/redo、DOM key 刷新。
- XLSX Viewer：补样本入口、错误态、HTTP status 检查、真实文件名、编辑初始值、Undo/Redo、zoom 列宽/行高、download、readOnly/Tab/Escape 行为。
- PDF Viewer：补真实 fetch、页数估算、错误态、page/zoom hash、search 反馈、download 错误处理。
- Components：补全真实 fixture demo，修复图片原始宽撑开与响应式布局。
- Home：状态更新为当前验证结果。
- 全局响应式：nav wrap、页面 min-width/overflow 防护。
