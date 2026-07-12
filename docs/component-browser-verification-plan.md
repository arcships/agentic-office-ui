# Vue Extend UI 组件浏览器验证测试方案

## 目标

使用内置浏览器对 Vue Extend UI demo 执行端到端验证，覆盖所有页面、组件、真实物料、交互状态、错误状态与视觉一致性。

验证维度：

1. 视觉样式：布局、间距、边框、颜色、hover、focus、disabled、滚动容器、响应式。
2. 核心功能：真实文件加载、交互事件、状态变化、错误状态、编辑行为。
3. API 行为：props、events、slots、exposed methods、composables 的页面可观察结果。
4. 真实数据兼容性：PDF、DOCX、XLSX、图片、OCR JSON、表单文件等真实物料。
5. 稳定性：控制台错误、网络错误、空状态、重复操作、页面切换后的资源释放。

## 测试入口

项目目录：`D:\code\ui\vue-extend`

启动命令：

```bash
pnpm dev
```

浏览器入口：

```text
http://localhost:5173
```

路由清单：

| 页面 | 路由 | 文件 | 测试重点 |
|---|---|---|---|
| Home | `/#/` | `apps/demo/src/pages/HomePage.vue` | 导航卡片、API 表格、路由跳转 |
| DOCX Viewer | `/#/docx-viewer` | `apps/demo/src/pages/DocxViewerPage.vue` | DOCX 加载、分页、文本样式、表格/图片 |
| DOCX Editor | `/#/docx-editor` | `apps/demo/src/pages/DocxEditorPage.vue` | toolbar、contenteditable、主题切换、undo/redo |
| XLSX Viewer | `/#/xlsx-viewer` | `apps/demo/src/pages/XlsxViewerPage.vue` | workbook、sheet tabs、cell edit、zoom |
| PDF Viewer | `/#/pdf-viewer` | `apps/demo/src/pages/PdfViewerPage.vue` | PDF 上传、工具栏、缩放、搜索、缩略图、下载 |
| Components | `/#/components` | `apps/demo/src/pages/ComponentsPage.vue` | SignaturePad、FileUpload、FileThumbnail、Spinner、BoundingBox、LayoutBlocks、Tooltip |

## 组件覆盖清单

| 包 | 组件 / API | 浏览器验证项 |
|---|---|---|
| `@arcships/vue-docx` | `DocxViewer` | 空态、loading、真实 DOCX、段落、标题、表格、图片、滚动、样式 |
| `@arcships/vue-docx` | `DocxEditorViewer` | toolbar 按钮、select、主题切换、contenteditable 输入 |
| `@arcships/vue-docx` | DOCX composables | 页面状态、编辑器状态、分页、主题、评论、修订、样式 |
| `@arcships/vue-xlsx` | `XlsxViewer` | workbook 加载、sheet tabs、cell selection、double click edit、zoom、undo/redo、download |
| `@arcships/vue-xlsx` | `XlsxViewerProvider` | provider 注入 controller 的共享状态 |
| `@arcships/vue-extend` | `PdfViewer` | engine loading、空态、PDF 渲染、页码、缩放、旋转、sidebar、下载 |
| `@arcships/vue-extend` | `SignaturePad` | 画线、签名 dataURL、clear、isEmpty、尺寸 |
| `@arcships/vue-extend` | `FileUpload` | click upload、accept 校验、错误提示、drag state、多文件、maxSize |
| `@arcships/vue-extend` | `FileThumbnail` | PDF、DOCX、XLSX、image、unknown 文件样式 |
| `@arcships/vue-extend` | `BoundingBoxCitations` | 多页框、置信度颜色、点击事件、空态 |
| `@arcships/vue-extend` | `LayoutBlocks` | OCR 图片叠框、类型颜色、列表、点击事件、空态 |
| `@arcships/vue-extend` | `Spinner` | 尺寸、颜色、动画、容器对齐 |
| `@arcships/vue-extend` | `Tooltip` | hover/focus 显示、定位、长文本、边界 |

## 真实物料测试集

物料目录：`apps/demo/public/samples`

### DOCX

| 文件 | 内容要求 | 覆盖 |
|---|---|---|
| `legal-contract.docx` | 多页合同、标题、编号列表、页眉页脚 | 分页、文字样式、滚动 |
| `invoice-table.docx` | 表格、合并单元格、金额、日期 | table 渲染、单元格样式 |
| `report-with-image.docx` | 图片、粗体、斜体、下划线、颜色 | run style、图片 |
| `chinese-mixed.docx` | 中文、英文、数字、标点 | 字体 fallback、换行 |
| `corrupted.docx` | 损坏文件 | 错误提示、console 记录 |

### PDF

| 文件 | 内容要求 | 覆盖 |
|---|---|---|
| `sample.pdf` | 3-10 页文本型 PDF | 翻页、缩略图、搜索 |
| `scanned-invoice.pdf` | 扫描件 | 大图性能、缩放 |
| `rotated-pages.pdf` | 横向/竖向混排 | 旋转、viewport |
| `large-contract.pdf` | 30+ 页 | 滚动性能、thumbnail buffer |
| `corrupted.pdf` | 损坏文件 | 错误提示、engine 容错 |

### XLSX

| 文件 | 内容要求 | 覆盖 |
|---|---|---|
| `financial-model.xlsx` | 多 sheet、公式、金额格式 | sheet tabs、格式 |
| `sales-table.xlsx` | 冻结表头、宽列、长文本 | grid、滚动、截断 |
| `charts-images.xlsx` | 图表、图片 | chart/image composables |
| `large-grid.xlsx` | 500+ 行、50+ 列 | 性能、虚拟范围 |
| `corrupted.xlsx` | 损坏文件 | 错误提示、加载失败状态 |

### 图片与 OCR 数据

| 文件 | 内容要求 | 覆盖 |
|---|---|---|
| `invoice.png` | 真实发票或表单截图 | BoundingBoxCitations、LayoutBlocks |
| `contract-page.png` | 合同页面图片 | OCR block overlay |
| `ocr-layout.json` | title、text、table、figure、header、footer、list blocks | LayoutBlocks 类型颜色与列表 |
| `field-citations.json` | 供应商、金额、税号、日期、低/中/高置信度 | BoundingBoxCitations 多置信度展示 |

## 全局浏览器验证流程

1. 启动 demo：`pnpm dev`。
2. 打开 `http://localhost:5173`。
3. 检查首屏加载、标题、导航、首页卡片。
4. 读取控制台与页面错误。
5. 逐个点击 nav 与首页卡片进入每个路由。
6. 每个路由刷新一次，确认 hash route 正常。
7. 每个路由执行 desktop、tablet、mobile 三档宽度截图检查。
8. 每个页面完成后读取 console errors 与 network 失败请求。

通过标准：

- 页面加载期间无 uncaught error。
- nav active 样式正确。
- 内容无异常横向溢出。
- 主区域与内部 viewer 滚动互相独立。
- 卡片 hover、按钮 hover、输入 focus 有可见反馈。

## DOCX Viewer 测试

### 步骤

1. 打开 `/#/docx-viewer`。
2. 检查空态文案。
3. 点击 `Load sample`。
4. 上传 `legal-contract.docx`。
5. 滚动到第二页、中段、末页。
6. 上传 `invoice-table.docx`，检查表格与合并单元格。
7. 上传 `report-with-image.docx`，检查图片与文字样式。
8. 上传 `chinese-mixed.docx`，检查中文、英文、数字混排。
9. 上传 `corrupted.docx`，检查错误行为。
10. 读取 console errors。

### 通过标准

- 空态、loading、成功态可见。
- 至少渲染 2 页真实内容。
- 标题、粗体、斜体、下划线、颜色、表格、图片可见。
- 页面宽高、白底、阴影、边距稳定。
- 损坏文件给出明确错误反馈。

### 已知风险

- `DocxViewerPage.vue` 依赖 `/samples/demo.docx`，物料目录当前需要补齐。
- `DocxViewer.vue` 解析失败路径只写 `console.error`，页面层需要可见错误态。

## DOCX Editor 测试

### 步骤

1. 打开 `/#/docx-editor`。
2. 检查 toolbar：undo、redo、B、I、U、Heading select、主题切换。
3. 检查 undo/redo 初始 disabled 状态。
4. 点击文档段落，输入中文英文混合文本。
5. 点击 Bold / Italic / Underline 后继续输入。
6. 切换 Heading 1、Heading 2、Heading 3。
7. 切换暗色主题，再切回亮色主题。
8. 尝试 undo/redo。
9. 读取 console errors。

### 通过标准

- toolbar 可点击，有明确 disabled 态。
- contenteditable 可输入，文本可见。
- 主题切换影响页面背景与文字。
- 页码显示合理。
- 编辑操作对 model 或页面状态有可观察结果。

### 已知风险

- `DocxEditor.vue` 的 `onParagraphInput` 当前只读取文本，尚未写回 editor model。
- toolbar 格式按钮调用 editor API，需验证状态是否真实影响内容。

## XLSX Viewer 测试

### 前置要求

当前 `XlsxViewerPage.vue` 主要展示类型验证信息，需要增加真实 `XlsxViewer` demo 场景，或在测试前临时补齐页面入口。

### 步骤

1. 打开 `/#/xlsx-viewer`。
2. 加载 `financial-model.xlsx`。
3. 检查 workbook 内容与 sheet tabs。
4. 切换多个 sheet。
5. 点击单元格，检查 selection 边框。
6. 双击单元格，输入新值，按 Enter 提交。
7. 双击单元格，输入新值，按 Escape 取消。
8. 使用 Tab 提交并移动。
9. 操作 zoom +、zoom -、100%。
10. 操作 undo、redo。
11. 下载 workbook。
12. 加载 `sales-table.xlsx`、`charts-images.xlsx`、`large-grid.xlsx`。
13. 加载 `corrupted.xlsx`，检查错误行为。

### 通过标准

- workbook 内容显示，sheet tab 数量正确。
- selection 边框位置正确。
- 编辑值留在单元格内。
- zoom 改变字体与列宽视觉。
- 只读模式禁止编辑。
- 大表滚动保持响应。

### 已知风险

- `XlsxViewerPage.vue` 当前没有渲染 `XlsxViewer`。
- `XlsxViewer.vue` 当前显示本地 `cellData`，真实 workbook cell 数据是否进入 UI 需要实测。
- `XlsxViewer.vue` 列名生成只覆盖 A-Z，AA 之后需要修复。

## PDF Viewer 测试

### 步骤

1. 打开 `/#/pdf-viewer`。
2. 检查空态文案。
3. 上传 `sample.pdf`。
4. 等待 PDF engine 与页面渲染完成。
5. 操作 zoom +、zoom -、固定比例选择。
6. 打开 sidebar，检查 thumbnails 数量与当前页同步。
7. 搜索真实文本关键词。
8. 页码输入、下一页、上一页。
9. 旋转页面。
10. 下载文件，检查文件名。
11. 上传 `scanned-invoice.pdf`、`rotated-pages.pdf`、`large-contract.pdf`。
12. 上传 `corrupted.pdf`，检查错误行为。
13. 读取 console errors。

### 通过标准

- PDF 页面可见，缩略图数量匹配页数。
- zoom 与 rotate 改变 viewport。
- 搜索命中有可见结果。
- 下载文件名正确。
- engine 失败时显示错误态。

### 已知风险

- `PdfViewerPage.vue` 当前只支持上传，建议增加 sample loader。
- `PdfViewer.vue` 页码状态主要靠本地值，需要验证与真实 scroll active page 同步。

## Components 页面测试

### SignaturePad

步骤：

1. 打开 `/#/components`。
2. 在 canvas 上画一条签名线。
3. 检查 `Signature captured (...)` 出现。
4. 点击组件内 Clear。
5. 再画一次。
6. 点击页面外 Clear。
7. 验证 canvas 视觉尺寸为 400×200。

通过标准：

- 鼠标或触摸轨迹可见。
- dataURL 长度随签名变化。
- clear 后预览文案消失。

### FileUpload

步骤：

1. 上传合法 PDF。
2. 上传合法 DOCX。
3. 上传非法 TXT。
4. 上传非法 PNG。
5. 拖拽文件到区域。
6. 测试 multiple false 时多选行为。
7. 补 maxSize 场景测试超限文件。
8. 补 disabled 场景测试禁用状态。

通过标准：

- 合法文件触发 `Uploaded: filename`。
- 非法文件显示错误。
- dragover 样式变化。
- disabled 场景禁止打开 picker。

### FileThumbnail

步骤：

1. 检查 PDF、XLSX、PNG 三个现有样式。
2. 补充 DOCX、unknown、长文件名、无 MIME type。
3. 检查文本截断、图标、颜色。

通过标准：

- 图标、扩展名、文件名清晰。
- 长文件名截断合理。
- unknown 类型有 fallback 样式。

### Spinner

步骤：

1. 检查默认 spinner。
2. 补充不同 size、颜色、容器对齐场景。
3. 检查动画持续运行。

通过标准：

- 动画平滑。
- 尺寸与颜色 props 生效。
- 在按钮、空态、整页 loading 中对齐正常。

### BoundingBoxCitations

前置要求：当前 Components 页面需要补充组件展示入口。

步骤：

1. 使用 `invoice.png` 或 `sample.pdf` 的字段 citation 数据。
2. 渲染多页字段。
3. 检查高、中、低置信度颜色。
4. 点击 overlay box。
5. 点击字段列表项。
6. 渲染空 fields 场景。

通过标准：

- 框位置与字段列表一致。
- 点击 overlay 与 list 都发出 `field-click`。
- 空态显示 `No fields to display`。

### LayoutBlocks

前置要求：当前 Components 页面需要补充组件展示入口。

步骤：

1. 使用 `contract-page.png` 与 `ocr-layout.json`。
2. 检查 title、text、table、figure、header、footer、list 类型颜色。
3. 点击图片 overlay block。
4. 点击 block list 项。
5. 渲染空 blocks 场景。

通过标准：

- OCR 图像与 bbox 对齐。
- 类型颜色正确。
- 点击 overlay 与 list 都发出 `block-click`。
- 空态显示 `No layout blocks detected`。

### Tooltip

前置要求：当前 Components 页面需要补充组件展示入口。

步骤：

1. hover 触发 tooltip。
2. focus 触发 tooltip。
3. mouseleave / blur 隐藏 tooltip。
4. 测试长文本。
5. 测试靠近页面边界的定位。

通过标准：

- tooltip 显示与隐藏稳定。
- 文本不溢出 viewport。
- 键盘 focus 可触发。

## 响应式与视觉回归

视口：

| 名称 | 尺寸 |
|---|---|
| Desktop | 1440 × 900 |
| Laptop | 1280 × 720 |
| Tablet | 768 × 1024 |
| Mobile | 390 × 844 |

每个页面截图点：

1. 初始空态。
2. 成功加载真实物料。
3. 交互后的状态。
4. 错误态。

视觉验收标准：

| 类别 | 标准 |
|---|---|
| 布局 | 四档视口下可用，无主内容遮挡 |
| 颜色 | 使用 CSS token，无明显裸 HTML 样式 |
| 状态 | empty、loading、error、disabled、hover、focus 均可观察 |
| 滚动 | viewer 容器内部滚动，页面主布局稳定 |
| 文档 | PDF/DOCX 页面比例与白底纸张感稳定 |
| 表格 | XLSX 网格行列头对齐，selection 不漂移 |
| 无障碍 | 核心交互可通过键盘完成 |

## 稳定性与错误处理

每个主流程后检查：

1. Browser console errors。
2. Unhandled promise rejection。
3. Network 404/500。
4. WASM、worker、PDF engine 加载状态。
5. 多次页面切换后的重复 error。
6. 上传损坏文件后的恢复能力。

通过标准：

| 类别 | 标准 |
|---|---|
| Console | 无 uncaught error；预期失败有明确错误信息 |
| Network | samples 文件 200；WASM、worker、pdf engine 加载成功 |
| 文件上传 | accept、错误、成功事件都可观察 |
| 文档解析 | DOCX、PDF、XLSX 至少各 3 个真实文件 |
| 编辑 | DOCX、XLSX 编辑动作有可见结果或明确记录实现缺口 |
| 资源 | 连续切换页面 5 次，无明显卡死或重复 error |

## 缺陷记录模板

```md
## [Severity] 标题

- 页面：
- 组件：
- 文件：
- 视口：
- 物料：
- 复现步骤：
  1.
  2.
  3.
- 期望结果：
- 实际结果：
- Console / Network：
- 截图：
- 影响范围：
- 建议修复：
```

严重级别：

| 级别 | 定义 |
|---|---|
| P0 | 页面无法加载、核心 viewer 崩溃、真实物料完全不可用 |
| P1 | 核心功能不可用、编辑/上传/渲染明显错误 |
| P2 | 样式、状态、兼容性问题影响正常使用 |
| P3 | 文案、细节、边界视觉问题 |

## 执行顺序

1. 补齐 `apps/demo/public/samples` 真实物料。
2. 补齐 demo 覆盖缺口：`XlsxViewer`、`BoundingBoxCitations`、`LayoutBlocks`、`Tooltip`、PDF sample loader。
3. 运行 `pnpm dev`。
4. 使用内置浏览器逐页执行 snapshot、交互、console/errors 检查。
5. 对关键状态截图。
6. 按缺陷模板记录问题。
7. 修复 P0/P1 后重新跑全量核心流程。

## 当前阻塞与准备项

| 项目 | 状态 | 处理 |
|---|---|---|
| `public/samples` 真实物料 | 待补齐 | 加入 DOCX、PDF、XLSX、PNG、JSON 样本 |
| XLSX demo 真实 viewer | 待补齐 | 在 `XlsxViewerPage.vue` 渲染 `XlsxViewer` 与 sample loader |
| BoundingBox demo | 待补齐 | 在 `ComponentsPage.vue` 增加真实字段数据入口 |
| LayoutBlocks demo | 待补齐 | 在 `ComponentsPage.vue` 增加 OCR 图片与 JSON 入口 |
| Tooltip demo | 待补齐 | 在 `ComponentsPage.vue` 增加 hover/focus 场景 |
| PDF sample loader | 待补齐 | 在 `PdfViewerPage.vue` 增加 `/samples/sample.pdf` 加载按钮 |

准备项完成后开始浏览器实测。