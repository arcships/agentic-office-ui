# Viewer / Editor 架构审计：上游组件树、私有边界与我们的分层（#185）

> 批次：0.8.1 / issue [#185](https://github.com/arcships/agentic-office-ui/issues/185)。
> 固定基准：[`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)（2026-11-17，MIT）。
> 关联：[#171](https://github.com/arcships/agentic-office-ui/issues/171) page setup、[#172](https://github.com/arcships/agentic-office-ui/issues/172) PDF surface、[#176](https://github.com/arcships/agentic-office-ui/issues/176) 对象矩阵、[#178](https://github.com/arcships/agentic-office-ui/issues/178) CSV、[#179](https://github.com/arcships/agentic-office-ui/issues/179) FileSystem/SchemaBuilder、[#182](https://github.com/arcships/agentic-office-ui/issues/182) L2 route parity、[#183](https://github.com/arcships/agentic-office-ui/issues/183) presenter、[#184](https://github.com/arcships/agentic-office-ui/issues/184) chart internals。
>
> 本文只做架构结论与实施计划，不改产品代码。逐文件归属与许可证见 [upstream/extend-ui.md](../upstream/extend-ui.md)，组件级对照见 [components/ui-component-map.md](../components/ui-component-map.md)。

## 0. 结论摘要

1. **上游七个 viewer/editor 的“私有边界”不是同一种东西。** PDF 100% 公开（`apps/v4/components/extend/pdf-*.tsx` + Apache-2.0 的 EmbedPDF npm 插件，无 `@extend-ai/*`）；DOCX / XLSX / PPTX 是“公开 chrome 文件 + 独立版本化的 MIT npm runtime”；CSV / FileSystem / SchemaBuilder 全部在 pinned 仓库内，没有私有 runtime。
2. **“私有包”不等于闭源。** `@extend-ai/react-docx@0.15.0`、`@extend-ai/react-xlsx@0.23.0`、`@extend-ai/react-pptx@0.5.0` 都是 MIT、公开发布、有公开 GitHub 仓库（2026-11-21 核对：`api.github.com/repos/extend-hq/react-{docx,xlsx,pptx}` 均 200）。真正的边界是**版本与仓库边界**：它们不在 pinned UI 仓库里，所以 pinned commit 不能单独 `pnpm install` 出对应实现。
3. **我们不是从零写引擎。** `docx-core`、`xlsx-core`、`vue-docx`、`vue-xlsx` 的源码头直接标注了上游移植来源（例如 [packages/docx-core/src/index.ts:3](../../packages/docx-core/src/index.ts#L3) `@extend-ai/react-docx @ commit 6f70b92`）。缺的主要是 **chrome 组合 + 若干宿主能力**，不是渲染内核。
4. **`unreachable` 只剩一条，而且理由错了。** `pptx-presenter` 的 manifest 理由写着 presenter 在 `@extend-ai/react-pptx` 内；实测 0.5.0 的 `dist` 里 `presenter` 出现 0 次，只有 `.rpv-notes` 与 `showNotes`。presenter 在公开层和 private 包都不存在 → 应归为 **office-reference**（#183）。
5. **FileSystem / SchemaBuilder 的“私有依赖”论据不成立。** `@pierre/trees@1.0.0`、`@pierre/diffs@1.6.4` 均为 Apache-2.0（npm registry，2026-11-21），`@dnd-kit/*` 为 MIT。这两个 block 是**产品范围/工作量缺口**，必须用产品理由关闭，不能用许可证或“私有引擎”关闭（#179）。
6. **CSV 不是 `not-used`，是 `product-gap`。** 上游 `csv-viewer.tsx` 975 行全公开；我们只有 XLSX 的 CSV 路径（[packages/vue-xlsx/src/composables/workbook-state.ts:585](../../packages/vue-xlsx/src/composables/workbook-state.ts#L585)），没有独立 block、没有 fixture、没有 parity surface（#178）。
7. **L2（#182）的差异根因已定位**：我们的 demo 页面有 `product-header`（样例选择、props 面板），上游 docs 页只有 `[data-slot="component-preview"]` 演示框；且两侧容器高度不一致（上游演示多为固定 620/640px，我们 XLSX 用 `clamp(420px, 70vh, 760px)`）。L2 必须先固定 region + 尺寸，再谈像素。

## 1. 术语与分类（先读这一节）

### 1.1 三个“边界”

| 边界 | 含义 | 判定方法 | 例子 |
|---|---|---|---|
| 文件边界 | 是否在 pinned UI 仓库的公开文件里 | 文件是否存在于 `apps/v4/**` | `docx-viewer.tsx:1623` 使用了 `<DocxEditorViewer>`，但该组件的实现在公开文件之外 |
| 包边界 | 能力来自哪个 npm 包、哪个版本、什么许可证 | `apps/v4/package.json` + npm registry | `@extend-ai/react-docx 0.15.0`（MIT）；`@embedpdf/plugin-* 2.15.0`（Apache-2.0） |
| DOM 边界 | 哪些 DOM/class 由公开 wrapper 写、哪些由包内 runtime 写 | 公开文件的 JSX vs 包的 `data-*`/`rpv-*`/canvas | 公开文件写 toolbar；私有 runtime 写 `[data-docx-page-surface]`、`.rpv-stage`、grid canvas |

### 1.2 可行性分类（每个 viewer/editor 的每一层必须落在其中之一）

| 分类 | 定义 | 可用证据 | 不能做什么 |
|---|---|---|---|
| `public-vendored` | 上游公开文件 + 公开依赖即可渲染并对齐 | 源码路径 + 已存在的 0px L1 surface | 不需要 stub，不需要 provider |
| `public-partial` | 上游公开文件存在，但模块级 import 了不在 pinned 仓库的实现，或只覆盖部分状态 | 公开文件 + 包版本；需要裁剪/替换该 import | 不能声称“整页已对齐” |
| `office-reference` | 公开层与私有层都没有该布局；只能以 Office/Keynote/LibreOffice/ONLYOFFICE 等产品为参考 | 产品文档、截图、可运行产品 | 不能伪造上游像素对照，也不能写 `unreachable` 收尾 |
| `provider-dependent` | 数据/几何只在 runtime 内产生，只能通过回调或 provider 暴露 | 包内类型（如 `getObjectChart`）+ 宿主数据来源 | 不能承诺像素门禁，只能承诺 provider 合同 + Surface 消费 |
| `product-gap` | 能力缺失，且是否要做属于产品范围决策 | 我们的包/组件清单 + 宿主消费方 | 不能用“私有引擎”当理由；必须写产品理由与 owner |

`unreachable` 与 `manual` 不再是分类，只允许作为**报告状态**出现，且必须能映射到上表之一。映射规则：

| 旧状态 | 新分类判定 |
|---|---|
| `unreachable` 且上游确实没有该布局 | `office-reference`（有产品参考）或 `product-gap`（无参考、无消费方） |
| `unreachable` 且上游有布局但依赖私有 runtime 数据 | `provider-dependent` |
| `manual` | 拆成 `office-reference`（有规格）或 `product-gap`（无规格） |

### 1.3 层与状态

| 层 | 包 | 职责 |
|---|---|---|
| `core` | `docx-core` / `xlsx-core` / `pptx-core` / `office-runtime` | 解析、模型、布局、Worker、导出、加载策略 |
| `interaction` | `office-interaction` + 各 Surface | 对象描述、引用、候选、确认/取消、几何 |
| `vue` | `vue-docx` / `vue-xlsx` / `vue-pptx` / `vue-pdf` | viewer/editor/surface 组件与 controller/composable |
| `ui` | `vue-ui` | primitive、共享 Office chrome、blocks |
| `reference` | `tools/extend-reference`、`tools/parity-fixtures`、`tools/vue-reference` | L1 像素对照与基线 |

状态定义：`done`（有实现 + 有门禁/测试证据）、`partial`（有实现，缺状态/门禁/格式覆盖）、`missing`（没有实现）。

## 2. 总表：format × layer × status × owner

### 2.1 矩阵

| 格式 | core | interaction | vue（viewer/editor） | ui / chrome | reference（L1） | 结论 | owner |
|---|---|---|---|---|---|---|---|
| **PDF viewer** | done（pdfium worker + render runtime，`vue-pdf/src/pdf/`） | done（page/text/region/visual/native/behavior 引用） | done（`PdfViewer`、`PdfSurface`） | done（`pdf-toolbar`、`pdf-sidebar` 0px） | done（`pdf-toolbar`/`pdf-sidebar` blocking） | public-vendored，已收口 | #162 |
| **PDF editor** | partial（annotation 映射、表单值；无编辑事务） | partial（引用可用，编辑命令缺失） | missing（无 `PdfEditor`） | partial（`PdfEditorEmptyState` → `OfficeEmptyState`） | missing（无 editor surface） | public-vendored，未实现 | #172 + 新 issue |
| **DOCX viewer/editor** | done（port @6f70b92） | done | done（`DocxViewer`/`DocxEditor`/`DocxDocumentSurface`） | done（`docx-toolbar`/`docx-sidebar` 0px） | partial（DOCX chrome 参考是**裁剪 vendored 副本**，不是原文件） | public-partial（需与 0.15.0 对齐） | #159 |
| **XLSX viewer/editor** | done（worker + duke-sheets wasm + chart 解析/authoring） | done | done（`XlsxViewer` 含 ribbon/sheet tabs/authoring） | done（3 条 surface 64/64） | partial（ribbon/公式行无上游对照） | public-vendored（chrome）+ product-gap（pivot/page setup） | #160 / #171 |
| **PPTX viewer** | done（解析 + playback + 对象索引） | done | done（`PptxViewer`/`PptxStage`/sorter） | done（`pptx-toolbar`/`pptx-sorter` 0px） | done（2 条 blocking；presenter pending） | public-vendored（chrome） | #161 |
| **PPTX presenter** | done（notes 自 OOXML 读取） | done（timer/next） | done（`PptxPresenterPanel`） | done（Extend primitives 重组） | **missing**（上游无布局，private 包也无） | **office-reference** | #183 |
| **PPTX chart internals** | partial（`getObjectChart` provider + snapshot） | missing（Surface 未接入） | missing | — | missing | **provider-dependent** | #184 / #176 |
| **CSV** | partial（仅经 XLSX 路径） | missing | missing（无 `CsvViewer`） | missing | missing | **product-gap** | #178 |
| **FileSystem** | missing（无文件树模型） | missing | missing | missing | missing | **product-gap** | #179 |
| **SchemaBuilder** | missing（无 schema 持久化模型） | missing | missing | missing | missing | **product-gap** | #179 |

### 2.2 上游规模与边界一览

| 格式 | 公开文件（行数） | 公开依赖 | 包边界 | 私有 runtime 规模 |
|---|---|---|---|---|
| PDF | `pdf-viewer.tsx` 3,032；`pdf-editor.tsx` 4,314；`pdf-editor-panels.tsx` 2,319；`-dialogs.tsx` 2,005；`-properties.tsx` 1,989；`-shared.tsx` 2,512；`-workspace.tsx` 125；`pdf-block-resizable-shell.tsx` 316 | `@embedpdf/*` 2.15（viewer 11 个插件、editor 25 个）、`@tanstack/react-virtual`、`react-resizable-panels` | **无**（0 个 `@extend-ai/*` import） | — |
| DOCX | `docx-viewer.tsx` 1,659；`docx-editor.tsx` 2,451；`docx-annotation-card.tsx` 185 | `@tanstack/react-virtual`、shadcn primitives | `@extend-ai/react-docx` 0.15.0（MIT） | `editor.tsx` 67,923 行 + `index.tsx` 1,249 行 + 8 个子包 |
| XLSX | `xlsx-viewer.tsx` 1,844；`xlsx-editor.tsx` 2,835；`xlsx-controls.tsx` 464；`xlsx-pivot-controls.tsx` 2,034；`xlsx-pivot-header.tsx` 503；`workbook-select.tsx` 8 | 同上 | `@extend-ai/react-xlsx` 0.23.0（MIT） | `XlsxViewer.tsx` 19,030 行 + `controller.tsx` 4,845 行 |
| PPTX | `pptx-viewer.tsx` 1,419 | `@extend-ai/react-pptx/styles.css`、`@tanstack/react-virtual` | `@extend-ai/react-pptx` 0.5.0（MIT） | `dist/index.js` 7,851 行 + wasm（`pptx_wasm_bg.wasm`） |
| CSV | `csv-viewer.tsx` 975 | `@glideapps/glide-data-grid` 6.0.4-alpha24（MIT）、`papaparse`（MIT） | 无 | — |
| FileSystem | `file-system.tsx` 5,381；`document-preview-card.tsx` 71 | `@pierre/trees` 1.0.0-beta.4（Apache-2.0） | 无 | — |
| SchemaBuilder | `schema-builder.tsx` 2,703 | `@dnd-kit/*`（MIT）、`@pierre/diffs` ^1.2.7（npm latest 1.6.4 同为 Apache-2.0）、`next-themes` | 无 | — |

### 2.3 缺口收口表（每个 missing 行的决定）

| 缺口 | 分类 | 结论 | 产品理由 / 前置 | owner |
|---|---|---|---|---|
| PDF editor 外壳与面板 | public-vendored | **实现** | 上游 8 个文件全公开、EmbedPDF 插件全 Apache-2.0；引用/表单/签名是 PDF 的核心编辑价值 | #172（先决策 Surface toolbar）→ 新 issue |
| PDF annotation/form/stamp/redaction/export | public-vendored | **实现（分期）** | 先注释 + 表单（读/写），再做 stamp/redaction/export | 新 issue |
| DOCX chrome 与 0.15.0 对齐 | public-partial | **调查 → 实现** | 我们的 port 基线 `6f70b92` 在本地浅克隆里不可解析；需与 0.15.0 逐文件 diff | #159 follow-up |
| XLSX page setup dialog | product-gap | **调查（#171 A/B）** | 上游 `xlsx-editor.tsx` 没有独立 dialog；只有 `XlsxPrintPreview` 读 pageSetup | #171 |
| XLSX pivot 控件 | product-gap | **不计划（0.8/0.9）** | 上游 pivot 是 2,034+503 行专用 UI，且需要 pivot 编辑能力；无宿主消费方 | #160 文档已声明 |
| PPTX presenter | office-reference | **调查 → 实现** | 公开层与 0.5.0 包内都没有 presenter；必须由 Office 产品规格定义（#183） | #183 |
| PPTX chart internals | provider-dependent | **定义 provider → 接入** | 几何只在 renderer 内产生；`pptx-core` 已有 `PptxChartSnapshot` 与 `getObjectChart` 钩子 | #184 / #176 |
| CSV 独立 block | product-gap | **调查（ADR #178）** | 上游全公开可 vendored；问题是要不要第二套 grid 运行时 | #178 |
| FileSystem | product-gap | **调查 → 决定（#179）** | 依赖公开（Apache-2.0），真正成本是文件模型 + 四个 viewer 集成 | #179 |
| SchemaBuilder | product-gap | **调查 → 决定（#179）** | 缺 schema 持久化契约与宿主入口，不是许可证问题 | #179 |

## 3. 分层总图

```
上游（pinned repo）                         我们的对应层
┌───────────────────────────────┐          ┌──────────────────────────────┐
│ public chrome                 │          │ vue-* components             │
│  toolbar / sidebar / dialogs  │  ──────▶ │  DocxViewer/DocxEditor       │
│  context menus / status       │  vendor  │  XlsxViewer/XlsxRibbon       │
│  (imports @extend-ai/* 或     │          │  PptxViewer/PptxStage        │
│   @embedpdf/* / glide / …)    │          │  PdfViewer/PdfSurface        │
└──────────────┬────────────────┘          └──────────────┬───────────────┘
               │ 包边界                                    │
┌──────────────▼────────────────┐          ┌──────────────▼───────────────┐
│ runtime package (MIT, 独立版本)│  ──────▶ │ *-core（移植/自研引擎）       │
│  react-docx / react-xlsx /    │  port    │  docx-core / xlsx-core /     │
│  react-pptx / EmbedPDF /      │          │  pptx-core / office-runtime  │
│  glide-data-grid / pierre     │          └──────────────┬───────────────┘
└───────────────────────────────┘                         │
                                            ┌─────────────▼──────────────┐
                                            │ office-interaction          │
                                            │ 引用 / 选择 / 描述 / 几何    │
                                            └─────────────┬──────────────┘
                                                          │
                                            ┌─────────────▼──────────────┐
                                            │ vue-ui                      │
                                            │ primitives / chrome / blocks│
                                            └─────────────┬──────────────┘
                                                          │
                                            ┌─────────────▼──────────────┐
                                            │ L1 surfaces + L2 routes     │
                                            │ tools/parity-fixtures       │
                                            └────────────────────────────┘
```

---

## 4. PDF

### 4.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `pdf-viewer.tsx` | 3,032 | `PDFViewer`（`React.forwardRef`，:2,866） | 查看器：toolbar（`PDFViewerToolbar` :1,067）、搜索（:800）、页码（:716）、缩略图（:1,405）、滚动/文本选择层（:1,772）、file actions（:623） | toolbar/侧栏 DOM 在公开文件；页面 canvas 由 `RenderLayer`/`TilingLayer` 写 |
| `pdf-editor.tsx` | 4,314 | `PDFEditor`（:4,069） | 编辑器外壳：25 个 EmbedPDF 插件注册、注释/表单/签名/图章/密文/导出/打印 | 公开文件写面板骨架；注释/表单图层由插件写 |
| `pdf-editor-panels.tsx` | 2,319 | `PdfEditorThumbnailsPanel`/`OutlinePanel`/`AttachmentsPanel`/`PagesPanel`/`CommentsPanel`/`RedactionsPanel`/`StampsPanel`/`SignaturesPanel` | 左右侧栏面板 | 公开 |
| `pdf-editor-dialogs.tsx` | 2,005 | 打印、截取、签名、安全、属性对话框 | — | 公开 |
| `pdf-editor-properties.tsx` | 1,989 | 注释/表单字段属性面板 | — | 公开 |
| `pdf-editor-shared.tsx` | 2,512 | `PdfEditorEmptyState`、toolbar 按钮、色板、滚动容器、字段标签 | — | 公开 |
| `pdf-editor-workspace.tsx` | 125 | `PdfEditorWorkspace`（resizable 三栏） | — | 公开 |
| `pdf-block-resizable-shell.tsx` | 316 | 文档站 block 外壳 | — | 公开 |

**边界结论**：`rg "@extend-ai" pdf-*.tsx` 为 0。PDF 没有私有包边界；唯一外部大依赖是 EmbedPDF（`@embedpdf/core`、`@embedpdf/models`、25 个 `@embedpdf/plugin-*`，Apache-2.0）。

### 4.2 状态 / 命令 / 渲染管线

| 环节 | 上游实现（公开） | 我们（`vue-pdf`） | 差异 |
|---|---|---|---|
| 文档会话 | `EmbedPDF` + `createPluginRegistration` + `DocumentManagerPluginPackage`；`useActiveDocument`/`useDocumentManagerCapability` | `createPdfRenderRuntime()`（`@embedpdf/engines/pdfium-worker-engine`）+ `pdf-render-runtime.ts` | 我们只取 engine + models，不注册 plugin registry |
| 状态 | `useScroll`、`useZoom`、`useSearch`、`useSelectionCapability`、`useThumbnailPlugin`、`useDocumentPermissions`、`useAnnotation`、`useFormCapability`、`useRedaction`、`useSignature*`、`useStampLibraries`、`useExport`、`usePrint` | `PdfViewer` 内部 ref（`searchDraft`/`searchHits`/`pageLayout`/`zoom`）+ `PdfSurface` expose | 我们是“受控组件 + expose”，上游是“插件 capability” |
| 选择 | Selection 插件给 glyph 级 `PageTextSlice`；Annotation 插件给注释选择 | `pdf-selection-range.ts`/`pdf-selection-interaction.ts`（同样用 `@embedpdf/plugin-selection` 的 `glyphAt`/`expandToWordBoundary`） | 已对齐 glyph 模型；注释选择未实现 |
| 渲染 / worker | `RenderLayer`/`TilingLayer` + PDFium worker engine；DPR/scale 由插件算 | `pdf-page-raster.ts` + `pdf-render-runtime.ts`（自管 DPR/scale/像素预算） | 上游用插件图层，我们自管光栅化 |
| 导出 / 打印 | `ExportPluginPackage`/`useExport`；`PrintPluginPackage`/`usePrint` | `showPrint` 走浏览器打印；无 PDF 导出 | 导出能力缺失 |
| 私有 API | 无 | — | — |

### 4.3 我们的层

| 层 | 覆盖 | 缺 |
|---|---|---|
| core | `vue-pdf/src/pdf/`：render runtime、页布局、页面签名、annotation 映射、outline、参考适配 | 编辑事务（注释增删改、表单写回）、导出 |
| interaction | page/text/region/visual/native/behavior 六类引用（`pdf-reference-adapter.ts`） | 注释对象引用 |
| vue | `PdfViewer`（toolbar/搜索/缩略图/outline/打印/旋转/下载）、`PdfSurface`（无 toolbar，见 #172） | `PdfEditor` 及面板 |
| ui | `pdf-toolbar` 40/40、`pdf-sidebar`；`OfficeEmptyState` 来自 `PdfEditorEmptyState` | 面板级组件（字段属性、签名、图章） |

### 4.4 参考可行性

| 能力 | 分类 | 证据 | 缺失项（精确到文件/章节） |
|---|---|---|---|
| Viewer chrome | `public-vendored` **done** | `pdf-viewer.tsx:1067`；`tools/parity-fixtures/manifest.json` `pdf-toolbar`/`pdf-sidebar` blocking | 无 |
| Editor 外壳 + 面板 | `public-vendored` **missing** | `pdf-editor.tsx:4069`、`pdf-editor-panels.tsx:144` 起 8 个 export | 我们缺 `PdfEditor.vue` 及 `PdfEditorPanels.vue`；上游 `pdf-editor-shared.tsx` 的 `PdfEditorToolButton`/`PdfEditorToolbarSeparator`/`PdfEditorPanelHeader` 组件族没有对应物 |
| 注释 / 表单 / 签名 / 图章 / 密文 | `public-vendored` **missing** | `pdf-editor.tsx` 中 `AnnotationPluginPackage`、`FormPluginPackage`、`Signature*`、`Stamp*`、`Redaction*` | 缺编辑命令层；`pdf-annotation.ts` 只做只读映射 |
| 导出 / 打印 | `public-vendored` **partial** | `ExportPluginPackage`、`PrintPluginPackage`（`pdf-editor.tsx:47`） | 只有浏览器打印；无 flatten/export |

### 4.5 实现方案

- **静态 vendor 集合**：`pdf-editor.tsx`、`pdf-editor-panels.tsx`、`pdf-editor-dialogs.tsx`、`pdf-editor-properties.tsx`、`pdf-editor-shared.tsx`、`pdf-editor-workspace.tsx`（6 文件 ≈ 13,000 行）+ `@embedpdf/plugin-{annotation,form,signature,stamp,redaction,export,print,attachment,bookmark,capture,fullscreen,history,pan,spread}`。
- **stub**：`@/components/ui/*` → `vue-ui` primitive；`react-resizable-panels` → `OfficeSplitPane`（需新增）；`@/components/icon-placeholder` → `HugeIcon`。
- **provider**：无（全部公开）；但 `@embedpdf` 版本要先对齐（上游 2.15.0，我们 2.14.4）。
- **工作量/风险**：L（8–12 人日）；风险：EmbedPDF 大版本升级会动 render/selection 代码；PDF 编辑是安全敏感面（表单写回需权限校验）。
- **顺序**：① #172 决定 `PdfSurface` 是否带 toolbar；② 升级 EmbedPDF；③ 建 `PdfEditor` 外壳 + 空态；④ 注释面板；⑤ 表单；⑥ 签名/图章；⑦ 导出。
- **验收**：L1 新增 `pdf-editor-toolbar`、`pdf-editor-panels`；L2 新增 `/pdf-editor`。

### 4.6 L1 / L2 验收

- L1：`pdf-editor-toolbar`（复用 `pdf-toolbar` 结构）、`pdf-editor-panels`（缩略图/outline/注释面板 3 状态）。
- L2：`/pdf-viewer`（已有）与 `/pdf-editor`（新增）对 `content/docs/components/pdf-viewer.mdx` / `pdf-editor.mdx`。

---

## 5. DOCX

### 5.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `docx-viewer.tsx` | 1,659 | `DocxViewerPreview`（:1,108） | toolbar（`DocxToolbar` :549）、缩略图侧栏（:784/:1,012）、文件菜单（:314）、页码（:464）、加载态（:302） | toolbar/侧栏 DOM 公开；页面由 `<DocxEditorViewer>`（:1,623）写 |
| `docx-editor.tsx` | 2,451 | `DocxEditorPreview`（:1,943） | 编辑 toolbar（`DocxEditorToolbar` :937，约 900 行）、右键菜单（:132）、缩略图（:1,870）、边框/行距/段落样式菜单 | 同上；`renderContextMenu`/`renderCommentCard`/`renderTrackedChangeCard` 由公开文件注入 |
| `docx-annotation-card.tsx` | 185 | `createDocxCommentCardRenderer` / `createDocxTrackedChangeCardRenderer` | 批注/修订卡片 | 公开 |

**私有边界（包）**：`@extend-ai/react-docx@0.15.0`。公开文件从它 import：

- 组件：`DocxEditorViewer`（`docx-viewer.tsx:1623`、`docx-editor.tsx:2426`）。
- hooks：`useDocxEditor`、`useDocxComments`、`useDocxPageLayout`、`useDocxTrackChanges`、`useDocxViewerThumbnails`、`useDocxBorders`、`useDocxLineSpacing`、`useDocxParagraphStyles`、`useDocxDocumentTheme`、`useDocxFormFields`、`useDocxImageWrapMenu`、`useDocxPagination`、`useDocxPageThumbnails`、`useDocxModel`。
- 类型：`DocxEditorController`、`DocxDocumentTheme`、`ViewerZoomLevel`、`ViewerZoomState`、`DocxContextMenuRenderProps` 等 70+ 个；公开 hook 共 14 个（含 `useDocxModel`）。

**私有 DOM**：`editor.tsx` 写 97 个 distinct `data-docx-*` 标记（`data-docx-page-surface`、`data-docx-paragraph-host`、`data-docx-table-index`、`data-docx-text-start/end`、`data-docx-form-field`、`data-docx-header-footer-overlay`…）。我们的 `vue-docx`/`docx-core` 已使用 127 个（含我们新增的注释/绘图/块标记）；数量已超过上游，但“上游每个标记都有对应”仍需逐项比对（放进 #159 follow-up）。

### 5.2 状态 / 命令 / 渲染管线

| 环节 | 上游 | 我们 | 差异 |
|---|---|---|---|
| 模型/布局 | wasm + `ooxml-core` + `doc-model` + `layout-engine`；字体度量 + pretext | `docx-core/engine` + `layout`（分页、section、page-segmentation） | 已移植；基线提交待核 |
| 状态 | `useDocxEditor()` 返回 `DocxEditorController`（model、status、selection、historyRevision、selectedFormField…，见 `editor.tsx:4430`）；`DocxEditorViewer` 消费 | `vue-docx/composables/useDocxEditor.ts`（1,108 行）+ `editor-*.ts` 20+ 个命令模块 | 结构不同（我们拆文件），行为按 #159 对齐 |
| 选择/引用 | `DocxEditorSelection`、`DocxTextRange`、`DocxFormFieldLocation` | `editor-selection.ts` + `docx-object-index.ts` + `reference-adapter.ts` | 我们有统一的 `OfficeObjectReference`（上游无） |
| 渲染 | React 页面树 + canvas 图片层 + `pageVirtualization` | `DocxPageBody.ts` + `DocxPageSurface.vue` + `DocxImageLayer.vue` + canvas | 我们走 Vue 组件树；分页数据同源 |
| Worker | `docx-import-worker`（解析） | `viewer/docx-import-worker.ts` + `docx-export-worker.ts` | 对齐；我们多了导出 worker |
| 导出 | serializer | `viewer/docx-export.ts` | 已实现 |

### 5.3 我们的层

| 层 | 覆盖 | 缺 |
|---|---|---|
| core | 189 文件 / 73,419 行：engine、layout、editor ops、viewer（import/export/thumbnail/fonts）、blocks、引用 | 与 0.15.0 的增量（见 5.4） |
| interaction | 文本 range/页/块/表单字段引用 + 修订/批注行为 | 嵌套表格父子层级、页眉/脚注字段（#176） |
| vue | `DocxViewer` 912 行、`DocxEditor` 1,116 行、`DocxToolbar`/`DocxViewerToolbar`、`DocxDocumentSurface`、批注/修订 gutter | — |
| ui | `docx-toolbar` 32/32、`docx-sidebar` 24/24 | editor toolbar 没有独立 parity surface（与 viewer toolbar 共用结构） |

### 5.4 参考可行性

| 能力 | 分类 | 证据 | 缺口 |
|---|---|---|---|
| Viewer/editor chrome | `public-vendored` **done（经裁剪）** | `tools/extend-reference/src/vendor/extend/docx-toolbar.tsx`、`docx-sidebar.tsx`；manifest 两条 blocking | 参考是**手工裁剪副本**，不是 `docx-viewer.tsx` 原文件 —— 因为原文件模块级 import 私有包 |
| 页面渲染/fidelity | `public-partial` **done（port）** | `packages/docx-core/src/index.ts:3` 标注 `6f70b92` | 该 commit 不在本地浅克隆（`git rev-parse --is-shallow-repository` = true）；需对 0.15.0 重新 diff |
| 注释/修订卡片 | `public-vendored` **done** | `docx-annotation-card.tsx`；我们 `DocxTrackedChangeGutter.vue` + comment 组件 | 卡片视觉细节未做 L1 surface |
| 公式/绘图/文本框 | `public-partial` | 私有 `math-content.tsx`、`drawing-*`、`textbox-content-surface.tsx` | 我们已有 drawing/textbox 子集；公式覆盖需逐项核 |

### 5.5 实现方案

1. **验证 port 基线**（调查，0.5 人日）：`git -C /tmp/react-docx fetch --unshallow` 后 `git show 6f70b92:packages/react-viewer/package.json`，确认与 0.15.0 一致；若不成立，以 0.15.0 为基线重跑 diff。
2. **建立“原文件参考”**（1–2 人日）：在 `tools/extend-reference` 增加一个 `route` 模式，直接 import `apps/v4/components/extend/docx-viewer.tsx` 并安装 `@extend-ai/react-docx@0.15.0`（公开 MIT），别名 `@/` → `apps/v4`，替换 `next/dynamic` 与 `next-themes`。这会把“裁剪副本”升级为“原文件 + 真实私有包”的参考。
3. **补齐 chrome 差异**：以 `docx-editor.tsx:937` 的 toolbar 分组逐项对照我们的 `DocxToolbar.vue`；缺失项进 #159 follow-up。
4. **风险**：私有包体积大（`editor.tsx` 67,923 行）会拖慢参考构建；建议只在 nightly 跑原文件参考，PR 阶段仍用裁剪副本。

### 5.6 L1 / L2 验收

- L1：`docx-toolbar`、`docx-sidebar`（已有）；新增 `docx-editor-toolbar`（同一 reference 文件的 editor 分支）。
- L2：`/docx-viewer` vs `/docs/components/docx-viewer`；`/docx-editor` vs `/docs/components/docx-editor`。

---

## 6. XLSX

### 6.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `xlsx-viewer.tsx` | 1,844 | `XlsxViewerPreview`（:1,490）、`XlsxWorkbookSurface`（:1,371）、`WorkbookToolbar`（:915）、`WorkbookSheetTabs`（:1,099）、`WorkbookTableHeaderMenu`（:592） | 查看器外壳：toolbar、sheet tabs、搜索 popover、表头菜单 | toolbar/sheet tabs/状态私有与公开混合：公开文件写，数据来自 `<XlsxViewer>` 内 hooks |
| `xlsx-editor.tsx` | 2,835 | `XlsxEditorSurface`（:2,458）、`XlsxEditorPreview`（:2,546）、`EditorToolbar`（:1,319） | 编辑 ribbon（字体/字号/颜色/边框/合并/数字格式/对齐/换行）、公式行 | ribbon 与公式行都在公开文件；但模块级 import 私有 hook，无法单独运行 |
| `xlsx-controls.tsx` | 464 | `renderWorkbookContextMenu`、`renderWorkbookChartLoading`、`renderWorkbookFormControl`、`renderWorkbookImageSelection` | 私有 runtime 的 render prop 实现 | 公开 |
| `xlsx-pivot-controls.tsx` | 2,034 | `WorkbookPivotControls` | 透视表控件 | 公开（数据来自私有 pivot hooks） |
| `xlsx-pivot-header.tsx` | 503 | `WorkbookPivotHeaderMenu` | 透视表头菜单 | 公开 |
| `workbook-select.tsx` | 8 | `WorkbookSelect` re-export | — | 公开 |

**私有边界**：`@extend-ai/react-xlsx@0.23.0`。公开文件 import：

- 组件：`XlsxViewer`（`xlsx-viewer.tsx:1439`，props 含 `experimentalCanvas`、`renderScroller`、`renderTableHeaderMenu`、`showDefaultToolbar`、`showImages`、`fileTooLargeState`、`loadingState`、`errorState`）、`XlsxViewerProvider`（:1,827）。
- hooks：`useXlsxViewer`、`useXlsxViewerController`、`useXlsxViewerEditing`、`useXlsxViewerSelection`、`useXlsxViewerZoom`、`useXlsxViewerThumbnails`、`useXlsxViewerPivots`、`useXlsxViewerCharts`、`useXlsxViewerImages`、`useXlsxViewerTables`、`useXlsxPivotTableEditor`。

**私有 DOM**：grid 由 canvas 渲染（`experimentalCanvas`），并写 `data-xlsx-cell`、`data-xlsx-row`、`data-xlsx-col-header`、`data-xlsx-chart-element-kind`、`data-xlsx-chart-series-index`、`data-xlsx-chart-legend-index`、`data-xlsx-pivot-toggle` 等标记；公开文件用 `[&_[data-xlsx-pivot-toggle]]:…` 反选样式（`xlsx-viewer.tsx:1442`）。

### 6.2 状态 / 命令 / 渲染管线

| 环节 | 上游 | 我们 | 差异 |
|---|---|---|---|
| 解析 | `@dukelib/sheets-wasm`（duke-sheets） | `xlsx-core` + `xlsx-worker.ts`（1,408 行）+ 同款 wasm | 同引擎（wasm 包名 `duke_sheets_wasm`） |
| 状态 | `useXlsxViewerController`（私有 4,845 行）+ `XlsxViewerProvider` context | `useXlsxViewerController`（2,010 行）+ `workbook-state.ts` | 同名不同实现；我们的 controller 是公开 API |
| 选择/编辑 | `useXlsxViewerSelection`/`useXlsxViewerEditing`；cell editor overlay | `selection.ts`/`editing.ts`/`cell-input.ts`；DOM/canvas 双模 grid | 我们默认 canvas，`renderMode="dom"` 可切真实 `role=gridcell` |
| 渲染 | canvas grid + SVG chart renderer（private `chart-renderer.tsx`） | `vue-xlsx/src/render/chart-*.tsx`（SVG，adapted from 上游 `chart-renderer.tsx`）+ `surface-regl.tsx` | 同源移植；3D surface 用 regl |
| Worker | 私有 runtime 内 | `XlsxWorkerClient` + worker | 对齐 |
| 导出/打印 | 私有 | `authoring.ts`、`print/`、`exportCsv` | 我们多了 pageSetup 读取与 print preview |

### 6.3 我们的层

| 层 | 覆盖 | 缺 |
|---|---|---|
| core | 70 文件 / 35,812 行：worker、charts 解析/导出/插入、authoring（merge/filter/structure/visibility/number format）、print、images/form controls、pivot 元数据、引用 | pivot 计算/编辑；page setup 写回 |
| interaction | 单元格/区域/对象引用、对象索引 | 表格行独立对象、z 序候选（#176） |
| vue | `XlsxViewer` 1,109 行 + `XlsxRibbon`/`XlsxSheetTabs`/`XlsxToolbar`/`XlsxStatusBar`/`XlsxAuthoringDialog`/`XlsxPrintPreview` | ribbon/公式行无上游像素对照；page setup dialog |
| ui | `xlsx-toolbar` 32/32、`xlsx-sheet-tabs` 24/24、`xlsx-status-bar` 8/8 | — |

### 6.4 参考可行性

| 能力 | 分类 | 证据 | 缺口 |
|---|---|---|---|
| toolbar / sheet tabs / status | `public-vendored` **done** | `tools/extend-reference/src/vendor/extend/xlsx-*.tsx`；64/64 pass | 无 |
| ribbon / 公式行 | `public-partial` | `xlsx-editor.tsx:1319` 的 `EditorToolbar` 与 :1355 起的公式行都在**公开文件**里，但模块级 import 了 `useXlsxViewerEditing`/`useXlsxViewerSelection` 等私有 hook，不能脱离 `@extend-ai/react-xlsx` 单独渲染 | 需要 L1 `xlsx-ribbon`：参考 = 安装 pinned 私有包后渲染同一公开文件（R2），或降级为结构断言 |
| authoring dialogs | `public-vendored` **done**（除 page setup） | `XlsxAuthoringDialog.vue` + 16 条组件测试 | page setup（#171） |
| pivot | `product-gap` | `xlsx-pivot-controls.tsx` 2,034 行；私有 pivot hooks | 0.8/0.9 不做（已有产品决策） |
| chart internals 几何 | `provider-dependent` | 私有 `chart-renderer.tsx`；我们 SVG 移植已有 axis/labels/legend | 精确 Office 布局规则（#184） |

### 6.5 实现方案

- **page setup（#171）**：先做 A/B 决策。若做：`XlsxPageSetupDialog.vue` = `OfficeDialog` + `xlsx-core/print/page-setup.ts` 读写 + 撤销；L1 `xlsx-page-setup`（参考为 `xlsx-editor.tsx` 无独立 dialog → **office-reference**，用 Excel 页面设置对话框截图）。
- **ribbon 对照**：在上游公开文件 `xlsx-editor.tsx` 外套一个只提供 controller 的 React 壳（安装 `@extend-ai/react-xlsx@0.23.0`，MIT），渲染 ribbon/公式行子集；若成本过高，明确标 `public-partial` 只做结构断言。
- **订正既有文档**：[components/ui-component-map.md:80](../components/ui-component-map.md) 写的“公式行是私有包内部的 `Group`”不准确：公式行在公开文件内，私有的是它依赖的 hook。
- **pivot**：保持 `not planned`，在 `docs/guide/xlsx.md` 与 #160 文档写明。
- **风险**：canvas 像素对照依赖字体度量；ribbon 参考取样需要私有包版本锁定。

### 6.6 L1 / L2 验收

- L1：新增 `xlsx-ribbon`（`public-partial` 或 office-reference）、`xlsx-page-setup`（若 #171 做）。
- L2：`/xlsx-viewer` vs `/docs/components/xlsx-viewer`；`/xlsx-editor`（demo 暂无，需新增）vs `xlsx-editor.mdx`。

---

## 7. PPTX

### 7.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `pptx-viewer.tsx` | 1,419 | `PptxViewerPreview`（:940，props :927） | toolbar（`PptxToolbar` :385）、缩略图 rail（`PptxSidebarThumbnail` :600、列表 :639/:859）、文件菜单（:230）、页码（:304） | toolbar/rail 公开；幻灯片舞台由 `<ReactPptxViewer>`（:1,375）写 |

**私有边界**：`@extend-ai/react-pptx@0.5.0`。公开文件 import `ReactPptxViewer`、`usePptxViewerThumbnails`、`ParsedPresentation`、`PptxViewerController`、`ViewerZoomLevel`、`PresentationSource`，并 `import "@extend-ai/react-pptx/styles.css"`（`pptx-viewer.tsx:17`）。

**私有 DOM**：`.rpv-root`、`.rpv-stage`、`.rpv-viewport`、`.rpv-workspace`、`.rpv-filmstrip`、`.rpv-thumbnail`、`.rpv-toolbar`、`.rpv-status`、`.rpv-notes`、`.rpv-diagnostics`、`.rpv-slide-label`、`.rpv-spinner`（`dist/index.css` 中 12 个类）。公开 wrapper 通过 `className="… [&_.rpv-stage]:min-h-full … [&_[data-rpv-list-item]]:…"` 反选（`pptx-viewer.tsx:1386`）。

### 7.2 状态 / 命令 / 渲染管线

| 环节 | 上游 | 我们 | 差异 |
|---|---|---|---|
| 解析 | wasm（`pptx_wasm_bg.wasm`） | `pptx-core/archive.ts` + `buildPresentation`（`@aiden0z/pptx-renderer`） | 渲染器不同：私有 rust-office wasm vs 开源 `@aiden0z/pptx-renderer` 1.2.4 + 本地 patch |
| 状态 | `ReactPptxViewer` 内部 controller（`PptxViewerController`）+ `usePptxViewerThumbnails` | `PptxViewer`/`PptxStage` + `usePptxDocument`（470 行）+ `pptx-core/browser` 会话 | 我们是“session + 受控 props”，上游是“组件内 controller” |
| 选择/引用 | 私有（SlideNode 层，无公开引用 API） | `pptx-object-index.ts`（slide/group/table-cell/chart-series/chart-point/chart-legend-entry）+ `reference-adapter.ts` | 我们领先：上游没有对象引用 |
| 播放 | 私有（`SlideTransition`/`SlideTiming` 类型；无公开播放控制器） | `pptx-core/playback/`（time-tree、track-compiler、transitions、identity、morph）+ `browser/track-player.ts` | 我们自研播放层 |
| 备注 | `showNotes` + `SlideNote[]` + `.rpv-notes` | `pptx-core/notes.ts` + `PptxPresenterPanel` | 我们从 OOXML 直读，不依赖私有包 |
| 导出 | 私有（无公开导出） | 无（仅下载原文件） | 导出缺失，不在 0.8 范围 |

### 7.3 我们的层

| 层 | 覆盖 | 缺 |
|---|---|---|
| core | 32 文件 / 9,434 行：archive/limits/media policy/notes/对象索引/引用/播放模型 | 导出；chart internals 接入 |
| interaction | 幻灯片/对象/表格行/图表系列/点/图例引用 | Surface 未消费 chart targets |
| vue | `PptxViewer` 1,720 行、`PptxStage` 915 行、`PptxViewerToolbar` 923 行、`PptxSlideCard`、`PptxThumbnail`、`PptxPresenterPanel` 90 行、`PptxLinkMenu` | — |
| ui | `pptx-toolbar` 40/40、`pptx-sorter` 24/24 | presenter（`office-reference`） |

### 7.4 参考可行性

| 能力 | 分类 | 证据 | 缺口 |
|---|---|---|---|
| toolbar / 缩略图 rail | `public-vendored` **done** | `tools/extend-reference/src/vendor/extend/pptx-toolbar.tsx`、`pptx-slide-card.tsx` | 无 |
| sorter 网格 | `public-vendored`（重组）**done** | 上游没有 sorter 网格；参考取 rail item + `file-thumbnail.tsx`，manifest 注明 | 上游无网格布局 |
| presenter | `office-reference` **missing** | `pptx-viewer.tsx` 无 presenter；`@extend-ai/react-pptx@0.5.0` 的 `dist` 中 `presenter` 出现 0 次；只有 `showNotes` 与 `.rpv-notes` | 需要 PowerPoint/Keynote/Google Slides/LibreOffice/ONLYOFFICE 的区域与尺寸规格（#183） |
| chart internals | `provider-dependent` | `pptx-core/src/reference-adapter.ts:54` `getObjectChart?`；`pptx-object-index.ts:218` 使用 `PptxChartSnapshot` | Surface 未接入；几何来源是 `@aiden0z/pptx-renderer` 的 chart 数据 |
| 备注 | `public-vendored`（自读 OOXML）**done** | `pptx-core/src/notes.ts`；`PptxPresenterPanel` | 上游 notes 无公开布局 |

### 7.5 实现方案

- **#183（presenter）**：产出 `docs/design/pptx-presenter-reference.md`（区域图、尺寸表、交互清单、截图来源），再决定 `pptx-presenter` 的验收方式：**office-reference 截图基线 + 组件测试**，不再用 `referenceStatus: "unreachable"`。
- **#184（chart internals）**：定义 `PptxChartGeometryProvider`（series/point rect、legend item、axis tick），由 renderer 适配器实现；`PptxStage` 消费 provider 输出做命中/高亮；L1 不承诺像素，只做 provider 合同测试。
- **#176（对象矩阵）**：把 PPTX 行从“provider 已有但未接入”改为“provider 合同 + Surface 接入”两步，附 fixture。
- **风险**：`@aiden0z/pptx-renderer` 需要本地 patch（`patches/@aiden0z__pptx-renderer@1.2.4.patch`），升级时要同步维护 patch 与 chart 几何适配。

### 7.6 L1 / L2 验收

- L1：`pptx-toolbar`、`pptx-sorter`（blocking）；`pptx-presenter` 改为 office-reference 基线。
- L2：`/pptx-viewer` vs `/docs/components/pptx-viewer`；演示模式（presenter）单独一页 `/pptx-present`，reference 为 Office 截图矩阵（非上游路由）。

---

## 8. CSV

### 8.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `csv-viewer.tsx` | 975 | `CsvViewer({ className, data, search })`（:622） | 解析（`parseDelimitedText`）、列定义、搜索 popover（:339）、缩放、上传、文件菜单（:259）、空/错误态 | 除 grid 外全部公开；grid 是 `glide.DataEditor`（:948，动态 import） |

**依赖**：`@glideapps/glide-data-grid` 6.0.4-alpha24（MIT，canvas 网格，动态 import + `dist/index.css`）、`papaparse`（MIT，静态 import，用于解析）。**没有私有包**。

### 8.2 状态 / 命令 / 渲染管线

| 环节 | 上游 | 我们 |
|---|---|---|
| 数据 | `data?: string`（CSV 文本）+ 上传文件 `file.text()`；只有 `headers/rows` 字符串数组，没有工作簿/单元格模型 | CSV → XLSX 转换：`convertCsvBufferToXlsx`（`vue-xlsx/src/composables/workbook-state.ts:585`）+ `Workbook.fromCsvString`（duke-sheets wasm） |
| 状态 | 组件内 `useState`（parsed/zoom/selection/search） | `XlsxViewer` controller（工作簿模型） |
| 渲染 | Glide canvas + 自定义 `Theme`（:700 起） | XLSX canvas grid |
| 搜索 | 内部 popover，`CSV_SEARCH_BATCH_ROW_COUNT = 500`、debounce 300ms | `useXlsxSearch`（XLSX 语义） |
| 导出 | 文件菜单下载原始 CSV | `exportCsv`（重新序列化，格式信息丢失） |

### 8.3 我们的层

| 层 | 覆盖 | 缺 |
|---|---|---|
| core | 仅 URL/格式识别（`xlsx-url-policy.ts:302`）；CSV→XLSX 转换实际住在 `vue-xlsx`（`workbook-state.ts:585`），**层次错位** | 独立 CSV 模型（分隔符、引号、编码探测） |
| interaction | 无 | CSV 单元格/区域引用 |
| vue | `XlsxViewer` 接受 `.csv` | `CsvViewer` |
| ui | 无 | CSV toolbar/搜索 |

### 8.4 参考可行性

| 能力 | 分类 | 证据 | 缺口 |
|---|---|---|---|
| CSV viewer block | `product-gap` | `apps/v4/components/extend/csv-viewer.tsx:622` 全公开（无 `@extend-ai/*`，grid 走 `glide.DataEditor` :948） | 我们没有该 block、没有 fixture、没有 surface；`docs/upstream/extend-ui.md` 仍写 `not-used`，与事实不一致 |
| CSV 读取能力 | `public-vendored`（经 XLSX）**partial** | `docs/guide/xlsx.md:449` “CSV 输入” | 无独立 block；无“只读/格式丢失/编码”限制说明 |

### 8.5 实现方案（#178 ADR 要回答的问题）

| 选项 | 内容 | 成本 | 风险 |
|---|---|---|---|
| A. 复用 `XlsxViewer` | 现状 + 文档化限制 | S | 视觉/交互与上游 CSV 不一致；无法做 CSV 专属 L1 |
| B. 独立 `CsvViewer`（推荐若要“与上游一致”） | vendor `csv-viewer.tsx`（975 行）+ `@glideapps/glide-data-grid` + `papaparse`；L1 `csv-viewer`；L2 `/csv` | M（3–5 人日） | 第二套 grid 运行时（canvas）；主题/缩放/搜索重复实现 |
| C. 不提供独立 block | 明确 `not planned` + 产品理由 | S | 与上游组件清单永久分叉 |

建议：**先做 A 的文档收口**（0.8.1），把 B 作为 0.9 候选；若 B，则 `CsvViewer.vue` 只依赖 `papaparse` + glide，不引入 XLSX controller。

### 8.6 L1 / L2 验收

- L1（若 B）：`csv-viewer`（default/search/empty/error 4 状态）。
- L2（若 B）：`/csv` vs `/docs/components/csv-tsv-viewer`。

---

## 9. FileSystem

### 9.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `file-system.tsx` | 5,381 | `FileSystem`（:1,440）；类型 `FileSystemItem`/`FileSystemProps`（:188–282） | Finder 风格浏览：icons/list/columns/gallery 四视图、面包屑、排序（:611）、筛选（:705）、搜索、重命名/新建/删除、拖放、缩略图、空态；文件打开对话框内嵌 PDF/DOCX/PPTX/XLSX 预览（`Lazy*Viewer` :167–186） | 根 `div[data-slot="file-system"]`（:2,173）；文件树由 `@pierre/trees` 的 `FileTree` 写 |
| `document-preview-card.tsx` | 71 | `DocumentPreviewCard` | 悬浮预览卡 | 公开 |

**依赖**：`@pierre/trees` 1.0.0-beta.4（Apache-2.0：`createFileTreeIconResolver`、`getBuiltInSpriteSheet`、`prepareFileTreeInput`、`FileTree`、`useFileTree`）+ `Command`/`Dialog`/`DropdownMenu`/`Popover`/`ContextMenu`/`Tabs` primitives + `FileThumbnail` + 四个 viewer 的 lazy 预览。

### 9.2 状态 / 命令 / 渲染管线

| 环节 | 上游 | 我们 |
|---|---|---|
| 数据模型 | `items: FileSystemItem[]`（扁平 manifest）+ `loadChildren({path,cursor})` 懒加载 + `getFileUrl` + `previewImageUrl(s)`/`loadPreviewImageUrl` | 无 |
| 状态 | 组件内 index（`FileSystemIndex` :295）+ 受控 `view`/`defaultPath`/`onSelectionChange` | 无 |
| 渲染 | `PierreFileTree` + 视图渲染器 + `FileThumbnail` | 无；`FileUpload`/`FileThumbnail` 覆盖上传/缩略图 |
| 预览 | 内嵌四个 viewer（懒加载） | 无 |
| 命令 | 重命名/新建/删除/移动/排序/筛选 | 无 |

### 9.3 我们的层

全部 `missing`。已有相邻能力：`FileUpload`（#163）、`FileThumbnail`（#163）、`DocumentSplits`（#164，页拆分）。

### 9.4 参考可行性

| 能力 | 分类 | 证据 | 缺口 |
|---|---|---|---|
| FileSystem block | `product-gap`（`reference-only`） | `file-system.tsx:1440`；`data-slot="file-system"`；`@pierre/trees` Apache-2.0 | `FileSystem.vue`、`DocumentPreviewCard.vue`、文件树模型、4 个视图、宿主 API |
| 依赖“私有”论据 | **不成立** | npm registry：`@pierre/trees@1.0.0` license=`apache-2.0`；`@pierre/diffs@1.6.4` license=`apache-2.0` | 文档需改为“工作量/产品范围”理由 |

### 9.5 实现方案（#179 决策）

- **不做的产品理由（若选不做）**：0.8/0.9 没有文件浏览宿主；`FileUpload` + `FileThumbnail` 已覆盖上传与缩略图；引入文件树要同时定义文件模型、权限/URL 解析、四个 viewer 的预览集成，属于新产品面而非 viewer 对齐。
- **若做**：拆 `file-system-core`（index/排序/筛选/懒加载）+ `FileSystem.vue` + `DocumentPreviewCard.vue`；provider：`getFileUrl`、`loadChildren`、`loadPreviewImageUrl`；L1 `file-system`（4 视图 × 空/加载/错误）；L2 `/file-system` vs `/docs/components/file-system`。
- **工作量/风险**：L（≥ #163 的 1.5 倍）；风险：`@pierre/trees` beta API、四 viewer 懒加载的包体积、拖放与键盘可达性。

### 9.6 L1 / L2 验收

- L1：`file-system`（icons/list/columns/gallery、空、加载、错误）。
- L2：`/file-system` vs `file-system.mdx` 演示（`file-system-docs.tsx:464`，容器内 `p-4 sm:p-6`）。

---

## 10. SchemaBuilder

### 10.1 上游组件树

| 文件 | 行数 | 公开入口 | 作用 | DOM 归属 |
|---|---:|---|---|---|
| `schema-builder.tsx` | 2,703 | `SchemaBuilderPanel`（:2,338）、`SAMPLE_SCHEMA`（:520）、类型 `SchemaBuilderSchema`/`Property`（:157–216） | 抽取 schema 编辑器：字段树、嵌套 object/array、enum、类型菜单、拖拽排序、JSON 预览 | 全部公开；行标记 `data-schema-builder-property-id`/`-row-id`（:233–234） |

**依赖**：`@dnd-kit/core`/`sortable`/`utilities`（MIT，拖拽）；`@pierre/diffs`（Apache-2.0，JSON 预览 + 虚拟化）；`next-themes`（主题）；`Collapsible`/`DropdownMenu`/`ScrollArea`/`Tabs` primitives。

### 10.2 状态 / 命令 / 渲染管线

| 环节 | 上游 | 我们 |
|---|---|---|
| 数据模型 | `SchemaBuilderSchema { properties: SchemaBuilderProperty[] }`；`SerializedSchema` 输出 JSON Schema | 无 |
| 状态 | 组件内 reducer（属性定位 `PropertyLocation`、移动预览 `PropertyMovePreview`） | 无 |
| 渲染 | 递归 `NestedEditor` + `ArrayItemsEditor` + `EnumEditor` + DndContext | 无 |
| 预览 | `@pierre/diffs` 的文件视图（worker pool） | 无 |

### 10.3 我们的层

全部 `missing`。

### 10.4 参考可行性

| 能力 | 分类 | 证据 | 缺口 |
|---|---|---|---|
| SchemaBuilder block | `product-gap` | `schema-builder.tsx:2338`；`@dnd-kit` MIT、`@pierre/diffs` Apache-2.0 | 无 schema 组件、无持久化契约、无宿主入口 |
| 许可证论据 | **不成立** | npm registry 同上 | 必须改用“缺 schema 持久化 + 无消费方” |

### 10.5 实现方案（#179 决策）

- **不做（推荐 0.9 前）**：产品没有 schema 编辑入口；`LayoutBlocks`/`BoundingBoxCitations` 只需要只读 review 数据。
- **若做**：先定义 `ExtractionSchema` 数据模型与宿主保存/发布流程；`SchemaBuilder.vue` 用 `@dnd-kit/vue`（或 `vue-draggable-plus`）替代 dnd-kit，JSON 预览用 `<pre>` + 轻量高亮替代 `@pierre/diffs`（避免第二套 worker pool）。
- **工作量/风险**：M–L（4–8 人日）；风险：嵌套拖拽 + 键盘可达性；schema 校验语义。

### 10.6 L1 / L2 验收

- L1：`schema-builder`（空/默认/嵌套/JSON 4 状态）。
- L2：`/schema-builder` vs `schema-builder.mdx`（`schema-builder-demo.tsx:9`，`h-[620px]`）。

---

## 11. #182：L1 与 L2 surface 计划

### 11.1 L1（组件格子）增补

| surface | reference 来源 | 状态 | owner |
|---|---|---|---|
| `docx-editor-toolbar` | `docx-editor.tsx:937`（vendored 或原文件参考） | 新增 | #159 |
| `xlsx-ribbon` | `xlsx-editor.tsx:1319` + 公式行 :1355（`public-partial`：公开文件依赖私有 hook） | 新增 | #160 |
| `xlsx-page-setup` | 无上游（office-reference） | 视 #171 | #171 |
| `pdf-editor-toolbar` / `pdf-editor-panels` | `pdf-editor.tsx:4069`、`pdf-editor-panels.tsx:144` | 新增 | #172 |
| `csv-viewer` | `csv-viewer.tsx:622` | 视 #178 | #178 |
| `file-system` | `file-system.tsx:1440` | 视 #179 | #179 |
| `schema-builder` | `schema-builder.tsx:2338` | 视 #179 | #179 |
| `pptx-presenter` | 无上游（office-reference 截图） | 改造 | #183 |

### 11.2 L2（route）计划

> reference route 指公开文档站 `https://www.extend.ai/ui/docs/components/<slug>`（本地对应 `apps/v4/app/(app)/docs/[[...slug]]` + `content/docs/components/<slug>.mdx`）；表中的路径省略站点前缀。

**Region 定义**

- Subject（我们的 demo）：`[data-testid="<format>-page"] .viewer-container`；`viewer-container` 目前不统一（PDF 页没有 `product-surface` 类，XLSX 页有），L2 前应先统一为 `data-route-region="viewer"`。
- Reference（上游 docs）：`[data-slot="component-preview"]` 内、**排除** `[data-slot="code"]` 的子节点；上游演示框固定 `rounded-xl border`，高度多为 620/640px。
- 两侧必须使用同一 fixture、同一主题、同一 viewport、**同一固定高度**（建议 720px，另加 620px 窄屏 case），否则几何比较无意义。

| route | reference route | reference 文件 | 排除的 demo 脚手架 | 阶段 |
|---|---|---|---|---|
| `/docx-viewer` | `/docs/components/docx-viewer` | `components/docx-viewer-docs.tsx:42` | 我们的 `product-header`（样例/props 面板）；上游 `[data-slot=code]` | P1 |
| `/docx-editor` | `/docs/components/docx-editor` | `components/docx-editor-docs.tsx:42` | 同上 | P1 |
| `/xlsx-viewer` | `/docs/components/xlsx-viewer` | `components/xlsx-viewer-docs.tsx:40` | 同上 + 我们的 props 面板 | P1 |
| `/pptx-viewer` | `/docs/components/pptx-viewer` | `components/pptx-viewer-demo.tsx:33` | 同上 | P1 |
| `/pdf-viewer` | `/docs/components/pdf-viewer` | `components/pdf-viewer-docs.tsx:207` | 同上 | P1 |
| `/pdf-editor` | `/docs/components/pdf-editor` | `components/pdf-editor-docs.tsx:100` | 新建 demo 页 | P3 |
| `/csv` | `/docs/components/csv-tsv-viewer` | `components/csv-viewer-docs.tsx:47` | 视 #178 | P3 |
| `/file-system` | `/docs/components/file-system` | `components/file-system-docs.tsx:464` | 视 #179 | P3 |
| `/schema-builder` | `/docs/components/schema-builder` | `components/schema-builder-demo.tsx:9` | 视 #179 | P3 |

**Reference 构建策略**

| 方案 | 内容 | 优点 | 缺点 |
|---|---|---|---|
| R1（现状） | `tools/extend-reference` 手工 vendored 组件 | 轻、PR 快 | 与上游原文件有漂移；无法覆盖完整 route |
| R2（推荐 P1） | 在 `tools/extend-reference` 增加 route 模式：**原样 import** `apps/v4/components/extend/*.tsx`，安装 pinned `@extend-ai/*`（MIT 公开），别名 `@/`→`apps/v4`，stub `next/dynamic`、`next-themes` | 真正“上游组件树”，能覆盖 L2 | 构建重（DOCX 私有包 67k 行）；需要 nightly |
| R3（备选） | 跑 pinned 上游 Next 应用（`pnpm --filter v4 dev`）并截图 | 最真实 | Next 16 + turbo + wasm，CI 成本最高 |

**实施步骤**

1. 统一 demo region 标记（`data-route-region="viewer"`）与固定高度；写 `docs/testing/ui-parity-ci.md` 的 route 章节。
2. 在 `tools/parity-fixtures` 新增 `routes.json`（route、region selector、viewport、theme、fixture）；`pixel-parity.mjs` 增加 `--routes` 模式。
3. 先 report-only 跑 P1 五条 route，输出差异分类：产品组件 / demo 脚手架 / 上游无对应。
4. 稳定后把 `/docx-viewer`、`/xlsx-viewer` 切 blocking；P3 随 #178/#179 推进。

**验收（沿用 #182）**：每条 route 有 reference/subject 截图；viewer 区域几何 0px、像素 ≤0.1%；demo 专属外壳有 exclusion 说明。

---

## 12. 新工作的决策树

```
要新增/补齐一个 viewer、editor 或 block 能力
│
├─ 1. pinned commit 有公开实现吗？
│   ├─ 有 ──▶ 1a. 公开文件能否在不安装 @extend-ai/* 的前提下渲染？
│   │         ├─ 能 ──▶ public-vendored
│   │         │         产出：static vendor 清单 + L1 surface + L2 route
│   │         └─ 不能 ──▶ 1b. 该 runtime 是否已移植到 *-core？
│   │                    ├─ 已移植 ──▶ public-partial
│   │                    │              产出：与 npm 版本逐文件 diff + 原文件参考（R2）
│   │                    └─ 未移植 ──▶ 先移植 core 或登记 product-gap（写明理由）
│   └─ 没有 ──▶ 2. 有可引用的产品规格（Office/Keynote/LibreOffice/ONLYOFFICE）吗？
│               ├─ 有 ──▶ office-reference
│               │         产出：区域/尺寸/交互规格 + 产品截图基线（不伪造上游像素）
│               └─ 没有 ──▶ 3. 数据/几何是否只在 runtime 内产生？
│                          ├─ 是 ──▶ provider-dependent
│                          │         产出：provider 接口 + Surface 消费 + 合同测试
│                          └─ 否 ──▶ product-gap
│                                    产出：产品理由 + owner + 明确 not planned/待需求
│
└─ 禁止：以 unreachable / manual 结束；以“私有引擎”作为唯一理由；
         在没有 L1 的情况下声称“整页已对齐”。
```

---

## 13. 仍缺 Office / 开源证据的调查项

| 项 | 已知 | 缺什么 | 证据来源 | owner |
|---|---|---|---|---|
| PPTX presenter 区域模型 | 上游与 `react-pptx@0.5.0` 都无 presenter；我们已有简化 rail（next/timer/notes） | current/next/notes/timer/controls/缩略图的区域、比例、断点；多显示器/全屏行为；CJK 长备注 | PowerPoint Presenter View、Keynote Presenter Display、Google Slides、LibreOffice Impress Console、ONLYOFFICE 截图/文档 | #183 |
| PPTX notes 映射 | `notes.ts` 直读 `notesMaster`/备注页 | 备注占位符继承、字体缩放、溢出策略 | OOXML `notesMaster` 规范 + Office 行为 | #183 |
| Chart internals 几何 | 我们有 `PptxChartSnapshot`（series/point/legend）+ XLSX SVG chart 布局 | plotArea/legend/axis/label 的碰撞与换行规则；百分比/日期轴 | LibreOffice chart、ONLYOFFICE、Univer、ECharts 布局（仅借鉴） | #184 |
| XLSX ribbon 参考 | ribbon/公式行都在公开 `xlsx-editor.tsx` 内，依赖私有 hook | 私有 controller 的最小可用喂数（哪些 hook 必需） | 安装 `@extend-ai/react-xlsx@0.23.0` 跑参考（R2） | #160 |
| DOCX port 基线 | 源码标注 `6f70b92` | 该 commit 与 0.15.0 的差异 | `git fetch --unshallow` + `git diff 6f70b92..<0.15.0 tag>` | #159 |
| FileSystem 文件模型 | 公开 props 完整 | 宿主是否真的需要浏览入口、权限/URL 契约 | 产品决策（demo 无消费方） | #179 |
| SchemaBuilder 持久化 | 公开 schema 类型 | `ExtractionSchema` 的保存/发布/校验契约 | 产品决策 | #179 |

---

## 14. 风险与顺序

| 风险 | 影响 | 缓解 |
|---|---|---|
| 把“私有包”当闭源 | 继续用 `unreachable` 关闭工作 | 本文 §1.1/§1.2：包边界 ≠ 闭源；所有包均 MIT/Apache |
| 手工 vendored 参考漂移 | L1 0px 但整页不一致 | R2 原文件参考；L2 route 门禁 |
| 两套 grid/chart/worker 运行时 | 包体与维护成本 | CSV 优先复用 XLSX；SchemaBuilder JSON 预览不用 `@pierre/diffs` |
| DOCX/XLSX port 基线过期 | 行为差异难归因 | 先做 0.15.0/0.23.0 的逐文件 diff，再补功能 |
| presenter 无参考 | 继续 `pending` | #183 产出 office-reference 规格 + 截图基线 |

**建议顺序**：#172（PDF surface 决策）→ #183（presenter 规格）→ #184（chart provider）→ #178（CSV ADR）→ #179（FileSystem/SchemaBuilder 决策）→ #171（page setup）→ #182 P1 routes → #176（对象矩阵收尾）。

---

## 附录 A：证据清单

| 证据 | 路径 / 命令 |
|---|---|
| 上游固定提交 | `/tmp/extend-ui` @ `70fe70c6`；`git -C /tmp/extend-ui log -1` |
| PDF 无私有 import | `rg "@extend-ai" /tmp/extend-ui/apps/v4/components/extend/pdf-*.tsx` → 0 |
| DOCX 私有 seam | `docx-viewer.tsx:1623`、`docx-editor.tsx:2426`；`@extend-ai/react-docx@0.15.0` |
| XLSX 私有 seam | `xlsx-viewer.tsx:1439`、`:1827`；`@extend-ai/react-xlsx@0.23.0` |
| PPTX 私有 seam | `pptx-viewer.tsx:1375`；`@extend-ai/react-pptx@0.5.0` |
| presenter 不存在 | `rg -i -o presenter /tmp/ref-react-pptx/package/dist/{index.js,index.d.ts,index.css} \| wc -l` → 0 |
| 私有包公开性 | `curl -o /dev/null -w '%{http_code}' https://api.github.com/repos/extend-hq/react-{docx,xlsx,pptx}` → 200 |
| XLSX 公式行归属 | `xlsx-editor.tsx:1355`（`formulaDraft`）、:1502（`commitFormula`）在公开文件内 |
| 许可证 | npm registry：`@pierre/trees` apache-2.0、`@pierre/diffs` apache-2.0、`@glideapps/glide-data-grid` MIT、`papaparse` MIT、`@dnd-kit/core` MIT、`@aiden0z/pptx-renderer` Apache-2.0、`@embedpdf/core` Apache-2.0 |
| 我们的 port 来源 | `packages/docx-core/src/index.ts:3`；`docs/xlsx-migration-architecture.md`（`f285a1c`） |
| CSV 路径 | `packages/vue-xlsx/src/composables/workbook-state.ts:585`；`docs/guide/xlsx.md:449` |
| PPTX renderer | `packages/pptx-core/src/browser.ts:8`（`@aiden0z/pptx-renderer`）+ `patches/@aiden0z__pptx-renderer@1.2.4.patch` |
| L1 surfaces | `tools/parity-fixtures/manifest.json`（26 条，25 blocking，1 pending） |
| L2 region | 上游 `[data-slot="component-preview"]`（`docx-viewer-docs.tsx:42` 等）；我们 `.viewer-container`（`apps/demo/src/pages/*.vue`） |
| demo 路由 | `apps/demo/src/main.ts`（hash router，19 条） |

## 附录 B：与 attribution 状态的关系

`docs/upstream/extend-ui.md` / `attribution.json` 的 `adapted/copied/reference-only/not-used` 是**归属状态**，不是可行性分类。两者映射：

| attribution | 本文章节里的可行性 |
|---|---|
| `adapted` / `copied` | `public-vendored` / `public-partial` |
| `reference-only` | `office-reference` / `provider-dependent` / `product-gap`（三选一，需在本文对应格式节写明） |
| `not-used` | `product-gap`（必须补产品理由）或 `office-reference` |

当前需要修正的 attribution 叙述（不在本 issue 改状态，由 owner issue 处理）：

- `csv-viewer.tsx` 写着 “CSV is out of scope for the 0.8 Office surfaces”，与 #178 的“重新评估”冲突 → 由 #178 决定后改成 `reference-only` + 产品理由，或实现。
- `file-system.tsx` / `schema-builder.tsx` 的 notes 把依赖描述成不可用；实际是 Apache-2.0，需要改写为工作量/产品范围理由（#179）。

## 附录 C：issue 更新建议（可直接贴到对应 issue）

- **#182**：L2 的前置条件不是“新建参考站”，而是 ① demo region 统一（`data-route-region="viewer"`）与固定高度；② reference 用 R2（原样 import 上游 wrapper + 安装 `@extend-ai/*`）。上游演示框 `[data-slot="component-preview"]` 不包含产品 header，我们的 `product-header` 必须排除。
- **#183**：`unreachableReason` 中“presenter 在私有 `@extend-ai/react-pptx` 内”不成立（0.5.0 dist 中 `presenter` 出现 0 次，只有 `showNotes`/`.rpv-notes`）。presenter 应归 `office-reference`；`pptx-presenter` 的验收改为 office-reference 截图 + 组件测试。
- **#184**：我们不是“没有 chart 布局”，而是 XLSX 有 SVG 移植、PPTX 只有 `PptxChartSnapshot` 数据；缺的是 provider 几何合同与 Surface 消费。建议把 #184 的产出拆成 “provider 接口草案” + “Office 几何规则文档”。
- **#176**：PPTX chart internals 行改 `provider-dependent`；PDF 附件/OCG/签名/权限保持 `provider-dependent`（视觉识别器需 provider）；DOCX 页眉/脚注字段与嵌套表格为 `product-gap`（可在 #159 follow-up 实现）。
- **#178**：CSV 的现状不是 `not-used`，是 `product-gap`；ADR 在 A/B/C 中选一个，B 需要新增 `csv-viewer` L1 + `/csv` L2。
- **#179**：FileSystem/SchemaBuilder 的“私有依赖”理由不成立（Apache-2.0）。若不做，写成 “无宿主消费方 + 文件模型/schema 契约缺失”；若做，按 §9.5/§10.5 拆 issue。
- **#171**：上游没有独立 page setup dialog；若做，参考是 office-reference（Excel 页面设置），不是上游像素。
- **#172**：PDF editor 的 toolbar 在 `pdf-editor.tsx` 内是公开的；`PdfSurface` 是否带 toolbar 与本 issue 的 A/B 决策应同时定，避免出现“editor 有 toolbar、Surface 没有”的双标准。
