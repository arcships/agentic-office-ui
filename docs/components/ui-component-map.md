# Extend UI → `@arcships/*` component map

> 0.8 批次（#165）。固定基准：[`extend-hq/ui@70fe70c6`](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)（MIT）。
> 逐文件来源、状态与许可证见 [upstream/extend-ui.md](../upstream/extend-ui.md)；
> 本表按**组件**重组同一份事实，并补上每个组件挂在哪条 `ui-parity` surface 与基线上。
>
> `node scripts/ci/verify-ui-docs.mjs` 会检查：`packages/vue-ui/src/components/ui/*.vue` 的每个导出、
> `tools/parity-fixtures/manifest.json` 的每个 surface 都在本表出现；新增组件必须同步更新这里。

命令、阈值与审批流程见 [docs/testing/ui-parity-ci.md](../testing/ui-parity-ci.md)；
从旧样式迁移到 token 的步骤见 [docs/guide/ui-style-migration.md](../guide/ui-style-migration.md)。

## 状态图例

| 标记 | 含义 |
|---|---|
| ✅ blocking | `gateStatus: blocking`：`pnpm ui:parity -- --only core --baseline` 在 PR 阶段即失败 |
| ⏳ #issue | `gateStatus: pending`：仍采集、仍进 artifact，但不阻断；issue 收尾后改 `gateStatus` 并补基线 |
| — | 没有可渲染的上游对照（静态容器或私有依赖），由 token 与结构断言覆盖 |

## 1. Primitive 层（`@arcships/vue-ui`，reka-ui）

上游 `apps/v4/components/ui/*.tsx` 是文档站入口，多数只是 `apps/v4/registry/new-york-v4/ui/*.tsx`
的 re-export；`button`、`toggle`、`scroll-area` 在 `components/ui` 内直接实现。

| 上游 Extend 文件 | 状态 | 本仓库导出 | parity surface | 基线 | 备注 |
|---|---|---|---|---|---|
| `apps/v4/components/ui/button.tsx` | adapted | `OfficeButton` | `button` | ✅ blocking | 6 variant × 4 size + loading/disabled；[ui-primitives.md](ui-primitives.md) |
| `apps/v4/components/ui/button.tsx` | adapted | `OfficeIconButton` | `button`（`icon-sm`） | ✅ blocking | 复用 Button 的 token 与 focus ring，必有可访问名称 |
| `apps/v4/components/ui/input.tsx` | adapted | `OfficeInput` | `input` | ✅ blocking | focus ring 在 wrapper 的 `:has(:focus-visible)` 上；#157 已 40/40 像素 0 |
| `apps/v4/components/ui/input.tsx` | adapted | `OfficeTextarea` | — | — | 上游没有 `textarea.tsx`，多行控件沿用 Input token |
| `apps/v4/components/ui/select.tsx` | adapted | `OfficeSelect` | `select` | ✅ blocking | reka-ui listbox；`class` 合并到 trigger |
| `apps/v4/components/ui/dialog.tsx` | adapted | `OfficeDialog` | `dialog` | ✅ blocking | header/description/footer 槽位 + close |
| `apps/v4/components/ui/tooltip.tsx` | adapted | `OfficeTooltip` | `tooltip` | ✅ blocking | hover 触发，气泡为被测目标 |
| `apps/v4/components/ui/badge.tsx` | adapted | `OfficeBadge` | — | — | toolbar/状态区使用 |
| `apps/v4/components/ui/card.tsx` | adapted | `OfficeCard` | — | — | 静态容器 |
| `apps/v4/components/ui/spinner.tsx` | adapted | `OfficeSpinner`（`Spinner` 委托） | `button`（loading） | 随 Button | loading 旋转器 |
| `apps/v4/components/ui/skeleton.tsx` | adapted | `OfficeSkeleton` | `document-sidebar`（skeleton） | ✅ blocking（#158） | 参考侧与对照侧同像素 |
| `apps/v4/components/ui/separator.tsx` | adapted | `OfficeSeparator` | — | — | toolbar 分组分隔线 |
| `apps/v4/components/ui/popover.tsx` | adapted | `OfficePopover` | — | — | 宿主自定义浮层 |
| `apps/v4/components/ui/dropdown-menu.tsx` | adapted | `OfficeDropdownMenu` | — | — | 菜单项由 `OfficeMenuItems` 渲染 |
| `apps/v4/components/ui/context-menu.tsx` | adapted | `OfficeContextMenu` | — | — | 共享 `useOfficeMenuKeyboard` |
| `apps/v4/components/ui/sheet.tsx` | adapted | `OfficeSheet` | — | — | overlay 侧栏基础件 |
| `apps/v4/components/ui/tabs.tsx` | adapted | `OfficeTabs` | — | — | ribbon 分组 |
| `apps/v4/components/ui/toggle.tsx` | adapted | `OfficeToggle` | — | — | 工具按钮按下态 |
| `apps/v4/components/ui/scroll-area.tsx` | adapted | `OfficeScrollArea` | — | — | 侧栏/列表滚动容器 |
| `apps/v4/components/ui/sonner.tsx` | adapted | `OfficeToastProvider` / `OfficeToastViewport` | — | — | 视觉沿用 token，行为来自 reka-ui Toast |
| `apps/v4/components/ui/alert.tsx` | reference-only | `OfficeErrorState` | `office-error-state` | ✅ blocking | 信息层级参考，内容宿主可配置 |

## 2. 共享 Office chrome（#158）

| 上游 Extend 文件 | 状态 | 本仓库导出 | parity surface | 基线 | 备注 |
|---|---|---|---|---|---|
| `apps/v4/components/extend/pdf-viewer.tsx`（toolbar 行） | reference-only | `OfficeToolbar` + `OfficeToolbarGroup` + `OfficeToolbarSeparator` | `pdf-toolbar` | ✅ blocking | 40/40 像素一致（#162） |
| `apps/v4/components/extend/pdf-viewer.tsx`（page/zoom） | reference-only | `OfficePageControl` / `OfficeZoomControl` | `pdf-toolbar`, `docx-toolbar` | 见 surface | 页码与缩放输入 |
| `apps/v4/components/extend/pdf-viewer.tsx`（search） | reference-only | `OfficeSearchPopover` / `OfficeFindBar` | `docx-toolbar` | ✅ blocking | 搜索浮层与 find bar；默认行保持上游 toolbar |
| `apps/v4/components/extend/document-viewer-sidebar.tsx` | reference-only | `DocumentViewerSidebar` / `DocumentViewerSidebarSkeleton` | `pdf-sidebar`, `document-sidebar` | ✅ blocking（#158：40/40） | 160px rail，inline/overlay，closed 用 `-ml-40` / `-translate-x-full` |
| `apps/v4/components/extend/pdf-editor-shared.tsx`（`PdfEditorEmptyState`） | reference-only | `OfficeEmptyState` | `office-empty-state` | ✅ blocking（#158：16/16） | glyph 可选，`gap-2 px-4 py-10`，`max-w-56 text-xs` 描述 |
| `apps/v4/components/extend/xlsx-viewer.tsx`（`loadError` 卡片） | reference-only | `OfficeErrorState` | `office-error-state` | ✅ blocking（#158：24/24） | `max-w-md rounded-lg border p-4 text-sm` + `mt-4` action；长消息按卡片宽度换行 |
| `apps/v4/components/extend/docx-viewer.tsx` / `xlsx-viewer.tsx`（`ViewerLoadingSurface`） | reference-only | `OfficeLoadingState` | `office-loading-state` | ✅ blocking（#158：16/16） | `grid min-h-52 place-items-center` + 16px spinner |
| `apps/v4/components/extend/pdf-viewer.tsx`（footer 区） | reference-only | `OfficeStatusBar` | `office-status-bar` | ✅ blocking（#158：8/8） | `border-t bg-muted/40 px-3 py-2 text-xs`；左侧页码/状态、右侧文件名 |

## 3. 文件组件（#163）

| 上游 Extend 文件 | 状态 | 本仓库导出 | parity surface | 基线 | 备注 |
|---|---|---|---|---|---|
| `apps/v4/components/extend/file-upload.tsx` | adapted | `FileUpload` | `file-upload` | ✅ blocking | 6 状态；默认 `showFileList=false` 与宿主列表去重 |
| `apps/v4/components/extend/file-thumbnail.tsx` | adapted | `FileThumbnail` | `file-thumbnail` | ✅ blocking | 方形预览卡 + loading/error 图层 |

## 4. XLSX chrome（#160）

XLSX 用到的上游 chrome 已按功能切分：`@extend-ai/react-xlsx` 私有引擎仍不可分发，但
toolbar / sheet tabs / status strip 的 markup 与 class 串全部来自固定 commit，可直接渲染。

| 上游 Extend 文件 | 状态 | 本仓库导出 | parity surface | 基线 | 备注 |
|---|---|---|---|---|---|
| `apps/v4/components/extend/xlsx-viewer.tsx`（`WorkbookToolbar`） | reference-only | `XlsxToolbar`（`OfficeToolbar` + `OfficeIconButton` + `HugeIcon`） | `xlsx-toolbar` | ✅ blocking | 4 viewport × light/dark 共 32 case 全部 0px；下载/导出/上传/打印/分页在 ⋯ 菜单 |
| `apps/v4/components/extend/xlsx-viewer.tsx`（`WorkbookSheetTabsInner`） | reference-only | `XlsxSheetTabs` | `xlsx-sheet-tabs` | ✅ blocking | 24 case 全部 0px；新增/删除/翻页按钮改成 `show-sheet-actions`（默认关） |
| `apps/v4/components/extend/file-system.tsx`（`aria-live` 状态条） | reference-only | `XlsxStatusBar`（`OfficeStatusBar`） | `xlsx-status-bar` | ✅ blocking | 上游 XLSX 没有状态栏，取 Extend 自己的状态条 markup 作参考；8 case 全部 0px |
| `apps/v4/components/extend/xlsx-editor.tsx`（格式行） | adapted（静态 resting 子集，`vendor/extend/xlsx-ribbon.tsx`） | `XlsxRibbon`（`OfficeToolbar` + `OfficeButton`/`OfficeIconButton` + `HugeIcon`） | `xlsx-ribbon` | ✅ blocking（#170：32 case 全 0px） | undo/redo、字体/字号、B/I/U/S、文字/填充色、对齐、换行、数字格式、边框、合并；17 个 inline `<svg>` 全部换成 Hugeicons |
| `apps/v4/components/extend/xlsx-editor.tsx`（公式行） | reference-only | `XlsxFormulaBar` | — | — | 公式行编辑态依赖私有 `@extend-ai/react-xlsx` hook，暂无独立静息态对照 |

## 5. Block 组件（文档内容块，#164）

| 上游 Extend 文件 | 状态 | 本仓库导出 | parity surface | 基线 | 备注 |
|---|---|---|---|---|---|
| `apps/v4/components/extend/bounding-box-citations.tsx` + `registry/new-york-v4/blocks/bounding-box-citations-block.tsx` | adapted | `BoundingBoxCitations` | `bounding-box-citations` | ✅ blocking | Form 标量字段、Expected/Modified、空/加载/错误；8 状态 0px（#164） |
| `apps/v4/components/extend/layout-blocks.tsx` + `registry/new-york-v4/blocks/layout-blocks-block.tsx` | adapted | `LayoutBlocks` | `layout-blocks` | ✅ blocking | OCR 面板、叠加层、markdown；7 状态 0px（#164） |
| `apps/v4/registry/new-york-v4/blocks/e-signature.tsx` | adapted | `ESignatureFields` / `SignaturePad` | `e-signature` | ✅ blocking | Unsigned/Signed/Exporting/Preview；6 状态 0px（#164） |
| `apps/v4/components/extend/document-splits.tsx` + `registry/new-york-v4/blocks/document-splits-block.tsx` | adapted | `DocumentSplits` | `document-splits` | ✅ blocking | 区间/缩略图/增删/键盘排序；8 状态 0px（#164） |
| `apps/v4/components/extend/csv-viewer.tsx` | adapted | `CsvViewer` | — | — | 语义表格 + 行虚拟化替代 Glide canvas；解析/搜索/缩放/上传/下载和空/错/loading；结构测试 `tests/component/csv-viewer.test.mjs`，视觉差异见 [CSV ADR](../design/csv-viewer-adr.md)（#178） |
| `apps/v4/components/extend/file-system.tsx` + `registry/new-york-v4/blocks/file-system-block.tsx` | reference-only | — | — | — | 延期：私有文件系统依赖，已记录在 [document-blocks.md](document-blocks.md) |
| `apps/v4/registry/new-york-v4/blocks/schema-builder.tsx` | reference-only | — | — | — | 延期：独立表单构建器，已记录在 [document-blocks.md](document-blocks.md) |
| Extend 选区/标注 blocks | reference-only | `OfficeRegionSelector` / `OfficeRegionAnnotator` / `OfficeObjectOutlineLayer` | — | — | 规范坐标选区原语，见[组件手册](README.md) |

## 6. Viewer 外壳

| 包 | 组件 | 上游 Extend 文件 | parity surface | 基线 |
|---|---|---|---|---|
| `@arcships/vue-pdf` | `PdfViewer` | `apps/v4/components/extend/pdf-viewer.tsx` | `pdf-toolbar`, `pdf-sidebar` | ✅ blocking（#162） |
| `@arcships/vue-docx` | `DocxViewer` / `DocxEditor` | `apps/v4/components/extend/docx-viewer.tsx`, `docx-editor.tsx` | `docx-toolbar`, `docx-sidebar` | ✅ blocking（#159：32/32 + 24/24） |
| `@arcships/vue-xlsx` | `XlsxViewer` | `apps/v4/components/extend/xlsx-viewer.tsx`, `xlsx-editor.tsx` | `xlsx-toolbar`, `xlsx-sheet-tabs`, `xlsx-status-bar`, `xlsx-ribbon` | ✅ blocking（#160：32 + 24 + 8 case；#170：32 case，全部 0px） |
| `@arcships/vue-pptx` | `PptxViewer` / `PptxStage` | `apps/v4/components/extend/pptx-viewer.tsx` | 无上游可渲染 surface（presenter 布局私有） | 见 [pptx-chrome.md](pptx-chrome.md) |
| `@arcships/vue-pptx` | `PptxViewer` / `PptxViewerToolbar` / `PptxSlideCard` / `PptxPresenterPanel` | `apps/v4/components/extend/pptx-viewer.tsx` | `pptx-toolbar`, `pptx-sorter`（presenter 无上游布局） | ✅ blocking（#161：40/40 + 24/24 像素 0；`pptx-presenter` 为 unreachable） |
| `@arcships/vue-ui` | `CsvViewer` | `apps/v4/components/extend/csv-viewer.tsx` | 无（Glide canvas 与语义表格不可像素比较） | 结构/行为测试；视觉差异见 [CSV ADR](../design/csv-viewer-adr.md)（#178） |


## 7. 面板与文档

- Primitive 属性、键盘与 ARIA 合同：[ui-primitives.md](ui-primitives.md)
- 共享 chrome：[office-chrome.md](office-chrome.md)
- File 组件：[file-components.md](file-components.md)
- PDF chrome：[pdf-chrome.md](pdf-chrome.md)
- Token 表：[design/office-theme.md](../design/office-theme.md) 与 [design/extend-tokens.md](../design/extend-tokens.md)
