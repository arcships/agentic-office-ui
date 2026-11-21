# Extend UI 上游索引与开源归属

> 适用批次：0.8（#166）。本文件是 `extend-hq/ui` 的唯一人工索引，与机器可读的 [`attribution.json`](attribution.json) 同步；`scripts/ci/verify-attribution.mjs` 会检查两者是否一致、许可证是否完整、adapted/copied 文件是否带来源头、上游 URL 是否可达。

## 固定基准

| 项目 | 值 |
|---|---|
| 上游仓库 | `https://github.com/extend-hq/ui` |
| 固定提交 | `70fe70c6f1e1eee53954c6923f3d9f583557a57e`（2026-11-17） |
| 许可证 | MIT |
| 版权 | Copyright (c) 2026 CrowdView Inc, dba Extend；Portions Copyright (c) 2023 shadcn |
| 许可证全文 | [`licenses/extend-ui-MIT.txt`](licenses/extend-ui-MIT.txt) |
| 0.8 之前的固定基准 | `f2ff2f90954acfed1f60b7fd070a9491d4113047`，见 [`../upstream-parity-gap-audit.md`](../upstream-parity-gap-audit.md)，自 0.8 起由本文件取代 |

### 文档与源码 URL

- 组件文档：https://www.extend.ai/ui/docs/components
- Blocks 文档：https://www.extend.ai/ui/blocks
- 固定提交目录：https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e
- 固定许可证：https://github.com/extend-hq/ui/blob/70fe70c6f1e1eee53954c6923f3d9f583557a57e/LICENSE.md
- 全量 token：https://github.com/extend-hq/ui/blob/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/app/globals.css
- 字体定义：https://github.com/extend-hq/ui/blob/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/lib/fonts.ts
- shadcn primitives：https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/components/ui
- Extend 文档组件：https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/components/extend
- Extend blocks：https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/registry/new-york-v4/blocks

组件级文档页（由 #159–#164 使用）：

- https://www.extend.ai/ui/docs/components/docx-viewer
- https://www.extend.ai/ui/docs/components/docx-editor
- https://www.extend.ai/ui/docs/components/xlsx-viewer
- https://www.extend.ai/ui/docs/components/xlsx-editor
- https://www.extend.ai/ui/docs/components/pptx-viewer
- https://www.extend.ai/ui/docs/components/pdf-viewer
- https://www.extend.ai/ui/docs/components/pdf-editor
- https://www.extend.ai/ui/docs/components/file-upload
- https://www.extend.ai/ui/docs/components/file-thumbnail
- https://www.extend.ai/ui/docs/components/document-viewer-sidebar
- https://www.extend.ai/ui/docs/components/bounding-box-citations
- https://www.extend.ai/ui/docs/components/layout-blocks
- https://www.extend.ai/ui/docs/components/e-signature
- https://www.extend.ai/ui/docs/components/document-splits
- https://www.extend.ai/ui/docs/components/file-system

`www.extend.ai` 页面属于会随产品变化的站点；离线或站点不可用时使用 [`attribution.json`](attribution.json) 中记录的 Wayback 归档 URL。GitHub 固定提交页面不可变，因此不要求归档。

## 许可证副本

| 上游 | 许可证 | 版权 | 本地副本 |
|---|---|---|---|
| Extend UI（`extend-hq/ui`） | MIT | Copyright (c) 2026 CrowdView Inc, dba Extend；Portions Copyright (c) 2023 shadcn | [`licenses/extend-ui-MIT.txt`](licenses/extend-ui-MIT.txt) |
| `@extend-ai/react-docx` | MIT | Copyright (c) 2026 CrowdView Inc, dba Extend | [`licenses/extend-react-docx-MIT.txt`](licenses/extend-react-docx-MIT.txt) |
| `@extend-ai/react-xlsx` | MIT | Copyright (c) 2026 CrowdView Inc, dba Extend | [`licenses/extend-react-xlsx-MIT.txt`](licenses/extend-react-xlsx-MIT.txt) |
| shadcn-vue（`unovue/shadcn-vue`，仅参考） | MIT | Copyright (c) 2023 unovue | [`licenses/shadcn-vue-MIT.txt`](licenses/shadcn-vue-MIT.txt) |
| reka-ui（运行时 primitive） | MIT | Copyright (c) 2023 radix-vue | [`licenses/reka-ui-MIT.txt`](licenses/reka-ui-MIT.txt) |
| Lucide（`@lucide/vue`） | ISC | Copyright (c) 2026 Lucide Icons and Contributors | [`licenses/lucide-ISC.txt`](licenses/lucide-ISC.txt) |
| Geist / Geist Mono | SIL Open Font License 1.1 | Copyright (c) 2023 Vercel, in collaboration with basement.studio | [`licenses/geist-OFL-1.1.txt`](licenses/geist-OFL-1.1.txt) |
| Hugeicons（`@hugeicons/core-free-icons`） | MIT | Copyright (c) 2025 Hugeicons | [`licenses/hugeicons-MIT.txt`](licenses/hugeicons-MIT.txt) |

许可证文件是上游原文的逐字副本，不添加本地文件头。根仓库仍为 Apache-2.0（[`../../LICENSE`](../../LICENSE)），第三方声明集中在 [`../../THIRD_PARTY_NOTICES.md`](../../THIRD_PARTY_NOTICES.md)。

## 状态定义

| 状态 | 含义 | 是否需要文件头 |
|---|---|---|
| `reference-only` | 只参考设计、token、尺寸、交互或行为，代码独立实现 | 不需要，但必须在本索引登记 |
| `adapted` | 以上游文件为基础改写（结构、样式、逻辑有实质沿用） | 必须 |
| `copied` | 逐字或近乎逐字复制上游文件 | 必须 |
| `not-used` | 已评估但不使用 | 不需要，仍保留记录以证明评估范围 |

## 文件头格式

adapted/copied 文件必须在文件顶部（前 15 行内）包含：

```text
// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
```

注释语法按文件类型调整，文字内容不变：

| 文件类型 | 写法 |
|---|---|
| `.ts` / `.mjs` / `.cjs` / `.tsx` | `// Adapted from ...` |
| `.css` | `/* Adapted from ... */` |
| `.vue`（`<script setup>` 内） | `// Adapted from ...` |
| `.vue`（模板之外的 HTML 注释） | `<!-- Adapted from ... -->` |
| 逐字复制的文件 | 可用 `Copied from` 开头，其余文字相同 |

校验器只要求四类稳定标记（大小写不敏感）：`extend-hq/ui` + 固定 commit、许可证标识（`(MIT)`、`MIT License` 或 `SPDX-License-Identifier: MIT`）、`Copyright (c) 2026 CrowdView Inc, dba Extend`、`portions Copyright (c) 2023 shadcn`，并要求出现在前 15 行的注释行内。因此以下变体均合法：

```text
// Adapted from extend-hq/ui <commit> (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
```

```text
/**
 * Vendored from extend-hq/ui @ <commit>
 * Source: apps/v4/components/ui/button.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License.
 */
```

第二种写法用于逐字复制（`copied`）的多行头，但许可证标识不可省略。

## 索引

### 设计 token 与字体（#155、#156）

| Extend 文件 | 用途 | 我们的目标文件 | 状态 | 备注 |
|---|---|---|---|---|
| `apps/v4/app/globals.css` | 语义 token：radius、色阶、ring、shadow、spacing | `packages/vue-ui/src/styles/tokens.css` | adapted | 只映射数值，不复制 CSS；文件头记录了固定 commit，按 #167 反向检查归为 adapted |
| `apps/v4/app/legacy-themes.css` | 主题变体、暗色作用域、字体变量 | `packages/vue-ui/src/styles/theme.css` | adapted | 我们用 `[data-office-theme]` 作用域，选择器策略独立；文件头记录固定 commit |
| `apps/v4/app/shadcn-tailwind.css` | Tailwind 与 CSS 变量桥接 | — | not-used | 仓库不使用 Tailwind 或 preflight |
| `apps/v4/lib/fonts.ts` | Geist / Geist Mono / 文档字体族 | `packages/vue-ui/src/styles/fonts.css` | reference-only | 只参考字体栈与可变字体意图，不使用 next/font |
| `apps/v4/lib/font-definitions.ts` | 文档站字体元数据 | `packages/vue-ui/src/styles/fonts.css` | reference-only | 只有 Geist / Geist Mono 条目相关 |
| `apps/v4/lib/utils.ts` | Tailwind class 合并（cn） | `tools/extend-reference/src/lib/cn.ts` | adapted | 只有 #155 harness 需要 cn()；Vue 包不使用 Tailwind 辅助函数 |
| `apps/v4/lib/pdf-thumbnail-utils.ts` | PDF 页缩略图生成与缓存 | `packages/vue-pdf/src/components/PdfThumbnailStrip.vue` | reference-only | 我们已有 PDF 引擎缩略图，只参考布局与缓存行为 |
| `apps/v4/lib/pdf-viewer-blocks.ts` | 文档站 PDF demo 元数据 | — | not-used | 文档站注册信息，没有产品对应物 |

### 通用 primitive（#157）

`#157` 计划把 primitive 放在 `packages/vue-ui/src/components/ui/`；目录级约定见下文，目录一旦存在，里面的 Vue/CSS 文件都必须带来源头。

上游 `apps/v4/components/ui/*.tsx` 是文档站引用入口：其中 `accordion`、`alert`、`aspect-ratio`、`badge`、`breadcrumb`、`card`、`collapsible`、`command`、`context-menu`、`dialog`、`dropdown-menu`、`input`、`kbd`、`popover`、`select`、`separator`、`sheet`、`sidebar`、`skeleton`、`sonner`、`spinner`、`switch`、`tabs`、`toast`、`tooltip` 只是 `apps/v4/registry/new-york-v4/ui/*.tsx` 的 re-export；`button`、`toggle`、`scroll-area`、`color-picker`、`group`、`resizable` 在 `components/ui` 内直接实现（`button.tsx` 与 registry 版本内容一致）。复制或对照时以具体实现文件为准，并在提交说明里写清实际路径。

| Extend 文件 | 用途 | 我们的目标文件 | 状态 | 备注 |
|---|---|---|---|---|
| `apps/v4/components/ui/button.tsx` | Button 变体与尺寸 | `packages/vue-ui/src/components/ui/OfficeButton.vue` | adapted | 用 reka-ui slot 行为 + scoped CSS 重写；IconButton 复用 icon 尺寸 |
| `apps/v4/components/ui/button.tsx` | 图标按钮尺寸、ghost/active 与触控区 | `packages/vue-ui/src/components/ui/OfficeIconButton.vue` | adapted | 复用 Button；新增必需的可访问名称与 `aria-pressed` |
| `apps/v4/components/ui/input.tsx` | Input 样式与 focus ring | `packages/vue-ui/src/components/ui/OfficeInput.vue` | adapted | 高度、圆角、边框 alpha、focus-visible ring 来自上游 class 列表 |
| `apps/v4/components/ui/input.tsx` | Textarea 复用 Input 样式 | `packages/vue-ui/src/components/ui/OfficeTextarea.vue` | adapted | 上游没有 textarea.tsx，多行控件沿用 Input 的 token |
| `apps/v4/components/ui/select.tsx` | Select：portal、碰撞定位、滚动按钮 | `packages/vue-ui/src/components/ui/OfficeSelect.vue` | adapted | 行为交给 reka-ui，视觉层对齐上游 |
| `apps/v4/components/ui/dropdown-menu.tsx` | 下拉菜单、item、separator、快捷键 | `packages/vue-ui/src/components/ui/OfficeDropdownMenu.vue` | adapted | — |
| `apps/v4/components/ui/context-menu.tsx` | 右键上下文菜单 | `packages/vue-ui/src/components/ui/OfficeContextMenu.vue` | adapted | item 样式与 DropdownMenu 共用 |
| `apps/v4/components/ui/dialog.tsx` | 模态框、header/footer、关闭按钮、backdrop | `packages/vue-ui/src/components/ui/OfficeDialog.vue` | adapted | — |
| `apps/v4/components/ui/sheet.tsx` | 边缘抽屉 | `packages/vue-ui/src/components/ui/OfficeSheet.vue` | adapted | 基于 Dialog primitive |
| `apps/v4/components/ui/popover.tsx` | 锚定浮层与箭头 | `packages/vue-ui/src/components/ui/OfficePopover.vue` | adapted | — |
| `apps/v4/components/ui/tooltip.tsx` | Tooltip 浮层、延迟、偏移 | `packages/vue-ui/src/components/ui/OfficeTooltip.vue` | adapted | 现有 `components/Tooltip.vue` 保持公开 API，内部委托 OfficeTooltip |
| `apps/v4/components/ui/tabs.tsx` | Tab list / trigger / panel | `packages/vue-ui/src/components/ui/OfficeTabs.vue` | adapted | — |
| `apps/v4/components/ui/toggle.tsx` | Toggle 变体与按下态 | `packages/vue-ui/src/components/ui/OfficeToggle.vue` | adapted | — |
| `apps/v4/components/ui/badge.tsx` | Badge 变体与尺寸 | `packages/vue-ui/src/components/ui/OfficeBadge.vue` | adapted | — |
| `apps/v4/components/ui/card.tsx` | Card 外壳、header、content、footer | `packages/vue-ui/src/components/ui/OfficeCard.vue` | adapted | — |
| `apps/v4/components/ui/separator.tsx` | 方向分隔线 | `packages/vue-ui/src/components/ui/OfficeSeparator.vue` | adapted | — |
| `apps/v4/components/ui/skeleton.tsx` | 加载占位 | `packages/vue-ui/src/components/ui/OfficeSkeleton.vue` | adapted | — |
| `apps/v4/components/ui/spinner.tsx` | 加载指示 | `packages/vue-ui/src/components/ui/OfficeSpinner.vue` | adapted | 现有 `components/Spinner.vue` 保持公开 API，内部委托 OfficeSpinner |
| `apps/v4/components/ui/scroll-area.tsx` | 自定义滚动条区域 | `packages/vue-ui/src/components/ui/OfficeScrollArea.vue` | adapted | — |
| `apps/v4/components/ui/sonner.tsx` | Toast 外观配置 | `packages/vue-ui/src/components/ui/OfficeToastProvider.vue` | adapted | 只沿用视觉 token，行为来自 reka-ui Toast（未使用 vue-sonner） |
| `apps/v4/components/ui/alert.tsx` | 内联提示容器 | `packages/vue-ui/src/components/OfficeErrorState.vue` | reference-only | #158 的空/错状态保持宿主可配置，只借用信息层级 |
| `apps/v4/components/ui/color-picker.tsx` | 带预设与 alpha 的取色器 | `packages/vue-xlsx/src/components/XlsxAuthoringDialog.vue` | reference-only | #160 计划用于填充色/字体色；真正复制前必须复核状态 |
| `apps/v4/components/ui/accordion.tsx` | Accordion | — | not-used | Office 界面不需要 |
| `apps/v4/components/ui/aspect-ratio.tsx` | 比例容器 | — | not-used | 需要时直接用 CSS aspect-ratio |
| `apps/v4/components/ui/breadcrumb.tsx` | 面包屑 | — | not-used | 后置的 FileSystem 暂无面包屑计划 |
| `apps/v4/components/ui/collapsible.tsx` | 折叠 | — | not-used | DocumentViewerSidebar 只有 inline/overlay 两种模式 |
| `apps/v4/components/ui/command.tsx` | 命令面板 | — | not-used | 搜索弹层使用普通 input + 列表，过滤权归还宿主 |
| `apps/v4/components/ui/group.tsx` | 分段按钮组 | — | not-used | OfficeToolbarGroup 覆盖工具栏分组 |
| `apps/v4/components/ui/kbd.tsx` | 快捷键标记 | — | not-used | 用 aria-keyshortcuts + tooltip 表达 |
| `apps/v4/components/ui/resizable.tsx` | 面板拖拽 | — | not-used | 侧栏宽度用 ResizeObserver |
| `apps/v4/components/ui/sidebar.tsx` | 应用外壳侧栏 | — | not-used | 属于应用外壳，不是文档 viewer primitive |
| `apps/v4/components/ui/switch.tsx` | Switch | — | not-used | 设置项使用带标签的 Toggle |
| `apps/v4/components/ui/toast.tsx` | 旧版 toast | — | not-used | 上游已被 sonner 取代，本地使用 vue-sonner |

### 文档组件（#158–#164）

| Extend 文件 | 用途 | 我们的目标文件 | 状态 | 备注 |
|---|---|---|---|---|
| `apps/v4/components/extend/pdf-viewer.tsx` | PDF viewer：toolbar、页码、缩放、搜索、侧栏、下载、打印 | `packages/vue-pdf/src/components/PdfViewer.vue` | reference-only | #158 定义共用 chrome，#162 对齐 PDF |
| `apps/v4/components/extend/pdf-editor.tsx` | PDF editor 工作区组合 | `packages/vue-pdf/src/components/PdfEditor.vue` | reference-only | #162 计划中的编辑外壳 |
| `apps/v4/components/extend/pdf-editor-workspace.tsx` | 可编辑工作区与页面布局 | `packages/vue-pdf/src/components/PdfEditor.vue` | reference-only | 与 pdf-editor.tsx 同一目标 |
| `apps/v4/components/extend/pdf-editor-panels.tsx` | 签名、表单、批注侧栏 | `packages/vue-pdf/src/components/PdfEditorPanels.vue` | reference-only | #162 |
| `apps/v4/components/extend/pdf-editor-dialogs.tsx` | 打印、截取、安全、属性对话框 | `packages/vue-pdf/src/components/PdfEditorDialogs.vue` | reference-only | #162 |
| `apps/v4/components/extend/pdf-editor-properties.tsx` | 表单字段与文档属性面板 | `packages/vue-pdf/src/components/PdfEditorProperties.vue` | reference-only | #162 |
| `apps/v4/components/extend/pdf-editor-shared.tsx` | 共享 context、feature flag、颜色预设 | `packages/vue-pdf/src/components/PdfEditor.vue` | reference-only | #162 |
| `apps/v4/components/extend/document-viewer-sidebar.tsx` | inline/overlay 文档侧栏、skeleton、动画 | `packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue` | reference-only | 规范组件由 #158 在 `components/chrome/` 实现；#163 只记录缩略图条集成 seam，不重复实现 |
| `apps/v4/components/extend/file-upload.tsx` | 拖拽区、accept chips、文件列表、ready/error | `packages/vue-ui/src/components/FileUpload.vue` | adapted | #163 用 scoped CSS + `--office-*` token 重写；Hugeicons 提供与参考一致的 glyph 数据 |
| `apps/v4/components/extend/file-upload.tsx` | FileUpload 默认 accept 列表与图标定义 | `packages/vue-ui/src/components/file-upload-defaults.ts` | adapted | #163 抽到模块级，供 `defineProps()` 默认值引用；不渲染标记 |
| `apps/v4/components/extend/file-thumbnail.tsx` | 方形/可配置预览卡、blur-in、错误回退 | `packages/vue-ui/src/components/FileThumbnail.vue` | adapted | #163 重写 preview 卡、reveal、loading 与 error fallback |
| `apps/v4/components/extend/docx-viewer.tsx` | DOCX viewer chrome 与翻页 | `packages/vue-docx/src/components/DocxViewer.vue` | reference-only | #159 |
| `apps/v4/components/extend/docx-editor.tsx` | DOCX editor chrome 与格式控件 | `packages/vue-docx/src/components/DocxEditor.vue` | reference-only | #159 |
| `apps/v4/components/extend/docx-annotation-card.tsx` | 评论与修订 gutter 卡片 | `packages/vue-docx/src/components/DocxAnnotationCard.vue` | reference-only | #159 |
| `apps/v4/components/extend/xlsx-viewer.tsx` | XLSX ribbon、sheet tabs、状态栏、菜单 | `packages/vue-xlsx/src/components/XlsxViewer.vue` | reference-only | #160 |
| `apps/v4/components/extend/xlsx-editor.tsx` | XLSX 编辑 ribbon 与作者入口 | `packages/vue-xlsx/src/components/XlsxViewer.vue` | reference-only | #160 |
| `apps/v4/components/extend/xlsx-editor.tsx` | XLSX 编辑 ribbon 格式行（静态 resting 子集） | `tools/extend-reference/src/vendor/extend/xlsx-ribbon.tsx` | adapted | #170：私有 `useXlsxViewer*` hook 改成 props，select/popover/dropdown 弹层不渲染；pivot/sheet 选择器/zoom/公式行不在该子集内 |
| `apps/v4/components/extend/xlsx-controls.tsx` | ribbon 按钮、输入、下拉 | `packages/vue-xlsx/src/components/XlsxRibbon.vue` | reference-only | #160 |
| `apps/v4/components/extend/workbook-select.tsx` | 工作簿/sheet 选择 | `packages/vue-xlsx/src/components/XlsxSheetTabs.vue` | reference-only | #160 |
| `apps/v4/components/extend/pptx-viewer.tsx` | PPTX viewer、presenter、sorter chrome | `packages/vue-pptx/src/PptxViewer.vue` | reference-only | #161 |
| `apps/v4/components/extend/bounding-box-citations.tsx` | 引用框浮层、框、侧栏列表 | `packages/vue-ui/src/components/BoundingBoxCitations.vue` | adapted | #164：Form 标签页、标量字段编辑、高亮与活动字段；JSON 标签页（pierre/diffs）与数组/对象网格编辑器保持 reference-only |
| `apps/v4/components/extend/layout-blocks.tsx` | Layout blocks viewer、OCR blocks、PDF 叠加 | `packages/vue-ui/src/components/LayoutBlocks.vue` | adapted | #164：页图/宿主 viewer slot、block 叠加几何、类型配色 |
| `apps/v4/components/extend/layout-blocks-panel.tsx` | block 列表、置信度、markdown | `packages/vue-ui/src/components/LayoutBlocks.vue` | adapted | #164：卡片列表、置信度/页码、空/加载/错误、安全 markdown 子集 |
| `apps/v4/components/extend/document-splits.tsx` | 拆分区间选择与结果列表 | `packages/vue-ui/src/components/DocumentSplits.vue` | adapted | #164：拆分卡片、页缩略图、增删、拖放与键盘重排 |
| `apps/v4/components/extend/file-system.tsx` | Finder 风格文件浏览 | `packages/vue-ui/src/components/FileSystem.vue` | reference-only | #164 评估后暂缓：上游 5,381 行，依赖完整文件模型、面包屑、预览卡与 DocumentPreviewCard，0.8 无宿主消费方；证据见 [`../components/document-blocks.md`](../components/document-blocks.md#filesystem-暂缓) |
| `apps/v4/components/extend/document-preview-card.tsx` | 文档引用悬浮预览卡 | `packages/vue-ui/src/components/DocumentPreviewCard.vue` | reference-only | 与 FileSystem 一并后置，先登记避免以后从头找来源 |
| `apps/v4/components/extend/pdf-block-resizable-shell.tsx` | 文档站 PDF block 可缩放外壳 | `packages/vue-ui/src/components/LayoutBlocks.vue` | reference-only | 只有 PDF 叠加块支持缩放时才相关 |
| `apps/v4/components/extend/document-scroll-content.tsx` | Radix ScrollArea content 的薄封装 | — | not-used | 没有可移植行为 |
| `apps/v4/components/extend/csv-viewer.tsx` | CSV 表格查看器 | — | not-used | CSV 不在 0.8 范围 |
| `apps/v4/components/extend/schema-builder.tsx` | 抽取 schema 编辑器 | — | not-used | #164 评估后延期：上游 2,703 行（嵌套字段/类型编辑、校验、抽取预览），0.8 无消费方；证据见 [`../components/document-blocks.md`](../components/document-blocks.md#schemabuilder-暂缓) |
| `apps/v4/components/extend/ui-toast.tsx` | Toast manager provider | — | not-used | 通知归宿主，库不内置 toast manager |
| `apps/v4/components/extend/xlsx-pivot-controls.tsx` | 数据透视表控件 | — | not-used | 透视表不在 0.8 范围 |
| `apps/v4/components/extend/xlsx-pivot-header.tsx` | 透视表表头菜单 | — | not-used | 透视表不在 0.8 范围 |

### Blocks（#159–#164）

| Extend 文件 | 用途 | 我们的目标文件 | 状态 | 备注 |
|---|---|---|---|---|
| `apps/v4/registry/new-york-v4/blocks/bounding-box-citations-block.tsx` | 引用框 block 组合 | `packages/vue-ui/src/components/BoundingBoxCitations.vue` | adapted | #164：review panel + 页高亮；PDF viewer 由宿主 slot 提供 |
| `apps/v4/registry/new-york-v4/blocks/layout-blocks-block.tsx` | Layout blocks block 组合 | `packages/vue-ui/src/components/LayoutBlocks.vue` | adapted | #164：页叠加 + OCR block 面板 |
| `apps/v4/registry/new-york-v4/blocks/document-splits-block.tsx` | 文档拆分 block 组合 | `packages/vue-ui/src/components/DocumentSplits.vue` | adapted | #164：拆分面板 + 宿主页选择回调 |
| `apps/v4/registry/new-york-v4/blocks/e-signature.tsx` | 电子签名流程、手写/打印签名预览 | `packages/vue-ui/src/components/SignaturePad.vue`、`packages/vue-ui/src/components/ESignatureFields.vue` | adapted | #164：手写/输入签名、清除/确认/预览；字段面板独立为 `ESignatureFields` |
| `apps/v4/registry/new-york-v4/blocks/file-system-block.tsx` | 文件系统 block 组合 | `packages/vue-ui/src/components/FileSystem.vue` | reference-only | #164 后置，与 FileSystem 同一结论 |
| `apps/v4/registry/new-york-v4/blocks/docx-editor.tsx` | DOCX editor block 组合 | `packages/vue-docx/src/components/DocxEditor.vue` | reference-only | #159 |
| `apps/v4/registry/new-york-v4/blocks/docx-layout-blocks.tsx` | DOCX + layout blocks 组合 | `packages/vue-ui/src/components/LayoutBlocks.vue` | adapted | #164：`#viewer` slot 可接入 DOCX surface；block 面板共用 |
| `apps/v4/registry/new-york-v4/blocks/excel-editor.tsx` | Excel editor block 组合 | `packages/vue-xlsx/src/components/XlsxViewer.vue` | reference-only | #160 |
| `apps/v4/registry/new-york-v4/blocks/pdf-dropzone.tsx` | PDF 拖拽上传 block | `packages/vue-ui/src/components/FileUpload.vue` | reference-only | #163 |
| `apps/v4/registry/new-york-v4/blocks/pdf-editor-block.tsx` | PDF editor block 组合 | `packages/vue-pdf/src/components/PdfViewer.vue` | reference-only | #162 |

### 目录级约定与参考快照（#155、#157、#158）

| Extend 文件 | 用途 | 我们的目标文件 | 状态 | 备注 |
|---|---|---|---|---|
| `apps/v4/registry/new-york-v4/ui/*.tsx` | 所有 primitive 共用的 token 样式规则 | `packages/vue-ui/src/styles/ui.css` | adapted | Tailwind utility 改写为 `office-*` class，只读取 `--office-*` token |
| `apps/v4/components/ui/*.tsx` | primitive 目录约定 | `packages/vue-ui/src/components/ui/` | adapted | 该目录下任何 `.vue` / `.ts` / `.css` 都源自对应 Extend primitive，必须带来源头 |
| `apps/v4/components/ui/*.tsx` | 共用 Office chrome 目录约定 | `packages/vue-ui/src/components/chrome/` | reference-only | #158 的 toolbar/sidebar/status/state 组件是独立实现，按 Extend PDF/DOCX chrome 尺寸与 token 对齐，不复制上游代码（因此不加来源头）；像素证据见 `document-sidebar` 与 `office-*` surface |
| `apps/v4 vendored reference snapshot (components)` | #155/#157 像素对照 harness 复制的上游组件 | `tools/extend-reference/src/vendor/` | adapted | 类名与结构对齐上游；改动清单见 `vendor/NOTICE.md`（render-prop 换成原生元素、去掉 border-beam 依赖、弹层改为可测量的静态 markup） |
| `apps/v4 vendored reference snapshot (styles)` | #155 harness 复制的上游 CSS | `tools/extend-reference/src/styles/` | adapted | `extend-globals.css` 只裁剪 `@import` 图，正文由 `extract-extend-tokens.mjs` 校验 sha256；`extend-shadcn-tailwind.css` 除来源头外逐字复制；`parity.css` 是 harness 自有代码，不在范围内 |

## 与 #155 协调：`tools/extend-reference` 路径约定

`#155` 的对照 harness 使用 vendoring 保证可重复。它以 2026-11-20 的工作目录为基准，校验器按以下实际路径检查：

```text
tools/extend-reference/
├── README.md                        # 启动参考页与对照页
├── UPSTREAM.md                      # 仓库、固定 commit、复制日期与文件清单
├── src/
│   ├── vendor/
│   │   ├── NOTICE.md                # Extend UI MIT 全文或指向 licenses/extend-ui-MIT.txt
│   │   ├── ui/                      # button/card/select/separator/spinner, toggle/group/color-picker
│   │   └── extend/                  # pdf/docx/pptx/xlsx chrome subsets, xlsx-ribbon
│   ├── styles/
│   │   ├── extend-globals.css       # vendored apps/v4/app/globals.css
│   │   ├── extend-shadcn-tailwind.css
│   │   └── parity.css              # harness 自有，不在 copied 范围
│   └── lib/cn.ts                    # adapted from apps/v4/lib/utils.ts
└── src/surfaces/                    # harness 自有组件，无来源头
```

规则：

1. `src/vendor/` 与 `src/styles/extend-*.css` 下的每个 `.ts`、`.tsx`、`.css` 文件必须在前 15 行内包含来源头，且必须出现许可证标识（`(MIT)` 或 `MIT License`）；仅有 commit 和 copyright 不够。
2. `src/vendor/NOTICE.md`（或 `tools/extend-reference/UPSTREAM.md`、`LICENSE-extend-ui.txt`）必须明确说明 MIT、包含 `CrowdView Inc, dba Extend`，并附上 MIT 全文或指向 `docs/upstream/licenses/extend-ui-MIT.txt`；否则 `verify-attribution.mjs` 直接失败。
3. `src/lib/cn.ts` 由 `apps/v4/lib/utils.ts` 改写而来，必须带 `adapted` 来源头（当前只有一段说明注释，缺 commit、copyright 和许可证标识，需要补齐）。
4. 只复制渲染所需的最小文件集；不要复制上游的 `package.json`、锁文件或字体二进制。字体通过依赖引入（OFL 见 Geist 许可证副本）。
5. 如果改为从仓库外 checkout 读取上游源码而不 vendoring，则 `src/vendor/` 不创建，本行保持 `copied` 但校验器只报告 pending；此时必须在 `UPSTREAM.md` 记录 checkout 的 commit 和获取方式。
6. `tools/extend-reference` 只用于对照验证，不进入任何公开包的构建产物。

2026-11-21 状态（#155/#157/#170）：上表路径已落地。`src/vendor/` 下 18 个 `extend/` 组件文件与 12 个 `ui/` 文件、`src/styles/extend-globals.css`、`src/styles/extend-shadcn-tailwind.css`、`src/lib/cn.ts` 均带来源头；`src/vendor/NOTICE.md`、`src/vendor/` 同级 `UPSTREAM.md` 与 `README.md` 记录改动清单。上游 CSS 正文的 sha256 由 `docs/design/extend-token-sources.json` 固定并在 `scripts/ci/extract-extend-tokens.mjs --extend-dir <checkout>` 时与上游 checkout 交叉校验。

## 维护流程

新增或修改 upstream 相关文件时：

1. 在 [`attribution.json`](attribution.json) 的 `index` 增加/更新一行（`upstreamFile`、`target`、`status`、`issues`）。
2. 在本文件对应表格增加/更新同一行，`Extend 文件`、`我们的目标文件`、`状态` 必须与 JSON 完全一致。
3. `adapted` / `copied` 文件写入来源头；`copied` 可用 `Copied from` 开头。
4. 运行 `node scripts/ci/verify-attribution.mjs`；本地无网络时用 `ATTRIBUTION_SKIP_URLS=1`，但 CI 必须保持网络校验。
5. 若改变固定 commit：更新 `extend.commit`、`header.required`、所有已有文件头、本文件基准表和 `THIRD_PARTY_NOTICES.md`，并在提交说明写明迁移原因。

CI 注册：

- `node scripts/ci/run-suite.mjs docs` → `documentation-contract` + `open-source-attribution`
- `node scripts/ci/run-acceptance-shard.mjs docs` → 同上（`documentation-contract` + `open-source-attribution`）

## 已验证范围

- 全部 31 个 `apps/v4/components/ui` 文件、31 个 `apps/v4/components/extend` 文件、10 个 `apps/v4/registry/new-york-v4/blocks` 文件，以及 `app/`、`lib/` 中与本批次相关的 8 个文件，均已在索引中逐项登记。
- 覆盖 #155、#156、#157、#158、#159、#160、#161、#162、#163、#164 的实际或计划使用。
- 当前没有任何 `adapted` / `copied` 目标文件存在，因此没有需要添加文件头的现存文件；校验器会把这些目标报告为 pending，并在文件创建后强制检查头部。
- `#155` 的 `tools/extend-reference` 合入后，校验器会立即检查 `src/vendor/` 与 `src/styles/extend-*.css` 的来源头、`src/vendor/NOTICE.md`，并反向检查是否存在未登记的来源头文件。
- 校验器自身的正向与反向路径由 `tests/unit/attribution-contract.test.mjs` 覆盖：完整 fixture 通过，缺头、未登记头、许可证缺失、索引不同步均失败。
