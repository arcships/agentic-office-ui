# PDF viewer / surface chrome（#162）

> 状态：已实现（issue #162）。`PdfViewer` 的默认 toolbar 行是 Extend
> `PDFViewerToolbar` 的逐像素复刻；PDF 额外的 layout、首/末页、上/下一页、文件名、
> 下载和打印命令收进右侧 `⋯` actions 菜单。缩略图/大纲侧栏、空/加载/错误/加密状态、
> 文本选择浮层、链接 hover card、批注对话框和状态栏继续使用共享 chrome。
>
> 视觉基准：`tools/extend-reference/src/vendor/extend/pdf-toolbar.tsx`（Extend UI
> `apps/v4/components/extend/pdf-viewer.tsx`，固定 commit
> `70fe70c6f1e1eee53954c6923f3d9f583557a57e`）。默认行没有 Tailwind：
> `PdfViewerToolbar.vue` 用专用 CSS 回到上游 computed 尺寸。

## 1. 默认 toolbar 行（与上游一致）

左侧：

- sidebar toggle（`SidebarLeftIcon`，28px 按钮 / 16px 图标）；
- `Page` 文本、页码按钮（原生 input 透明覆盖，可输入/↑/↓/Home/End）、`of N`。

右侧：

- rotate counterclockwise + rotate clockwise（`showRotateControls` 时，
  `RotateClockwiseIcon`，第二个水平镜像）；
- 1×16 separator；
- zoom out、104px zoom select、zoom in（`CircleMinusIcon` / `CirclePlusIcon`）；
- 1×16 separator；
- search 按钮（`Search01Icon`，点击打开搜索浮层）；
- 1×16 separator；
- `⋯` actions 按钮（`MoreHorizontalIcon`）。

图标全部来自 `@hugeicons/core-free-icons`，由 `@arcships/vue-ui` 的 `HugeIcon`
渲染；尺寸 16px，chevron 在 <640px 为 18px，和上游 `sm:` 变体一致。

### 尺寸合同

| 项目 | 值 | 来源 |
|---|---|---|
| toolbar | `min-height: 48px`、`padding: 8px 12px`、`gap: 8px`、1px 底边 | Extend `min-h-12 px-3 py-2 gap-2 border-b` |
| 根字号/行高 | 16px / 24px | Extend toolbar 根节点 |
| 左侧组 | `gap: 8px` | `gap-2` |
| 右侧组 | `gap: 4px`、`justify-end` | `gap-1` |
| icon button | 28×28（<640px 为 32×32），radius 10px，ghost hover 4% accent | `size-icon-sm` / `sm:size-7` |
| 页码按钮 | 文本宽 + 18px padding + 2px border；<640px 字号 16/24、按钮 32px | `size-sm` / `sm:h-7` |
| zoom select | 104×28（<640px 高 32px），padding 0 9px，chevron 16px（<640px 18px） | `w-[104px] min-w-[104px] size-sm` |
| separator | 1×16，左右 margin 4px | `mx-1 h-4` |
| disabled | 每个控件 opacity 0.64，shadow/edge 关闭 | `disabled:opacity-64` |

页码输入使用透明原生 input 覆盖在可见文本上：默认状态与上游按钮逐像素一致，
同时保留 `pdf-page-input` 的 fill/Enter 和 `value` 合同。zoom 使用可见 trigger +
透明原生 `<select>` 覆盖，保留 `select_option("2")` 的分数值合同。

## 2. `⋯` actions 菜单

默认不渲染（因此不影响像素对照），点击 `pdf-actions` 后显示：

| 菜单项 | test id | 说明 |
|---|---|---|
| 文件名 | — | 只读 label |
| layout: continuous / single / spread | `pdf-page-layout`（group）、`pdf-page-layout-continuous` 等 | 保留 `update:pageLayout` |
| previous / next page | `pdf-previous-page`、`pdf-next-page` | 保留翻页命令 |
| first / last page | `pdf-first-page`、`pdf-last-page` | 保留首/末页命令 |
| print | `pdf-print` | 打印全部页面快照 |
| download | `pdf-download` | 使用已验证字节 |

菜单关闭时不渲染，`pdf-download`/`pdf-print` 等仍保留原 test id；测试需要先点击
`pdf-actions` 再操作。文档加载错误或控件禁用时 actions 按钮 disabled。

## 3. 侧栏、状态和浮层

| PDF 区域 | 共享组件 |
|---|---|
| 缩略图/大纲侧栏 | `DocumentViewerSidebar`（160px，inline/overlay auto，PDF 使用 `flush` 让缩略图列表拥有整条 rail） |
| 侧栏骨架 | `DocumentViewerSidebarSkeleton`（`variant="thumbnail"`：上游 80×112 页面占位 + 40×12 标签） |
| 空 / 加载 / 错误 / 加密 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` |
| 可见状态栏 | `OfficeStatusBar`（`show-status-bar`，默认关闭） |
| 文本选择浮层 | `OfficePopover` + `OfficeButton`（浏览器，fixed 锚点） |
| 链接 hover card | `OfficePopover`（link 批注，显示 URL 或目标页） |
| 批注详情 | `OfficeDialog`（`show-annotation-dialog`，默认关闭） |
| 表单字段提示 | `OfficeTooltip`（`PdfFieldTooltip`，浏览器 only） |

`PdfSurface` 仍是最小嵌入 Surface：不新增内部 toolbar（宿主拥有搜索/缩放/翻页控件，
这是公开合同）；状态、侧栏和可选 `OfficeStatusBar` 与查看器使用同一套组件。

## 4. 保留的公共合同

连续/单页/双页、对页渲染、缩略图跟随可见页、outline 键盘树、页标签、搜索命中导航、
旋转、缩放锚点、拖选/双击选词/复制、文本层、打印全页快照、拖放导入、URL policy、
诊断事件和 `expose` 方法均未改变。新增 props 均为可选：

| 组件 | 新增可选 prop | 默认 | 说明 |
|---|---|---|---|
| `PdfViewer` | `showPrint` | `true` | actions 菜单打印项 |
| `PdfViewer` | `showStatusBar` | `false` | 可见 `OfficeStatusBar` |
| `PdfSurface` | `showStatusBar` | `false` | 可见 `OfficeStatusBar` |
| `PdfSurface` | `showAnnotationDialog` | `false` | 点击批注打开 `OfficeDialog` |

新增消息键（en / zh-CN）：`pdf.viewer.print`、`pdf.viewer.moreActions`、
`pdf.viewer.rotateCounterclockwise`、`pdf.viewer.passwordRequired`、
`pdf.viewer.selectionActions`、`pdf.viewer.copySelection`、
`pdf.viewer.copiedSelection`、`pdf.viewer.linkActions`、`pdf.viewer.openLink`、
`pdf.viewer.annotationDetails`。

## 5. Extend 像素对照

```bash
pnpm ui:parity -- --only pdf-toolbar
pnpm ui:parity -- --only pdf-sidebar
```

PDF toolbar：**PASS — 40/40**（1440×900、1280×720、768×1024、390×844 × light/dark ×
default/hover/focus/active/disabled），`diffRatio 0.0`、`maxChannelDelta 0`。

PDF sidebar：**PASS — 24/24**（1440/1280/768/390 × light/dark × open/closed/skeleton），
`diffRatio 0.0`、`maxChannelDelta 0`；几何全部通过。两侧都是 160px rail、四页
600×800 PDF、同一缩略图 PNG、数字页标签、active page 1；1440/1280/768 走 inline，
390 走 overlay；closed 状态比较隐藏后的空 rail，skeleton 比较上游骨架。

为达到该结果，参考与 subject 使用同一组 fixture 值；subject 直接渲染生产组件
`PdfViewerToolbar` 和 PdfViewer 使用的 `DocumentViewerSidebar` + `PdfThumbnailStrip`，
而不是加载真实 PDF。

未修改阈值。

## 6. 截图 story

`scripts/ci/capture-pdf-chrome.py`（`pnpm test:pdf:chrome`）在进程内启动
`apps/demo/dist`，覆盖 390x844 / 768x1024 / 1440x900 × light/dark，驱动真实
`#/pdf-viewer` 与 `#/pdf-surface`：

- 断言 toolbar ≥48px、无横向溢出、不覆盖页面；
- 断言侧栏 160px、inline/overlay、关闭后内容不平移、缩略图仍在；
- 断言 toolbar 可见控件（页码/缩放/旋转/search/actions）有可访问名称；
- 连续/单页/双页各截一张 viewer 和 toolbar；
- 断言可见状态栏位于页面下方且不覆盖页面；
- coarse pointer 下 10 个可见控件 ≥44px。

最近一次结果：6/6 通过，0 console error；产物在 `output/pdf-chrome/`。
触摸 profile：1440 48px、768 61px（44px 命中区）、390 161px（4 行）；非触摸
parity/桌面测量为 1440/768 48px、390 125px。侧栏在窄容器为 overlay，1440 为
inline；状态栏 32px；coarse pointer 下 10 个可见控件均 ≥44px。

## 7. 验证

```bash
pnpm --filter @arcships/vue-pdf build
node --test --test-concurrency=1 tests/component/pdf-chrome.test.mjs
node --test --test-concurrency=1 tests/component/pdf-navigation-aids.test.mjs
node --test --test-concurrency=1 tests/component/pdf-view-modes.test.mjs
node --test --test-concurrency=1 tests/component/pdf-page-lifecycle.test.mjs
node --test --test-concurrency=1 tests/component/pdf-surface-lazy-render.test.mjs
node --test --test-concurrency=1 tests/component/pdf-object-selection.test.mjs
node --test --test-concurrency=1 tests/component/viewer-keyboard.test.mjs
node --test --test-concurrency=1 tests/component/viewer-print-theme.test.mjs
node scripts/ci/verify-test-registration.mjs
node scripts/ci/verify-public-api.mjs
node scripts/ci/verify-docs.mjs
node scripts/ci/verify-attribution.mjs
```

已知既有失败（不在本 issue 范围内，base 分支同样失败）：
`tests/component/vue-components.test.mjs` 的
`PdfSurface exposes the unified controlled object/region reference contract`、
`tests/component/office-a11y.test.mjs` 的 XLSX context-menu `role="menu"` 断言、
`tests/component/i18n-source.test.mjs` 的 `labels.ts` / DOCX toolbar 两个基线、
`file-drop-import` 的 PPTX `Element is not defined`、`touch-interactions` 的 XLSX
`@media (pointer: coarse)` 断言。
