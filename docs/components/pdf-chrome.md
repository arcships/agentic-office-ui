# PDF viewer / surface chrome（#162）

> 状态：已实现（issue #162）。`PdfViewer` 的 toolbar、缩略图/大纲侧栏、空/加载/错误/加密
> 状态、文本选择浮层、链接 hover card、批注对话框和状态栏改用 `@arcships/vue-ui` 的
> 共享 chrome；PDFium 会话、翻页/缩放/旋转、搜索、选择、文本层、打印、连续/单页/双页布局
> 和公共 API 保持不变。
>
> 视觉基准：Extend UI 固定 commit
> `70fe70c6f1e1eee53954c6923f3d9f583557a57e` 的 PDF toolbar 与 document sidebar。
> 所有视觉值来自 `--office-*` token，没有 Tailwind/preflight/reset。

## 1. 组件映射

| PDF 区域 | 共享组件 |
|---|---|
| 工具栏骨架 | `OfficeToolbar` + `OfficeToolbarGroup` + `OfficeToolbarSeparator` |
| 侧栏按钮、翻页、缩放、旋转、搜索、打印、下载 | `OfficeIconButton`（28px `icon-sm`） |
| 页码输入 | `OfficeInput`（native input，`data-testid="pdf-page-input"`，Enter/change 提交） |
| 页码总数 | 保留 `pdf-page-count` 文本节点，`of N` 来自 `common.page.of` |
| 页码/缩放/布局原生控件 | 原生 `<select>` + `.office-ui` 与 `--office-*` 度量 |
| 搜索框 | 原生 `input` + `.office-ui office-input`（保留 form submit/Enter 合同） |
| 缩略图/大纲侧栏 | `DocumentViewerSidebar`（160px，inline/overlay auto，`Escape` 关闭关闭） |
| 侧栏骨架 | `DocumentViewerSidebarSkeleton` |
| 空 / 加载 / 错误 / 加密 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState`（加密用 `lock` 图标 + `pdf.viewer.passwordRequired`） |
| 可见状态栏 | `OfficeStatusBar`（`show-status-bar`，默认关闭） |
| 文本选择浮层 | `OfficePopover` + `OfficeButton`（浏览器；零尺寸 fixed 锚点） |
| 链接 hover card | `OfficePopover`（`link` 批注；显示清洗后的 URL + 打开动作） |
| 批注详情 | `OfficeDialog`（`show-annotation-dialog`，默认关闭） |
| 表单字段提示 | `OfficeTooltip`（`PdfFieldTooltip`，浏览器 only） |

`PdfViewer` 保留 `.pdf-toolbar` 类作为打印/触屏的公共钩子；工具栏根本身是
`OfficeToolbar`（`role="toolbar"`、48px 行高、`--office-*` token）。根元素同时带
`data-office-theme="light|dark"`，确保共享组件和弹层解析到同一主题。

## 2. 工具栏合同

| 项目 | 值 | 来源 |
|---|---|---|
| 桌面高度 | 48px（≥1100px 容器；含 1px 底边） | Extend `min-h-12` |
| 桌面字号/行高 | 16px / 24px | Extend toolbar 根节点 |
| 内边距/间距 | 12px 水平、8px 垂直、8px gap | Extend `px-3 py-2 gap-2` |
| 图标按钮 | 28px（`icon-sm`）；coarse pointer 命中区 ≥44px | Extend `size-7` |
| 原生 select/搜索框 | 28px（`--office-control-height-xs`） | 与 toolbar 行高一致 |
| 页码 | 保留 `pdf-first-page` / `pdf-previous-page` / `pdf-page-input` / `pdf-page-count` / `pdf-next-page` / `pdf-last-page` | 现有测试与黑盒 |
| 缩放 | `pdf-zoom-out` / `pdf-zoom-select`（0.5–2 的倍率值）/ `pdf-zoom-in` | 保留既有 value 合同 |
| 布局 | `pdf-page-layout` 原生 select（continuous/single/spread） | 保留 `update:pageLayout` |
| 搜索 | `form role="search"` + `pdf-search-input` / `pdf-search-submit` / `pdf-search-prev` / `pdf-search-next` / `pdf-search-result` | 保留 Enter/submit 与 live region |
| 其他 | `pdf-sidebar-toggle`、`pdf-rotate`、`pdf-print`、`pdf-download` | 打印为 #162 新增可选按钮 |

### 响应式策略

- **≥1100px 容器**：显示文件名、第一页/最后一页，全部命令一行。
- **≤1100px**：隐藏文件名和第一页/最后一页（PageUp/PageDown、Home/End、页码输入仍可导航），
  压缩搜索框和缩放 select，桌面/平板命令保持单行。
- **≤520px**：搜索独占一行，工具栏按 `flex-wrap` 换行；`pdf-page-layout` 保留但压缩。
- **≤390px**：进一步压缩页码总数和布局 select；无横向溢出。
- 侧栏：容器 ≥768px 为 inline；更窄为 overlay，关闭时页面内容不平移。

## 3. 状态与交互

- `pdf-status` 视觉隐藏的 polite/atomic live region 保持不变，`data-state` 仍为
  `idle | loading | error | ready`。
- 空/加载/错误/加密状态改用 `Office*State`；错误节点保留
  `data-testid="load-error"`（查看器）和 `pdf-surface-load-error`（Surface）以及
  `data-error-code`。加密状态额外带 `data-encrypted="true"` 和 `lock` 图标。
- `show-status-bar` 在查看器/Surface 中渲染 `OfficeStatusBar`，包含页码、文件名和
  加载/错误状态；默认关闭，不改变既有高度。
- `PdfSurface` 仍是最小嵌入 Surface：不新增内部 toolbar（宿主拥有搜索/缩放/翻页控件，
  这是公开合同）；它的状态、侧栏和可选 `OfficeStatusBar` 与查看器使用同一套组件。
- 文本选择浮层只在浏览器挂载（Node/host renderer 无 DOM 时不挂载），复制走
  `navigator.clipboard.writeText`；`Escape`/关闭 popover 清空选择。

浏览器交互抽查（Chromium 149，`#/pdf-viewer` 与 `#/pdf-surface`）：拖选后浮层出现，
点击复制写入剪贴板并显示“已复制”；上传 `pdf-acroform-annotations.pdf` 后悬停内部
链接显示“第 2 页”卡片，点击 highlight 批注打开 `OfficeDialog` 显示
“Highlighted contract clause”；以上均无 console/page error。
- 链接 hover card 复用 Surface 已有的 `linkActivate` 事件和 URL policy；
  `show-annotation-dialog` 打开 `OfficeDialog`，批注/表单字段的既有
  `annotationActivate` 事件仍先发出。
- 表单控件在浏览器中由 `OfficeTooltip` 包裹字段名；`as-child` 触发，不增加影响
  绝对定位的盒子。

## 4. 保留的行为

连续/单页/双页布局、对页渲染、缩略图跟随可见页、outline 键盘树、页标签、搜索命中导航、
旋转、缩放锚点、拖选/双击选词/复制、文本层、打印全页快照、拖放导入、URL policy、
诊断事件和所有公开 props/events/expose 方法均未改变。新增 props 均为可选：

| 组件 | 新增可选 prop | 默认 | 说明 |
|---|---|---|---|
| `PdfViewer` | `showPrint` | `true` | 工具栏打印按钮 |
| `PdfViewer` | `showStatusBar` | `false` | 可见 `OfficeStatusBar` |
| `PdfSurface` | `showStatusBar` | `false` | 可见 `OfficeStatusBar` |
| `PdfSurface` | `showAnnotationDialog` | `false` | 点击批注时打开 `OfficeDialog` |

新增消息键：`pdf.viewer.print`、`pdf.viewer.passwordRequired`、
`pdf.viewer.selectionActions`、`pdf.viewer.copySelection`、
`pdf.viewer.copiedSelection`、`pdf.viewer.linkActions`、`pdf.viewer.openLink`、
`pdf.viewer.annotationDetails`（en / zh-CN）。

## 5. 截图 story

`scripts/ci/capture-pdf-chrome.py` 在进程内启动 `apps/demo/dist`，覆盖
390x844 / 768x1024 / 1440x900 × light/dark，驱动真实 `#/pdf-viewer` 与
`#/pdf-surface`：

- 断言 toolbar ≥48px、无横向溢出、不覆盖页面；
- 断言侧栏 160px、inline/overlay 模式、关闭后内容不平移、缩略图仍在；
- 断言页码/缩放/旋转/搜索/打印/下载/侧栏按钮存在且有可访问名称；
- 连续/单页/双页各截一张 viewer 和 toolbar；
- 断言可见状态栏位于页面下方且不覆盖页面。

```bash
pnpm --filter demo build
pnpm test:pdf:chrome
# 等价于：
# PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright \
#   node scripts/ci/run-python.mjs scripts/ci/capture-pdf-chrome.py --serve apps/demo/dist
```

最近一次结果：6/6 通过，0 console error；产物在 `output/pdf-chrome/`。
在 touch/coarse 指针 profile 下（390、768 两个 viewport）：390x844 toolbar 211px（4 行，
44px 命中区，15 个可见控件）、768x1024 115px（2 行）、1440x900 48px（1 行）；侧栏在
390/768 为 overlay、1440 为 inline；状态栏 32px。非触摸的 parity/桌面测量在 390 为
125px、768 为 81px。

## 6. Extend 像素对照（`pnpm ui:parity -- --only pdf-toolbar`）

参考 surface 是 Extend 的静态 PDF toolbar 子集（无布局 select、无内联搜索框、无打印、
下载/操作合并为菜单）。本仓库的 toolbar 必须保留这些既有命令和测试钩子，因此
**几何在桌面宽度对齐，像素差异来自被保留的额外交互控件**，不能在不删功能的前提下
降到 0.1%。

| viewport | reference 高度 | subject 高度 | geometry | 像素差异（light / dark） | 最大通道差 |
|---|---|---|---|---|---|
| 1440x900 | 48px | 48px | PASS | 25.3049% / 32.7167% | 217 / 236 |
| 1280x720 | 48px | 48px | PASS | 25.3049% / 32.7167% | 217 / 236 |
| 768x1024 | 48px | 81px | FAIL (+33px) | 11.2476% / 22.2753% | 217 / 236 |
| 390x844 | 125px | 161px | FAIL (+36px) | 11.7837% / 24.4957% | 218 / 236 |

桌面几何（48px、16/24 子、12/8 内边距、8px gap）已与参考完全一致；768/390 的差异是
保留的布局 select、内联搜索和打印按钮导致的额外换行。阈值仍为
`geometryTolerancePx: 0`、`maxDiffPixelRatio: 0.001`，未修改。剩余差异需要产品决定是否
把布局/搜索/打印移入溢出菜单后才能消除。

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
