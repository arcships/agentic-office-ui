# PPTX viewer / presenter / sorter chrome（#161）

> 状态：已实现（issue #161）。`PptxViewer` 的浏览工具栏、缩略图网格、演讲者面板、
> 链接右键菜单和空/加载/错误/受保护状态全部改用 `@arcships/vue-ui` 的共享 chrome
> 与 primitive；播放状态机、动画、转场、媒体、Morph、书签、presenter/sorter、
> 键盘/Escape/全屏和公共 API 保持不变。
>
> 视觉基准：Extend UI 固定 commit
> `70fe70c6f1e1eee53954c6923f3d9f583557a57e` 的 PPTX toolbar 与 sorter；
> 所有视觉值来自 `--office-*` token，没有 Tailwind/preflight/reset。

## 1. 组件映射

| PPTX 区域 | 共享组件 |
|---|---|
| 工具栏骨架 | `OfficeToolbar` + `OfficeToolbarGroup` + `OfficeToolbarSeparator` |
| 上/下页、全屏、备注、视图模式 | `OfficeIconButton`（28px `icon-sm`） |
| 页码 | `OfficePageControl`（输入框 + `of N`，`update:modelValue` → `goTo()`） |
| 缩放 | `OfficeZoomControl`（浏览器；50–200%，25% 预设） |
| 隐藏页标记 | `OfficeBadge` |
| 缩略图 loading / error / hidden | `OfficeSkeleton` / `OfficeBadge` + `OfficeIcon` |
| sorter 格子 | `OfficeCard`（每页一张卡片，`data-slide-index` / `data-active`） |
| presenter 面板 | `OfficeCard` + `OfficeSeparator` + `OfficeScrollArea` |
| 空 / waiting / loading / 取消 / 错误 / 受保护 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` |
| 单页渲染失败条重试/跳过 | `OfficeButton` |
| 链接右键菜单 | `OfficeDropdownMenu`（`PptxLinkMenu`，指针锚点） |

`PptxLinkMenu` 作为低层组件从根入口导出，供宿主在自建舞台时复用同一菜单；
完整查看器仍按 `contextMenu` / `action` 事件驱动它。

## 2. 工具栏合同

| 项目 | 值 | 来源 |
|---|---|---|
| 桌面高度 | 49px（`min-height: 48px` + 1px 底边） | Extend `min-h-12` |
| 图标按钮 | 28px（`icon-sm`）；coarse pointer 命中区 ≥44px | Extend `size-7`/`size-8` |
| 页码输入 | `inputmode="numeric"`、`pattern="[0-9]*"`，Enter/Esc/↑/↓/Home/End | `OfficePageControl` |
| 缩放 | `OfficeZoomControl`，`min=50`、`max=200`、`options=[50…200]`、`showFit=false` | Extend zoom cluster |
| 无重叠 | 工具栏在正常流中，`toolbar.bottom <= stageWrap.top` | 截图脚本断言 |
| 390px | 分组换行，不横向溢出；搜索独占一行；trailing 独占一行 | `chrome.css` + viewer CSS |

实测（`output/pptx-chrome/summary.json`，DPR 1，coarse pointer）：

| viewport | toolbar 外高 | 是否换行 | sorter 列 | sorter 滚动 |
|---|---|---|---|---|
| 1440×900 | 49px | 否 | 5 | 无（12 页全可见） |
| 768×1024 | 165px | 是 | 4 | 无 |
| 390×844 | 213px | 是（4 行：页码/视图模式/搜索/缩放） | 2 | 413px（可滚动） |

390px 下工具栏高是因为 coarse pointer 把所有控件提升到 44px 命中区、搜索独占一行；
它仍在正常流中，不覆盖舞台（舞台顶部 y=546）。如需更紧凑，宿主可关闭
`show-search` 或 `show-toolbar`。

### 2.1 视图模式与备注

- 工具栏左侧固定提供 `single` / `sorter` / `present` 三个 `OfficeIconButton`；
  `active` 写 `data-active="true"` 和 `aria-pressed="true"`。
- `mode` 仍可受控（`v-model:mode`）；省略时工具栏内部切换，并始终发
  `update:mode`。演示模式隐藏工具栏，播放控制条新增“退出演示”（
  `data-testid="pptx-exit-present"`）。
- `showNotes` 仍可受控（`v-model:show-notes`）；工具栏备注按钮切换 presenter rail，
  并始终发 `update:showNotes`。
- `showPresentToggle` / `showNotesToggle` / `showFullscreenControl` 默认 `true`，
  都可单独关闭。

### 2.2 DOM 合同

以下 test id 保留或新增，供宿主 e2e 使用：

| testid | 说明 |
|---|---|
| `pptx-viewer` | 根节点；`data-state`、`data-encrypted`、`data-office-theme` |
| `office-toolbar` | 工具栏根（`role="toolbar"`） |
| `pptx-previous-slide` / `pptx-next-slide` | 上下页 |
| `pptx-page-counter` | `OfficePageControl` 根；输入框为 `office-page-input` |
| `pptx-hidden-badge` | 隐藏页 badge |
| `pptx-view-mode-single` / `-sorter` / `-present` | 视图模式 |
| `pptx-notes-toggle` / `pptx-toolbar-fullscreen` | 备注 / 全屏 |
| `pptx-zoom-out` / `pptx-zoom-reset` / `pptx-zoom-in` | DOM-free 渲染器的缩放回退 |
| `pptx-sorter` / `pptx-sorter-grid` / `pptx-sorter-item` | sorter（格子带 `data-slide-index` / `data-active`） |
| `pptx-thumbnail-skeleton` / `pptx-thumbnail-error` / `pptx-thumbnail-hidden` | 缩略图状态 |
| `pptx-presenter` / `pptx-presenter-timer` / `pptx-presenter-notes` / `pptx-presenter-last` | presenter |
| `pptx-context-menu` | 菜单内容（`role="menu"`，`aria-label` 为 `link.menu`） |
| `pptx-empty` / `pptx-waiting` / `pptx-loading` / `pptx-cancelled` / `pptx-encrypted` / `pptx-error` | 状态 |

## 3. 缩略图与 slide sorter

- `PptxThumbnail` 增加 `data-thumbnail-state="idle|loading|ready|error"`；
  `renderThumbnail().ready` 解决前显示 `OfficeSkeleton`，拒绝时显示
  `OfficeIcon(triangle-alert)` + `pptx.thumbnail.failed`，并通过
  `aria-busy`/`role="img"` 播报。
- 隐藏页的 `Hidden` 文案改为 `OfficeBadge`（不再硬编码琥珀色）。
- sorter 的每个格子是 `OfficeCard`：`OfficeCard` 提供边框/圆角/背景/阴影；
  `data-active="true"` 用 accent ring 标示当前页。虚拟窗口、上下 spacer、
  `gridTemplateColumns`、滚动/键盘（PageUp/PageDown/Home/End 只滚网格）不变。
- 390px 下 2 列、卡片不横向溢出；12 页 deck 可滚动 413px，滚动不改变当前页。

## 4. Presenter 面板

`PptxPresenterPanel` 的结构改为：

```
aside[data-testid=pptx-presenter]
└─ OfficeCard
   ├─ meta：next slide 标签 + 下一页缩略图（或“最后一页”）+ clock 图标 + 计时
   ├─ OfficeSeparator（仅演示模式）
   └─ body[data-testid=pptx-presenter-notes]
      └─ OfficeScrollArea → 备注段落 / notes 插槽
```

计时、`role="timer"`、`aria-label`、`data-presenter-mode`、空备注文案和
`notes` 插槽都保持原行为；备注区在面板内滚动，不再让整个 aside 滚动。

## 5. 链接右键菜单

- 菜单本体是 `PptxLinkMenu` = `OfficeDropdownMenu`，内容 test id 为
  `pptx-context-menu`，`aria-label` 为 `pptx.viewer.link.menu`。
- 右键位置由 0×0 的 fixed 锚点传给 Reka；浮层仍按视口碰撞翻转/位移，替代了旧的
  手写 `left/top` 钳制。
- 键盘、typeahead、禁用项、Escape、外部点击关闭由 shared menu 提供；关闭时
  `restoreFocus` 把焦点还给 `PptxViewer` 根节点，不再落在隐藏锚点上。
- 菜单项：`open`（外部链接“打开链接”，内部跳转“跳到第 N 页”，带 `link` /
  `arrow-right` 图标）、`copy`（仅 URL 目标，`copy` 图标）。
- 无链接的对象、`showContextMenu="false"`、非 ready 状态、文档/页码/模式变化
  都会关闭菜单；`context-menu` 事件始终先抛出。
- 无 DOM 的组件测试渲染器不会挂载 Reka 浮层（`canUseDom` 守卫）；
  `tests/component/pptx-context-menu.test.mjs` 在 Happy DOM 中直接测
  `PptxLinkMenu` 的 pointer anchor、select/Escape/focus/disabled 合同。

## 6. 状态与受保护演示文稿

| 状态 | 组件 | 说明 |
|---|---|---|
| `empty` | `OfficeEmptyState`（`file-text`） | 未选择文件；`empty` 插槽替换描述 |
| `waiting` | `OfficeLoadingState` | 舞台尚未挂载 |
| `loading` | `OfficeLoadingState` | 打开文档；`loading` 插槽替换描述 |
| `cancelled` | `OfficeErrorState`（`circle-x`）+ 重新加载 | 取消不是错误，仍可重试 |
| `error` | `OfficeErrorState` + `error` 插槽 + 重试 | 解析/超时失败 |
| `encrypted` | `OfficeErrorState`（`lock`） | 见下 |

`pptx-core` 在 ZIP 解析前检查 OLE/CFB 头（`D0 CF 11 E0 A1 B1 1A E1`）。PowerPoint
的密码保护 `.pptx` 和旧版 `.ppt` 改名文件都是 CFB 容器，zip reader 只能给出
`PARSE_FAILED`；现在返回 `UNSUPPORTED_CONTENT`，`PptxViewer` 把它映射为
`pptx-encrypted`（锁图标 + `pptx.viewer.encryptedTitle/encryptedDescription`），
同时在根节点写 `data-encrypted="true"`。`data-state` 仍是 `error`，已有宿主的
状态判断不受影响。

## 7. 测试与截图证据

组件测试（逐文件、DOM-free renderer）：

```bash
node --test tests/component/pptx-preview.test.mjs
node --test tests/component/pptx-hyperlinks.test.mjs
node --test tests/component/pptx-presenter-view.test.mjs
node --test tests/component/pptx-slide-sorter.test.mjs
node --test tests/component/pptx-context-menu.test.mjs   # Happy DOM
```

截图 story（真实 demo 页面 + 12 页 fixture 上传，进程内静态服务器）：

```bash
pnpm --filter demo build
PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright \
  node scripts/ci/run-python.mjs scripts/ci/capture-pptx-chrome.py --serve apps/demo/dist
```

2026-11-20 实测：6/6 用例 PASS、0 console error（390/768/1440 × light/dark），
覆盖 browse-single / sorter / present+presenter；断言包括 toolbar ≥48px、
无横向溢出、`toolbar.bottom <= stage.top`、sorter 卡片不溢出且可滚动、
presenter 的 Card/ScrollArea/Separator/next/timer 存在。产物：
`output/pptx-chrome/<viewport>/<theme>/` 与 `summary.json`。

### 与 Extend 参考的剩余 delta

#155 的固定参考 app 目前只有 `pdf-toolbar` 与 `document-sidebar` surface，
没有可对齐的 PPTX toolbar/sorter React 参考：
`tools/parity-fixtures/manifest.json` 中没有 `pptx-*` 条目，新增 surface 需要
先把 Extend `PptxToolbar` 中本库没有的 sidebar/file 菜单也 vendored 进来，
否则 compare 只能得到结构性失败。因此本 issue **不伪造像素 pass**，改为：

- 工具栏/页码/缩放/图标尺寸对齐 `docs/components/office-chrome.md` 的共享合同
  （49px、28/32px、48px min-height）；
- 用上面的 story 截图 + 几何断言提供响应式证据；
- 等 #165 或后续 issue 增加 `pptx-toolbar` 参考 surface 后再接入像素 gate。

已知视觉 delta：

1. 搜索仍是内联 `<input type="search">`（Extend 用 popover trigger），
   因此 390px 多占一行；它不属于本 issue 的 toolbar 迁移清单，行为保持不变。
2. `OfficeZoomControl` 只在浏览器渲染；DOM-free 测试渲染器使用等价的
   icon-button + 百分比按钮回退，像素证据来自浏览器截图。
3. sorter 的卡片样式（边框/圆角/阴影）来自 `OfficeCard`，卡片内边距比旧
   `PptxThumbnail` 的 7px 少 2px；虚拟窗口高度不变。
