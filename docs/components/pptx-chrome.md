# PPTX viewer / presenter / sorter chrome（#161）

> 状态：已实现（issue #161 / #161b）。`PptxViewer` 的浏览工具栏、缩略图网格、
> 演讲者面板、链接右键菜单和空/加载/错误/受保护状态全部使用 `@arcships/vue-ui`
> 的共享 chrome 与 primitive；播放状态机、动画、转场、媒体、Morph、书签、
> presenter/sorter、键盘/Escape/全屏和公共 API 保持不变。
>
> 视觉基准：Extend UI 固定 commit
> `70fe70c6f1e1eee53954c6923f3d9f583557a57e` 的 `PptxToolbar` 与缩略图卡片。
> 所有视觉值来自 `--office-*` token，没有 Tailwind/preflight/reset。

## 1. 组件映射

| PPTX 区域 | 生产组件 | 视觉来源 |
|---|---|---|
| 工具栏 | `PptxViewerToolbar.vue`（`data-pptx-toolbar`） | Extend `PptxToolbar`、`PptxSlideNumberControl`、`PptxFileActionsMenu` trigger |
| 上/下页、缩略图开关、缩放、操作菜单 | `OfficeIconButton`（`icon-sm`）+ Hugeicons | Extend 同名按钮与 Hugeicons |
| 页码 | 工具栏内 “Slide N of M” 按钮 + 编辑输入 | Extend `Slide {N} of {M}` |
| 缩放 | 104px 触发框 + 原生 `<select>` | Extend `SelectTrigger size=sm` |
| 操作菜单 | `pptx-toolbar__menu`（`role="menu"`） | Extend DropdownMenu 几何 |
| 隐藏页标记 | 工具栏内 `pptx-hidden-badge` | PPTX-only 状态（上游无隐藏页） |
| 搜索 | 工具栏下方 find bar（`pptx-viewer__searchbar`） | PPTX-only（上游无搜索控件） |
| 缩略图 loading / error / hidden | `OfficeSkeleton` / `OfficeIcon` / `OfficeBadge` | `FileThumbnail` loading fallback |
| sorter 格子 | `PptxSlideCard.vue` + `PptxThumbnail` | Extend `PptxSidebarThumbnail` 卡片 |
| presenter 面板 | `OfficeCard` + `OfficeSeparator` + `OfficeScrollArea` | Extend primitives（无上游 presenter 布局，见 §4） |
| 空 / waiting / loading / 取消 / 错误 / 受保护 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` | 共享状态组件 |
| 链接右键菜单 | `OfficeDropdownMenu`（`PptxLinkMenu`） | Extend 菜单 primitive |

`PptxLinkMenu` 作为低层组件从根入口导出，供宿主在自建舞台时复用同一菜单；
完整查看器仍按 `contextMenu` / `action` 事件驱动它。

## 2. 工具栏合同

默认工具栏只显示上游 `PptxToolbar` 的行：

```
[缩略图开关] | [上一页] Slide [3] of 12 [下一页]        [缩小] [100% ▾] [放大] | [...]
```

| 项目 | 值 | 来源 |
|---|---|---|
| 行高 | `min-height: 48px` + 1px 底边 | Extend `min-h-12` |
| 内边距 / 间距 | `8px 12px`、`gap: 8px`；左侧组 `gap: 4px` | `px-3 py-2 gap-2` / `gap-1` |
| 图标按钮 | 32px（<640），28px（≥640）；图标 16px | `size-8 sm:size-7`、`size-4` |
| 页码按钮 | 32/28px 高，`padding: 0 9px`，`font-weight: 400`，`text-primary` | `Button size=sm font-normal` |
| 缩放触发框 | `104px × 32/28px`，9px 内边距，`size=sm` 边框/阴影 | `w-[104px] min-w-[104px]` |
| 分隔线 | 1×16px，`margin: 0 4px` | `Separator mx-1 h-4` |
| 换行 | 390px 下左右两组各占一行（矩阵 89px；coarse pointer 44px 命中区时 chrome story 为 113px）；无横向溢出 | `flex-wrap` |
| 搜索 | 不在工具栏内；操作菜单“搜索幻灯片”或 `Ctrl+F` 打开 find bar | #161b |

### 2.1 操作菜单

上游工具栏右侧的 `⋯`（`aria-label="Open PowerPoint actions"`）在 PPTX 里承载
本库有、上游 PPTX viewer 没有的命令，顺序固定：

| 菜单项 | testid | 说明 |
|---|---|---|
| 单页视图 | `pptx-view-mode-single` | `menuitemcheckbox`，`aria-checked` |
| 幻灯片浏览 | `pptx-view-mode-sorter` | `menuitemcheckbox` |
| 演示 | `pptx-view-mode-present` | `menuitem`，进入演示模式 |
| 演讲者备注 | `pptx-notes-toggle` | `menuitemcheckbox`，切换 presenter rail |
| 全屏 | `pptx-toolbar-fullscreen` | `menuitem`，进入/退出全屏 |
| 搜索幻灯片 | `pptx-search-toggle` | 打开下方 find bar（`aria-expanded`） |

菜单关闭时不渲染内容；按钮/菜单项的 testid 保持不变，宿主 e2e 需要先点
`pptx-actions`（或直接发 `Ctrl+F` 打开搜索）。`showPresentToggle` /
`showNotesToggle` / `showFullscreenControl` / `showSearch` 仍可单独关闭；
关闭全部时 `⋯` 也不渲染，工具栏只剩上游控件。

### 2.2 视图模式与备注

- `mode` 仍可受控（`v-model:mode`）；省略时菜单内部切换，并始终发
  `update:mode`。演示模式隐藏工具栏，播放控制条新增“退出演示”
  （`data-testid="pptx-exit-present"`）。
- `showNotes` 仍可受控（`v-model:show-notes`）；菜单备注项切换 presenter rail，
  并始终发 `update:showNotes`。
- 单页/浏览布局由 `viewMode` prop 或 `update:viewMode` 驱动；菜单选中态用
  `aria-checked`，不再用 `aria-pressed`。
- 缩略图开关是上游工具栏的第一项；`showSidebar` 是初始/受控值，点击后 viewer
  内部切换（不改变公共 API）。

### 2.3 DOM 合同

以下 test id 保留或新增，供宿主 e2e 使用：

| testid | 说明 |
|---|---|
| `pptx-viewer` | 根节点；`data-state`、`data-encrypted`、`data-office-theme` |
| `pptx-viewer-toolbar` | 工具栏根（`role="toolbar"`，同时带 `data-pptx-toolbar`） |
| `pptx-sidebar-toggle` | 缩略图侧栏开关 |
| `pptx-previous-slide` / `pptx-next-slide` | 上/下页 |
| `pptx-page-counter` | “Slide N of M”；数字为 `office-page-input`（显示按钮，点击后变输入），总数为 `office-page-total` |
| `pptx-hidden-badge` | 隐藏页 badge |
| `pptx-actions` / `pptx-actions-menu` | 操作菜单触发器 / 菜单内容 |
| `pptx-view-mode-single` / `-sorter` / `-present` | 视图模式（菜单项） |
| `pptx-notes-toggle` / `pptx-toolbar-fullscreen` | 备注 / 全屏（菜单项） |
| `pptx-zoom-out` / `pptx-zoom-reset` / `pptx-zoom-in` / `pptx-zoom-select` | 缩放按钮、可见数值、原生选择框 |
| `pptx-search-toggle` / `pptx-search-input` / `pptx-search-close` | find bar |
| `pptx-sorter` / `pptx-sorter-grid` / `pptx-sorter-item` | sorter（格子带 `data-slide-index` / `data-active`） |
| `pptx-thumbnail-skeleton` / `pptx-thumbnail-error` / `pptx-thumbnail-hidden` | 缩略图状态 |
| `pptx-presenter` / `pptx-presenter-timer` / `pptx-presenter-notes` / `pptx-presenter-last` | presenter |
| `pptx-context-menu` | 菜单内容（`role="menu"`，`aria-label` 为 `link.menu`） |
| `pptx-empty` / `pptx-waiting` / `pptx-loading` / `pptx-cancelled` / `pptx-encrypted` / `pptx-error` | 状态 |

## 3. 缩略图与 slide sorter

- `PptxThumbnail` 增加 `data-thumbnail-state="idle|loading|ready|error"` 与
  `show-label`（sorter 卡片自己渲染页码）；`renderThumbnail().ready` 解决前显示
  `OfficeSkeleton`，拒绝时显示 `OfficeIcon(triangle-alert)` +
  `pptx.thumbnail.failed`，并通过 `aria-busy`/`role="img"` 播报。
- 隐藏页的 `Hidden` 文案使用 `OfficeBadge`（不再硬编码琥珀色）。
- sorter 每格是 `PptxSlideCard`，结构和取值来自上游缩略图卡片：
  `p-2`、`gap-2`、`rounded-md`（8px）、`text-xs`（12px/16px）、
  hover/active `bg-sidebar-accent`；预览框 `rounded-sm`（6px）、背景
  `bg-background`、`shadow-xs`（active `shadow-sm`）、白色 preview 层。
  `data-active="true"` 的选中态用同样的 accent 背景（不再用 ring）。
- 虚拟窗口、上下 spacer、`gridTemplateColumns`、滚动/键盘
  （PageUp/PageDown/Home/End 只滚网格）不变。卡片高度按
  `8 + 缩略图高度 + 8 + 16 + 8` 计算，与 `--pptx-sorter-item-height` 对齐。
- 390px 下 2 列，卡片不横向溢出；12 页 deck 可滚动，滚动不改变当前页。

## 4. Presenter 面板

`PptxPresenterPanel` 的结构：

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

**上游参照不可达**：固定 commit 的 `apps/v4/components/extend/pptx-viewer.tsx`
只包含工具栏、缩略图 rail 和 loading/empty/error；presenter 视图与演讲者备注
在私有 `@extend-ai/react-pptx` 引擎内部，本仓库不 vendor 该引擎。用 Extend
primitive 拼一个“参考”只会重复我们自己的组合，因此
`tools/parity-fixtures/manifest.json` 把 `pptx-presenter` 登记为
`referenceStatus: "unreachable"`，报告会列出原因而不是伪造像素 pass；该面板由
共享 primitive、组件测试和 §7 的 chrome story 验证。

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

## 7. 测试与像素证据

组件测试（逐文件、DOM-free renderer）：

```bash
node --test tests/component/pptx-preview.test.mjs
node --test tests/component/pptx-hyperlinks.test.mjs
node --test tests/component/pptx-presenter-view.test.mjs
node --test tests/component/pptx-slide-sorter.test.mjs
node --test tests/component/pptx-context-menu.test.mjs   # Happy DOM
```

### 7.1 Extend 像素矩阵

`pptx-toolbar` 与 `pptx-sorter` 已进入 `pnpm ui:parity` 矩阵（DPR 2，
1440×900 / 1280×720 / 768×1024 / 390×844 × light/dark）。参考端是 vendored
上游标记（`tools/extend-reference/src/vendor/extend/pptx-toolbar.tsx`、
`pptx-slide-card.tsx`），被测端是生产组件：

```bash
pnpm ui:parity -- --only pptx-toolbar,pptx-sorter,pptx-presenter
```

2026-11-20 实测（`output/ui-parity/report.json`）：

| surface | 状态 | 视口×主题×状态 | 结果 | geometry | 像素 |
|---|---|---|---|---|---|
| `pptx-toolbar` | default/hover/focus/active/disabled | 4×2×5 = 40 | 40 pass | 0px | 0% |
| `pptx-sorter` | default/hover/selected | 4×2×3 = 24 | 24 pass | 0px | 0% |
| `pptx-presenter` | — | — | unreachable（§4） | — | — |

矩阵阈值未修改（`maxChannelDelta: 2`、`maxDiffPixelRatio: 0.001`、0px 几何）。两个 surface 目前是 `gateStatus: "pending"`（`gatePendingIssue: "#161"`）：像素已经为 0，但 #165 的 committed baseline/approval 仍待补；`pptx-presenter` 以 `referenceStatus: "unreachable"` 出现在 report 的 unreachable 表，不参与 pass/fail 计数。

### 7.2 响应式 story

真实 demo 页面 + 12 页 fixture 上传，进程内静态服务器：

```bash
pnpm --filter demo build
PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright \
  node scripts/ci/run-python.mjs scripts/ci/capture-pptx-chrome.py --serve apps/demo/dist
```

覆盖 390/768/1440 × light/dark 的 browse-single / sorter / present+presenter；
断言包括 toolbar ≥48px、无横向溢出、`toolbar.bottom <= stage.top`、
sorter 卡片不溢出且可滚动、presenter 的 Card/ScrollArea/Separator/next/timer
存在。产物：`output/pptx-chrome/<viewport>/<theme>/` 与 `summary.json`。
