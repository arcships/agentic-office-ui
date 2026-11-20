# 共享 Office chrome（toolbar / sidebar / status / states）

> 状态：已实现（issue #158）。所有 viewer/editor 共用的 chrome 层位于
> `@arcships/vue-ui`；DOCX 已在本批 #159 完成迁移，XLSX/PPTX/PDF 的接入见
> #160–#162。
>
> 视觉基准：Extend UI 固定 commit
> `70fe70c6f1e1eee53954c6923f3d9f583557a57e` 的 PDF toolbar 与
> `DocumentViewerThumbnailSidebar`。所有视觉值来自 `--office-*` token；不引入
> Tailwind、preflight 或全局 reset。

## 1. 安装与主题

```ts
// main.ts
import "@arcships/vue-ui/style.css"
import "@arcships/vue-ui/theme.css"
```

组件必须位于 `[data-office-theme]` 子树内；推荐用 `OfficeThemeProvider`：

```vue
<script setup lang="ts">
import { ref } from "vue"
import { OfficeThemeProvider, OfficeToolbar } from "@arcships/vue-ui"
const theme = ref<"light" | "dark" | "system">("system")
</script>

<template>
  <OfficeThemeProvider :theme="theme">
    <OfficeToolbar label="文档工具栏">…</OfficeToolbar>
  </OfficeThemeProvider>
</template>
```

chrome 样式只作用于 `.office-*` 类，并在组件内部声明 `box-sizing: border-box`，
因此宿主没有全局 reset 时控件高度、48px toolbar、160px sidebar 仍然保持合同值。

## 2. 布局合同

| 项目 | 值 | 来源 |
|---|---|---|
| toolbar 最小高度 | `3rem` / 48px（min-height） | Extend `min-h-12` |
| toolbar 内边距 | `8px 12px` | Extend `px-3 py-2` |
| toolbar 组间距 | 4px / 8px | Extend `gap-1` / `gap-2` |
| toolbar 分隔 | `border-block-end` + `bg-background` | Extend `border-b bg-background` |
| icon button | 28px（`icon-sm`）/ 32px（`icon`） | Extend `size-8 sm:size-7` / `size-9 sm:size-8` |
| coarse pointer 命中区 | ≥ 44px | `--office-ui-touch-target: 2.75rem` |
| sidebar 宽度 | 160px（`10rem`） | Extend `w-40` |
| sidebar 断点 | 容器宽度 < 768px → overlay | `breakpoint` 默认 768 |
| sidebar 动画 | 200ms ease-out；`prefers-reduced-motion` 关闭 | Extend `duration-200` |

说明：desktop 下 toolbar 的实测外高是 **49px**（32px 控件 + 上下各 8px padding +
1px 底边）。`min-height: 48px` 是下限；Extend 参考在 `border-box` 下也是 49px。
coarse pointer 下 page input 提升到 44px 命中区，toolbar 实测约 61px。

## 3. 组件

### 3.1 `OfficeToolbar` / `OfficeToolbarGroup` / `OfficeToolbarSeparator`

```vue
<OfficeToolbar label="演示文档工具栏">
  <OfficeToolbarGroup gap="sm" grow>
    <OfficeIconButton label="显示页面" icon="file-text" size="icon-sm" />
    <OfficePageControl v-model="page" :total-pages="12" />
  </OfficeToolbarGroup>
  <template #trailing>
    <OfficeToolbarGroup gap="sm">
      <OfficeSearchPopover v-model="query" :result-count="4" />
      <OfficeToolbarSeparator />
      <OfficeZoomControl v-model="zoom" />
    </OfficeToolbarGroup>
  </template>
</OfficeToolbar>
```

| 组件 | Props | Slots / 事件 |
|---|---|---|
| `OfficeToolbar` | `label`、`wrap`（默认 `true`）、`sticky`、`class` | `default`/`leading`（左）、`center`、`trailing` |
| `OfficeToolbarGroup` | `label`、`gap: "sm" \| "md"`、`align`、`grow`、`wrap`、`class` | 默认 slot |
| `OfficeToolbarSeparator` | `class` | 无 |

根节点写 `role="toolbar"`、`data-slot="toolbar"`、`data-testid="office-toolbar"`。
分组写 `role="group"`、`data-slot="toolbar-group"`；分隔符写
`data-slot="toolbar-separator"`。`wrap` 关闭时根和分组都保持单行；默认允许换行。
390px 下 trailing 分组会自动占据整行（`@media (max-width: 30rem)`），不会压缩
点击区或覆盖正文。

### 3.2 `OfficeIconButton`

复用 #157 的 primitive，不重复实现。`variant` 默认 `ghost`，`size` 默认
`icon`（32px），`icon-sm` 为 28px；`active` 输出 `data-active="true"` 和
`aria-pressed="true"`；`disabled` / `loading` 与 `OfficeButton` 一致。
`label` 是必填的可访问名称。`@media (pointer: coarse)` 用伪元素把命中区扩展到
至少 44×44px。

### 3.3 `OfficePageControl`

```vue
<OfficePageControl v-model="page" :total-pages="12" />
```

| Prop | 默认 | 说明 |
|---|---|---|
| `modelValue` | — | 1 起始当前页；省略时组件自管，仍发事件 |
| `defaultPage` | `1` | 非受控初始页 |
| `totalPages` | 必填 | `0` 时显示 `–` 并禁用输入 |
| `disabled` | `false` | 禁用输入和键盘 |
| `label` | `common.page.current` | 输入框可访问名称 |
| `nativeInput` | `false` | 渲染原生 `<input>`，供宿主表单集成或非 DOM 渲染器使用 |
| `inputTestId` | `office-page-input` | 输入框的 `data-testid` |
| `class` | `""` | 合并到根 |

事件：`update:modelValue`、`change`。
键盘：`Enter` 提交、`Escape` 还原、`ArrowUp` / `ArrowDown` 步进、`Home` / `End`
到首/末页；输入越界会钳制到 `1..totalPages`，非法文本恢复当前值；失焦提交。

### 3.4 `OfficeZoomControl`

```vue
<OfficeZoomControl v-model="zoom" v-model:fit-mode="fitMode" />
```

`modelValue` 是百分比（100 = 实际大小），`fitMode` 为 `none` / `width` / `page`。
`options` 默认 `[50, 75, 100, 125, 150, 200]`；当前值不在预设中时会插入并按数值排序。
`showFit` 默认 `true`，select 中依次提供 Fit width / Fit page / Actual size 与百分比。
事件：`update:modelValue`、`update:fitMode`、`change: { zoom, fitMode }`。
`min` / `max` 默认 25 / 400，到达边界时对应按钮禁用。`disabled` 同时禁用三个控件。
`nativeSelect` 为 `true` 时用原生 `<select>` 替换 Reka select（百分比选项，fit
模式由宿主单独提供）；`selectTestId` / `selectTitle` 分别设置它的 `data-testid`
和原生 title。`OfficeToolbar` 内的 DOCX 工具栏使用这一模式以保留原生表单契约。

### 3.5 `OfficeSearchPopover`

```vue
<OfficeSearchPopover
  v-model="query"
  v-model:open="searchOpen"
  v-model:case-sensitive="caseSensitive"
  v-model:whole-word="wholeWord"
  v-model:regex="regex"
  :result-count="resultCount"
  :active-index="activeIndex"
  :status="searchStatus"
  @next="nextResult"
  @previous="previousResult"
  @submit="runSearch"
  @close="onSearchClose"
/>
```

| Prop | 默认 | 说明 |
|---|---|---|
| `modelValue` | — | 查询文本，省略时自管 |
| `resultCount` / `activeIndex` | `0` / `0` | 命中数和 0 起始活动命中 |
| `status` | `"idle"` | `idle` / `searching` / `ready` / `error` |
| `caseSensitive` / `wholeWord` / `regex` | `false` | 选项状态 |
| `showOptions` | `true` | 显示 Aa / ab\| / .\* |
| `open` | — | 受控展开；省略时自管 |
| `inline` | `false` | 不渲染触发器/portal，直接在原地渲染常驻搜索控件 |
| `inputType` | `"text"` | 原生输入类型；viewer 可设为 `search` |
| `statusText` | — | 覆盖本地化计数行（例如 DOCX 的 `1/4`） |
| `testIdPrefix` | `"office-search"` | 生成 `-input` / `-status` / `-previous` / `-next` / 选项 test id |
| `disabled` | `false` | 禁用触发器与内容 |

事件：`update:modelValue`、`update:open`、`update:caseSensitive`、
`update:wholeWord`、`update:regex`、`next`、`previous`、`submit`、`close`。
`submit` 在“有查询但无命中/尚未搜索”时按 Enter 触发，宿主可据此启动搜索。
打开后焦点进入输入框；`Enter` 下一个、`Shift+Enter` 上一个、`Escape` 关闭。
计数使用 `common.search.resultOf`（`{current}` / `{total}`），状态区
`role="status" aria-live="polite"`。触发器默认是 `OfficeIconButton`（search 图标，
28px），可用 `trigger` slot 替换。`inline` 模式保留同一输入、计数、选项、导航和
键盘契约，只是省去 trigger/portal，供工具栏常驻搜索使用；选项按钮在该模式下是
带 `office-toggle` 类的原生按钮，以兼容没有全局 `Element` 的自定义渲染器。

### 3.6 `OfficeStatusBar`

```vue
<OfficeStatusBar
  :page="3"
  :total-pages="12"
  file-name="Q3-review.pdf"
  status="Ready"
  warning="部分图片已优化"
  error="第 5 页渲染失败"
/>
```

`page` + `totalPages` 使用 `common.page.indicator`；`busy` 时显示 spinner 并把空
`status` 回退到 `common.loading`。warning 用 `role="status"`，error 用
`role="alert"`，两者都带图标和本地化前缀。`start` / `end` slot 可替换左右区域。
根节点是 `<footer>`，带 `aria-label`（默认 `common.chrome.statusBar`）。

### 3.7 `OfficeEmptyState` / `OfficeErrorState` / `OfficeLoadingState`

三个状态共享 `.office-state` 视觉（居中、图标、标题、描述、操作区），支持
`compact` 和 `class`。

| 组件 | 关键 props | 可访问性 / 事件 |
|---|---|---|
| `OfficeEmptyState` | `title`、`description`、`icon`（默认 `folder-open`） | `role="group"` + labelledby/describedby；`actions` slot |
| `OfficeErrorState` | `title`、`description`、`error`、`icon`（`triangle-alert`）、`showRetry`、`retryLabel` | `role="alert"`；默认 Retry 按钮发 `retry` |
| `OfficeLoadingState` | `title`、`description`、`progress`（0–100）、`showSpinner` | `role="status" aria-live="polite" aria-busy="true"`；`progress` 时输出 `role="progressbar"` |

`ErrorState` 的描述优先级：`description` → `error` → 目录默认文案。
`LoadingState` 的进度条宽度只由 `progress` 决定，可重复截图。

### 3.8 `DocumentViewerSidebar`

这是 #163 及后续 viewer 的**唯一** sidebar 实现，API 可复用：

```vue
<div class="viewer-shell"> <!-- position: relative -->
  <DocumentViewerSidebar
    v-model:open="sidebarOpen"
    :loading="thumbnailsLoading"
    :width="160"
    label="文档页面"
    @mode-change="onModeChange"
  >
    <template #header>页面</template>
    <ThumbnailList />
    <template #footer>共 12 页</template>
  </DocumentViewerSidebar>
  <div class="viewer-body">…</div>
</div>
```

| Prop | 默认 | 说明 |
|---|---|---|
| `open` | — | 受控展开；省略时自管 |
| `defaultOpen` | `true` | 非受控初始值 |
| `mode` | `"auto"` | `auto` / `inline` / `overlay` |
| `breakpoint` | `768` | auto 模式 inline/overlay 分界（容器宽度，含 768） |
| `width` | `160` | 数字按 px，字符串原样写入 CSS 变量 |
| `loading` | `false` | 用内置 skeleton 替换默认 slot |
| `label` | `common.chrome.sidebar` | `<aside>` 可访问名称 |
| `closeOnEscape` | `true` | overlay 打开时 Esc 关闭并发 `update:open=false` |
| `container` | `null` | auto 模式测量的元素；默认 `parentElement` |
| `class` | `""` | 合并到根 |

事件：`update:open`、`open-change`、`mode-change`（`inline` / `overlay`）。
Slots：`header`、默认内容、`footer`。`defineExpose`：`inline`、`mode`、`open`、
`close()`。

DOM 标记：`data-document-sidebar`、`data-sidebar-mode`、`data-sidebar-open`、
`data-sidebar-inline`、`data-transitions-ready`，以及
`data-slot="document-sidebar"` / `data-testid="document-viewer-sidebar"`。
关闭时写 `aria-hidden="true"` 和 `inert`。

行为：

- `auto` 用 `useElementWidth()` 观察 `container`（默认父元素）。隐藏/脱离文档时保留
  最后一个非零宽度，避免 keep-alive 或 `display:none` 让布局来回跳。
- inline 打开占据 160px；关闭用负 `margin-inline-start` 收起并让内容平滑扩展。
- overlay 绝对定位在最近的 positioned 祖先内；关闭用 `translateX(-100%)` 并禁用指针。
- 初始两帧关闭 transition（`data-transitions-ready="false"`），之后才启用 200ms
  动画；`prefers-reduced-motion: reduce` 下完全没有 transition。
- `DocumentViewerSidebarSkeleton` 可单独使用，默认 3 张 80×112 卡片，`role="status"`。

`useElementWidth(target)` 也作为公开 composable 导出，供 #163 等需要按容器宽度
决策的宿主复用。

## 4. 受控 / 非受控约定

`OfficePageControl`、`OfficeZoomControl`、`OfficeSearchPopover`（查询与 open）和
`DocumentViewerSidebar`（open）都遵循同一套约定：

- 传了 `modelValue` / `open` / `fitMode` 等 prop → 受控，组件只发 `update:*`，
  不擅自改内部值；
- 不传 → 非受控，组件内部维护状态，同时仍发 `update:*`；
- 所有交互事件都同时带语义事件（`change`、`next`、`retry`、`mode-change`），
  宿主不必从 `update:*` 反推用户意图。

## 5. i18n

新增键全部在 `@arcships/office-interaction` 的 `OfficeMessages` 中，en / zh-CN
同步提供；`OfficeChromeMessages` 已导出，便于宿主写完整目录。

| 路径 | 用途 |
|---|---|
| `common.chrome.toolbar` | 工具栏默认可访问名称 |
| `common.chrome.sidebar` / `sidebarToggle` | sidebar 与 toggle 名称 |
| `common.chrome.statusBar` | 状态栏可访问名称 |
| `common.chrome.fileName` | 文件名前缀 |
| `common.chrome.error` / `warning` | 状态栏消息前缀 |
| `common.chrome.retry` | 错误状态重试按钮 |
| `common.chrome.emptyTitle` / `emptyDescription` | 空状态默认文案 |
| `common.chrome.errorTitle` / `errorDescription` | 错误状态默认文案 |
| `common.chrome.loadingTitle` / `loadingDescription` | 加载状态默认文案 |
| `common.zoom.fitWidth` / `fitPage` / `actualSize` / `percent` | 缩放预设与百分比模板 |

组件不包含硬编码的用户可见文案；宿主可以用 `OfficeThemeProvider` 的 `messages`
或 viewer 的 `messages` prop 覆盖任意键。

## 6. 键盘 / 焦点 / 读屏

| 位置 | 合同 |
|---|---|
| toolbar | `role="toolbar"` + `aria-label`；分组 `role="group"` |
| icon button | 必须有 `label`；active 写 `aria-pressed`；coarse pointer ≥44px |
| page input | `inputmode="numeric"`、`pattern="[0-9]*"`、`aria-label`；Enter/Escape/↑/↓/Home/End |
| zoom | select 使用 reka-ui combobox/listbox；方向键/Enter/Esc 由 primitive 提供 |
| search | 触发器 `aria-haspopup="dialog"` + `aria-expanded`；打开后焦点进入输入框；Enter/Shift+Enter/Escape；计数 `role=status` |
| status bar | page/status 文本；warning `role=status`、error `role=alert` |
| states | empty `role=group`、error `role=alert`、loading `role=status` + `aria-busy`；标题/描述用 `useId` 关联 |
| sidebar | `<aside>` + `aria-label`；关闭时 `aria-hidden` + `inert`；overlay 打开时 Esc 关闭 |

## 7. 响应式与视觉证据

截图脚本在 `scripts/ci/capture-office-chrome.py`。它在本进程内启动一个临时静态
服务器（不会泄漏 Vite/preview 子进程），跑完在 `finally` 里关闭：

```bash
pnpm --filter demo build
PYTHON=/tmp/bb-venv/bin/python RUN_PYTHON_MODULES=playwright \
  node scripts/ci/run-python.mjs scripts/ci/capture-office-chrome.py --serve apps/demo/dist
```

产物：`output/office-chrome/<viewport>/<theme>/`（`*-full-page.png`、`*-toolbar.png`、
`*-sidebar-open.png`、`*-sidebar-closed.png`、`*-sidebar-skeleton.png`、
`*-states.png`、`*-search-popover.png`、`*-zoom-options.png`、`*-status-error.png`）
和 `summary.json`。

2026-11-20 在本分支的实测结果（Chromium headless，DPR 1，1440/1280/768/390 ×
light/dark 共 8 例，全部 PASS、0 console error）：

| viewport | toolbar 外高 | 换行 | sidebar | container | close 后内容位移 | pointer |
|---|---|---|---|---|---|---|
| 1440×900 | 49px | 否（单行） | inline 160px | 1390px | −160px | fine |
| 1280×720 | 49px | 否（单行） | inline 160px | 1230px | −160px | fine |
| 768×1024 | 61px | 否（coarse 控件 44px） | overlay 160px | 742px | 0 | coarse，44px |
| 390×844 | 113px | 是（trailing 独占一行） | overlay 160px | 364px | 0 | coarse，44px |

每个用例还断言：页面和 toolbar 无横向溢出、控件不越出 toolbar、toolbar 与正文不
重叠、search popover 不超出视口、inline/overlay 切换 shell 高度不变、过渡时长为
0.2s、coarse pointer 命中区 44×44、page input 键盘提交/钳制。`summary.json` 的
`accessibility` 字段记录手工扫描结果（无未命名按钮/表单、无重复 id、无空 live
region）。

## 8. 测试与 CI

新增并被 `scripts/ci/run-suite.mjs` 注册：

```bash
node --test --test-concurrency=1 tests/component/vue-ui-chrome.test.mjs
node --test --test-concurrency=1 tests/unit/vue-ui-chrome-contract.test.mjs
pnpm --filter @arcships/vue-ui typecheck
pnpm --filter @arcships/vue-ui build
```

- 组件测试：toolbar 结构、icon button 状态、page/zoom/search 键盘、status 与
  empty/error/loading、sidebar auto/inline/overlay、ResizeObserver、skeleton、
  open/close、zh-CN。
- 合同测试：Extend 几何、token-only CSS、公开导出、i18n 键、gzip 预算和 public
  API fingerprint。
- 包体积：`dist/index.js` 43,391 B gzip + `style.css` 11,567 B gzip，相对 #157 前
  基线增加 35,818 B，仍在 35,840 B 预算内。
- `public-api-contract.json`：`@arcships/vue-ui` fingerprint 从 199/`8a1aff85…`
  更新为 233/`1479c668…`；`@arcships/office-interaction` 因新增导出类型
  `OfficeChromeMessages` 从 154/`e1b91beb…` 更新为 155/`5732cb82…`。

## 9. 接入顺序与已知边界

- #159–#162 按 DOCX / XLSX / PPTX / PDF 顺序替换各自 toolbar、status 与
  thumbnail sidebar；本 issue 不改四个 viewer 的公开 props 或内部样式。
- `OfficeFindBar` 保留给旧 host 和已有的 `Surface` 搜索条；新 chrome 使用
  `OfficeSearchPopover`。
- overlay sidebar 需要 positioned 容器；否则会相对更外层定位。文档和 demo 使用
  `.chrome-demo__body`（`position: relative`）。
- 暂不提供右侧 sidebar、拖动改变宽度或多面板；这些可由宿主在现有组件外层组合。
- axe-core 由 #165 接入；本 issue 的手工检查覆盖角色、名称、重复 id 和 live region，
  组件测试覆盖键盘与焦点行为。

## 10. DOCX 迁移（issue #159，2026-11-20 重开补齐）

参考面与产品面都以固定 commit `70fe70c6` 的 Extend DOCX 组件为唯一布局来源：

- 参考面 `tools/extend-reference/src/vendor/extend/docx-toolbar.tsx` 是
  `apps/v4/components/extend/docx-viewer.tsx` 的 `DocxToolbar`（含
  `DocxPageNumberControl`、`DocxFileActionsMenu`）逐类名拷贝；
  `tools/extend-reference/src/vendor/extend/docx-sidebar.tsx` 是
  `DocxThumbnailSidebarContent`/`DocxThumbnailSidebarList` 的静态版本。
  旧的手写「Extend 风格完整控件集」参考面（文件/搜索/版式/修订/主题/上传/下载）
  已删除，因为它不是上游默认行。
- `DocxViewerToolbar` 默认节点就是上游行：48px `min-h-12 px-3 py-2` 根节点，
  左侧缩略图开关 + `Page`/页码/`of N`，右侧缩放减/档位/加、分隔线和 `⋯` 动作菜单。
  图标使用上游 Hugeicons（`HugeIcon`），不再用 Lucide 的等价图标。
- 上游行里没有的产品命令收进动作菜单（`.docx-viewer-toolbar__menu`）：
  Find、Comments、Edits、主题、上传、下载；`docx-show-comments`、
  `docx-show-tracked-changes`、`docx-theme-toggle`、`docx-upload`、`docx-download`、
  `docx-find` 的 test id 不变，但只在菜单打开时存在。
  版式/适配/审阅显示是上游没有的 host 能力，改为可选 props
  （`showViewControls` / `showReviewControls`，默认 `false`），同样只在菜单里出现；
  文件名字段是 `showFileName`（默认 `false`），内联搜索是 `showSearch`（默认
  `false`）。`DocxViewer` 透传为 `toolbarSearch` / `toolbarFileName` /
  `toolbarViewControls` / `toolbarReviewControls`，全部默认关闭。
- 原生表单契约保留：`docx-page-current` 仍是 `<input>`（`type=text`
  + `inputmode=numeric`，聚焦时草稿不被 `currentPage` 覆盖），`docx-zoom-select`
  仍是覆盖在触发器之上的原生 `<select>`；`docx-zoom-out` / `docx-zoom-in` 为按钮。
  `Ctrl/Cmd+F` 仍打开同一个 find 弹层并聚焦输入框。
- `DocxThumbnailPanel` 与上游一致：160px `w-40` rail、无 header、16px 轨道留白、
  12px 行内缩进、92px `FileThumbnail` 页卡 + `text-xs` 页码 + `pb-3` 行距、
  `bg-sidebar-accent` 选中态和 `p-4` 骨架；行使用上游的绝对定位 + `translateY`
  （同时保证嵌套 canvas 的圆角栅格化与参考面一致）。`data-testid="docx-thumbnail"`、
  `docx-thumbnail-label`、`data-thumbnail-page-index/-label/-parity-blank`、
  `data-thumbnail-state` 保持不变；页数改由 listbox 的 `data-page-count` 读取。
- `DocxToolbar`（编辑器工具栏）不在本次参考面范围内，保持既有两行结构、类名和
  `editor-zoom` 原生 select。

证据：`pnpm ui:parity -- --only docx-toolbar,docx-sidebar` 全矩阵
（4 视口 × 2 主题 × toolbar 4 状态 / sidebar 3 状态 = 56 例）56/56 PASS，
diff 像素 0、diff ratio 0、几何 delta 0（详见
[docx-ui-stories.md](./docx-ui-stories.md#与-155-parity-harness-的关系)）。
`tests/component/docx-*.test.mjs` 与 `tests/blackbox/docx_*.py` 已同步到新的
默认行（菜单内命令、`toolbarSearch` 选项）。

## 11. PPTX 迁移（issue #161）

PPTX 的生产工具栏是 `PptxViewerToolbar`，行结构、控件顺序、32/28px 控件、
`Slide N of M` 和 104px 缩放触发框逐项对齐 Extend `PptxToolbar`；它没有套
`OfficeToolbar` 的 leading/center/trailing 外壳，否则多出的 wrapper 会改变
390px 下的换行位置（像素矩阵会直接暴露 2px 偏差）。工具栏内的按钮仍复用
`OfficeIconButton` 与 Hugeicons，菜单复用共享 dropdown 几何。

布局/演示/备注/全屏/查找不在上游默认行里，统一收进右侧 `⋯` 操作菜单；搜索
`Ctrl/Cmd+F` 打开工具栏下方的 find bar。证据与像素数字见
[pptx-chrome.md](./pptx-chrome.md#7-测试与像素证据)。
