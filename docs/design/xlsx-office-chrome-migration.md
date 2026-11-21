# XLSX chrome 迁移到统一设计系统（#160）

> 批次：0.8 / issue [#160](https://github.com/arcships/agentic-office-ui/issues/160)。
> 依赖：[#156](https://github.com/arcships/agentic-office-ui/issues/156) token + `OfficeIcon`、
> [#157](https://github.com/arcships/agentic-office-ui/issues/157) primitive 层、
> [#158](https://github.com/arcships/agentic-office-ui/issues/158) 共享 Office chrome。
> 参考基准：[extend-hq/ui@70fe70c6](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)。

本文记录 XLSX viewer/authoring 的 toolbar、ribbon、菜单、对话框和状态栏如何落到共享层。
菜单与对话框现在**真正运行在 Reka 版 `OfficeDropdownMenu` / `OfficeDialog` 上**（第三轮 0.8 复审）。

**0.8 重开结论（第四轮）**：上游 chrome markup 是自包含的，因此 `xlsx-toolbar`、
`xlsx-sheet-tabs`、`xlsx-status-bar` 三条 `ui-parity` surface 已经建成并在
4 viewport × light/dark × 全部 state 上**全部 0px**（64/64 pass），默认行只保留上游控件，
其余命令移进上游本来就有的 ⋯ 菜单。细节见 §4.2。

**#170 补完（第五轮）**：`xlsx-editor.tsx` 的 `EditorToolbar` 格式行（undo/redo、字体/字号、
B/I/U/S、文字/填充色、水平/垂直对齐、换行、数字格式、边框、合并）已经以**静态 resting 子集**
vendored 到 `tools/extend-reference/src/vendor/extend/xlsx-ribbon.tsx`，并新增 `xlsx-ribbon`
surface：4 viewport × light/dark × default/hover/focus/disabled = **32/32 pass、0 AA、0px**。
ribbon 的 17 个 inline `<svg>`、8 个 Lucide registry 图标（undo/redo、三个水平对齐、link、table、
check）以及 B/I/U/S 的文本字形全部换成同一套 Hugeicons（上游 `IconPlaceholder` 的 hugeicons 字形），
`XlsxRibbon.vue` 里不再有 inline `<svg>`；`xlsx-office-chrome.test.mjs` 断言 29 个 Hugeicons
（24x24 viewBox + 自带 stroke 数据）。
仍未覆盖的只有**公式行**（`Group` + 私有 `@extend-ai/react-xlsx` 编辑态），以及上游格式行里
本仓库 ribbon 没有的命令（pivot、sheet 选择器、zoom）——它们分别属于 sheet-tabs/toolbar
surface 或产品缺口。

## 0. 相对 issue 验收清单的审计（0.8 复审）

| issue 清单项 | 状态 | 证据 |
|---|---|---|
| `XlsxToolbar` / `XlsxRibbon` 改用 primitives；48px toolbar、28/32px icon、gap/separator | ✅ | §1 表；§4.3 toolbar=48px、`office-button--size-icon-sm` |
| 缩放/搜索/sheet tabs/冻结/打印/导出/Worker 状态控件对齐 | ✅ | 缩放/搜索/冻结/打印/导出：§1；**sheet tabs：本轮补 provider + token（10 处三元色值清零）**；Worker 状态：viewer 无独立控件，由 `OfficeLiveStatus`（`xlsx-live-status`）+ 只读/编辑锁 gate + grid `data-worker-backed` 暴露（与上游一致） |
| 右键/表头/图表/绘图菜单统一用 DropdownMenu/ContextMenu | ✅ 真正使用 `OfficeDropdownMenu`（Reka），含子菜单/键盘/焦点/关闭 | `xlsx-menus.test.mjs`（happy-dom，12 条）、§4.3 |
| authoring dialog 用 Dialog/Sheet | ✅ 真正使用 `OfficeDialog`（Reka），含焦点陷阱/Esc/backdrop/关闭后还焦 | `xlsx-authoring-dialog.test.mjs`（happy-dom，16 条）、§4.3 |
| 状态栏用 `OfficeStatusBar`；空/加载/错误用 `Office*State` | ✅ | `xlsx-office-chrome.test.mjs` |
| 保留公式编辑/AutoFilter/冻结/图表/绘图/Worker/导出/RTL/DOM render | ✅ | §5 + 750 条 XLSX 测试 |
| 截图 story：viewer/authoring、light/dark、390/768/1440 | ✅ | §4.3 每 viewport × theme 输出 viewer/toolbar/menu/**authoring** 截图 |
| XLSX toolbar/ribbon 与 Extend 参考 diff ≤0.1% | ✅ toolbar + sheet tabs + status bar + ribbon 全部 0px | §4.2：`xlsx-toolbar` 32/32、`xlsx-sheet-tabs` 24/24、`xlsx-status-bar` 8/8、`xlsx-ribbon` 32/32，合计 96/96 pass；公式行仍缺上游对照（见 §4.2 例外） |
| 现有 XLSX unit/component/blackbox 全绿 | ✅（本轮实测） | unit 275 + component 475（含新增 happy-dom 套件）+ touch blackbox 20/20 + xlsx 导出/Worker 黑盒 |
| 菜单键盘/焦点/关闭不回归 | ✅ | §4.3 浏览器断言 + 组件测试 |
| 移动端 toolbar 换行且 grid 可用高度 | ✅（改为“单行不裁切”） | 390px：上游行不换行（`wrapped=false`，48px），grid 360px；`verify-xlsx-chrome-interaction.py` 断言见 §4.2.1 |

## 1. 结论速览

| 区域 | 迁移后 | 说明 |
|---|---|---|
| `XlsxToolbar` | `OfficeToolbar` + `OfficeToolbarSeparator` + `OfficeIconButton` + `HugeIcon` + ⋯ 菜单 | 默认行 = 上游 `[zoom out][100%][zoom in] │ [search] │ [⋯]`；下载/导出/上传/打印/分页在菜单里 |
| `XlsxRibbon` | 第二个 `OfficeToolbar` 行 + 每组 `OfficeToolbarGroup` + `OfficeButton` / `OfficeIconButton` | 原来的 disabled `<fieldset>` 换成显式 `controlsDisabled` |
| `XlsxStatusBar` | `OfficeStatusBar` + Extend 状态条外观（28px / 居中 / muted） | 选区地址、统计、note 的 `data-testid` 全部保留；段分隔用 `::before` 不污染文本 |
| 空/加载/错误 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` | 仍可被 `#empty` / `#loading` / `#error` 插槽覆盖 |
| 右键/表头/图表/绘图菜单 | `OfficeDropdownMenu`（Reka）：`items` 模型 + 指针锚点定位 | 键盘/typeahead/子菜单/焦点/关闭全部由 primitive 层提供（见 §3） |
| AutoFilter 面板 | 同上（checkbox 行用 `keepOpen` 保持多选） | |
| Authoring 对话框 | `OfficeDialog`（Reka）：overlay + header/title/body + close | 焦点陷阱、Esc、backdrop、关闭还焦由 primitive 层提供（见 §3） |
| Sheet tabs | 上游 `TabsList`/`TabsTrigger` 几何（`bg-muted`、`p-0.5`、`max-w-48`、`sm:h-8`）+ token | 无 inline 色值；新增/删除/翻页按钮 = `show-sheet-actions`（默认关） |
| 主题 | viewer/各 chrome 组件用 `OfficeThemeProvider` 提供 `data-office-theme` | 原有 `.xlsx-*--dark` 标记保留给黑盒/矩阵，颜色一律走 token |

`XlsxToolbar` / `XlsxRibbon` / `XlsxStatusBar` / `XlsxContextMenu` / `XlsxAuthoringDialog`
的 **public props/emits 一字未改**。新增两个低层出口 `XlsxHeaderMenu` / `XlsxAutoFilterMenu`
（`XlsxGrid` 已自动挂载），因此 `scripts/ci/public-api-contract.json` 中 `@arcships/vue-xlsx`
的声明指纹按预期从 103 更新到 105 项（sha256 `8663ede6…`），见
[api/public-api-contract.md](../api/public-api-contract.md)。

## 2. 为什么要 `OfficeThemeProvider`

设计 token 全部声明在 `[data-office-theme]` 作用域下（见 `packages/vue-ui/src/styles/tokens.css`）。
XLSX 的 `isDark` 是 prop，不是全局主题，所以每个装配点自己提供主题：

- `XlsxViewer` 包一层 provider（`theme = isDark ? 'dark' : 'light'`，同时传 `locale` / `messages`）；
- `XlsxToolbar` / `XlsxRibbon` / `XlsxStatusBar` / `XlsxAuthoringDialog` 各自也包一层，
  这样它们作为公开组件被宿主单独挂载时仍然拿到 token；
- provider 宿主是 `display: contents`，不参与布局，flex 高度不会变。

两条容易回归的宿主约束（都有浏览器断言兜底，见 §4.3）：

1. **chrome 行不可被压缩**：`.xlsx-viewer` 是 flex column，若 toolbar/ribbon/statusbar 没有
   `flex: 0 0 auto`，390px 下它们会被压到 `min-block-size` 并裁掉换行出来的第二行
   （实测修复前 toolbar 48px 高但内容溢出、ribbon 51px 而内容 132px+）。现在三者都固定
   `flex: 0 0 auto`，toolbar 在 390px 长到 84px。
2. **原生控件不参与宿主页面样式**：缩放/字号/数字格式/边框下拉都带 `.office-ui` 类，
   因此 demo 页面里 `.page :is(button, select, …):not(.office-ui)` 的 `min-height: 2.25rem`
   不会把它们撑到 36px（修复前 toolbar 实测 53px，修复后 48px）。
   同理，图标按钮里的内联 SVG 显式声明 `fill: none; stroke: currentColor`，
   否则会被默认填充吞掉（页分隔预览按钮曾整颗图标不可见）。

## 3. Reka 覆盖层接法：指针锚点 + 共享 item 模型

菜单与对话框现在直接渲染共享 primitive。为了让既有的命令式 API、`data-testid` 契约和
测试选择器都保持不变，接法如下：

### 3.1 菜单：`OfficeDropdownMenu` + 不可见指针锚点

`XlsxContextMenu` / `XlsxHeaderMenu` / `XlsxAutoFilterMenu` 都渲染
`<OfficeDropdownMenu :items="…">`，并在 `#trigger` 插槽里放一个 0×0、`position: fixed` 的
锚点元素，定位到指针坐标。于是：

- 打开方式仍是命令式的（网格右键命中、图表/绘图覆盖层、表头按钮、AutoFilter 触发按钮），
  但内容、子菜单、typeahead、方向键/Home/End、Esc/Tab 关闭与焦点归还全部由 Reka 提供；
- 菜单始终通过 `data-testid`（`xlsx-context-menu`、`xlsx-column-menu-list`、
  `xlsx-row-menu-list`、`xlsx-autofilter-menu`）与原有 item testid 暴露，黑盒选择器不变；
- 子菜单（chart type 家族/类型、paste special）改用 `type: "sub"`，由 Reka 负责展开与收起，
  不再有手写的 `ArrowRight`/`ArrowLeft`/`mouseenter` 分支；
- 菜单是 `modal` 的（与旧的全屏 backdrop 等价）：长按合成菜单后 `touchend` 不会立刻把它关掉
  （`TOUCH-XLSX-LONG-PRESS` 通过），点击菜单外关闭。
- 关闭后焦点交给 `restoreFocus` 回调：单元格菜单回到网格元素，表头/AutoFilter 菜单回到网格容器。

为了承载这些契约，primitive 层做了**向后兼容的增量扩展**（默认行为不变）：

| 扩展 | 位置 | 用途 |
|---|---|---|
| `OfficeMenuEntry.testId`（item/checkbox/label/separator/group/sub） | `components/ui/types.ts` + `OfficeMenuItems.ts` | 把历史 `data-testid` 固定到 Reka 渲染的行上 |
| `OfficeMenuSubmenu.contentTestId` | 同上 | 子菜单容器仍需 `xlsx-chart-type-menu` 等 testid |
| `OfficeMenuItem/CheckboxItem.keepOpen` | 同上 | AutoFilter/多选行选中后不关闭菜单（`event.preventDefault()`） |
| `OfficeDropdownMenu.contentTestId` / `contentData` / `contentAriaLabel` / `modal` / `restoreFocus` | `OfficeDropdownMenu.vue`（`OfficeContextMenu.vue` 同步） | 菜单 surface 的 testid、`data-menu-kind`、无障碍名称、非模态与焦点归还 |
| checkbox 行改用 `modelValue` | `OfficeMenuItems.ts` | reka-ui 2.11 的 `DropdownMenuCheckboxItem` 只认 `modelValue`；此前 `checked` 被忽略，勾选状态从不显示（顺带修掉的 primitive bug） |

### 3.2 对话框：`OfficeDialog`

`XlsxAuthoringDialog` 现在渲染 `<OfficeDialog :open="true" …>`：

- `title` 是原来的 `dialogTitle`，`closeTestId` 保持 `xlsx-authoring-close`，
  `contentTestId`/`contentData` 保持 `xlsx-authoring-dialog` + `data-kind`；
- 表单（数字格式、条件格式、数据验证、图表、批注、超链接、表格、粘贴选项、名称管理器）
  原样放在默认插槽里，状态/校验逻辑一字未改；
- 焦点陷阱、Esc、点击遮罩关闭由 Reka 提供；`XlsxViewer.closeAuthoring()` 在关闭后把焦点还给
  打开者（`OfficeDialog` 的 `update:open(false)` → 我们 emit `close`）。
- 旧的手写 `Teleport`/backdrop/header/close 按钮已删除；面板宽度/最大高度仍由
  `.xlsx-authoring.office-dialog__content` 收紧到 520px / 82vh。

**#171 补充**：页面设置没有做成上面的“一个 `kind` 一个 section”，而是独立组件
`XlsxPageSetupDialog.vue`（同样渲染 `<OfficeDialog :open="true">`，复用同一个
`.xlsx-authoring` 字段样式与宽度约束），因为它的字段是工作表级打印状态而不是选区对象。
`XlsxAuthoringDialog` 在 `kind="pageSetup"` 时直接转交该组件，所以 `XlsxViewer` /
`XlsxSheetSurface` 的挂载点与焦点归还逻辑都没有重复；`content-test-id` 为
`xlsx-page-setup-dialog`，关闭按钮为 `xlsx-page-setup-close`。

### 3.3 测试怎么迁的（不遮罩、不退自定义）

最小 host renderer 没有真实 DOM，Reka 在它与 `document` 之间无法工作。因此：

- **新增** `tests/component/xlsx-menus.test.mjs`（happy-dom，12 条）：单元格/表头/行头/AutoFilter/
  图表/绘图菜单的 testid、命令、只读禁用、子菜单、`keepOpen` 多选、右键命中 kind、
  ArrowDown/焦点归还、Esc 关闭、以及"确实使用 OfficeDropdownMenu 而不是本地标记"的源码断言；
- **新增** `tests/component/xlsx-authoring-dialog.test.mjs`（happy-dom，16 条）：shell/标题/焦点陷阱/
  Esc/backdrop/关闭按钮、数字格式、批注、超链接、表格、条件格式、数据验证、图表、粘贴选项、
  名称管理器的表单 → 控制器调用；
- 旧文件里这些"假 renderer 驱动 Reka 组件"的用例被移除（覆盖已迁移，不是删掉了事）：
  `xlsx-structure`、`xlsx-authoring-ui`、`xlsx-auto-filter`、`xlsx-chart-type-menu`、
  `xlsx-drawing-clipboard`、`xlsx-cut`、`xlsx-menu-state`、`xlsx-office-chrome`；
  纯控制器/模型断言全部保留；
- 真实浏览器路径（`pnpm test:xlsx:chrome` + `touch_mobile.py`）继续覆盖集成行为。

已知基线（**不是** #160 引入，未在本分支修复，留给对应 issue）：

- `tests/component/i18n-source.test.mjs` 的 "TypeScript sources contain no hard-coded
  translatable copy" 在 `packages/vue-ui/src/components/ui/labels.ts` 上报中文
  `关闭`——#157 的 primitive 本地化 seam；
- 同文件 "the previously hard-coded strings now resolve through the catalog" 在
  `vue-docx/src/components/DocxViewerToolbar.vue` 上要求 `m.docx.toolbar.trackedChanges`——#159 的合并遗留。

两者均与本轮文件无关（`git diff 2df2ca36 -- <file>` 为空）。

## 4. 视觉对照证据

### 4.1 共享 chrome 截图矩阵

`scripts/ci/capture-office-chrome.py` 覆盖 1440x900 / 1280x720 / 768x1024 / 390x844 × light/dark，
新增 `xlsx_story_flow`：截图 `xlsx-chrome` / `xlsx-ribbon`，并断言

- toolbar ≥ 48px、toolbar/ribbon 无水平溢出（`scrollWidth - clientWidth <= 1`）；
- toolbar 在 390px 换行（`trailing.top > leading.top`）；
- toolbar ≥ 5 个、ribbon ≥ 10 个共享 `.office-icon-button`；
- story 内没有未命名的 button。

运行方式（demo 需要先构建，脚本自起静态服务器）：

```bash
pnpm --filter demo build
python3 scripts/ci/run-python.mjs scripts/ci/capture-office-chrome.py --serve apps/demo/dist
```

产物：`output/office-chrome/<viewport>/<theme>/<prefix>-xlsx-chrome.png`
与 `<prefix>-xlsx-ribbon.png`。390px 下 ribbon 换行后高度上限 132px（可滚动），
保证 grid 仍有可用高度；1440px 下 ribbon 为两行（Excel 风格分组换行）。

### 4.2 与 Extend 参考的像素对照（#160 重开后的结论）

**结论**：`xlsx-toolbar` / `xlsx-sheet-tabs` / `xlsx-status-bar` 三条 surface **已经可达且全部 0px**。
上一轮“不可达”的判断只对**私有引擎**成立，对 **chrome markup** 不成立：
`apps/v4/components/extend/xlsx-viewer.tsx` 里 `WorkbookToolbar`、`WorkbookSheetTabsInner`、
`WorkbookFileActionsMenu`、`WorkbookSearchPopover` 的 JSX 与 class 串都是自包含的
（shadcn primitives + Hugeicons + 两个 `useXlsxViewer*` hook），只有 `<XlsxViewer>` 网格本身依赖
`@extend-ai/react-xlsx`。因此本轮把 **chrome 子集**（约 260 行）vendored 进参考侧，
用 props 替代两个 hook，不改 class 串、不改图标：

| 新增文件 | 上游来源 | 精简内容 |
|---|---|---|
| `tools/extend-reference/src/vendor/extend/xlsx-toolbar.tsx` | `xlsx-viewer.tsx`（`WorkbookToolbar` / `ToolbarTooltip` / 两个 trigger） | tooltip/popover/dropdown 弹层不渲染，只保留闭合行的 trigger |
| `tools/extend-reference/src/vendor/extend/xlsx-sheet-tabs.tsx` | `xlsx-viewer.tsx`（`WorkbookSheetTabsInner`）+ registry `tabs.tsx` / `scroll-area.tsx` class 串 | 去掉缩略图预览（仅 hover 出现），Base UI `Tabs`/`ScrollArea` 换成同 class 的静态元素 |
| `tools/extend-reference/src/vendor/extend/xlsx-status-bar.tsx` | `file-system.tsx` 的 `aria-live` 状态条 | 上游 XLSX **没有**状态栏，取 Extend 自己的状态条作参考（见 §4.4） |

实测（`PYTHON=/tmp/bb-venv/bin/python pnpm ui:parity -- --only xlsx-toolbar,xlsx-sheet-tabs,xlsx-status-bar`）：

| surface | case | 结果 |
|---|---|---|
| `xlsx-toolbar` | 4 viewport × light/dark × 4 state = 32 | **32 pass，0 fail，0px**（`pass-with-AA` 也为 0） |
| `xlsx-sheet-tabs` | 4 × 2 × 3 = 24 | **24 pass，0 fail，0px** |
| `xlsx-status-bar` | 4 × 2 × 1 = 8 | **8 pass，0 fail，0px** |
| `xlsx-ribbon` | 4 × 2 × 4 = 32 | **32 pass，0 fail，0px** |
| 合计 | 96 | **96/96 pass** |

矩阵与阈值（`geometryTolerancePx: 0`、`maxDiffPixelRatio: 0.001`）一字未改。
#165 在 0.8 集成分支上复跑同一命令仍为 **64/64 pass、0 AA**，随后把三条 surface 切为
`gateStatus: blocking` 并提交对照侧基线（8 state × 2 viewport × 2 theme = 32 张）；
采集/审批流程见 [ui-parity-ci.md §5](../testing/ui-parity-ci.md)。`xlsx-ribbon` 在 #170 补上
参考/对照与基线（16 张，审批见 `tests/visual/extend/APPROVALS.md`）。

仍然**不可达**的对照（逐项列出，不再用“ribbon unreachable”概括）：

- `xlsx-formula-bar`：上游公式行（`xlsx-editor.tsx` L1355 起）的 `Group`/`GroupText` 来自私有
  `@extend-ai/react-xlsx` 编辑态，且没有独立静息态可渲染；留给后续 issue。
- 上游格式行里本仓库 ribbon 没有的命令：pivot insert/fields（产品缺口，`not planned`）、
  add/remove sheet + active-sheet select（由 `xlsx-sheet-tabs` surface 覆盖）、zoom cluster
  （由 `xlsx-toolbar` surface 覆盖）。
- 本仓库 ribbon 有、上游格式行没有的 cluster：clipboard（cut/copy/paste/paste special）、
  structure（插入/删除行列）、authoring（图表/批注/超链接/表格/条件格式/数据验证/冻结/名称管理器）、
  read-only 开关。这些由 `tests/component/xlsx-*.test.mjs` 与 `xlsx-chrome` 截图覆盖，
  `xlsx-ribbon` surface 不伪造上游对照。
- 颜色控件的 **text colour** well 保留 Excel 的 “A” 字形（fill/border 用 Hugeicons）；
  它不在像素对照的控件集合内。

#### 4.2.1 默认 chrome 只剩上游控件

按“默认只显示上游有的控件”的要求，生产默认行现在是：

```
[zoom out] [100% ▾] [zoom in] │ [search] │ [⋯]
```

- 缩放/搜索/⋯ 触发器：`OfficeIconButton size="icon-sm"` + **Hugeicons** 原始图标（与参考同字形）；
- 缩放选择：把上游 `SelectTrigger size="sm" w-[84px]` 的闭合形态画出来，再叠一个透明原生
  `<select>` 承接键盘/平台弹出（`data-testid="xlsx-zoom-select"` 不变）；
- 上游本来就有的 **⋯ 文件操作菜单**：下载、导出、上传、打印预览、分页预览全部进去，
  外加文件名 label；`data-testid` 全部保留（`xlsx-toolbar-download`/`-export`/`-upload`/
  `-print-preview`/`-page-breaks`），只是从行内移到菜单里；
- 文件名不再占行内位置（上游行内没有 identity）；
- sheet tabs 默认只渲染 tab 本身；Excel 的新增/删除/左右翻页按钮改为 `show-sheet-actions`
  （默认 `false`，host 显式打开才出现）；tab 右键菜单、拖拽排序、重命名、隐藏/取消隐藏不变。

旧几何断言（toolbar 在 390px 换行、≥5 个 icon button）也按新行更新：
上游行在 390px **不换行**（内容只有 ~250px），现在是 `wrapped=false` + `rowControls>=4`。

#### 4.2.2 参考侧 class 串到 Vue 的落点

| 层 | 指标 | 参考值 | 实测 | 状态 |
|---|---|---|---|---|
| geometry | toolbar 高度（各 viewport） | 48px（`min-h-12`） | 48px | ✅ 0px |
| geometry | 图标按钮 | `size-8 sm:size-7` | 28px（≥640）/ 32px（<640） | ✅ 0px |
| geometry | 缩放 trigger | `min-h-8 sm:min-h-7` / `w-[84px]` / `px-[calc(2.5-1px)]` | 同值，16px/18px 箭头 | ✅ 0px |
| geometry | separator | `mx-1 h-4` + `w-px` | 1 × 16px，margin 4px | ✅ 0px |
| geometry | sheet tab strip | `border-t bg-muted/40 px-3 py-2` + 40px（溢出 50px）viewport | 同值 | ✅ 0px |
| geometry | tab | `sm:h-8`、`max-w-48`、`rounded-md`、list `p-0.5` | 同值 | ✅ 0px |
| geometry | 状态条 | 28px、居中、`text-xs`、`gap-1` | 同值（`OfficeStatusBar` 外壳 + scoped 覆盖） | ✅ 0px |
| token | 颜色/圆角/阴影 | `--office-*` | 无裸 hex / 裸 px（`lint:style` 0 new） | ✅ |

深色模式踩到的两个 1/255 细节（已固定，避免以后回归）：

1. 上游是 `not-dark:bg-clip-padding`：浅色下 border 压在行背景上，深色下压在 trigger 自身填充上
   → `.xlsx-toolbar__zoom-trigger` 用 `[data-office-theme="dark"]` 切换 `background-clip`；
2. 深色的 1px inset 是 `--alpha(white/6%)`，用 `color-mix(in srgb, …)` 会低 1/255
   → 必须用 `color-mix(in oklab, white 6%, transparent)`。

#### 4.2.3 状态栏：上游没有 XLSX 状态栏（§4.4）

固定的 70fe70c6 里 `XlsxWorkbookSurface` = toolbar + grid + sheet tabs，**没有状态栏**；
`xlsx-editor.tsx` 也没有。只有 file-system block 有一条 `aria-live` 状态条：

```tsx
<div aria-live="polite"
     className="flex h-7 shrink-0 items-center justify-center gap-1 border-t bg-muted/40 px-3 text-xs text-muted-foreground">
```

本仓库的 XLSX 状态栏保留“地址 + 统计 + 说明”的语义，但外观按这条 Extend 状态条做：28px、居中、
`text-xs`、muted；段间分隔用 `::before { content: "· " }` 画，**不进 `textContent`**，
所以 `xlsx-status-range` / `xlsx-status-count` 等 testid 的文本断言不变。
manifest 里这条 surface 的 `notes` 已写明“参考是 Extend 状态条而不是 XLSX 专属组件”。

行为层证据（`verify-xlsx-chrome-interaction.py`、`capture-office-chrome.py`、组件/黑盒套件）
与像素 evidence 并行存在。

### 4.3 浏览器端交互/几何证据（补组件测试盖不到的部分）

最小 host renderer 没有布局引擎、也没有 `document`，所以下面两点只能在真实浏览器里验：

```bash
pnpm --filter demo build
pnpm test:xlsx:chrome
# 等价：python3 scripts/ci/run-python.mjs scripts/ci/verify-xlsx-chrome-interaction.py
```

每个 viewport × theme 断言并通过：

- toolbar 在 390px 换行（`trailing.top > leading.top`），在 768/1440 单行、高度 48px（或 coarse 44px 触控尺寸）；
- toolbar/ribbon 无水平溢出、无垂直裁切（换行时 toolbar 自己长高，ribbon 在 ≤900px 容器内按 132px 上限滚动）；
- `.xlsx-viewer__body` ≥ 360px、grid ≥ 300px，页面无横向溢出；
- 右键菜单（Reka `OfficeDropdownMenu`）：`ArrowDown` 先把焦点放到第一个 `menuitem`
  （`xlsx-context-cut`），再按一次走到 `xlsx-context-copy`；`Escape` 关闭并把焦点还给表格；
- authoring 对话框：点击 Ribbon「Number format…」打开，表面是 `office-dialog__content`，
  表单仍在，`Escape` 关闭后焦点回到 viewer（`closeAuthoring()` 恢复调用者焦点，本轮新增）。

产物：`output/xlsx-chrome/<viewport>/<theme>/` 下的 `-xlsx-viewer.png`、`-xlsx-toolbar.png`、
`-xlsx-context-menu.png`、`-xlsx-authoring.png`，以及 `summary.json`。

## 5. 行为保留清单

| 行为 | 证据 |
|---|---|
| 公式编辑 / 公式栏 | `tests/component/xlsx-cell-editing.test.mjs`、`xlsx-formula-reference-highlight.test.mjs` |
| AutoFilter | `tests/component/xlsx-auto-filter.test.mjs` |
| 冻结窗格 | `tests/component/xlsx-grid-freeze-panes.test.mjs` |
| 图表 / 绘图 / 图表类型菜单 | `xlsx-chart-*.test.mjs`、`xlsx-drawing-*.test.mjs` |
| Worker / 大文件延迟解析 | `xlsx-worker-*.test.mjs`、`xlsx-deferred-load.test.mjs` |
| 导出 | `xlsx-export.test.mjs`、`tests/blackbox/xlsx_export_openpyxl.py` |
| RTL / DOM render | `xlsx-grid-rtl.test.mjs`、`xlsx-dom-render.test.mjs` |
| 触摸 | `tests/blackbox/touch_mobile.py`（选择器 `.xlsx-toolbar__icon-button`、`.xlsx-toolbar__zoom-select`、`.xlsx-ribbon__icon-button` 均保留） |
| ribbon 图标 | `xlsx-office-chrome.test.mjs`：ribbon 源码无 `<svg>`、无 Lucide 名称图标，渲染出 29 个 Hugeicons（`xlsx-ribbon` surface 用同一字形做像素对照） |
| 键盘/焦点/菜单关闭 | 组件：`xlsx-menu-state.test.mjs`、`xlsx-office-chrome.test.mjs`（Escape/Tab/只读项）；浏览器：§4.3（焦点进入菜单、ArrowDown 走项、Escape 还焦） |
| 触摸 44px | `tests/blackbox/touch_mobile.py` → `TOUCH-XLSX-TARGETS-44PX` PASS |
| Sheet tabs 主题/anchor | `xlsx-office-chrome.test.mjs`（token + provider）、`xlsx-sheet-tabs.test.mjs`（键盘/ARIA） |
