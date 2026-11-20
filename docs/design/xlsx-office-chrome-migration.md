# XLSX chrome 迁移到统一设计系统（#160）

> 批次：0.8 / issue [#160](https://github.com/arcships/agentic-office-ui/issues/160)。
> 依赖：[#156](https://github.com/arcships/agentic-office-ui/issues/156) token + `OfficeIcon`、
> [#157](https://github.com/arcships/agentic-office-ui/issues/157) primitive 层、
> [#158](https://github.com/arcships/agentic-office-ui/issues/158) 共享 Office chrome。
> 参考基准：[extend-hq/ui@70fe70c6](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)。

本文记录 XLSX viewer/authoring 的 toolbar、ribbon、菜单、对话框和状态栏如何落到共享层。
菜单与对话框现在**真正运行在 Reka 版 `OfficeDropdownMenu` / `OfficeDialog` 上**（第三轮 0.8 复审），
本篇同时记录为此做的 primitive 扩展、测试迁移与仍然不可达的像素对照项。

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
| XLSX toolbar/ribbon 与 Extend 参考 diff ≤0.1% | ❌ 不可达（参考侧私有依赖） | §4.2 结论表 |
| 现有 XLSX unit/component/blackbox 全绿 | ✅（本轮实测） | unit 275 + component 475（含新增 happy-dom 套件）+ touch blackbox 20/20 + xlsx 导出/Worker 黑盒 |
| 菜单键盘/焦点/关闭不回归 | ✅ | §4.3 浏览器断言 + 组件测试 |
| 移动端 toolbar 换行且 grid 可用高度 | ✅ | 390px：toolbar 84px + `wrapped=true` + grid 360px |

## 1. 结论速览

| 区域 | 迁移后 | 说明 |
|---|---|---|
| `XlsxToolbar` | `OfficeToolbar` + `OfficeToolbarGroup` + `OfficeToolbarSeparator` + `OfficeIconButton` | 48px 外壳、`gap: 4px`、28px 图标按钮；缩放/搜索面板走 token |
| `XlsxRibbon` | 第二个 `OfficeToolbar` 行 + 每组 `OfficeToolbarGroup` + `OfficeButton` / `OfficeIconButton` | 原来的 disabled `<fieldset>` 换成显式 `controlsDisabled` |
| `XlsxStatusBar` | `OfficeStatusBar`（`#start` / `#end` 插槽） | 选区地址、统计、note 的 `data-testid` 全部保留 |
| 空/加载/错误 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` | 仍可被 `#empty` / `#loading` / `#error` 插槽覆盖 |
| 右键/表头/图表/绘图菜单 | `OfficeDropdownMenu`（Reka）：`items` 模型 + 指针锚点定位 | 键盘/typeahead/子菜单/焦点/关闭全部由 primitive 层提供（见 §3） |
| AutoFilter 面板 | 同上（checkbox 行用 `keepOpen` 保持多选） | |
| Authoring 对话框 | `OfficeDialog`（Reka）：overlay + header/title/body + close | 焦点陷阱、Esc、backdrop、关闭还焦由 primitive 层提供（见 §3） |
| Sheet tabs | `OfficeThemeProvider` 包裹 + 全部 inline 样式走 `var(--office-*)` | 10 处 `isDark ? … : …` 三元色值全部移除 |
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

### 4.2 与 Extend 参考的像素对照：可测项与不可达项

**结论**：XLSX toolbar/ribbon 的 `diff ≤ 0.1%` 这一条**无法在本仓库达成**，也不应该用放宽阈值
或测试遮罩来“达成”。原因与 #159 相同：参考侧不可渲染。可达的是 **geometry / structural / token**
三层一致性，已用浏览器断言逐项固定：

| 层 | 指标 | 参考值 | 实测（`verify-xlsx-chrome-interaction.py`） | 状态 |
|---|---|---|---|---|
| geometry | toolbar 高度（≥640px，精细指针） | 48px（`min-h-12`） | 48px | ✅ |
| geometry | toolbar 高度（390px，换行两行） | 换行不裁切 | 84px，`wrapped=true` | ✅ |
| geometry | 图标按钮 | 28px 正方形 | `office-button--size-icon-sm` = 28px | ✅ |
| geometry | cluster 间距 | 4px / 8px | `--office-ui-space-1` / `-2`（gap-sm/md） | ✅ |
| geometry | 分隔线 | 1px × 16px | `--office-border-width` × `--office-ui-space-4` | ✅ |
| geometry | 状态栏 / ribbon 行高 | 紧凑 24–32px | 状态栏 24px；ribbon 每行 28px | ✅ |
| geometry | 网格可用高度（390px） | 不被 chrome 挤掉 | grid = 360px（= `.xlsx-viewer__body` 的 `min-height`） | ✅ |
| structure | 共享组件接线 | toolbar/group/separator/icon-button | `tests/component/xlsx-office-chrome.test.mjs` | ✅ |
| structure | 菜单/对话框表面 | `.office-menu__*` / `.office-dialog__*` | 同上 + 浏览器断言 `office-dialog__content` | ✅ |
| token | 颜色/圆角/阴影 | `--office-*` | 组件里只剩 `var(--office-*, <fallback>)`；无裸 hex | ✅ |
| pixel | 与 Extend `xlsx-viewer.tsx` 截图像素差 | ≤0.1% | **不可测** | ❌ |

不可达的原因（不是“暂时没做”，而是结构性阻断）：

- `apps/v4/components/extend/xlsx-viewer.tsx`（1844 行）与 `xlsx-editor.tsx`（2835 行）
  直接 import 私有包 `@extend-ai/react-xlsx@0.23.0`；该包不在 MIT 仓库内、也不在离线环境里，
  参考侧页面无法构建；
- parity harness（`#155`）因此只登记了自包含的 `pdf-toolbar` 等 surface，没有 `xlsx-*` surface；
  `pnpm ui:parity -- --only xlsx-…` 没有可跑的 scene（强行登记只会得到一个空/伪造的比较）。

替代证据（可重复）：

1. `scripts/ci/verify-xlsx-chrome-interaction.py`：真实 demo `#/xlsx-viewer`，1440/768/390 × light/dark，
   断言上表 geometry + 菜单/对话框键盘与焦点行为，输出 `output/xlsx-chrome/summary.json` 与截图；
2. `scripts/ci/capture-office-chrome.py`（`xlsx_story_flow`）：共享 chrome 画廊里的 toolbar/ribbon/status
   构图与无障碍断言（4 viewport × 2 theme）；
3. `tests/component/xlsx-office-chrome.test.mjs`（6 条）固定 toolbar/ribbon/status bar/状态组件的共享接线；
4. `tests/component/xlsx-menus.test.mjs`（12 条，happy-dom）与
   `tests/component/xlsx-authoring-dialog.test.mjs`（16 条，happy-dom）覆盖 Reka 菜单/对话框的
   testid、命令、键盘/焦点、只读禁用与表单；`tests/component/xlsx-*.test.mjs` 合计 475 条、
   `tests/unit/xlsx-*.test.mjs` 275 条行为回归；
5. `tests/blackbox/touch_mobile.py` 20/20（含 `TOUCH-XLSX-TARGETS-44PX`、长按菜单 → Reka 菜单保持打开、双指缩放）。

建议收口：与 #159 一致，**接受 geometry/structural/token parity + 记录像素 delta**；
像素级判定留给 `#165`（若届时能引入 `@extend-ai/react-xlsx` 或上游提供自包含的 xlsx toolbar 组件，
再补 `xlsx-toolbar` surface，不需要改本次实现）。

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
| 键盘/焦点/菜单关闭 | 组件：`xlsx-menu-state.test.mjs`、`xlsx-office-chrome.test.mjs`（Escape/Tab/只读项）；浏览器：§4.3（焦点进入菜单、ArrowDown 走项、Escape 还焦） |
| 触摸 44px | `tests/blackbox/touch_mobile.py` → `TOUCH-XLSX-TARGETS-44PX` PASS |
| Sheet tabs 主题/anchor | `xlsx-office-chrome.test.mjs`（token + provider）、`xlsx-sheet-tabs.test.mjs`（键盘/ARIA） |
