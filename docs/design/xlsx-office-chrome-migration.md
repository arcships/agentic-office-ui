# XLSX chrome 迁移到统一设计系统（#160）

> 批次：0.8 / issue [#160](https://github.com/arcships/agentic-office-ui/issues/160)。
> 依赖：[#156](https://github.com/arcships/agentic-office-ui/issues/156) token + `OfficeIcon`、
> [#157](https://github.com/arcships/agentic-office-ui/issues/157) primitive 层、
> [#158](https://github.com/arcships/agentic-office-ui/issues/158) 共享 Office chrome。
> 参考基准：[extend-hq/ui@70fe70c6](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)。

本文记录 XLSX viewer/authoring 的 toolbar、ribbon、菜单、对话框和状态栏如何落到共享层，
以及**没有**直接使用 Reka 版 `OfficeContextMenu` / `OfficeDropdownMenu` / `OfficeDialog`
的边界与后续动作。

## 1. 结论速览

| 区域 | 迁移后 | 说明 |
|---|---|---|
| `XlsxToolbar` | `OfficeToolbar` + `OfficeToolbarGroup` + `OfficeToolbarSeparator` + `OfficeIconButton` | 48px 外壳、`gap: 4px`、28px 图标按钮；缩放/搜索面板走 token |
| `XlsxRibbon` | 第二个 `OfficeToolbar` 行 + 每组 `OfficeToolbarGroup` + `OfficeButton` / `OfficeIconButton` | 原来的 disabled `<fieldset>` 换成显式 `controlsDisabled` |
| `XlsxStatusBar` | `OfficeStatusBar`（`#start` / `#end` 插槽） | 选区地址、统计、note 的 `data-testid` 全部保留 |
| 空/加载/错误 | `OfficeEmptyState` / `OfficeLoadingState` / `OfficeErrorState` | 仍可被 `#empty` / `#loading` / `#error` 插槽覆盖 |
| 右键/表头/图表/绘图菜单 | 共享 `.office-menu__content` / `.office-menu__item` 表面 + `useOfficeMenuKeyboard` | 定位层保留（见 §3） |
| AutoFilter 面板 | 共享 `.office-popover__content` + `.office-menu__*` 行 | |
| Authoring 对话框 | 共享 `.office-dialog__content` 外壳（header/title/body/footer/close）+ `.office-button` 操作 | 见 §3 |
| 主题 | viewer/各 chrome 组件用 `OfficeThemeProvider` 提供 `data-office-theme` | 原有 `.xlsx-*--dark` 标记保留给黑盒/矩阵，颜色一律走 token |

`XlsxToolbar` / `XlsxRibbon` / `XlsxStatusBar` / `XlsxContextMenu` / `XlsxAuthoringDialog`
的 **public props/emits 一字未改**，`scripts/ci/public-api-contract.json` 的声明指纹因此不变。

## 2. 为什么要 `OfficeThemeProvider`

设计 token 全部声明在 `[data-office-theme]` 作用域下（见 `packages/vue-ui/src/styles/tokens.css`）。
XLSX 的 `isDark` 是 prop，不是全局主题，所以每个装配点自己提供主题：

- `XlsxViewer` 包一层 provider（`theme = isDark ? 'dark' : 'light'`，同时传 `locale` / `messages`）；
- `XlsxToolbar` / `XlsxRibbon` / `XlsxStatusBar` / `XlsxAuthoringDialog` 各自也包一层，
  这样它们作为公开组件被宿主单独挂载时仍然拿到 token；
- provider 宿主是 `display: contents`，不参与布局，flex 高度不会变。

## 3. 已知差异：Reka 覆盖层没有直接接线

`OfficeContextMenu` / `OfficeDropdownMenu` / `OfficeDialog` / `OfficeSheet` / `OfficeSelect`
建在 Reka UI 上，**需要真实 DOM**（`document`、`Element`、pointer/portal/focus scope）。
本仓库的 XLSX 组件测试用的是 `tests/component/vue-test-renderer.mjs` 的最小 host renderer，
在其中挂载这些组件会直接抛错：

```
OfficeDialog  -> TypeError: Cannot read properties of undefined (reading 'display')
OfficeContextMenu -> ReferenceError: Element is not defined
OfficeSelect  -> ReferenceError: Element is not defined
```

（`OfficeIconButton` / `OfficeButton` / `OfficeToolbar` / `OfficeStatusBar` / `Office*State`
只用 Reka 的 `Primitive` / 纯 Vue，最小 renderer 下正常，所以上述 chrome 组件可以真接线。）

XLSX 的菜单还有两个 Reka 菜单覆盖不了的需求：

1. 定位来自原生 `contextmenu` 事件的坐标（包括长按合成事件、图表/绘图对象覆盖层），
   不是来自声明式 trigger；
2. 子菜单（图表类型 / 粘贴选项）与"先关闭再派发 `authoring`"的顺序被组件测试逐条固定。

因此本轮采用**共享表面 + 共享键盘语义**的接法，而不是把菜单整体换成 Reka 组件：

- DOM 类名与 `OfficeDropdownMenu` / `OfficeContextMenu` 渲染出来的完全一致
  （`.office-menu__content`、`.office-menu__viewport`、`.office-menu__item`、
  `.office-menu__separator`、`.office-menu__item-shortcut`、`.office-menu__sub-trigger`），
  视觉/间距/键盘提示由 `styles/ui.css` 提供，本地只保留定位与 hover/focus 高亮；
- 键盘语义继续走 `useOfficeMenuKeyboard` / `resolveOfficeMenuKeyCommand`
  （Escape/Tab 关闭并还焦、方向键/Home/End 循环、表头菜单的 type-ahead 保留）。

对话框同理：`XlsxAuthoringDialog` 复用 `.office-dialog__content`、`.office-dialog__header/_title/_body/_footer/_close`
和 `.office-button` 的视觉契约，但保留自己的 `Teleport`、`Escape`、`pointerdown` 关闭与
`tabindex` 焦点行为，因为它的打开/关闭由 `XlsxViewer` 的 `v-if` 驱动，且测试用
`findByTestId(...).props.onClick()` 直接驱动按钮。

**后续（建议放到 #165 或独立 issue）**：把 XLSX 组件测试从最小 renderer 迁到 happy-dom 后，
这些表面可以无缝换成 Reka 版 `OfficeContextMenu` / `OfficeDialog`——本轮的类名契约就是为这次替换准备的。

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

### 4.2 与 Extend 参考的像素对照

`#155` 的 parity harness 目前只有 `button`、`input`、`select`、`dialog`、`tooltip`、
`file-upload`、`file-thumbnail`、`pdf-toolbar`、`document-sidebar` 等 surface，
**没有 `xlsx-*` surface**。新增的障碍是上游参考实现不可 vendored：

- `apps/v4/components/extend/xlsx-viewer.tsx`（1844 行）/ `xlsx-editor.tsx`（2835 行）
  直接依赖私有包 `@extend-ai/react-xlsx@0.23.0`，该包不随 MIT 仓库分发，
  离线环境无法渲染参考侧；
- 与之相对，`pdf-toolbar` surface 之所以可行，是因为上游 `pdf-toolbar.tsx` 是自包含组件。

因此本轮 **没有** `pnpm ui:parity -- --only xlsx-...` 可跑，替代证据是：

1. §4.1 的共享 chrome 截图矩阵（旧结构基线见 `tests/baseline/visual/structure/xlsx-viewer/**`；
   该基线由 `scripts/ci/capture-visual-matrix.py` 维护，#74 的 240 场景最终基线仍在补）；
2. `tests/component/xlsx-*.test.mjs`（465 条）对 DOM 契约、命令、i18n 与状态的断言；
3. token 级别对齐：XLSX chrome 的每个颜色/圆角/间距都取自 `--office-*`，不再有硬编码
   `#2563eb` / `#e4e4e7` 之类的局部值。

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
| 键盘/焦点/菜单关闭 | `xlsx-menu-state.test.mjs`、`xlsx-chart-type-menu.test.mjs`、`xlsx-structure.test.mjs` |
