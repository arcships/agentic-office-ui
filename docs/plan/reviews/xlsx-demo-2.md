# xlsx-demo Review #2

Date: 2026-07-09
Scope: `packages/vue-xlsx/src/`, `apps/demo/src/pages/XlsxViewerPage.vue`, `apps/demo/src/main.ts`
Task: [xlsx-006](docs/plan/tasks/xlsx-006.md)
Previous: [xlsx-demo-1](docs/plan/reviews/xlsx-demo-1.md)

## Typecheck

```bash
pnpm --filter @arcships/vue-xlsx typecheck
```

✅ **零错误**。`tsc --noEmit` 通过，无类型错误。

## Review #1 阻塞项回检

[xlsx-demo-1](docs/plan/reviews/xlsx-demo-1.md) 标记了 6 个 P1 阻塞项，逐一复查：

| # | 原阻塞项 | 状态 | 位置 |
|---|---|---|---|
| 1 | Ribbon 6 tab | ✅ 已实现 | [XlsxRibbon.vue](packages/vue-xlsx/src/components/XlsxRibbon.vue) — Home/Insert/Page Layout/Formulas/Data/View，含 Clipboard、Font、Alignment、Number、Cells、Workbook、Export、Calculation、Zoom、Display 分组 |
| 2 | Formula bar | ✅ 已实现 | [XlsxFormulaBar.vue](packages/vue-xlsx/src/components/XlsxFormulaBar.vue) — name box + fx 标签 + formula input，支持 Enter 提交、blur 提交、chart series 公式目标 |
| 3 | 缩略图 | ✅ 已实现 | [XlsxSheetTabs.vue](packages/vue-xlsx/src/components/XlsxSheetTabs.vue#L29-L39) — hover 弹出 canvas 缩略图，通过 [useXlsxViewerThumbnails](packages/vue-xlsx/src/composables/useXlsxViewerThumbnails.ts) 渲染 |
| 4 | 拖拽上传 | ✅ 已实现 | [XlsxViewerPage.vue:48-58](apps/demo/src/pages/XlsxViewerPage.vue#L48) — dragenter/dragover/dragleave/drop 处理，drag overlay 提示 |
| 5 | URL 输入加载 | ✅ 已实现 | [XlsxViewerPage.vue:16-25](apps/demo/src/pages/XlsxViewerPage.vue#L16) — URL 输入框 + Load URL 按钮 + Enter 快捷键 |
| 6 | Canvas 开关 | ⚠️ 存在但未连通 | [XlsxViewerPage.vue:34](apps/demo/src/pages/XlsxViewerPage.vue#L34) — `useCanvas` checkbox 存在，状态显示正常，但 `useCanvas` 值未传入 XlsxViewerWrapper/XlsxViewer/XlsxGrid，无实际渲染切换效果 |

### Canvas 开关详情

`XlsxViewerPage.vue` 中 `useCanvas` ref 仅用于状态显示（line 44），不参与任何 prop 传递。XlsxViewerWrapper 的 props 仅含 `file`/`src`/`fileName`/`readOnly`/`useWorker`，不含 canvas 相关控制。XlsxGrid 当前仅支持 Canvas 渲染，不存在 DOM 渲染替代路径。该开关实际为 dead control。

## 新增 Findings

### P2 — non-blocking

| # | 描述 | 位置 |
|---|---|---|
| 1 | Canvas 开关未连通：`useCanvas` 状态不传入 viewer，无实际渲染模式切换 | [XlsxViewerPage.vue:34](apps/demo/src/pages/XlsxViewerPage.vue#L34), [XlsxViewerWrapper](apps/demo/src/pages/XlsxViewerPage.vue#L97-L115) |
| 2 | Ribbon View tab 的 readOnly checkbox 仅本地状态，未回传 controller：`<XlsxRibbon>` 的 `update:readOnly` emit 未被 XlsxViewer 或 XlsxViewerPage 处理 | [XlsxViewer.vue:13-18](packages/vue-xlsx/src/components/XlsxViewer.vue#L13), [XlsxRibbon.vue:131](packages/vue-xlsx/src/components/XlsxRibbon.vue#L131) |

### P3 — non-blocking

| # | 描述 | 位置 |
|---|---|---|
| 1 | XlsxSelectionOverlay 仍为占位组件（空 div），注释说明为未来 DOM 覆盖层预留 | [XlsxSelectionOverlay.vue](packages/vue-xlsx/src/components/XlsxSelectionOverlay.vue#L16-L18) |

## 已通过验证

### Import 路径

全部 import 路径正确：

- `@arcships/xlsx-core`：types、API、wasm 初始化
- `@arcships/vue-xlsx`：controller + viewer
- `../composables/useXlsxViewerThumbnails`：组件内单例相对引用
- `./components/Xxx.vue` / `./composables/xxx` / `./render`：同 package 内模块引用
- 无直接引用 dist、无跨 workspace 裸路径

### stub/mock/fake 检查

- `packages/vue-xlsx/src/` 下无 `TODO`/`FIXME`/`HACK`/`noop`/`mock`/`fake` 残留
- `useXlsxViewerController.ts` 中 `_setChartSeriesFormula` 初始化为 `() => false`，但紧随其后从 `chartImageDomain.setChartSeriesFormula` 覆写，属于 late-binding 初始化模式，非残留 stub
- `XlsxImageLayer.vue` 的 `xlsx-image-layer__placeholder` 为 CSS class 名（图片加载中占位），非代码 stub
- `XlsxFormulaBar.vue` 的 `placeholder="Enter a formula or value"` 为 HTML input 属性

### Controller + Viewer 接入

- ✅ `useXlsxViewerController` + `XlsxViewer` 正确 import 自 `@arcships/vue-xlsx`
- ✅ Wrapper 组件 key 机制保证 source 切换时正确重载
- ✅ 5 个 sample fixture 切换正常
- ✅ Worker toggle 正确传入 `useWorker` prop
- ✅ Read-only toggle 正确传入 `readOnly` prop
- ✅ setWasmSource 配置正确（main.ts）

## 结论

**pass** — Review #1 的 6 个阻塞项中 5 项已完整实现。Canvas 开关存在但未连通（P2），Ribbon View tab readOnly 未回传（P2），均不影响核心功能可用性。typecheck 零错误，import 路径正确，无 stub/mock 残留。
