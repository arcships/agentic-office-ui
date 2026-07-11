# 上游 XLSX 功能对齐清单与迁移操作方案

> [!WARNING]
> **历史资料。** 本文保存 2026-07-06 的上游分析和最初迁移方案。文中的本地绝对路径、目标文件、行数和“必须复刻”均不是当前产品状态或发布承诺。当前实现与验收请从[文档索引](INDEX.md)进入。

Date: 2026-07-06

## 目的

本文档基于对上游 `@extend-ai/react-xlsx`（commit `f285a1c`）全部 14 个源文件的逐行深读，产出：

1. **功能对齐清单**：上游每个模块实际做什么、怎么实现（不是接口签名），Vue port 必须复刻什么。
2. **迁移分类**：每个文件按"直接复制 / 机械改写 / 局部改写 / 重写"分类。
3. **操作方案**：通过拷贝 + 脚本减少工作量的具体步骤。

上游源码位于 `/Users/eric8810/Code/extend-ui-upstream/react-xlsx/packages/react-xlsx/src/`。

---

## 一、文件迁移分类总表

| 文件 | 行数 | React 依赖 | 迁移策略 | 目标位置 |
|---|---:|---|---|---|
| `wasm.ts` | 82 | 0 | **直接复制** | `xlsx-core/src/wasm.ts` |
| `safe-calculate.ts` | 99 | 0 | **直接复制** | `xlsx-core/src/safe-calculate.ts` |
| `worker-client.ts` | 230 | 0 | **直接复制** | `xlsx-core/src/worker-client.ts` |
| `xlsx-worker.ts` | 835 | 0 | **直接复制** | `xlsx-core/src/xlsx-worker.ts` |
| `colors.ts` | 224 | 0 | **直接复制** | `xlsx-core/src/colors.ts` |
| `images.ts` | 3870 | 0 | **直接复制** | `xlsx-core/src/images.ts` |
| `charts.ts` | 4366 | 0 | **直接复制** | `xlsx-core/src/charts.ts` |
| `types.ts` | 1505 | 29 | **复制后清理** | `xlsx-core/src/types.ts` + `vue-xlsx/src/types.ts` |
| `controller.tsx` | 4987 | 166 | **机械改写** → `composables.ts` | `vue-xlsx/src/composables.ts` |
| `chart-renderer.tsx` | 7174 | 22 | **局部改写** | `vue-xlsx/src/chart-renderer.ts` |
| `surface-regl.tsx` | 1185 | 10 | **局部改写** | `vue-xlsx/src/surface-regl.ts` |
| `XlsxViewer.tsx` | 16615 | 420 | **重写**（Vue 模板/canvas） | `vue-xlsx/src/XlsxViewer.vue` |
| `index.ts` | 76 | 0 | **直接复制后调整导出** | `xlsx-core/src/index.ts` + `vue-xlsx/src/index.ts` |
| `wasm-url.d.ts` | 4 | 0 | **直接复制** | `xlsx-core/src/wasm-url.d.ts` |

**统计**：
- 直接复制：~9700 行（wasm + safe-calculate + worker-client + xlsx-worker + colors + images + charts）
- 复制后清理：~1500 行（types.ts，删 React 组件类型）
- 机械改写：~5000 行（controller → composables）
- 局部改写：~8400 行（chart-renderer + surface-regl）
- 重写：~16600 行（XlsxViewer.tsx → Vue）

---

## 二、逐文件对齐清单

### 2.1 `wasm.ts`（82行）— 直接复制

**上游实现**：WASM 引擎单例懒加载。`getSheetsWasmModule()` 首次调用时 `import("@dukelib/sheets-wasm")` → `mod.default()` 初始化，后续返回缓存 Promise。`setWasmSource()` 把用户配置的源规范化为 worker 可用形式（string/URL→href，BufferSource→**深拷贝 ArrayBuffer**，Response→不可用→禁用 worker）。

**必须复刻**：
- 单例 Promise 模式（全局只 init 一次）
- BufferSource 源深拷贝（transfer 后主线程不失效）
- `canUseConfiguredWasmSourceInWorker()`（Response 源 → false → 禁用 worker）
- worker 源逐消息同步（worker 是独立 realm，每条 load/parseCharts 消息携带 wasmSource）

**迁移操作**：直接 `cp`，零修改。

---

### 2.2 `safe-calculate.ts`（99行）— 直接复制

**上游实现**：封装 `Workbook.calculate()` 的安全调用。Rust 引擎遇到公式引用不存在的 sheet 时会 panic 进 wasm trap，毒化 Workbook 实例。两层防护：
1. **预扫描**：`hasUnresolvedSheetReferences` 用正则遍历所有 sheet 的 formulaCells，收集被引用 sheet 名，与 `sheetNames` 比对，命中则跳过计算。
2. **trap 兜底**：`try/catch` + `options.reparse()` 重新 `fromBytes` 拿新实例。

`tryRecalculate` 不预扫不 reparse，trap 时不刷新状态（避免读毒化实例崩溃）。

**必须复刻**：
- 两层防护机制
- `tryRecalculate` 的激进模式（不预扫不 reparse）
- 正则 `SHEET_REF_REGEX`（带 `g` flag，使用前 `lastIndex = 0`）

**迁移操作**：直接 `cp`，零修改。

---

### 2.3 `worker-client.ts`（230行）— 直接复制

**上游实现**：主线程侧 Worker RPC 客户端。自增 id、pending map、四种消息（load/parseCharts/getCellSnapshot/getRowsBatch）。load/parseCharts 先 `buffer.slice(0)` clone 再 transfer clone。dispose 时 reject 所有 in-flight 为 AbortError。

**必须复刻**：
- 四种消息协议
- ArrayBuffer transfer clone（`buffer.slice(0)` 再 transfer）
- dispose 的 AbortError 语义
- worker error 时 reject `"Worker request failed."`

**迁移操作**：直接 `cp`，零修改。Worker URL 构建方式 `new URL("./xlsx-worker.js", import.meta.url)` 需确保 Vue 构建能解析（Vite 原生支持）。

---

### 2.4 `xlsx-worker.ts`（835行）— 直接复制

**上游实现**：Worker 主体。持有模块级状态（workbook/sheets/charts/tabs），load 后长期驻留。`loadWorkbook` 完整流程：
1. `Workbook.fromBytes(bytes)`
2. 统计 formulaCount，≤1000 才 `safeCalculate`
3. ≥5MB 且公式≤1000 → 跳过完整 XML 结构解析，只用轻量 `parseWorkerSheetLayoutAssets`
4. legacy .xls（魔数 `D0CF11E0`）→ 强制跳过 XML
5. `buildSheetList` 合并 structureAssets + mergeMetadata
6. 图表分两阶段：load 时 quick，后续 parseCharts 精解析

**关键常量**：`FORMULA_COUNT_THRESHOLD=1000`、`FAST_STRUCTURE_PARSE_THRESHOLD_BYTES=5MB`、`DEFAULT_ROW_HEIGHT=24`、`DEFAULT_COL_WIDTH=80`。

`getCellDisplayValue` 优先级：**格式化值 > 缓存公式值 > 计算值**，error 时回退缓存，HTML entity decode。

**必须复刻**：
- 完整 loadWorkbook 流程（阈值判断、快速结构解析、legacy 检测）
- `getCellDisplayValue` 优先级链
- 模块级状态长期驻留
- `canParseXmlInWorker` 检查 DOMParser

**迁移操作**：直接 `cp`，零修改。依赖 `images.ts`/`charts.ts`/`safe-calculate.ts`/`wasm.ts`/`types.ts`，需一起复制。

---

### 2.5 `colors.ts`（224行）— 直接复制

**上游实现**：纯颜色解析。`resolveWorkbookColor` 解析 hex/rgb/argb/theme + tint（HSL 调亮度）。`resolveWorkbookFillStyle` 返回 `{backgroundColor, backgroundImage}`，支持 gradient（linear/path）。

**注意**：indexed color 未实现（已知 gap）。tint 用 HSL 实现，与 images.ts 的 DrawingML 变换（lumMod/lumOff/shade）是**不同体系**，不能混用。

**迁移操作**：直接 `cp`，零修改。

---

### 2.6 `images.ts`（3870行）— 直接复制

**上游实现**：图片/形状/表单控件/工作表元数据的解析+序列化引擎。**纯 zip+XML 操作**，不依赖 wasm Workbook 做解析（只用 `isRowHidden`/`mergedRegions` 等读取辅助）。

**核心功能**：
- `parseWorkbookImageAssets(bytes)`：unzipSync 整个 xlsx → 解析 Content_Types/workbook/theme/styles/drawing XML
- 三种 anchor（two-cell/one-cell/absolute），EMU(9525/px) 与 px 互转
- 列宽换算用 OOXML 公式 + canvas `measureText("0")` 实测 digitWidth
- `updateWorkbookImageAnchor`：通过 imageOriginsById Map 定位，原地改 XML 节点
- `mergeWorkbookImageAssets`：导出时重新注入 drawing XML+rels+media，重建 sheet `<drawing>` 引用 + Content_Types，全程 try/catch 容错
- form controls：三源合并（sheet XML `<control>` + VML + ctrlProp XML）
- shapes：prstGeom 预设几何 + custGeom 自定义 SVG path（M/L/C/Z）；分组递归 + chOff/chExt 缩放
- DrawingML 颜色变换：alpha/lumMod/lumOff/shade/tint（与 colors.ts 不同体系）
- sheet 元数据：hiddenRows/Cols、mergeMetadata、parseSheetState（row/col 宽高覆盖、style id、sparklines、conditional formatting、cached formula values、zoomScale、showGridLines）

**必须复刻**：全部。这是图片/形状/表单控件/工作表元数据的完整实现，零 React 依赖。

**迁移操作**：直接 `cp` 到 `xlsx-core/src/images.ts`。注意依赖 `DOMParser`/`XMLSerializer`（浏览器原生，SSR 需 polyfill）。

---

### 2.7 `charts.ts`（4366行）— 直接复制

**上游实现**：图表数据模型解析 + XML 样式应用 + 锚点/定义更新。**强依赖 wasm Workbook**（读单元格值）。

**核心功能**：
- `loadWorkbookChartAssets`：遍历 sheetCount，classic charts（`worksheet.charts`）+ modern charts（`worksheet.chartsEx`）
- ChartEx layout 映射：boxWhisker/clusteredColumn/funnel/paretoLine/regionMap/sunburst/treemap/waterfall
- histogram：Scott 规则 + nice step 算 bin 宽；pareto 累积百分比
- chart 类型识别：`PRIMARY_CHART_TYPE_LOCAL_NAMES` 正则匹配 plotArea 子节点，switch 映射含 grouping + barDir + scatterStyle
- `parseChartSeriesFormula`/`buildChartSeriesFormula`/`applyChartSeriesFormula`：`SERIES(name,cat,val,order[,bubble])` 解析/构造/应用
- `resolveReferenceValues`：从 workbook 读单元格值（`getCalculatedValueAt`/`getFormattedValueAt`），支持 named range、跨 sheet
- `applyChartStyleFromXml`：DOM 解析 + 正则 fallback；chart rels 读 colorStyle.xml + style.xml
- `updateWorkbookChartAnchor`/`updateWorkbookChartDefinition`：改 drawing XML / chartN.xml
- chartsheet 处理 + buildTabs

**必须复刻**：全部。零 React 依赖。

**迁移操作**：直接 `cp` 到 `xlsx-core/src/charts.ts`。

---

### 2.8 `types.ts`（1505行）— 复制后清理

**上游实现**：80 个类型导出。React 依赖 29 处，主要是 `React.ReactNode`/`React.CSSProperties`/`React.PointerEvent`/`React.Ref` 等组件 prop 类型。

**迁移操作**：
1. 直接复制到 `xlsx-core/src/types.ts`
2. 删除 React 组件 prop 类型（`XlsxViewerProps`、`DefaultToolbarProps`、render slot 类型等，约 200 行）
3. 保留数据模型类型（`XlsxSheetData`/`XlsxCellAddress`/`XlsxCellRange`/`XlsxChart`/`XlsxImage`/`XlsxTable`/`XlsxConditionalFormatRule`/`XlsxDataValidation`/`XlsxSparkline`/`XlsxWorkbookTab` 等，约 1300 行）
4. `React.CSSProperties` → `Record<string, string | number>` 或 Vue 的 `CSSProperties`
5. `React.ReactNode` → `unknown` 或删除
6. 在 `vue-xlsx/src/types.ts` 中单独定义 Vue 组件 props 类型

**可脚本化**：用 sed 删除含 `React.` 的行，手动检查剩余类型完整性。

---

### 2.9 `controller.tsx`（4987行）— 机械改写为 `composables.ts`

**上游实现**：纯逻辑 hook，唯一导出是 `useXlsxViewerController`。React API 使用：`useState`×28、`useCallback`×97、`useRef`×12、`useEffect`×6、`useMemo`×20。**唯一"JSX"是个布尔表达式**（`return (message.includes(...))`），无 DOM 渲染。

**核心数据流**：
```
UI 事件 → controller 方法 → worksheet wasm 方法 → maybeRecalculate → refreshWorkbookState → React state → 重渲染
```

**关键实现逻辑（必须复刻）**：

#### A. 历史模型（双层）
1. **snapshot 历史**（整份 bytes 重建）：resize/merge/unmerge/addSheet/removeSheet/defineNamedRange/image/chart rect 移动、paste 含 merge 时。每次快照 `saveXlsxBytes` + sanitize + merge images。撤销 = `Workbook.fromBytes` 整体重建。
2. **cell-edit 历史**（单格 before/after）：setCellValue/setCellFormula/setCellStyle。`CellMutationState = {formula, style, value}`。
3. **range-edit 历史**（多格 mutations）：pasteText/pasteStructured/fillSelection/sortTable/clearSelectedCells。
4. `applyCellMutationState`：恢复时**优先 formula**（有 formula 走 setFormula，否则走 setCell(value)），style 仅在是对象时写回。
5. `HISTORY_LIMIT=100`，FIFO 淘汰。
6. `isApplyingHistoryRef` 防止 undo/redo 自身再记历史。

#### B. selection 模型
7. `selectCell(cell, {extend})`：extend 时 anchor 不变，selection = normalize(anchor, cell)。
8. `selectRange`：activeCell 落在 end（右下）。
9. `normalizeRange`：所有读写 selection 的入口先 normalize。

#### C. cell edit
10. `coerceUserEnteredValue`：`'` 前缀→纯文本；`true/false`→boolean；纯数字→Number；否则字符串。
11. `setCellValue`：capture before → `worksheet.setCell(a1, coerce(value))` → capture after → recalc + refresh + recordCellEditHistory。
12. `setCellFormula`：空→`setCell(a1,"")`；非空→`setFormula(a1, formula)`。

#### D. clipboard/paste/fill
13. `getClipboardData`：三种格式（text TSV + html table + structured JSON），跳过 mergedSecondary。
14. `pasteText`：`rawValue.startsWith("=")` → `setFormula`；否则 `coerceUserEnteredValue` → `setCell`。
15. `pasteStructuredClipboardData`：有 merge 时走 snapshot 历史；value 直接用 payload.value（不再 coerce）。
16. `fillSelection`：**平铺/重复填充**（`offset % sourceSize`），公式原样复制（不做 R1C1 偏移），样式也复制。逐格 capture+apply。

#### E. sort
17. `sortTable`：capture 每格 `{formula, style, value}`（完整 CellMutationState），sort key 用 `getCalculatedValueAt` + `getFormattedValueAt`（计算值，不是原始公式）。排序：empty 沉底 → number → boolean → localeCompare(numeric:true)。无变化时只 setSortState 不写 cell。历史走 range-edit。

#### F. merge/unmerge/resize
18. merge/unmerge 走 **snapshot 历史**，调 `worksheet.mergeCells`/`unmergeCells`。
19. resize：worker-backed 或 readOnly → `applyReadOnlyResizeOverride`（纯前端，不碰 workbook）；正常 → `worksheet.setColumnWidth`/`setRowHeight` + snapshot 历史。
20. `pxToSheetColumnWidth` / `pxToSheetRowHeight` 换算公式。

#### G. style
21. `setCellStyle`：单格 `worksheet.setCellStyleAt`，走 cell-edit 历史。
22. `setRangeStyle`：范围 `worksheet.setRangeStyle(a1)`（单次 wasm 批量调用），走 range-edit 历史。
23. `cloneCellStyle`：`structuredClone` 优先，回退 JSON。

#### H. formula/recalculate
24. `maybeRecalculateWorkbook`：仅 `shouldAutoCalculate` 为真才 `tryRecalculate`，trap 时关 `shouldAutoCalculate`（一次性降级）。
25. `recalculate`（手动）：成功→`refreshWorkbookState`；trap→跳过 refresh（避免崩溃）。
26. `refreshWorkbookState`：`buildSheetList` + `loadWorkbookChartAssets` + `setRevision+1`。

#### I. sheet 管理
27. `addSheet`：去重命名（`workbook.sheetIndex` 循环加后缀）→ `workbook.addSheet` → buildSheetList 重建。
28. `removeActiveSheet`：`workbook.removeSheet(index)` → filter sheetOrigins/images/shapes → buildSheetList。

#### J. image/chart selection
29. selectChart/selectImage 互斥（与 selectCell 也互斥）。
30. `setChartRect`/`setImageRect`：`rectToImageAnchor`（px→EMU）→ `updateWorkbookChartAnchor`/`updateWorkbookImageAnchor` → snapshot 历史。
31. `moveImageBy`/`resizeImageBy`：算 currentRect → `resizeImageRect(rect, handle, deltaX, deltaY)` → `setImageRect`。

#### K. zoom
32. 纯前端 state，不写 workbook。`clampZoomScale`（10-400）。`zoomScaleOverridesByTabId` 按 tab 存。

#### L. getCellStyle
33. controller 内的 `getCellStyleAt`（wasm）是数据层；viewer 的 `getCellStyle` prop 是渲染层。**完全不同**。

**迁移操作（机械改写）**：

| React API | Vue 等价 | 说明 |
|---|---|---|
| `useState(x)` | `const x = ref(x)` | 28 处 |
| `useCallback(fn, [deps])` | `function fn() { ... }` | 97 处，依赖数组删除 |
| `useRef(x)` | `const xRef = ref(x)` 或局部 `let` | 12 处 |
| `useEffect(fn, [deps])` | `watch(deps, fn)` / `onMounted(fn)` | 6 处 |
| `useMemo(fn, [deps])` | `const x = computed(() => fn())` | 20 处 |
| `React.useCallback` | 普通函数 | 同上 |

**可脚本化程度：高**。核心是正则替换 + 手动调整 effect。详见第四节操作方案。

---

### 2.10 `chart-renderer.tsx`（7174行）— 局部改写

**上游实现**：纯 SVG 渲染（React JSX），仅 3D surface 用 WebGL canvas。支持 17+ 图表类型。React 依赖仅 22 处（`React.memo`×1、`useCallback`×2、JSX return×90）。

**支持的图表类型**：bar/column/line/area/scatter/bubble/radar/pie/doughnut/pie3d/barOfPie/surface/stock/waterfall/funnel/boxwhisker/sunburst/treemap/regionMap/combo/unsupported。

**依赖**：d3-scale/d3-shape/d3-hierarchy/d3-geo/topojson-client/us-atlas/world-atlas。

**迁移操作**：
1. 90 处 `return (<svg>...</svg>)` JSX → Vue render function（`h()`）或保留 SVG 生成逻辑改为字符串模板
2. `React.memo` → `defineComponent` + Vue 内置 memo
3. 2 个 `useCallback` → 普通函数
4. d3 逻辑、颜色解析、chart 数据读取等纯逻辑函数（~6000 行）直接复制
5. surface chart 的 SVG contour fallback 直接复制；WebGL 部分（surface-regl.tsx）单独处理

**可脚本化程度：中**。JSX→render function 可脚本辅助，但 SVG 元素属性映射需手动检查。

---

### 2.11 `surface-regl.tsx`（1185行）— 局部改写

**上游实现**：用 regl (WebGL) 渲染 3D surface chart。`buildSurfaceMesh` 做网格细分、3D 投影（rotY/rotX + 透视）、法线光照、墙体、等高线。contour(2D) 时回退 SVG。

**迁移操作**：
1. 1 个 `useEffect` → `onMounted`/`watch`
2. 1 个 `useRef` → `ref`
3. 1 个 `React.memo` → `defineComponent`
4. WebGL 逻辑（`buildSurfaceMesh`、regl command、shader）直接复制
5. 失败兜底（try/catch 设 `failed=true` 显示 SVG fallback）保留

**可脚本化程度：高**。WebGL 逻辑零 React 依赖。

---

### 2.12 `XlsxViewer.tsx`（16615行）— 重写

**上游实现**：双模式（canvas/DOM）混合渲染。**默认 canvas 模式**（`experimentalCanvas=true`）。自实现虚拟化、4 窗格分层（scroll/top/left/corner）、脏矩形增量重绘、blit 滚动缓冲。

**核心渲染架构**：
- 4 个 body canvas + 5 个 header canvas（按冻结窗格分层）
- 可见范围：前缀和二分查找 ± 480px overscan + `expandMergeAwareWindow` 补合并区
- 增量重绘：签名比对跳过 + blit 滚动缓冲 + 脏矩形剔除
- DPR：`canvas.width = cssWidth * dpr` + `setTransform(dpr,...)`

**cell 值读取**：
- `getCellDisplayValue`：formatted(WASM 已套 numFmt) > cached formula value > calculated value；error 时回退缓存
- **viewer 不解析 numFmt**——格式化在 WASM `getFormattedValueAt` 内部完成

**样式解析链**：
inheritedStyle(col⊕row) → cellStyle → tableStyle → `buildCellStyle`(CSS) → `scaleCssProperties`(zoom) → **getCellStyle override** → `buildCanvasCellStyleCache`

**条件格式**：dataBar（矩形条+渐变）、colorScale（全 cell 背景线性插值）、iconSet（Path2D）。colorScale 与 getCellStyle backgroundColor 互斥。

**图片/图表 overlay**：EMU_PER_PIXEL=9525；3 种 anchor；无自定义 renderImage 时 bake 进 canvas，否则 DOM overlay；图表始终 DOM（SVG）。8 方向 resize handle。

**交互**：pointer 选区拖拽(阈值 4px) + fill handle + 行/列轴选；键盘完整导航（Arrows/Home/End/PageUp-Down/Tab/Enter/Delete/F2/Ctrl+Z-Y）；双击/可打印键编辑；wheel+ctrl/pinch zoom（rAF 节流）；无内置右键菜单。

**缩略图**：独立 canvas，复用纯函数样式/取值逻辑；axis 裁剪到 200×80 / 900×1440；输出 192px 等比缩。

**toolbar**：极简（文件名 + zoom + download + sheet tabs），非 ribbon，无格式化按钮。

**dark mode**：LIGHT/DARK palette；dark 下 cell surface = `resolveDarkModeSurface`（theme bg 暗混合），文字用 `resolveReadableTextColor`（contrastRatio≥4.5）。

**worker/直读双路径**：worker-backed 走 `getRowsBatchAsync`（异步），直读走同步 `worksheet.getRowsBatch`；两者都 ±48 行 overscan。

**迁移操作**：这是唯一需要重写的文件。建议：
1. **canvas 渲染管线**（虚拟化/分层/blit/脏矩形/DPR）：逻辑可参考，但需用 Vue 的生命周期和响应式系统重写。canvas 2D API 调用直接复用。
2. **纯函数**（`buildCellStyle`/`resolveInheritedCellStyle`/`getCellDisplayValue`/条件格式解析/EMU 换算/`expandMergeAwareWindow`/`buildStickyOffsets`/`findIndexForOffsetPrefix` 等）：直接复制到 `xlsx-core` 或 `vue-xlsx` 的 utils。
3. **交互逻辑**：事件处理改用 Vue 的 `@event` 绑定，但算法（选区/键盘导航/编辑/zoom）直接复用。
4. **toolbar**：用 Vue 模板重写，结构对齐上游极简设计。
5. **缩略图**：canvas paint 逻辑直接复用。

---

## 三、功能对齐要点汇总（按领域）

### 3.1 引擎层（wasm + safe-calculate + worker + controller 加载/导出）

| # | 要点 | 来源文件 |
|---|---|---|
| 1 | WASM 单例懒加载 Promise | wasm.ts |
| 2 | BufferSource 源深拷贝（transfer 安全） | wasm.ts |
| 3 | worker 源逐消息同步（独立 realm） | wasm.ts |
| 4 | `canUseConfiguredWasmSourceInWorker`（Response→禁用 worker） | wasm.ts |
| 5 | `safeCalculate` 两层防护（预扫描 + trap 兜底 + reparse） | safe-calculate.ts |
| 6 | `tryRecalculate` 不预扫不 reparse，trap 时不刷新 | safe-calculate.ts |
| 7 | 公式总数阈值 1000（超过不自动计算） | xlsx-worker.ts / controller.tsx |
| 8 | 5MB 快速结构解析（跳过完整 XML，只抓行高） | xlsx-worker.ts |
| 9 | legacy .xls 自动跳过 XML（魔数检测） | xlsx-worker.ts |
| 10 | `getCellDisplayValue` 优先级：formatted > cached > calculated | xlsx-worker.ts |
| 11 | Worker 四消息协议 + ArrayBuffer transfer clone | worker-client.ts |
| 12 | Worker 模块级状态长期驻留 | xlsx-worker.ts |
| 13 | preflight 不解压（直接解析 zip central directory） | controller.tsx |
| 14 | 三阈值：单 worksheet XML >200MB / sharedStrings >50MB / 合计 >256MB | controller.tsx |
| 15 | 文件大小上限默认 25MB | controller.tsx |
| 16 | 延迟加载（deferLoadingAboveBytes 阈值） | controller.tsx |
| 17 | worker 仅用于只读 | controller.tsx |
| 18 | worker 错误回退主线程（DOMParser/XMLSerializer 未定义等） | controller.tsx |
| 19 | 图表 hydration：先 quickAssets → 1500ms 超时 → 精解析 | controller.tsx |
| 20 | `saveXlsxBytes` + sanitize（修 styles.xml 双重转义）+ merge images | controller.tsx |
| 21 | `saveCsvString` 只导当前 sheet | controller.tsx |
| 22 | 历史快照 = `saveXlsxBytes` + sanitize + merge；撤销 = `fromBytes` 整体重建 | controller.tsx |

### 3.2 编辑能力（controller）

| # | 要点 | 来源 |
|---|---|---|
| 23 | 双层历史：snapshot（结构变更）/ cell-edit（单格）/ range-edit（多格） | controller.tsx |
| 24 | `applyCellMutationState` 优先 formula（有 formula 走 setFormula，否则 setCell(value)） | controller.tsx |
| 25 | `HISTORY_LIMIT=100` FIFO | controller.tsx |
| 26 | `isApplyingHistoryRef` 防止 undo/redo 自身记历史 | controller.tsx |
| 27 | `coerceUserEnteredValue`：`'`前缀→文本；true/false→boolean；纯数字→Number | controller.tsx |
| 28 | paste 区分 formula/value（`startsWith("=")` → setFormula） | controller.tsx |
| 29 | pasteStructured value 不再 coerce（copy 时存的就是 display value） | controller.tsx |
| 30 | fillSelection 平铺/重复填充，公式不做 R1C1 偏移，样式也复制 | controller.tsx |
| 31 | sortTable capture 完整 CellMutationState，sort key 用计算值 | controller.tsx |
| 32 | sort 无变化时只 setSortState 不写 cell | controller.tsx |
| 33 | merge/unmerge/resize 走 snapshot 历史 | controller.tsx |
| 34 | resize worker/readOnly 走 applyReadOnlyResizeOverride（纯前端） | controller.tsx |
| 35 | setRangeStyle 单次 wasm 批量调用 `worksheet.setRangeStyle(a1)` | controller.tsx |
| 36 | maybeRecalculate trap 时一次性降级 shouldAutoCalculate | controller.tsx |
| 37 | image/chart rect 移动走 snapshot 历史，px→EMU 换算 | controller.tsx |
| 38 | zoom 纯前端 state，不写 workbook，按 tab 存 override | controller.tsx |

### 3.3 渲染层（XlsxViewer）

| # | 要点 | 来源 |
|---|---|---|
| 39 | 双模式：默认 canvas，DOM+virtualizer 兜底 | XlsxViewer.tsx |
| 40 | 4 窗格分层（scroll/top/left/corner）+ 5 header canvas | XlsxViewer.tsx |
| 41 | 自实现虚拟化：前缀和二分查找 ± 480px overscan | XlsxViewer.tsx |
| 42 | `expandMergeAwareWindow` 补合并区 | XlsxViewer.tsx |
| 43 | 增量重绘：签名比对 + blit 滚动 + 脏矩形剔除 | XlsxViewer.tsx |
| 44 | DPR：`canvas.width = cssWidth * dpr` + `setTransform` | XlsxViewer.tsx |
| 45 | viewer 不解析 numFmt（WASM `getFormattedValueAt` 内部完成） | XlsxViewer.tsx |
| 46 | 样式解析链：inherited→cell→table→buildCellStyle→scale→getCellStyle override→canvas cache | XlsxViewer.tsx |
| 47 | 13 种 Excel border→CSS 映射 | XlsxViewer.tsx |
| 48 | gridline 用 boxShadow(DOM)/path stroke(canvas)，有 border 的边移除 gridline | XlsxViewer.tsx |
| 49 | mergedSecondary 跳过绘制；锚定 cell 用跨格 displayRect | XlsxViewer.tsx |
| 50 | 条件格式：dataBar(矩形+渐变)/colorScale(全背景插值)/iconSet(Path2D) | XlsxViewer.tsx |
| 51 | colorScale 与 getCellStyle backgroundColor 互斥 | XlsxViewer.tsx |
| 52 | 图片无 renderImage 时 bake 进 canvas，否则 DOM overlay | XlsxViewer.tsx |
| 53 | 图表始终 DOM（SVG） | XlsxViewer.tsx |
| 54 | 8 方向 resize handle | XlsxViewer.tsx |
| 55 | 键盘完整导航（Arrows/Home/End/PageUp-Down/Tab/Enter/Delete/F2/Ctrl+Z-Y） | XlsxViewer.tsx |
| 56 | wheel+ctrl/pinch zoom（rAF 节流 + 48ms idle commit） | XlsxViewer.tsx |
| 57 | 缩略图独立 canvas，复用纯函数，192px 等比缩 | XlsxViewer.tsx |
| 58 | toolbar 极简（文件名+zoom+download+sheet tabs） | XlsxViewer.tsx |
| 59 | dark mode：resolveDarkModeSurface + resolveReadableTextColor(≥4.5) | XlsxViewer.tsx |
| 60 | worker/直读双路径，±48 行 overscan | XlsxViewer.tsx |
| 61 | getCellStyle prop：最后一层样式覆盖，DOM/canvas 共享 | XlsxViewer.tsx |
| 62 | 网格动态增长（200/50 起步，阈值触发扩展） | XlsxViewer.tsx |

### 3.4 图表/图片/形状

| # | 要点 | 来源 |
|---|---|---|
| 63 | 图片 zip 全量解析，三种 anchor，EMU(9525/px) 互转 | images.ts |
| 64 | 列宽 OOXML 公式 + canvas measureText 实测 digitWidth | images.ts |
| 65 | `mergeWorkbookImageAssets` 导出合并（重新注入 drawing+rels+media） | images.ts |
| 66 | form controls 三源合并（sheet XML + VML + ctrlProp） | images.ts |
| 67 | shapes 预设几何 + 自定义 SVG path；分组递归 + chOff/chExt 缩放 | images.ts |
| 68 | DrawingML 颜色变换（alpha/lumMod/lumOff/shade/tint）与 colors.ts 不同体系 | images.ts |
| 69 | 图表双来源：classic charts + modern chartsEx | charts.ts |
| 70 | SERIES 公式 parse/build/apply 三件套 | charts.ts |
| 71 | chart 数据强依赖 wasm `getCalculatedValueAt`/`getFormattedValueAt` | charts.ts |
| 72 | chart XML 样式：DOM 解析 + 正则 fallback | charts.ts |
| 73 | histogram Scott 规则 + nice step；pareto 累积百分比 | charts.ts |
| 74 | chart-renderer 纯 SVG，17+ 类型，d3 全家桶 | chart-renderer.tsx |
| 75 | 地图 chart：内置 us-atlas + world-atlas topojson | chart-renderer.tsx |
| 76 | 3D surface：regl WebGL，自定义旋转+透视+法线光照+墙体 | surface-regl.tsx |
| 77 | contour(2D) 走 SVG fallback，不用 WebGL | surface-regl.tsx |
| 78 | colors.ts indexed color 未实现（已知 gap） | colors.ts |

---

## 四、迁移操作方案（拷贝 + 脚本）

### Phase 0：准备

```bash
# 上游源码目录
UPSTREAM=/Users/eric8810/Code/extend-ui-upstream/react-xlsx/packages/react-xlsx/src
# 本地目标
CORE=packages/xlsx-core/src
VUE=packages/vue-xlsx/src
```

### Phase 1：直接复制零 React 依赖文件（~9700 行）

```bash
# 以下文件零 React 依赖，直接复制
cp $UPSTREAM/wasm.ts             $CORE/wasm.ts
cp $UPSTREAM/safe-calculate.ts   $CORE/safe-calculate.ts
cp $UPSTREAM/worker-client.ts    $CORE/worker-client.ts
cp $UPSTREAM/xlsx-worker.ts      $CORE/xlsx-worker.ts
cp $UPSTREAM/colors.ts           $CORE/colors.ts
cp $UPSTREAM/images.ts           $CORE/images.ts
cp $UPSTREAM/charts.ts           $CORE/charts.ts
cp $UPSTREAM/wasm-url.d.ts       $CORE/wasm-url.d.ts
cp $UPSTREAM/xlsx-worker.js      $CORE/xlsx-worker.js   # 如果存在预编译版本
```

**验证**：`pnpm --filter @arcships/xlsx-core typecheck`（修 import 路径后应直接通过）。

**需要的 import 路径调整**：
- `xlsx-worker.ts` 依赖 `./images`、`./charts`、`./safe-calculate`、`./wasm`、`./types` — 已在同一目录
- `images.ts` 依赖 `./colors`、`./types` — 已在同一目录
- `charts.ts` 依赖 `./images`、`./types` — 已在同一目录

### Phase 2：types.ts 复制后清理（~1500 行）

```bash
cp $UPSTREAM/types.ts $CORE/types.ts
```

然后脚本清理 React 类型：

```bash
# 删除 React import
sed -i '' 's/^import type \* as React from "react";//' $CORE/types.ts

# React.CSSProperties → Record<string, string | number>
sed -i '' 's/React\.CSSProperties/Record<string, string | number>/g' $CORE/types.ts

# React.ReactNode → unknown
sed -i '' 's/React\.ReactNode/unknown/g' $CORE/types.ts

# React.PointerEvent<HTMLElement> → Record<string, unknown>
sed -i '' 's/React\.PointerEvent<[^>]*>/Record<string, unknown>/g' $CORE/types.ts

# React.Ref<HTMLDivElement> → unknown
sed -i '' 's/React\.Ref<[^>]*>/unknown/g' $CORE/types.ts

# React.ButtonHTMLAttributes → Record<string, unknown>
sed -i '' 's/React\.ButtonHTMLAttributes<[^>]*>/Record<string, unknown>/g' $CORE/types.ts

# React.HTMLAttributes → Record<string, unknown>
sed -i '' 's/React\.HTMLAttributes<[^>]*>/Record<string, unknown>/g' $CORE/types.ts
```

手动检查：删除纯 React 组件 prop 接口（`XlsxViewerProps`、`DefaultToolbarProps`、render slot 类型等），保留数据模型类型。

### Phase 3：controller.tsx 机械改写为 composables.ts（~5000 行）

这是工作量最大的改写。分两步：**脚本预处理 + 手动调整**。

#### 步骤 3a：脚本预处理

```bash
cp $UPSTREAM/controller.tsx $VUE/composables.ts

# 1. 替换 import
sed -i '' 's/import \* as React from "react";/import { ref, computed, watch, onMounted, onUnmounted, shallowRef } from "vue"/' $VUE/composables.ts

# 2. useState → ref（需要手动检查初始值类型）
# 示例模式：const [x, setX] = React.useState(init) → const x = ref(init); const setX = (v) => { x.value = v }
# 这个正则较复杂，建议用 Node 脚本处理

# 3. useCallback → 普通函数
# React.useCallback((args) => { body }, [deps]) → function name(args) { body }
# 97 处，建议用 Node 脚本

# 4. useRef → ref 或 let
# const xRef = React.useRef(init) → const xRef = ref(init) 或 let xRef = init

# 5. useMemo → computed
# const x = React.useMemo(() => expr, [deps]) → const x = computed(() => expr)

# 6. useEffect → watch/onMounted
# React.useEffect(() => { ... }, [deps]) → watch([deps], () => { ... }) 或 onMounted(() => { ... })
```

#### 步骤 3b：建议的 Node 预处理脚本

```js
// scripts/migrate-controller-to-composables.mjs
import { readFileSync, writeFileSync } from 'node:fs'

const file = 'packages/vue-xlsx/src/composables.ts'
let s = readFileSync(file, 'utf-8')

// 1. import 替换
s = s.replace(/import \* as React from "react";/, 'import { ref, computed, watch, onMounted, onUnmounted, shallowRef } from "vue"')

// 2. useState: const [x, setX] = React.useState(init) → const x = ref(init)
//    需要保留 setX 作为赋值函数
s = s.replace(
  /const \[(\w+), set(\w+)\] = React\.useState\(([^)]*)\)/g,
  'const $1 = ref($3)'
)
// setX 调用 → x.value = ...
// 这部分需要更复杂的 AST 变换，建议用 jscodeshift 或手动

// 3. useCallback → 普通函数（保留函数名）
s = s.replace(
  /const (\w+) = React\.useCallback\(\(([^)]*)\) => \{/g,
  'function $1($2) {'
)
// 对应的 }, [deps]) → }
s = s.replace(/\}, \[[^\]]*\]\)/g, '}')

// 4. useRef → ref
s = s.replace(
  /const (\w+) = React\.useRef\(([^)]*)\)/g,
  'const $1 = ref($2)'
)

// 5. useMemo → computed
s = s.replace(
  /const (\w+) = React\.useMemo\(\(\) => (.*?), \[[^\]]*\]\)/g,
  'const $1 = computed(() => $2)'
)

// 6. useEffect → watch (简单情况) / onMounted (空依赖)
// 需要手动判断依赖数组
s = s.replace(
  /React\.useEffect\(\(\) => \{/g,
  '/* TODO: convert to watch/onMounted */ (() => {'
)

writeFileSync(file, s)
```

#### 步骤 3c：手动调整

脚本处理后需要手动处理的点：
- `setState(x)` → `state.value = x`（所有 setX 调用）
- `useEffect` 的依赖数组 → `watch` 的监听源
- `useEffect` 空依赖 → `onMounted`
- `useEffect` 返回 cleanup → `onUnmounted`
- `useRef` 的 `.current` → `.value`（如果是 ref）或保持不变（如果是局部变量）
- `useCallback` 的依赖数组语义：Vue 闭包不需要手动声明依赖，但需确保闭包捕获的是 ref 而非解构值
- 返回值从 React state 对象改为 getter/setter 或直接暴露 ref

#### 步骤 3d：导出函数签名

```ts
// 上游：export function useXlsxViewerController(options): XlsxViewerController
// Vue：export function useXlsxViewerController(options): XlsxViewerController
// 签名不变，内部实现从 React hooks 改为 Vue reactivity
```

### Phase 4：chart-renderer.tsx 局部改写（~7174 行）

```bash
cp $UPSTREAM/chart-renderer.tsx $VUE/chart-renderer.ts
```

改写点：
1. `import * as React from "react"` → `import { defineComponent, h, memo } from "vue"`
2. 90 处 `return (<svg>...</svg>)` → `return h('svg', { ...attrs }, [children])` 或保留 SVG 字符串生成
3. `React.memo(Component)` → `memo(defineComponent({ ... }))`
4. 2 个 `useCallback` → 普通函数
5. d3 逻辑、颜色解析、chart 数据读取等纯函数（~6000 行）零修改

**可脚本化**：JSX→`h()` 可用 Vue compiler 的 JSX transform 辅助，但 SVG 属性名差异需手动检查。

### Phase 5：surface-regl.tsx 局部改写（~1185 行）

```bash
cp $UPSTREAM/surface-regl.tsx $VUE/surface-regl.ts
```

改写点：
1. `useEffect` → `onMounted`/`watch`
2. `useRef` → `ref`
3. `React.memo` → `defineComponent`
4. WebGL 逻辑直接复制

### Phase 6：XlsxViewer.tsx 重写（~16615 行）

这是唯一需要完全重写的文件。建议策略：

1. **提取纯函数到 xlsx-core**：
   - `buildCellStyle`、`resolveInheritedCellStyle`、`mergeResolvedCellStyle`
   - `getCellDisplayValue`、`getCellNumericValue`、`getCellBooleanValue`
   - 条件格式解析（`resolveConditionalDataBar/ColorScale/IconForCell`）
   - `expandMergeAwareWindow`、`buildStickyOffsets`、`findIndexForOffsetPrefix`
   - EMU 换算、`resolveAnchoredRect`
   - 缩略图 paint 逻辑
   - 键盘导航算法、选区算法

2. **Vue 组件重写**：
   - `<template>`：viewport div + canvas 元素（4 body + 5 header）+ overlay div（图片/图表/编辑 input）+ toolbar + sheet tabs
   - `<script setup>`：调用 `useXlsxViewerController`，canvas paint 用 `requestAnimationFrame` + 依赖监听
   - canvas 2D API 调用直接复用上游逻辑
   - 事件处理用 Vue `@event` 绑定

3. **分步实现优先级**：
   - 先实现 canvas 模式（上游默认）
   - 再实现 DOM 模式（兜底）
   - 最后实现交互（选区/键盘/编辑/zoom）
   - 缩略图和 toolbar 最后

### Phase 7：index.ts 调整

```bash
# xlsx-core/src/index.ts — 导出引擎层
cp $UPSTREAM/index.ts $CORE/index.ts
# 调整导出：删 React 组件导出，保留 useXlsxViewerController/initWasm/setWasmSource 等

# vue-xlsx/src/index.ts — 导出 Vue 组件 + composables
# 手动编写，导出 XlsxViewer.vue + useXlsxViewerController + useXlsxViewerThumbnails
```

### Phase 8：依赖安装

```bash
# 上游依赖
pnpm --filter @arcships/xlsx-core add @dukelib/sheets-wasm fflate
pnpm --filter @arcships/vue-xlsx add @tanstack/react-virtual d3-geo d3-hierarchy d3-scale d3-shape regl topojson-client us-atlas world-atlas
pnpm --filter @arcships/vue-xlsx add -D @types/d3-geo @types/d3-hierarchy @types/d3-scale @types/d3-shape @types/geojson @types/topojson-client
```

注意：`@tanstack/react-virtual` 是 React 专用，Vue 等价是 `@tanstack/vue-virtual`。但如果只实现 canvas 模式（不使用 DOM virtualizer），可以不装。

### Phase 9：License 声明

上游 `@extend-ai/react-xlsx` 是 MIT license。复制源码需在 `packages/xlsx-core/LICENSE` 和 `packages/vue-xlsx/LICENSE` 中保留版权声明。

---

## 五、验证路径

每个 Phase 完成后的验证：

| Phase | 验证命令 |
|---|---|
| 1 | `pnpm --filter @arcships/xlsx-core typecheck` |
| 2 | `pnpm --filter @arcships/xlsx-core typecheck` |
| 3 | `pnpm --filter @arcships/vue-xlsx typecheck` + 手动检查 controller 方法行为 |
| 4 | `pnpm --filter @arcships/vue-xlsx typecheck` + 图表渲染验证 |
| 5 | `pnpm --filter @arcships/vue-xlsx typecheck` + surface chart 验证 |
| 6 | 浏览器验证 `http://localhost:5000/#/xlsx-viewer` |
| 7 | `pnpm typecheck && pnpm build` |
| 8 | `pnpm install` 无错误 |
| 9 | `git diff --check` |

最终验证：用 `apps/demo/public/samples/` 下的 fixture（financial-model.xlsx、charts-images.xlsx 等）在浏览器中测试：
- 加载/显示/选区/编辑/公式重算
- 导出 XLSX 用 openpyxl 验证
- 冻结窗格/合并/条件格式/验证/sparkline 显示
- 图表/图片显示和拖拽
- 大文件 worker 路径

---

## 六、工作量估算

| Phase | 文件 | 行数 | 策略 | 预估工作量 |
|---|---|---:|---|---|
| 1 | 7 个 .ts 文件 | ~9700 | 直接复制 | 极小（cp + import 路径） |
| 2 | types.ts | ~1500 | 复制+脚本清理 | 小（sed 脚本） |
| 3 | controller→composables | ~5000 | 脚本+手动 | 中大（脚本预处理 + effect 手动调整） |
| 4 | chart-renderer | ~7174 | 局部改写 | 中（JSX→h() + 手动检查） |
| 5 | surface-regl | ~1185 | 局部改写 | 小 |
| 6 | XlsxViewer→Vue | ~16615 | 重写 | 大（最大工作量） |
| 7 | index.ts | ~150 | 调整 | 极小 |
| 8 | 依赖 | - | 安装 | 极小 |
| 9 | License | - | 声明 | 极小 |

**总计**：~41000 行源码，其中 ~17000 行直接复制/脚本处理，~8400 行局部改写，~16600 行重写（XlsxViewer）。

---

## 七、与当前手写实现的对比

当前 `vue-xlsx` 的手写实现 vs 上游对齐后的预期：

| 领域 | 当前手写 | 上游对齐后 |
|---|---|---|
| workbook engine | `workbook = null`，手写 cellText 字典 | 真实 `@dukelib/sheets-wasm` Workbook 实例 |
| 公式计算 | 手写 5 个函数（SUM/AVG/MIN/MAX/COUNT），无依赖图 | Rust 引擎完整计算 |
| 导出 | 最小 OOXML writer（~189行） | `workbook.saveXlsxBytes()` + sanitize + merge images |
| 样式 | 基础 cellStyles patch | 完整 styleById/themePalette/namedCellStyle + 13 种 border + numFmt |
| 图片/图表 | `images: [], charts: []` | 完整 zip+XML 解析 + anchor + render |
| 条件格式 | 基础 parser + dataBar hint | 完整 dataBar/colorScale/iconSet 渲染 |
| 虚拟化 | `MAX_VISIBLE_COLS=100, MAX_VISIBLE_ROWS=500` 硬截断 | 前缀和二分查找 + 480px overscan + blit + 脏矩形 |
| worker | 无 | 完整四消息协议 + transfer + 回退 |
| 历史 | snapshot 字典 diff | 双层（snapshot bytes + cell/range-edit mutations） |
| clipboard | TSV only | 三格式（text + html + structured JSON） |
| sort | localeCompare 字符串 | 计算值 + 类型感知（number/boolean/text/empty） |
| 缩略图 | 基础 canvas grid | 完整样式 + 图片 + 192px 等比缩 |

---

## 八、上游 XLSX Playground 对齐

### 8.1 Fixture

只有一个 fixture 文件：`public/examples/welcome.xlsx`（417KB）。通过常量 `PLAYGROUND_SAMPLE_URL = "/examples/welcome.xlsx"` 加载。**没有 fixture dropdown**，只有一个"Sample"按钮硬编码加载。默认不加载任何文件，显示空状态。

### 8.2 功能清单

| 功能 | 行号 | 说明 |
|---|---|---|
| 本地文件上传 | 1297-1303 | 隐藏 `<input type="file">`，accept `.xls/.xlsx` |
| 拖拽打开文件 | 1244-1293 | 全屏 drop overlay，dragDepth 计数 |
| URL 加载远程文件 | 1219-1233 | Input + Load 按钮，Enter 提交 |
| 加载内置样例 | 1235-1238 | Sample 按钮加载 welcome.xlsx |
| Worker 模式开关 | 1126, 466-469 | title bar + View ribbon 双入口 |
| 只读模式开关 | 1125, 956-959 | `isReadOnly` |
| 5MB 自动只读阈值 | 39, 1151 | `AUTO_READ_ONLY_THRESHOLD_BYTES = 5MB` |
| 只读下允许调行列宽 | 1127, 969-971 | `allowResizeInReadOnly` |
| 实验性 Canvas 渲染器 | 1124, 947-951 | `experimentalCanvas`（默认 true） |
| 文档暗色模式 | 1123, 824-830 | `isDocumentDark` 独立于 shell 主题 |
| 自定义单元格高亮 | 1131-1142 | `getCellStyle` 回调，奇数行加底色 |
| 单元格样式编辑 | 377-397, 512-727 | font/alignment/number-format/border/fill |
| 撤销/重做 | 337/331, 503-508 | `undo`/`redo` |
| 合并/拆分 | 328/338, 732-737 | `mergeSelection`/`unmergeSelection` |
| 公式编辑栏 | 996-1028 | Name box + fx + 公式 Input，支持图表 SERIES |
| 命名区域定义 | 431-439, 856-872 | `defineNamedRange` |
| 公式重算 | 318, 850-855 | `recalculate`（Formulas + Data 两个 tab） |
| 工作表增删 | 324/329, 782-790 | `addSheet`/`removeActiveSheet` |
| 工作表切换+缩略图 | 1033-1074 | hover Tooltip 显示 canvas 缩略图 |
| 缩放控制 | 340, 900-924 | zoomIn/out/reset，预设 50-200% |
| 加载后自动重置缩放 | 1170-1193 | sourceKey 变化时 setZoomScale(100) |
| 导出 | 470-473, 831-844 | download（源文件）/exportXlsx/exportCsv |
| 表格表头排序菜单 | 1351-1372 | 自定义 DropdownMenu（A→Z / Z→A） |
| Shell 主题定制面板 | 475 | Sheet 抽屉：主题/radius/icon/font/appearance |
| Worker 脚本加载调试 | 233-259 | Performance API 计数 worker 加载次数 |
| 运行时状态指示 | 347-371 | 圆点+Tooltip，区分 Worker/Main/Deferred/Loading |

### 8.3 Ribbon 按钮完整列表（渲染顺序）

**6 个 Ribbon Tab**：Home · Insert · Page Layout · Formulas · Data · View

**Home**：Undo/Redo → Font family(6种)/Font size(11种) → Bold/Italic/Underline/Strikethrough/Font color/Fill color → 左中右/上中下/Wrap/Rotate → Number format(6种)/$/%/000 → Good/Neutral/Bad/Heading → Merge/Unmerge/Borders

**Insert**：Add sheet/Delete → Open/Sample → URL Input/Load

**Page Layout**：Document dark/Theme → Source/XLSX/CSV

**Formulas**：Recalc → Named range Input/Define

**Data**：排序提示文字 → Recalc → Clear

**View**：Zoom out/Select/Zoom in/Reset → Previous/Next/Sheet Select → Canvas Switch/Highlight Switch/Read only Switch → Use worker Switch/Resize read-only Switch/Badges

### 8.4 组件结构

使用 **shadcn/ui**（`components.json` style=`base-mira`，iconLibrary=`hugeicons`）。`src/components/ui/` 下 60+ 个标准 shadcn 组件。

布局：
```
<div h-[100dvh]>
  <input type=file hidden>
  <div max-w-[1800px] mx-auto flex-col>
    <div rounded-2xl border shadow>
      <XlsxViewerProvider controller isDark>
        <WorkbookToolbar />     // title bar + ribbon + formula bar
        <div flex-1 p-2>
          <XlsxViewer ... />     // showDefaultToolbar={false}（playground 自带 ribbon）
        </div>
        <SheetTabs />
      </XlsxViewerProvider>
    </div>
  </div>
</div>
```

**入口 main.tsx**：先 `setWasmSource("/duke_sheets_wasm_bg.wasm")`，再 `ThemeProvider → PlaygroundCustomizerProvider → App`。

### 8.5 controller 使用方式

`useXlsxViewerController`（App.tsx:1144-1169）根据 source 类型传不同 options：

```ts
// 公共字段
{ allowResizeInReadOnly, readOnly, readOnlyAboveBytes: 5MB, useWorker }

// file 类型：{ file: ArrayBuffer, fileName, ...公共字段 }
// url 类型：{ src: string, fileName?, ...公共字段 }
// null：{ ...公共字段 }（仅配置，不加载）
```

`<XlsxViewer>` props：`className`、`emptyState`、`fileTooLargeState`、`getCellStyle`、`height="100%"`、`allowResizeInReadOnly`、`isDark`、`loadingState`、`experimentalCanvas`、`readOnly`、`renderTableHeaderMenu`、`rounded`、`showDefaultToolbar={false}`。

### 8.6 本地 demo 对齐要点

- 本地 demo 的 `apps/demo/src/pages/XlsxViewerPage.vue` 应复刻上述 ribbon 结构和功能
- 本地已有多个 fixture（financial-model.xlsx、charts-images.xlsx 等），上游只有一个 welcome.xlsx——本地 fixture 更丰富，保持现有即可
- 上游 `showDefaultToolbar={false}`，playground 自己实现 ribbon；本地需要决定是 viewer 内置 toolbar 还是外部 ribbon
- 上游入口先 `setWasmSource`，本地也必须对齐

---

## 九、构建配置对齐

### 9.1 上游构建工具

上游用 **tsup** 构建包，**Vite** 构建 playground。

#### XLSX tsup.config.ts 关键点

```ts
// 两个 entry：主包 + worker
entry: ["src/index.ts"],           // 主包
entry: ["src/xlsx-worker.ts"],     // worker 独立打包
external: ["react", "react-dom"],  // React 外部化
noExternal: ["us-atlas", "world-atlas"],  // 地图数据内联
skipNodeModulesBundle: true,       // 不打包 node_modules
// 构建后复制 wasm 二进制到 dist/
```

`scripts/copy-duke-wasm.mjs`：从 `@dukelib/sheets-wasm` 解析 wasm 路径，复制到 `dist/duke_sheets_wasm_bg.wasm`。

#### XLSX playground vite.config.ts 关键点

```ts
worker: { format: "es" },    // Worker 用 ES module 格式
plugins: [react(), tsconfigPaths(), tailwindcss()],
// dev 启动前先 copy-duke-wasm.mjs
```

#### DOCX tsup.config.ts 关键点

```ts
entry: ["src/index.tsx", "src/docx-import-worker.ts"],  // 主包 + worker
external: ["react", "react-dom"],
noExternal: bundledWorkspacePackages,  // 7 个 workspace 包内联
shims: true,                            // CJS 构建 shim import.meta.url
dts: { resolve: true },                 // DTS 解析 workspace 包
onSuccess: "cp ../wasm/src/docx_wasm_bg.wasm dist/docx_wasm_bg.wasm"  // 复制 wasm
```

#### DOCX playground vite.config.ts 关键点

```ts
// 把 8 个 workspace 包 alias 到源码（dev 模式直接读 src，不走 dist）
resolve: {
  dedupe: ["react", "react-dom"],
  alias: {
    "@extend-ai/react-docx": ".../packages/react-viewer/src/index.tsx",
    "@extend-ai/react-docx-doc-model": ".../packages/doc-model/src/index.ts",
    // ... 其余 6 个包同理
  }
},
optimizeDeps: { exclude: Object.keys(workspaceSourceAliases) }  // 不预构建
```

### 9.2 本地 Vite 构建配置适配

本地 demo 用 Vite，包用 tsup。迁移后需要适配：

#### xlsx-core / docx-core 的 tsup.config

```ts
// 需要增加 worker entry 和 wasm 复制
import { defineConfig } from "tsup"
import { copyFileSync, mkdirSync } from "node:fs"
import { createRequire } from "node:module"
import { dirname, join } from "node:path"

export default defineConfig([
  {
    entry: ["src/index.ts", "src/xlsx-worker.ts"],  // 主包 + worker
    format: ["esm"],
    dts: true,
    clean: true,
    external: ["@dukelib/sheets-wasm"],  // wasm 包外部化
    skipNodeModulesBundle: true,
    onSuccess: () => {
      // 复制 wasm 二进制到 dist
      const require = createRequire(new URL("./package.json", import.meta.url))
      const wasmPath = join(dirname(require.resolve("@dukelib/sheets-wasm")), "duke_sheets_wasm_bg.wasm")
      copyFileSync(wasmPath, new URL("dist/duke_sheets_wasm_bg.wasm", import.meta.url))
    }
  }
])
```

#### demo 的 vite.config.ts 适配

```ts
import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

export default defineConfig({
  plugins: [vue()],
  server: { port: 5000 },
  worker: {
    format: "es",  // Worker 用 ES module 格式（上游对齐）
  },
  // wasm 文件处理：Vite 原生支持 ?url 后缀导入
  // worker 文件：Vite 原生支持 new URL("./worker.ts", import.meta.url)
  optimizeDeps: {
    // 大数据包不预构建（d3/us-atlas/world-atlas 体积大）
    exclude: ["us-atlas", "world-atlas"],
  },
})
```

#### 关键 Vite 特性依赖

| 特性 | Vite 支持 | 说明 |
|---|---|---|
| `new URL("./worker.ts", import.meta.url)` | ✅ 原生 | Vite 自动处理 worker 打包 |
| `new URL("./wasm.wasm", import.meta.url)` | ✅ 原生 | Vite 自动处理 wasm 资源 |
| `worker: { format: "es" }` | ✅ 原生 | ES module worker |
| `?url` 后缀导入 | ✅ 原生 | `import wasmUrl from "..."` |
| `import.meta.url` | ✅ 原生 | ESM 标准 |
| Dynamic import | ✅ 原生 | `import("@dukelib/sheets-wasm")` |
| Top-level await | ⚠️ 需配置 | `build.target: "esnext"` |

#### wasm 加载策略

**方案 A（推荐）：public 目录 + setWasmSource**

```ts
// main.ts
import { setWasmSource } from "@arcships/xlsx-core"
// wasm 文件放在 public/ 下，用绝对路径
setWasmSource("/duke_sheets_wasm_bg.wasm")
```

这是上游 playground 的做法——wasm 放 `public/`，启动时 `setWasmSource` 指定 URL。

**方案 B：import ?url**

```ts
import wasmUrl from "@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm?url"
setWasmSource(wasmUrl)
```

Vite 会把 wasm 作为资源处理，输出带 hash 的 URL。适合生产构建。

**DOCX wasm 同理**：`docx_wasm_bg.wasm`（1MB）也需要复制到 public 或用 `?url` 导入。

#### @tanstack/vue-virtual 替代

上游用 `@tanstack/react-virtual`。Vue 版用 `@tanstack/vue-virtual`（同一库的 Vue 适配）：

```ts
import { useVirtualizer } from "@tanstack/vue-virtual"
// API 基本一致，但需要传 ref 而非直接用
```

如果只实现 canvas 模式（XLSX）或 DOM 虚拟化（DOCX），也可以自实现 visible range + spacer div（上游的核心逻辑就是 `visiblePageIndexes` + `pageStackVirtualSpacers`，不复杂）。

#### d3 / topojson / regl 构建

这些是 chart-renderer 和 surface-regl 的依赖。Vite 原生支持 ESM 导入。注意：
- `us-atlas`（~2MB JSON）和 `world-atlas`（~1MB JSON）会被内联到 bundle。上游 tsup `noExternal: ["us-atlas", "world-atlas"]` 强制内联。Vite 下默认就是内联的。
- `regl` 是 CommonJS 包，Vite 会自动处理 CJS→ESM 转换。
- `@chenglou/pretext`（DOCX 文本布局核心依赖）是 ESM，MIT license，npm 0.0.8 可用。

---

## 十、验证清单

### 10.1 Phase 验证（每个迁移阶段完成后）

| Phase | 验证命令 | 预期结果 |
|---|---|---|
| 1. 直接复制 | `pnpm --filter @arcships/xlsx-core typecheck` | 零错误（修 import 路径后） |
| 2. types 清理 | `pnpm --filter @arcships/xlsx-core typecheck` | 零 React 类型残留 |
| 3. controller 改写 | `pnpm --filter @arcships/vue-xlsx typecheck` + `pnpm --filter @arcships/vue-xlsx build` | 零错误 |
| 4. chart-renderer | `pnpm --filter @arcships/vue-xlsx typecheck` | 零错误 |
| 5. surface-regl | `pnpm --filter @arcships/vue-xlsx typecheck` | 零错误 |
| 6. XlsxViewer 重写 | `pnpm --filter @arcships/vue-xlsx build` | 构建通过 |
| 7. index 调整 | `pnpm typecheck && pnpm build` | 全 workspace 通过 |
| 8. 依赖安装 | `pnpm install` | 无 peer warning |
| 9. License | `git diff --check` | 无空白错误 |

### 10.2 功能验证（每个领域迁移完成后）

#### 引擎层验证

```bash
# XLSX wasm 引擎 smoke
node --input-type=module - <<'NODE'
import init, { Workbook } from '@dukelib/sheets-wasm'
await init()  // 或传 wasm buffer
const wb = Workbook.fromBytes(new Uint8Array(/* xlsx bytes */))
wb.calculate()
console.log(wb.sheetNames, wb.getSheet(0).getCalculatedValueAt(0,0).toString())
const out = wb.saveXlsxBytes()
console.log('saved', out.byteLength)
NODE

# 验证导出
uv run --with openpyxl python -c "
from openpyxl import load_workbook
wb = load_workbook('/tmp/exported.xlsx', data_only=False)
ws = wb.active
assert ws['C1'].value == '=SUM(A1:B1)'
print('OK', ws['C1'].value)
"
```

#### 编辑能力验证清单

| 能力 | 验证方法 | 通过标准 |
|---|---|---|
| 文件加载 | 浏览器打开 welcome.xlsx | sheet 数/名称/单元格值正确 |
| 公式重算 | 修改公式后 Recalc | 依赖值更新 |
| 单元格编辑 | 双击编辑、Enter 提交 | 值持久化到 workbook |
| 样式编辑 | 改字体/颜色/对齐/边框 | 渲染正确 + 导出保留 |
| 合并/拆分 | 选中区域 Merge/Unmerge | 渲染正确 + 导出保留 |
| 行列调整 | 拖拽 header resize | 宽高持久化 + 导出保留 |
| 复制/粘贴 | 选中→Copy→Paste 到目标 | TSV + structured 正确 |
| 填充 | 选中→拖拽 fill handle | 值/公式/样式复制正确 |
| 排序 | 表头菜单 A→Z / Z→A | 行数据重排 + undo 正确 |
| 撤销/重做 | 编辑后 Undo/Redo | 恢复到正确状态 |
| 导出 XLSX | exportXlsx → openpyxl 验证 | cell/merge/formula/style 保留 |
| 导出 CSV | exportCsv → 文本检查 | TSV 格式正确 |
| 冻结窗格 | 加载含冻结的 fixture | 滚动时冻结区固定 |
| 条件格式 | 加载含 dataBar/colorScale 的 fixture | 渲染提示可见 |
| 数据验证 | 加载含 validation 的 fixture | 下拉/提示可见 |
| 图表 | 加载含 charts 的 fixture | 图表渲染 + 选中/拖拽 |
| 图片 | 加载含 images 的 fixture | 图片显示 + 定位正确 |
| worker 模式 | 开启 worker 开关 + 加载大文件 | 不阻塞 UI + worker 状态正确 |

#### DOCX 编辑能力验证清单

| 能力 | 验证方法 | 通过标准 |
|---|---|---|
| 文件导入 | Import .docx | 段落/表格/图片/样式正确渲染 |
| 文本编辑 | 双击段落编辑 | contentEditable + draft 正确 |
| 格式化 | 改字体/颜色/对齐/行距 | 渲染正确 + sourceXml 清除 |
| 表格 | 插入/删除行列 + resize | 结构正确 + 跨页切分 |
| 图片 | 插入/resize/wrap mode | 定位 + wrap exclusion 正确 |
| 分页 | 加载多页文档 | 页数/分页断点正确 |
| 缩略图 | 打开缩略图面板 | canvas 快照渲染 |
| 撤销/重做 | 编辑后 Undo/Redo | 快照恢复正确 |
| 导出 .docx | exportDocx → openpyxl/docx 验证 | 段落/样式/表格/图片保留 |
| 修订/评论 | 加载含修订的文档 | gutter 卡片显示 |
| 表单域 | 双击 form field | Dialog 配置正确 |
| basePackage 导出 | 导入后导出 | styles/numbering/headers 保留 |

### 10.3 验收验证（最终）

对照 `docs/visual-acceptance-handoff.md` 的验收标准：

| 路由 | 验证内容 | 通过标准 |
|---|---|---|
| `/#/` Home | 页面加载 | 无错误 |
| `/#/xlsx-viewer` | fixture 切换 + 编辑 + 导出 | 交互完整 + 导出保真 |
| `/#/docx-viewer` | fixture 切换 + 分页 + 渲染 | 页数正确 + 样式保真 |
| `/#/docx-editor` | 编辑 + 格式化 + 导出 | 编辑正确 + 导出保真 |
| `/#/pdf-viewer` | 分页 + 搜索 + 缩放 | shadcn 一致性 |
| `/#/components` | 组件展示 | shadcn 一致性 |

视口矩阵：`1440×900` / `1280×720` / `768×1024` / `390×844` 逐路由检查。

全量 gate 命令：
```bash
pnpm typecheck && pnpm build && \
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && \
git diff --check
```
