# 上游 DOCX 功能对齐清单与迁移操作方案

> [!WARNING]
> **历史资料。** 本文保存 2026-07-06 的上游分析和最初迁移方案。文中的本地绝对路径、目标文件、行数和“必须复刻”均不是当前产品状态或发布承诺。当前实现与验收请从[文档索引](INDEX.md)进入。

Date: 2026-07-06

## 目的

本文档基于对上游 `@extend-ai/react-docx`（commit `6f70b92`）全部 8 个子包 + playground 共 ~68000 行源码的逐行深读，产出：

1. **功能对齐清单**：上游每个模块实际做什么、怎么实现（不是接口签名），Vue port 必须复刻什么。
2. **迁移分类**：每个文件按"直接复制 / 复制后清理 / 机械改写 / 重写"分类。
3. **操作方案**：通过拷贝 + 脚本减少工作量的具体步骤。

上游源码位于 `/Users/eric8810/Code/extend-ui-upstream/react-docx/`。

---

## 一、文件迁移分类总表

### 1.1 引擎层 + 布局层 + 辅助模块（零 React 依赖，直接复制）

| 包 | 文件 | 行数 | 迁移策略 | 目标位置 |
|---|---|---:|---|---|
| wasm | `src/index.ts` | 211 | **直接复制** | `docx-core/src/wasm.ts` |
| wasm | `src/generated/docx_wasm.d.ts` | 59 | **直接复制** | `docx-core/src/generated/` |
| wasm | `src/generated/docx_wasm_bg.wasm.d.ts` | 16 | **直接复制** | `docx-core/src/generated/` |
| ooxml-core | `src/index.ts` | 88 | **直接复制** | `docx-core/src/ooxml-core.ts` |
| serializer | `src/index.ts` | 32 | **直接复制** | `docx-core/src/serializer.ts` |
| doc-model | `src/types.ts` | 438 | **直接复制** | `docx-core/src/types.ts` |
| doc-model | `src/clone.ts` | 406 | **直接复制** | `docx-core/src/clone.ts` |
| doc-model | `src/normalize.ts` | 102 | **直接复制** | `docx-core/src/normalize.ts` |
| doc-model | `src/index.ts` | 35 | **直接复制后调导出** | `docx-core/src/index.ts` |
| editor-ops | `src/index.ts` | 1329 | **直接复制** | `docx-core/src/editor-ops.ts` |
| layout-core | `src/pagination.ts` | 689 | **直接复制** | `docx-core/src/pagination.ts` |
| layout-core | `src/page-segmentation.ts` | 1223 | **直接复制** | `docx-core/src/page-segmentation.ts` |
| layout-core | `src/index.ts` | 382 | **直接复制后调导出** | `docx-core/src/layout-core.ts` |
| layout-engine | `src/index.ts` | 359 | **直接复制** | `docx-core/src/layout-engine.ts` |
| react-viewer | `pretext-layout.ts` | 1389 | **直接复制** | `docx-core/src/pretext-layout.ts` |
| react-viewer | `thumbnail-raster.ts` | 1239 | **直接复制** | `docx-core/src/thumbnail-raster.ts` |
| react-viewer | `layout-snapshot.ts` | 469 | **直接复制** | `docx-core/src/layout-snapshot.ts` |
| react-viewer | `docx-import.ts` | 238 | **直接复制** | `docx-core/src/docx-import.ts` |
| react-viewer | `docx-import-worker.ts` | 75 | **直接复制** | `docx-core/src/docx-import-worker.ts` |
| react-viewer | `section-layout.ts` | 298 | **直接复制** | `docx-core/src/section-layout.ts` |
| react-viewer | `pagination-breaks.ts` | 225 | **直接复制** | `docx-core/src/pagination-breaks.ts` |
| react-viewer | `image-render.ts` | 260 | **直接复制** | `docx-core/src/image-render.ts` |
| react-viewer | `page-count-reconciliation.ts` | 278 | **直接复制** | `docx-core/src/page-count-reconciliation.ts` |
| react-viewer | `content-signature.ts` | 155 | **直接复制** | `docx-core/src/content-signature.ts` |
| react-viewer | `canvas/layout-diagnostics.ts` | 156 | **直接复制** | `docx-core/src/canvas/layout-diagnostics.ts` |
| react-viewer | `canvas/types.ts` | 44 | **直接复制** | `docx-core/src/canvas/types.ts` |
| react-viewer | `wasm-source.ts` | 69 | **直接复制** | `docx-core/src/wasm-source.ts` |
| react-viewer | `utif.d.ts` | 18 | **直接复制** | `docx-core/src/utif.d.ts` |
| react-viewer | `core/state.ts` | 227 | **直接复制*** | `docx-core/src/core/state.ts` |

\* `core/state.ts` import 了 `DocxEditorSelection`/`DocxTextRange` 类型 from `../editor`——需先把这些类型提取到独立类型文件，或随 editor.tsx 一起迁移。

**小计：~10500 行可直接复制。**

### 1.2 editor.tsx 拆分（56454 行单文件必须先拆分）

| 区域 | 行号 | 行数 | 性质 | 策略 |
|---|---|---:|---|---|
| 纯常量/helper/布局计算/类型 | 1–24953 | 24953 | 纯函数 + 类型定义；73 处 React 引用全为 `React.CSSProperties` 类型标注 | **复制后清理**（sed 替换 React 类型） |
| `useDocxEditor` + composables | 24954–32870 | 7917 | hook 调用 + controller 逻辑；72 useState, 228 useCallback, 79 useRef, 60 useEffect, 84 useMemo | **机械改写** → `composables.ts` |
| `DocxEditorViewer` + 渲染 | 32871–56454 | 23584 | 巨型 JSX 组件，DOM 渲染 + 虚拟化 + 交互 | **重写** → `DocxEditor.vue` |

### 1.3 其他需改写/重写的文件

| 文件 | 行数 | React 依赖 | 策略 | 目标 |
|---|---:|---|---|---|
| `react-viewer/index.tsx` | 590 | 15 | **重写**（useDocxModel hook + ReactDocxViewer JSX） | `vue-docx/src/DocxViewer.vue` |
| `react-viewer/canvas/viewer.tsx` | 379 | 19 | **重写/跳过**（canvas 编辑模式，深度耦合 controller） | 可选 |
| `editor.tsx` 中的 `renderParagraphRuns` | ~1455 | JSX | **重写**（生成 `<span>` JSX → Vue render） | `vue-docx/src/` |
| `editor.tsx` 中的 `renderStaticHtml` | ~20 | `renderToStaticMarkup` | **重写**（用 Vue `renderToString` 或跳过） | `vue-docx/src/` |

### 1.4 统计

| 迁移策略 | 行数 | 占比 |
|---|---:|---:|
| 直接复制 | ~10500 | ~15% |
| 复制后清理（React 类型标注） | ~24953 | ~37% |
| 机械改写（hook → Vue reactivity） | ~7917 | ~12% |
| 重写（JSX 渲染 + 虚拟化 + 交互） | ~24553 | ~36% |
| **合计** | **~68000** | 100% |

---

## 二、逐模块对齐清单

### 2.1 引擎层（wasm + ooxml-core + serializer + doc-model）

**总行数**：1371 行，**全部零 React 依赖**。

#### wasm（211行）

**实现**：Rust/WASM DOCX 引擎的 JS 封装。单例懒加载 `initWasm()`，浏览器用 `fetch(new URL("./docx_wasm_bg.wasm", import.meta.url))`，Node 用 `process.getBuiltinModule("node:fs/promises")`（避免打包器看到 `node:` specifier）。提供 `setWasmSource()` 覆盖入口。SIMD 报错增强。

**关键**：DOCX 引擎是**无状态函数式**的——传入 bytes/package/model 返回结果，不持有对象（与 XLSX 的 Workbook 对象不同）。

**核心 API**：
- `wasmParseDocx(bytes)` → OOXML package
- `wasmBuildDocModelFromBytes(bytes)` → `{package, model}`
- `wasmSerializeDocx(model, basePackage?)` → ArrayBuffer
- `wasmModelToDocumentXml(model, basePackage?)` → string

**必须复刻**：单例 Promise、Node/浏览器双路径、SIMD 报错、`setWasmSource`。

#### ooxml-core（88行）

**实现**：OOXML 包操作薄层。`parseDocx(bytes)` → `OoxmlPackage`（Map 形式），`packageToArrayBuffer(pkg)` 反向。`createMinimalDocxPackage()` 构造最小合法 .docx。`withPart()` 不可变 part 操作。

#### serializer（32行）

**实现**：DocModel → OOXML 序列化。三个函数：
- `modelToDocumentXml(model, basePackage?)` → document.xml 字符串
- `serializeDocModel(model, basePackage?)` → 完整 OoxmlPackage
- `serializeDocx(model, basePackage?)` → ArrayBuffer

**关键**：`basePackage` 机制——序列化时传原始包给 wasm，保留 styles/numbering/headers 等非 document part。这是"模板编辑后导出"的关键。

#### doc-model（981行）

**实现**：核心文档模型类型 + 构建/克隆/归一化。

**类型体系**（types.ts 438行）：
- `DocModel = { nodes: (ParagraphNode | TableNode)[]; metadata: {...} }`
- 段落 children：`TextRunNode | ImageRunNode | FormFieldRunNode`
- 单元格内可嵌套表格
- 样式：`TextStyle`（run 级）、`ParagraphStyle`（段落级）、`TableStyle`/`TableRowStyle`/`TableCellStyle`
- 单位：twips（段落/表格/间距）+ eighthPt（border）+ Px（CSS 像素）
- 元数据：sections、headerSections、footerSections、paragraphStyles、numberingDefinitions、footnotes、endnotes、comments、compatibility

**clone.ts（406行）**：手工逐字段深拷贝（非 structuredClone），`new Uint8Array(data)` 复制图片二进制。用于编辑操作前的不可变快照。

**normalize.ts（102行）**：修复 JSON 反序列化的 Uint8Array 类型丢失（wasm 传 model 用 JSON 字符串，`number[]` → `Uint8Array`）。

**必须复刻**：全部。types.ts 是整个引擎的契约层，最高优先级。

---

### 2.2 编辑操作（editor-ops，1329行）

**零 React 依赖**。提供段落文本编辑操作：
- `splitParagraphChildrenAtTextOffsets`
- `updateParagraphText`
- `mutateParagraphTextStyleInRange`
- 等

被 editor.tsx 的 `commitParagraphText`/`applySelectedStyleChange` 调用。

---

### 2.3 布局层（layout-core + layout-engine）

**总行数**：2653 行，**全部零 React 依赖**。

#### layout-engine（359行）

**实现**：简易几何装箱式布局。`layoutDocument(model, options)` 把节点流式摆放到页面块，按页高溢出断页。**不做段落跨页切分**——整段不拆的装箱。是"尽力摆放"的基线。

#### layout-core/pagination.ts（689行）

**实现**：分页素材收集层。识别硬分页符（`<w:br w:type="page">`、`<w:pageBreakBefore>`、section break）、表格内分页符、lastRenderedPageBreak 提示。section 解析与 header/footer 继承。spacing 转换（twips→px）。

**关键**：
- `sourceXml` 缓存（同 XML 只解析一次）
- `inheritSectionReferences`（header/footer 按类型继承上一节）
- `resolveParagraphBeforeSpacingPx`（页首 spacing 抑制）
- `collectTableExplicitPageBreakInfo`（表格内分页符，含签名表特殊处理）

#### layout-core/page-segmentation.ts（1223行）

**实现**：真正的分页切分引擎。`buildDocumentPageNodeSegments(model, ...)` 按高度测量切分节点流为 `DocumentPageNodeSegment[][]`。

**核心算法**：
- 段落跨页（按行切分，`ParagraphLineRange`）
- 表格行跨页（`TableRowRange`/`TableRowSlice`）
- keepNext 链（链头推页后链内不再触发）
- widow/orphan（`minLinesPerSegment = widowControl ? 2 : 1`，≤3 行不切分）
- margin collapse（相邻段落间距取 `min(prevAfter, currBefore)`）
- 表格孤行保护

**关键设计**：测量与切分彻底解耦——`PageSegmentationCallbacks` 注入所有高度测量函数，本包不实现测量。

**必须复刻**：切分算法骨架；Vue port 自行实现 callbacks（Canvas measureText / pretext layout）。

---

### 2.4 react-viewer 辅助模块（~5100行，14/17 文件零 React 依赖）

#### docx-import.ts + docx-import-worker.ts（238 + 75行）

**实现**：DOCX 加载流程。`importDocxBuffer(buffer, options)` 决定走 worker 还是主线程：
- worker：`postMessage({type:"import-docx", buffer, wasmSource})` → 返回 `{package, model, timings}`
- 主线程：`parseDocx(buffer)` → `buildDocModel(pkg)`
- AbortSignal 支持

**必须复刻**：worker/主线程双路径、消息协议、AbortSignal。

#### pretext-layout.ts（1389行）⭐

**实现**：基于 `@chenglou/pretext` 库的文本排版引擎。**editor.tsx 文本渲染的核心度量层**。

**核心功能**：
- `layoutTextWithPretextAroundExclusions`：单 font 纯文本绕排除区布局
- `layoutItemsWithPretextAroundExclusions`：多 item（多 font/run）布局
- `resolveOffsetAtPoint`：点击→偏移（命中测试）
- `resolveCaretRectAtOffset`：偏移→光标矩形
- `resolveSelectionRects`：选区→矩形数组
- `sliceLayoutToLineRange`：行范围切片（跨页段落片段）

**缓存**：preparedText(8192) + layout(4096) + lineCount(16384) LRU。

**必须复刻**：全部。重度依赖 `@chenglou/pretext` 库。

#### thumbnail-raster.ts（1239行）

**实现**：DOCX 页面缩略图光栅化，两条路径：
1. **快照直绘**：从紧凑布局数据直接 Canvas 2D 绘制（跳过 DOM）
2. **DOM→SVG foreignObject 光栅化**：克隆页面 DOM → SVG data URI → img decode → drawImage

**辅助**：`DocxThumbnailSurfaceCache`（LRU）、`SerialIdleTaskQueue`（串行空闲任务队列）。

#### section-layout.ts（298行）

**实现**：解析 `<w:sectPr>` XML 为 `DocumentLayoutMetrics`（页面宽高、margin、header/footer 距离）。`twipsToPixels`（TWIPS_PER_PIXEL=15）。页面边框解析。

#### 其他辅助模块

| 文件 | 行数 | 功能 |
|---|---|---|
| `layout-snapshot.ts` | 469 | 布局快照数据结构 + 选区/光标命中操作 |
| `pagination-breaks.ts` | 225 | 显式分页符检测 |
| `page-count-reconciliation.ts` | 278 | 页数校准（缩放高度逼近目标页数） |
| `image-render.ts` | 260 | TIFF→PNG 转换、EMF/WMF 占位 |
| `content-signature.ts` | 155 | FNV-1a 内容签名（缓存友好） |
| `canvas/layout-diagnostics.ts` | 156 | 布局诊断（Canvas 2D 绘制用） |
| `wasm-source.ts` | 69 | wasm 源配置（worker 可用形式） |
| `core/state.ts` | 227 | 编辑器状态机（model+selection+history） |

---

### 2.5 editor.tsx 前半部分（1–24953行，纯函数/类型区）

**React 依赖**：73 处，**全部为 `React.CSSProperties` 类型标注**，零 hook 调用。

**核心纯函数（可直接复制）**：

| 领域 | 关键函数 | 行号 |
|---|---|---|
| 图片 wrap 几何 | `resolveDualWrappedFloatingImageGeometry` | 7232 |
| | `resolveParagraphDualWrappedTextLayout` | 7388 |
| pretext layout | `buildParagraphPretextLayoutSource` | 6586 |
| 行高估算 | `estimateParagraphLineHeightPx` | 10316 |
| | `paragraphLineCountWithinWidth` | 9982 |
| 表格高度 | `estimateTableRowHeightsPx` | 10850 |
| 分页核心 | `buildDocumentPageNodeSegments` | 12759 |
| | `buildRenderColumnSegmentsForPageSection` | 12466 |
| numbering | `buildParagraphNumberingLabels` | 19909 |
| 页眉页脚 | `resolveHeaderPaginationReservePx` | 2217 |
| 测量 | `resolveMeasuredPageContentHeightPx` | 2571 |
| drop cap | `resolveDropCapFontSizePx` | 7695 |
| 默认模型 | `defaultStarterModel` | 4546 |

**需替换 React 类型标注的函数（~25 个）**：所有返回 `React.CSSProperties` 的函数——`wrappedFloatingImageStyle`、`absoluteFloatingImageStyle`、`runStyleToCss`、`paragraphBlockStyle` 等。函数体无需改动，只需 `React.CSSProperties → Record<string, string | number | undefined>`。

**需重写的函数（2 个）**：
- `renderParagraphRuns`（18126–19581，~1455 行）：生成 `<span>` JSX，需改为 Vue render
- `renderStaticHtml`（4902）：调 `renderToStaticMarkup`，需改为 Vue `renderToString` 或跳过

---

### 2.6 editor.tsx 后半部分（24954–56454行，hooks + 组件区）

#### useDocxEditor hook（24954–29580，~4626 行）

**核心状态**（useState × ~20）：
- `model: DocModel` — 不可变文档模型
- `basePackage: OoxmlPackage` — 导入时的原始包（导出回写用）
- `selection: DocxEditorSelection` — 结构化选区（paragraph/table-cell）
- `activeTextRange: DocxTextRange` — 文本选区（location + offset）
- `pendingRunStyle` — **光标处待应用样式**（toolbar 改样式但无选区时缓存）
- `history: {past, future}` — 快照式撤销栈（上限 100），每快照含 model+selection+range
- `historyRestoreRequest` — nonce 驱动的 DOM 选区恢复请求

**核心事务分发器**：`dispatchEditorTransaction`（25545）——**整个编辑器的唯一修改入口**。
```
resolver(ctx) → patch {model?, selection?, range?, pendingRunStyle?, status?}
→ normalizeEditorCursorState → pushHistory → setModel/setSelection/...
```

**编辑能力**：

| 能力 | 实现 |
|---|---|
| 文本编辑 | contentEditable + draft 缓存（`paragraphDraftsRef`）→ `commitParagraphText` → `updateParagraphText`（editor-ops） |
| 格式化 | `applySelectedStyleChange`：展开选区改 run style，折叠改 pendingRunStyle |
| 段落样式 | `applyToSelectedParagraphNode`：clone model → 改 style → 清 sourceXml |
| 表格 | insert/delete row/column、resize（pointer 事件）、跨页切分 |
| 图片 | insert/resize/move/wrap mode、浮动定位、首字下沉 |
| 表单域 | checkbox/text/date/dropdown，SDT + legacy 两种来源 |
| 修订/评论 | **从 model 只读派生**（`collectTrackedChangesFromModel`/`collectCommentsFromModel`） |
| 列表 | toggleList、adjustDepth、numbering labels |
| 撤销/重做 | 快照式，每次 model 变化一个快照 |
| 复制/粘贴 | 纯文本（navigator.clipboard） |
| 导入/导出 | importDocxFile（wasm worker）、exportDocx（serializeDocx + basePackage） |

#### DocxEditorViewer 组件（32871–56454，~23584 行）

**渲染方式**：**纯 DOM**（非 canvas），canvas 仅用于缩略图。

**DOM 结构**：
```
viewerRoot → before spacer（虚拟化占位）
  → page wrapper → page surface
    → cover layer / border overlay / header / body / footer / gutter
  → after spacer
  → context menu / drag overlay
```

**分页**：`pageSegmentationPlan` useMemo → `buildDocumentPageNodeSegments` + 测量驱动迭代（先估算→渲染→测量→重新分页）。

**虚拟化**：`@tanstack/react-virtual`（overscan=2），`visiblePageIndexes` + before/after spacer div。

**交互**：contentEditable 段落编辑、DOM 选区恢复不变量（nonce + draft + suppress 机制）、wheel zoom、表格 resize handle。

**缩略图**：三级策略（快照直绘 → 活页栅格化 → 离屏渲染栅格化），通过 `DocxViewerPageSurfaceRegistry` 与 viewer 解耦。

---

### 2.7 Playground（~10600行）

**关键发现**：**没有 fixture 列表**。默认载入内存中的空白 starter model。用户通过 Import 按钮导入自己的 .docx。

**展示的 26 项功能**：文档主题、应用主题、撤销/重做、段落样式（含 HoverCard 预览）、字体族、字号、行距、字符格式（B/I/U/S/上下标）、文字颜色、高亮色、超链接、段落对齐、列表（bullet/numbered）、分栏显示、页面缩略图（虚拟滚动）、边框（13 种预设）、插入图片、插入表格、缩放、导入 .docx、导出 .docx、修订显示、批注显示、只读模式、右键菜单、表单域配置。

---

## 三、功能对齐要点汇总

### 3.1 引擎层

| # | 要点 | 来源 |
|---|---|---|
| 1 | WASM 单例懒加载 + Node/浏览器双路径 | wasm/index.ts |
| 2 | `setWasmSource` 覆盖 + SIMD 报错增强 | wasm/index.ts |
| 3 | DOCX 引擎是无状态函数式（bytes→model→bytes，不持有对象） | wasm/index.ts |
| 4 | `basePackage` 机制（序列化时保留原包的 styles/numbering/headers） | serializer/index.ts |
| 5 | DocModel 类型体系（discriminated union by type，twips 单位） | doc-model/types.ts |
| 6 | 手工深拷贝（cloneDocModel，非 structuredClone） | doc-model/clone.ts |
| 7 | JSON 归一化（Uint8Array 类型修复） | doc-model/normalize.ts |
| 8 | OoxmlPackage Map 表示 + 最小包构造 | ooxml-core/index.ts |

### 3.2 布局/分页

| # | 要点 | 来源 |
|---|---|---|
| 9 | 测量与切分彻底解耦（PageSegmentationCallbacks 注入） | page-segmentation.ts |
| 10 | 段落跨页按行切分（ParagraphLineRange） | page-segmentation.ts |
| 11 | 表格行跨页（TableRowRange/TableRowSlice） | page-segmentation.ts |
| 12 | keepNext 链（链头推页后链内不再触发） | page-segmentation.ts |
| 13 | widow/orphan（minLinesPerSegment=2，≤3 行不切分） | page-segmentation.ts |
| 14 | margin collapse（相邻段落间距取 min） | page-segmentation.ts |
| 15 | 表格孤行保护 | page-segmentation.ts |
| 16 | 硬分页符识别 + sourceXml 缓存 | pagination.ts |
| 17 | section header/footer 继承 | pagination.ts |
| 18 | 页首 spacing 抑制 | pagination.ts |
| 19 | 表格内分页符（含签名表特殊处理） | pagination.ts |
| 20 | lastRenderedPageBreak 段首提示 | pagination.ts |
| 21 | 页数校准（缩放高度逼近目标，scale 0.2-1.3，二分 10 步） | page-count-reconciliation.ts |
| 22 | 测量驱动迭代分页（先估算→渲染→测量→重新分页） | editor.tsx viewer |
| 23 | pretext 变宽文本布局（绕排除区、光标/选区命中） | pretext-layout.ts |

### 3.3 编辑

| # | 要点 | 来源 |
|---|---|---|
| 24 | `dispatchEditorTransaction` 是唯一修改入口 | editor.tsx useDocxEditor |
| 25 | 不可变模型（每次修改 cloneDocModel + 清 sourceXml） | editor.tsx |
| 26 | 快照式历史（每次 model 变化一个快照，上限 100） | editor.tsx |
| 27 | `pendingRunStyle`（光标处待应用样式） | editor.tsx |
| 28 | contentEditable + draft 缓存（避免 React 重渲染破坏输入 DOM） | editor.tsx viewer |
| 29 | `applySelectedStyleChange`：展开选区改 run style，折叠改 pendingRunStyle | editor.tsx |
| 30 | `historyRestoreRequest` nonce 驱动 DOM 选区恢复 | editor.tsx |
| 31 | selectionSession（pointer/keyboard/composition）控制选区恢复抢占 | editor.tsx |
| 32 | 修订/评论从 model 只读派生（不主动产生标记） | editor.tsx |
| 33 | 导入用 wasm worker（有主线程 fallback） | docx-import.ts |
| 34 | 导出用 `serializeDocx(model, basePackage)` | editor.tsx |

### 3.4 渲染

| # | 要点 | 来源 |
|---|---|---|
| 35 | 纯 DOM 渲染（非 canvas），canvas 仅用于缩略图 | editor.tsx viewer |
| 36 | 虚拟化：`@tanstack/react-virtual`，overscan=2，before/after spacer | editor.tsx viewer |
| 37 | 每页 `contain: "layout style"` 隔离重算 | editor.tsx viewer |
| 38 | `dangerouslySetInnerHTML` 渲染段落 runs | editor.tsx viewer |
| 39 | 浮动图片 absolute 定位，wrap exclusion 由 pretext-layout 计算 | editor.tsx viewer |
| 40 | 表格 resize handle（pointer 事件） | editor.tsx viewer |
| 41 | 缩略图三级策略（快照直绘 → 活页栅格化 → 离屏渲染栅格化） | editor.tsx + thumbnail-raster.ts |
| 42 | `DocxViewerPageSurfaceRegistry` 跨组件共享页面元素 | editor.tsx |
| 43 | `React.CSSProperties` 类型标注（25 个函数，运行时等价普通对象） | editor.tsx 前半 |

### 3.5 加载/导入

| # | 要点 | 来源 |
|---|---|---|
| 44 | `importDocxBuffer` worker/主线程双路径 + AbortSignal | docx-import.ts |
| 45 | worker 消息协议（import-docx request/response） | docx-import.ts |
| 46 | 嵌入字体加载（`loadEmbeddedFontsFromPackage`） | editor.tsx |
| 47 | 默认 starter model（空白文档 + Normal + Heading 1-6 样式） | editor.tsx |

---

## 四、迁移操作方案（拷贝 + 脚本）

### Phase 0：准备

```bash
UPSTREAM=/Users/eric8810/Code/extend-ui-upstream/react-docx/packages
CORE=packages/docx-core/src
VUE=packages/vue-docx/src
```

### Phase 1：直接复制零 React 依赖文件（~10500 行）

```bash
# 引擎层 4 包
cp $UPSTREAM/wasm/src/index.ts              $CORE/wasm.ts
cp $UPSTREAM/wasm/src/generated/*.d.ts      $CORE/generated/
cp $UPSTREAM/ooxml-core/src/index.ts        $CORE/ooxml-core.ts
cp $UPSTREAM/serializer/src/index.ts        $CORE/serializer.ts
cp $UPSTREAM/doc-model/src/types.ts         $CORE/types.ts
cp $UPSTREAM/doc-model/src/clone.ts         $CORE/clone.ts
cp $UPSTREAM/doc-model/src/normalize.ts     $CORE/normalize.ts
cp $UPSTREAM/doc-model/src/index.ts         $CORE/doc-model.ts

# 编辑操作
cp $UPSTREAM/editor-ops/src/index.ts        $CORE/editor-ops.ts

# 布局层 2 包
cp $UPSTREAM/layout-core/src/pagination.ts       $CORE/pagination.ts
cp $UPSTREAM/layout-core/src/page-segmentation.ts $CORE/page-segmentation.ts
cp $UPSTREAM/layout-core/src/index.ts            $CORE/layout-core.ts
cp $UPSTREAM/layout-engine/src/index.ts          $CORE/layout-engine.ts

# react-viewer 辅助模块（14 个零 React 依赖文件）
cp $UPSTREAM/react-viewer/src/pretext-layout.ts          $CORE/pretext-layout.ts
cp $UPSTREAM/react-viewer/src/thumbnail-raster.ts        $CORE/thumbnail-raster.ts
cp $UPSTREAM/react-viewer/src/layout-snapshot.ts          $CORE/layout-snapshot.ts
cp $UPSTREAM/react-viewer/src/docx-import.ts              $CORE/docx-import.ts
cp $UPSTREAM/react-viewer/src/docx-import-worker.ts       $CORE/docx-import-worker.ts
cp $UPSTREAM/react-viewer/src/section-layout.ts           $CORE/section-layout.ts
cp $UPSTREAM/react-viewer/src/pagination-breaks.ts        $CORE/pagination-breaks.ts
cp $UPSTREAM/react-viewer/src/image-render.ts             $CORE/image-render.ts
cp $UPSTREAM/react-viewer/src/page-count-reconciliation.ts $CORE/page-count-reconciliation.ts
cp $UPSTREAM/react-viewer/src/content-signature.ts        $CORE/content-signature.ts
cp $UPSTREAM/react-viewer/src/canvas/layout-diagnostics.ts $CORE/canvas/layout-diagnostics.ts
cp $UPSTREAM/react-viewer/src/canvas/types.ts             $CORE/canvas/types.ts
cp $UPSTREAM/react-viewer/src/wasm-source.ts              $CORE/wasm-source.ts
cp $UPSTREAM/react-viewer/src/utif.d.ts                   $CORE/utif.d.ts
# core/state.ts 需先抽类型（见 Phase 2）
```

**验证**：`pnpm --filter @arcships/docx-core typecheck`（修 import 路径后应通过）。

**import 路径调整**：所有 `@extend-ai/react-docx-*` 的 import 需改为相对路径（如 `@extend-ai/react-docx-doc-model` → `./types` + `./clone` + `./normalize`）。

### Phase 2：editor.tsx 拆分 + 类型清理（~24953 行）

```bash
# 步骤 1：复制 editor.tsx 前半部分到 docx-core
sed -n '1,24953p' $UPSTREAM/react-viewer/src/editor.tsx > $CORE/editor-helpers.ts

# 步骤 2：清理 React 类型标注
sed -i '' 's/React\.CSSProperties/Record<string, string | number | undefined>/g' $CORE/editor-helpers.ts
sed -i '' 's/React\.Dispatch<React\.SetStateAction<string>>/(value: string | ((prev: string) => string)) => void/g' $CORE/editor-helpers.ts
sed -i '' 's/React\.KeyboardEvent<HTMLElement>/{ key: string; shiftKey: boolean; ctrlKey: boolean; metaKey: boolean; preventDefault: () => void }/g' $CORE/editor-helpers.ts
sed -i '' 's/React\.PointerEvent<HTMLElement>/Record<string, unknown>/g' $CORE/editor-helpers.ts
sed -i '' 's/React\.MouseEvent<HTMLElement>/Record<string, unknown>/g' $CORE/editor-helpers.ts
sed -i '' 's/React\.ReactNode/unknown/g' $CORE/editor-helpers.ts

# 步骤 3：删除 import * as React from "react"
sed -i '' 's/import \* as React from "react";//' $CORE/editor-helpers.ts

# 步骤 4：抽出 renderParagraphRuns 和 renderStaticHtml（需重写的 JSX 函数）
# 手动从 editor-helpers.ts 中移到 vue-docx/src/ 待重写
```

### Phase 3：editor.tsx hooks 区机械改写 → composables.ts（~7917 行）

```bash
# 步骤 1：提取 hooks 区
sed -n '24954,32870p' $UPSTREAM/react-viewer/src/editor.tsx > $VUE/composables.ts
```

**脚本预处理**（同 XLSX 的 controller 迁移）：

| React API | Vue 等价 | 数量 |
|---|---|---|
| `useState(x)` | `const x = ref(x)` | ~72 |
| `useCallback(fn, [deps])` | `function fn() { ... }` | ~228 |
| `useRef(x)` | `const xRef = ref(x)` | ~79 |
| `useEffect(fn, [deps])` | `watch(deps, fn)` / `onMounted(fn)` | ~60 |
| `useMemo(fn, [deps])` | `const x = computed(() => fn())` | ~84 |

**特殊处理**：
- state 镜像 ref（`modelRef`/`selectionRef`）可删除——Vue 的 ref 本身就是 `.value`
- `dispatchEditorTransaction` 核心事务逻辑完全保留，只把 `setModel(x)` 改成 `model.value = x`
- `pendingRunStyle` 机制保留
- `historyRestoreRequest` nonce 保留
- `selectionSession` 机制保留

### Phase 4：editor.tsx viewer 区重写 → DocxEditor.vue（~23584 行）

这是最大工作量。建议策略：

1. **提取纯函数到 docx-core**：`pageSegmentationPlan` 的编排逻辑、`visiblePageIndexes` 计算、`pageStackVirtualSpacers` 计算等
2. **Vue 组件重写**：
   - `<template>`：viewer root + spacer divs + page wrapper/surface/header/body/footer/gutter + context menu
   - `<script setup>`：调用 `useDocxEditor`，分页 computed，虚拟化
   - `v-html` 替代 `dangerouslySetInnerHTML`（注意 contentEditable 冲突——需复刻 draft 缓存）
3. **虚拟化**：`@tanstack/vue-virtual` 替代 `@tanstack/react-virtual`
4. **DOM 选区恢复**：`watchEffect`/`nextTick` 重新编排时序

### Phase 5：renderParagraphRuns 重写（~1455 行）

上游用 JSX 生成 `<span>` 元素。Vue 等价：
- 用 render function（`h()`）生成虚拟节点
- 或用 Vue JSX（`@vue/babel-plugin-jsx`）
- 或生成 HTML 字符串 + `v-html`

### Phase 6：依赖安装

```bash
# 上游依赖
pnpm --filter @arcships/docx-core add @chenglou/pretext fast-png utif
pnpm --filter @arcships/vue-docx add @tanstack/vue-virtual
```

注意：`@chenglou/pretext` 是 pretext-layout.ts 的核心依赖，必须可用。

### Phase 7：License 声明

上游 `@extend-ai/react-docx` 是 MIT license。复制源码需保留版权声明。

---

## 五、与 XLSX 迁移的对照

| 维度 | react-xlsx | react-docx |
|---|---|---|
| 总行数 | ~41000 | ~68000 |
| 直接复制 | ~9700 行（7 文件） | ~10500 行（28 文件） |
| 机械改写 | ~5000 行（controller.tsx） | ~7900 行（editor.tsx hooks 区） |
| 重写 | ~16600 行（XlsxViewer.tsx） | ~24500 行（editor.tsx viewer 区 + renderParagraphRuns） |
| 最大单文件 | XlsxViewer.tsx 16615 行 | **editor.tsx 56454 行（必须先拆分）** |
| 引擎模型 | 有状态 Workbook 对象 | 无状态函数式（bytes→model→bytes） |
| 渲染方式 | canvas（默认）+ DOM 兜底 | 纯 DOM |
| 虚拟化库 | 自实现 | @tanstack/react-virtual |
| 文本度量 | WASM getFormattedValueAt | @chenglou/pretext |

**关键差异**：react-docx 把几乎所有代码塞进单个 `editor.tsx`（56454 行），不像 xlsx 按文件分离。迁移**第一步必须是按"纯 helper / hooks / viewer"三区拆分** editor.tsx。

---

## 六、与当前手写实现的对比

| 领域 | 当前手写 | 上游对齐后 |
|---|---|---|
| docx-core 行数 | 2114 行 | ~35000 行（引擎+布局+辅助+helper） |
| vue-docx 行数 | 1090 行 | ~32000 行（composables + viewer + render） |
| 文档模型 | 简化 types.ts | 完整 DocModel 类型体系（438行）+ clone + normalize |
| 解析/序列化 | 手写 XML | wasm 引擎（Rust） |
| 编辑 | stub 方法 | 完整事务分发器 + 60+ 编辑操作 |
| 分页 | 简单 layout.ts | 完整分页引擎（段落/表格跨页、keepNext、widow/orphan、margin collapse） |
| 文本布局 | 无 | pretext 变宽布局（绕排除区、光标/选区命中） |
| 渲染 | 简单 DOM | 完整 DOM 渲染 + 虚拟化 + contentEditable 编辑 + 浮动图片 + 缩略图 |
| 图片 | 基础 | wrap 几何 + 三种 anchor + EMF/TIFF 处理 |
| 表格 | 基础 | resize handle + 跨页切分 + 列宽 |
| 修订/评论 | 无 | 从 model 只读派生 |
| 表单域 | 无 | checkbox/text/date/dropdown |
| 导入/导出 | 手写 | wasm worker + basePackage 机制 |

---

## 七、构建配置对齐

### 7.1 上游构建工具

上游 DOCX 用 **tsup** 构建包，**Vite** 构建 playground。

#### react-viewer tsup.config.ts 关键点

```ts
entry: ["src/index.tsx", "src/docx-import-worker.ts"],  // 主包 + worker
external: ["react", "react-dom"],
noExternal: bundledWorkspacePackages,  // 7 个 workspace 包内联进产物
shims: true,                            // CJS 构建 shim import.meta.url
dts: { resolve: true },                 // DTS 解析 workspace 包类型
onSuccess: "cp ../wasm/src/docx_wasm_bg.wasm dist/docx_wasm_bg.wasm"  // 复制 wasm
```

#### playground vite.config.ts 关键点

```ts
// dev 模式：把 8 个 workspace 包 alias 到源码（不走 dist）
resolve: {
  dedupe: ["react", "react-dom"],
  alias: {
    "@extend-ai/react-docx": ".../packages/react-viewer/src/index.tsx",
    "@extend-ai/react-docx-doc-model": ".../packages/doc-model/src/index.ts",
    "@extend-ai/react-docx-editor-ops": ".../packages/editor-ops/src/index.ts",
    "@extend-ai/react-docx-layout-core": ".../packages/layout-core/src/index.ts",
    "@extend-ai/react-docx-layout-engine": ".../packages/layout-engine/src/index.ts",
    "@extend-ai/react-docx-ooxml-core": ".../packages/ooxml-core/src/index.ts",
    "@extend-ai/react-docx-serializer": ".../packages/serializer/src/index.ts",
    "@extend-ai/react-docx-wasm": ".../packages/wasm/src/index.ts",
  }
},
optimizeDeps: { exclude: Object.keys(workspaceSourceAliases) }  // 不预构建
```

### 7.2 本地 Vite 构建适配

#### docx-core 的 tsup.config

迁移后 docx-core 需要增加 worker entry 和 wasm 复制：

```ts
import { defineConfig } from "tsup"
import { copyFileSync } from "node:fs"

export default defineConfig([
  {
    entry: ["src/index.ts", "src/docx-import-worker.ts"],  // 主包 + worker
    format: ["esm"],
    dts: true,
    clean: true,
    external: ["@chenglou/pretext", "fast-png", "utif"],
    skipNodeModulesBundle: true,
    onSuccess: () => {
      copyFileSync(
        new URL("../../upstream/react-docx/packages/wasm/src/docx_wasm_bg.wasm", import.meta.url),
        new URL("dist/docx_wasm_bg.wasm", import.meta.url)
      )
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
  server: { port: 5173 },
  worker: { format: "es" },
  // wasm: 放 public/ + setWasmSource("/docx_wasm_bg.wasm")
  // worker: Vite 原生支持 new URL("./docx-import-worker.ts", import.meta.url)
})
```

### 7.3 关键 Vite 特性

| 特性 | Vite 支持 | 说明 |
|---|---|---|
| `new URL("./worker.ts", import.meta.url)` | ✅ | Worker 打包 |
| `new URL("./wasm.wasm", import.meta.url)` | ✅ | wasm 资源 |
| `worker: { format: "es" }` | ✅ | ES module worker |
| Dynamic import | ✅ | `import("@arcships/docx-core")` |
| `@chenglou/pretext` | ✅ | ESM, MIT, npm 0.0.8 |
| `fast-png` | ✅ | ESM |
| `utif` | ✅ | CJS, Vite 自动转换 |

### 7.4 wasm 加载策略

与 XLSX 相同，推荐 `public/` + `setWasmSource`：

```ts
import { setWasmSource } from "@arcships/docx-core"
setWasmSource("/docx_wasm_bg.wasm")
```

`docx_wasm_bg.wasm`（1MB）从上游 `packages/wasm/src/` 复制到本地 `apps/demo/public/`。

### 7.5 @tanstack/vue-virtual

上游 DOCX viewer 用 `@tanstack/react-virtual` 做页面虚拟化。Vue 版用 `@tanstack/vue-virtual`，或自实现 visible range + spacer div（核心逻辑就是 `visiblePageIndexes` + `pageStackVirtualSpacers`）。

---

## 八、上游 DOCX Playground 对齐

### 8.1 Fixture

**没有 fixture 文件**。默认载入内存中的空白 starter model（`defaultStarterModel`：1 个空段落 + Normal + Heading 1-6 样式，Calibri 字体）。用户通过 Import 按钮导入自己的 .docx。

### 8.2 功能清单（26 项）

1. 文档主题 light/dark
2. 应用主题 light/dark/system（next-themes）
3. 撤销/重做
4. 段落样式选择（含 HoverCard 实时预览）
5. 字体族（Calibri/Arial/Times New Roman/Georgia/Helvetica/Courier New）
6. 字号（8–48pt）
7. 行距（1/1.15/1.2/1.5/2/2.5/3）
8. 字符格式 Bold/Italic/Underline/Strike/Superscript/Subscript
9. 文字颜色 + 高亮色（含 7 个高亮预设）
10. 超链接编辑/移除/悬停卡片
11. 段落对齐 Left/Center/Right/Justify
12. 列表 Bullet/Numbered
13. 分栏显示
14. 页面缩略图（虚拟滚动，canvas 快照渲染）
15. 边框（13 种预设）
16. 插入图片
17. 插入表格
18. 缩放 50–200%
19. 导入 .docx/.doc
20. 导出 .docx
21. 修订显示开关
22. 批注显示开关
23. 只读模式开关
24. 右键菜单（cut/copy/paste、表格行列增删、图片层级）
25. 表单域配置（双击打开 Dialog）
26. DEV 测试钩子

### 8.3 Toolbar 按钮顺序

`Undo/Redo` → [段落样式 Select+HoverCard] → [字体族 Select] → [字号 Select] → [行距 Select] → `Bold/Italic/Underline/Strike/Superscript/Subscript` → [Text ColorPicker] → [Highlight ColorPicker] → `Link` → `Left/Center/Right/Justify` → `Bullet/Numbered` → [分栏显示] → `Pages`(缩略图) → [Borders 下拉] → `Image/Table` → `Zoom Out/Select/In` → `Import/Download` → [Show edits Switch] → [Show comments Switch] → [Read only Switch] → `Document/Theme`

### 8.4 组件结构

使用 **shadcn/ui**（62 个 `components/ui/*.tsx`）。布局：

```
<div h-[100dvh]>
  <DocxEditorViewer editor={controller} ... />
  // toolbar 在 DocxEditorViewer 内部（通过 props 传入 render 函数）
</div>
```

`useDocxEditor()` 拿到 controller，消费全部 10 个 hook：`useDocxEditor`、`useDocxDocumentTheme`、`useDocxPageLayout`、`useDocxPageThumbnails`、`useDocxParagraphStyles`、`useDocxLineSpacing`、`useDocxBorders`、`useDocxFormFields`、`useDocxTrackChanges`、`useDocxComments`。

### 8.5 本地 demo 对齐要点

- 本地 `apps/demo/src/pages/DocxEditorPage.vue` 应复刻上述 26 项功能
- 本地有 fixture（.docx 样例文件），上游没有——本地可以更丰富
- 上游 toolbar 在 viewer 内部通过 render props 注入；本地需要决定是 viewer 内置还是外部 toolbar

---

## 九、验证清单

### 9.1 Phase 验证

| Phase | 验证命令 | 预期结果 |
|---|---|---|
| 1. 直接复制 28 文件 | `pnpm --filter @arcships/docx-core typecheck` | 零错误（修 import 路径后） |
| 2. editor.tsx 拆分+清理 | `pnpm --filter @arcships/docx-core typecheck` | 零 React 类型残留 |
| 3. hooks 改写 | `pnpm --filter @arcships/vue-docx typecheck` + `build` | 零错误 |
| 4. viewer 重写 | `pnpm --filter @arcships/vue-docx build` | 构建通过 |
| 5. renderParagraphRuns | `pnpm --filter @arcships/vue-docx typecheck` | 零错误 |
| 6. 依赖安装 | `pnpm install` | 无 peer warning |
| 7. License | `git diff --check` | 无空白错误 |

### 9.2 功能验证

#### 引擎层验证

```bash
# DOCX wasm 引擎 smoke
node --input-type=module - <<'NODE'
import { initWasm, wasmBuildDocModelFromBytes, wasmSerializeDocx } from '@arcships/docx-core'
await initWasm()
const fs = await import('node:fs')
const bytes = fs.readFileSync('/tmp/sample.docx')
const { model, package: pkg } = wasmBuildDocModelFromBytes(new Uint8Array(bytes))
console.log('nodes:', model.nodes.length, 'styles:', model.metadata.paragraphStyles?.length)
const out = wasmSerializeDocx(model, pkg)
console.log('saved:', out.byteLength)
NODE

# 验证导出
uv run --with python-docx python -c "
from docx import Document
doc = Document('/tmp/exported.docx')
print('paragraphs:', len(doc.paragraphs))
print('tables:', len(doc.tables))
for p in doc.paragraphs[:5]:
    print(repr(p.text), p.style.name)
"
```

#### 编辑能力验证清单

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
| 导出 .docx | exportDocx → python-docx 验证 | 段落/样式/表格/图片保留 |
| 修订/评论 | 加载含修订的文档 | gutter 卡片显示 |
| 表单域 | 双击 form field | Dialog 配置正确 |
| basePackage 导出 | 导入后导出 | styles/numbering/headers 保留 |
| numbering | 加载含列表的文档 | 编号标签正确 |
| 浮动图片 | 加载含 wrap 图片的文档 | 文本绕排正确 |
| 脚注/尾注 | 加载含脚注的文档 | 脚注区域渲染 |
| contentEditable draft | 连续输入文本 | 光标不跳 + draft 正确提交 |
| DOM 选区恢复 | 编辑后 undo | 光标回到正确位置 |
| 页数校准 | 加载有 documentPageCount 的文档 | 页数匹配（±3 内缩放收敛） |

### 9.3 验收验证

对照 `docs/visual-acceptance-handoff.md`：

| 路由 | 验证内容 | 通过标准 |
|---|---|---|
| `/#/docx-viewer` | fixture 切换 + 分页 + 渲染 | 页数正确 + 样式保真 |
| `/#/docx-editor` | 编辑 + 格式化 + 导出 | 编辑正确 + 导出保真 |

视口矩阵：`1440×900` / `1280×720` / `768×1024` / `390×844` 逐路由检查。

全量 gate：
```bash
pnpm typecheck && pnpm build && \
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && \
git diff --check
```
