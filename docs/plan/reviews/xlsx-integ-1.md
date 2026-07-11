# xlsx-integ Review #1

Date: 2026-07-07

## Findings

### P1 — blocking

无。

### P2 — non-blocking

无。

### P3 — non-blocking

| # | 描述 | 位置 |
|---|---|---|
| 1 | Console 日志仍使用 `[react-xlsx]` 前缀，应更新为 `[xlsx-core]` / `[vue-xlsx]` | [packages/xlsx-core/src/safe-calculate.ts:79](packages/xlsx-core/src/safe-calculate.ts#L79), [safe-calculate.ts:84](packages/xlsx-core/src/safe-calculate.ts#L84), [safe-calculate.ts:96](packages/xlsx-core/src/safe-calculate.ts#L96); [packages/vue-xlsx/src/composables.ts:3657](packages/vue-xlsx/src/composables.ts#L3657), [composables.ts:3672](packages/vue-xlsx/src/composables.ts#L3672), [composables.ts:3692](packages/vue-xlsx/src/composables.ts#L3692) |
| 2 | 内部剪贴板 MIME 仍为 `application/x-react-xlsx-range+json` | [packages/vue-xlsx/src/composables.ts:76](packages/vue-xlsx/src/composables.ts#L76) |
| 3 | JSDoc 注释中引用 `react-xlsx`（示例代码块标记为 `tsx`） | [packages/xlsx-core/src/types.ts:739](packages/xlsx-core/src/types.ts#L739), [types.ts:809](packages/xlsx-core/src/types.ts#L809), [types.ts:818](packages/xlsx-core/src/types.ts#L818) |
| 4 | `@arcships/xlsx-core` 声明了 d3-*/regl/topojson/us-atlas/world-atlas 为依赖，但 xlsx-core 源码未 import 这些包（它们属于 chart-renderer/surface-regl，应在 vue-xlsx） | [packages/xlsx-core/package.json:16-24](packages/xlsx-core/package.json#L16) |

## 集成点验证

### X1: wasm 引擎调用 → ✅ 通过

- `parseWorkbookBuffer` 路径：`getSheetsWasmModule()` → `Workbook.fromBytes()` → `safeCalculate()`（含 `reparse` 兜底）— [composables.ts:1218-1243](packages/vue-xlsx/src/composables.ts#L1218)
- `setCellFormula` → `worksheet.setFormula(cellAddressToA1(cell), formula)` — [composables.ts:4022-4037](packages/vue-xlsx/src/composables.ts#L4022)
- `setCellValue` → `worksheet.setCell()`（空公式）或 `worksheet.setFormula()`（有公式）— [composables.ts:1132-1145](packages/vue-xlsx/src/composables.ts#L1132)
- `maybeRecalculateWorkbook` → `tryRecalculate(targetWorkbook)` — [composables.ts:2672-2681](packages/vue-xlsx/src/composables.ts#L2672)
- `exportXlsx` → `createSavedWorkbookBytes` → `workbook.saveXlsxBytes()` → `sanitizeSavedWorkbookBytes()` → `mergeWorkbookImageAssets()` → `downloadBytes` — [composables.ts:3198-3201](packages/vue-xlsx/src/composables.ts#L3198), [composables.ts:3483-3489](packages/vue-xlsx/src/composables.ts#L3483)
- undo/redo snapshot 路径：`Workbook.fromBytes(cloneBytes(entry.bytes))` — [composables.ts:3230-3232](packages/vue-xlsx/src/composables.ts#L3230)
- `buildSheetList` 使用 `workbook.getSheet()` / `worksheet.getColumnWidth()` / `worksheet.getRowHeight()` / `worksheet.usedRange()` — [composables.ts:537-616](packages/vue-xlsx/src/composables.ts#L537)

无 stub/noop 替代真实调用。所有 `return null` / `return false` 均为合法 guard clause（无 workbook、无 worksheet、非 worker 模式等场景）。

### X2: worker 调用 → ✅ 通过

- `XlsxWorkerClient` 创建：`new Worker(new URL("./xlsx-worker.js", import.meta.url), { type: "module" })` — [worker-client.ts:112](packages/xlsx-core/src/worker-client.ts#L112)
- 消息协议完整：`load`、`parseCharts`、`getCellSnapshot`、`getRowsBatch` 四种消息 — [worker-client.ts:4-78](packages/xlsx-core/src/worker-client.ts#L4)
- ARB transfer clone：`buffer.slice(0)` — [worker-client.ts:228-229](packages/xlsx-core/src/worker-client.ts#L228)
- worker 加载路径：`getWorkerClient().loadWorkbook(buffer, ...)` — [composables.ts:2333](packages/vue-xlsx/src/composables.ts#L2333)
- `getCellSnapshotAsync` → `getWorkerClient().getCellSnapshot(...)` — [composables.ts:2697-2705](packages/vue-xlsx/src/composables.ts#L2697)
- `getRowsBatchAsync` → `getWorkerClient().getRowsBatch(...)` — [composables.ts:2708-2713](packages/vue-xlsx/src/composables.ts#L2708)
- worker 异常回退到主线程（DOMParser/XMLSerializer 缺失、chart payload 不完整）— [composables.ts:2354-2364](packages/vue-xlsx/src/composables.ts#L2354)
- worker dispose 时 reject 所有 in-flight 为 AbortError — [worker-client.ts:117-131](packages/xlsx-core/src/worker-client.ts#L117)

### X3: 图片/图表 assets → ✅ 通过

- `loadWorkbookImageAssets` → `parseWorkbookImageAssets(bytes)` — [composables.ts:1828-1849](packages/vue-xlsx/src/composables.ts#L1828)
- `ensureChartAssetsHydrated` → `loadWorkbookChartAssets(targetWorkbook, imageAssetsRef.value, ...)` — [composables.ts:2055-2080](packages/vue-xlsx/src/composables.ts#L2055)
- `startChartDisplayHydration` → worker `parseCharts` + main thread fallback — [composables.ts:2082-2137+](packages/vue-xlsx/src/composables.ts#L2082)
- `createSavedWorkbookBytes` → `mergeWorkbookImageAssets(sanitizedBytes, ...)` — [composables.ts:3198-3201](packages/vue-xlsx/src/composables.ts#L3198)
- `restoreHistoryEntry` → `loadWorkbookImageAssets` + `loadWorkbookChartAssets` 两阶段恢复 — [composables.ts:3230-3251](packages/vue-xlsx/src/composables.ts#L3230)
- 图片 assets 正确 revoke（`revokeWorkbookImageAssets`）— [composables.ts:1983-1990](packages/vue-xlsx/src/composables.ts#L1983), [composables.ts:2233](packages/vue-xlsx/src/composables.ts#L2233)

## 结论

**pass** — 三个集成点（X1-X3）全部通过真实调用路径验证。composables.ts 对 xlsx-core 的依赖均为 `@arcships/xlsx-core` 包导入，无 stub/noop 替代真实调用。worker 消息协议完整，导出路径覆盖 `saveXlsxBytes` + sanitize + merge images。P3 findings 均为品牌命名残留，不影响功能。
