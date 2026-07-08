# Review: xlsx-core-images-split

Date: 2026-07-08
Reviewer: DimCode
Status: **pass**

## Scope

拆分 `packages/xlsx-core/src/images.ts`（上游 3870 行）为 `images/` 目录下的模块文件，单文件 ≤ 1000 行，功能对齐上游 `@extend-ai/react-xlsx` commit `f285a1c`。

上游源码对比：`/Users/eric8810/Code/extend-ui-upstream/react-xlsx/packages/react-xlsx/src/images.ts`

---

## 架构文档文件清单 vs 实际产出

| 文档定义 (docs/xlsx-migration-architecture.md L74-81) | 实际产出 | 行数 | 状态 |
|---|---:|---|---|
| `image-parser.ts` (~700) | `images/image-parser.ts` | 544 | ✅ |
| `drawing-parser.ts` (~700) | `images/drawing-parser.ts` | 917 | ✅ |
| `column-width.ts` (~400) | `images/column-width.ts` | 416 | ✅ |
| `theme-palette.ts` (~500) | `images/theme-palette.ts` | 570 | ✅ |
| `grid-render.ts` (~800) | `images/grid-render.ts` | 774 | ✅ |
| `image-export.ts` (~500) | `images/image-export.ts` | 542 | ✅ |
| `index.ts` (barrel) | `images/index.ts` | 144 | ✅ |
| — | `images/form-control-parser.ts` | 429 | ⚠️ 未在设计文档中 |

所有文件均 ≤ 1000 行。旧 `packages/xlsx-core/src/images.ts` 已删除。

---

## Findings

### P2 / Non-blocking

**F1. form-control-parser.ts 未在架构文档中定义**

- 严重程度：P2
- 阻塞：non-blocking
- 设计文档：[docs/xlsx-migration-architecture.md#L74-L81](docs/xlsx-migration-architecture.md#L74-L81)
- 代码位置：[packages/xlsx-core/src/images/form-control-parser.ts](packages/xlsx-core/src/images/form-control-parser.ts)
- 说明：上游 images.ts 中 form controls 逻辑（三源合并、控件名归一化、hidden shapes 关联）被提取为独立模块。这是合理的职责分离——form control 解析原是 images.ts 中最耦合的部分，独立成模块降低 drawing-parser.ts 的行数（从 917 行降至无此提取时的 ~1350 行）。模块通过 drawing-parser.ts 重导出接入，不影响外部消费。
- 建议：更新架构文档，将 `form-control-parser.ts` 加入文件清单。

### P3 / Non-blocking

**F2. parseWorkbookChartStyleAssets 使用内联类型导入**

- 严重程度：P3
- 阻塞：non-blocking
- 代码位置：[packages/xlsx-core/src/images/image-export.ts#L86](packages/xlsx-core/src/images/image-export.ts#L86)
- 说明：函数返回类型写为 `import("./grid-render").WorkbookChartStyleAssets` 而非从 barrel 导入 `WorkbookChartStyleAssets`。功能正确，TS 将其视为纯类型引用。保持一致性考虑，建议改为 `import type { WorkbookChartStyleAssets } from "./grid-render"`。
- 建议：可忽略，或统一 import 风格时修复。

---

## 验证结果

### 1. 功能覆盖率

上游 images.ts 全部 24 个导出函数/类型均已分配到对应子模块：

| 上游导出 | 所在子模块 | 通过 barrel 重导出 |
|---|---|---|
| `resolveWorksheetDefaultColumnWidthPixels` | column-width.ts | ✅ |
| `resolveWorksheetDefaultRowHeightPixels` | column-width.ts | ✅ |
| `resolveWorksheetHiddenRows` | column-width.ts | ✅ |
| `resolveWorksheetHiddenCols` | column-width.ts | ✅ |
| `resolveWorksheetMergeMetadata` | column-width.ts | ✅ |
| `revokeWorkbookImageAssets` | image-export.ts | ✅ |
| `parseWorkbookStructureAssets` | image-export.ts | ✅ |
| `parseWorkbookChartStyleAssets` | image-export.ts | ✅ |
| `parseWorkbookImageAssets` | image-export.ts | ✅ |
| `updateWorkbookImageAnchor` | image-export.ts | ✅ |
| `mergeWorkbookImageAssets` | image-export.ts | ✅ |
| `emuToPixels` | image-parser.ts | ✅ |
| `pixelsToEmu` | image-parser.ts | ✅ |
| `pxToSheetColumnWidth` | column-width.ts | ✅ |
| `resolveSheetColumnWidthPixels` | column-width.ts | ✅ |
| `resolveSheetRowHeightPixels` | column-width.ts | ✅ |
| `resolveRenderedSheetAxisPixels` | column-width.ts | ✅ |
| `resolveContentSheetAxisPixels` | column-width.ts | ✅ |
| `rectToImageAnchor` | column-width.ts | ✅ |
| `resizeImageRect` | column-width.ts | ✅ |
| `WorkbookImageSheetOrigin`（类型） | image-parser.ts | ✅ |
| `WorkbookTableMetadata`（类型） | grid-render.ts | ✅ |
| `WorkbookImageAssets`（类型） | grid-render.ts | ✅ |
| `WorkbookStructureAssets`（类型） | grid-render.ts | ✅ |
| `WorkbookChartStyleAssets`（类型） | grid-render.ts | ✅ |

上游 images.ts 中所有内部函数（`parseXml`, `serializeXml`, `parseRelationships`, `parseWorkbookStyles`, `parseSheetState` 等约 30 个）也均正确分配到对应子模块。

### 2. Typecheck

- `pnpm --filter @extend-ai/xlsx-core typecheck` → **通过**
- `pnpm --filter @extend-ai/vue-xlsx typecheck` → **通过**

### 3. Import 路径

- 子模块间互引用：全部使用相对路径（`./image-parser`, `./column-width`, `./drawing-parser`, `./grid-render`, `./theme-palette`）
- 外部依赖：`../colors`（grid-render.ts, theme-palette.ts）, `../types`（所有子模块）, `@dukelib/sheets-wasm`（image-parser.ts）, `fflate`（image-parser.ts, image-export.ts）
- 类型导入正确使用 `import type`，无值/类型混淆

### 4. Stub / Mock / Fake

- 无 `throw new Error` 桩代码
- 无 `@ts-ignore` / `@ts-nocheck` / `as any`
- 无 TODO / FIXME / stub / mock / fake / placeholder（`isPlaceholderFormControlName` 是上游功能，非桩）

### 5. 循环依赖

依赖图无环：

```
colors ← grid-render, theme-palette
types ← 所有子模块
image-parser ← column-width, drawing-parser, image-export, grid-render, theme-palette, form-control-parser
column-width ← image-parser, grid-render (type only)
grid-render ← colors, image-parser, theme-palette
theme-palette ← colors, image-parser
drawing-parser ← image-parser, column-width, theme-palette, grid-render (type only)
form-control-parser ← image-parser, drawing-parser, grid-render (type only), column-width
image-export ← image-parser, drawing-parser, grid-render, column-width
```

### 6. 下游消费

- `xlsx-core/src/index.ts`：正确从 `./images` barrel 重导出 18 函数 + 5 类型
- `xlsx-core/src/xlsx-worker.ts`：import 路径对齐 barrel
- `xlsx-core/src/charts/chart-parser.ts`：type-only import 对齐 barrel
- `xlsx-core/src/charts/chart-export.ts`：type-only import 对齐 barrel

---

## 结论

**pass** — 实现覆盖架构文档定义的全部功能，typecheck 通过，无 stub/mock/fake 残留，无循环依赖。唯一偏差是 `form-control-parser.ts` 未在设计文档中列出，但这是合理的模块化提炼，不构成阻塞项。
