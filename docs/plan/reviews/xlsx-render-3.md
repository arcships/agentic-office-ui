# xlsx-render Review #3

Date: 2026-07-09
Scope: `packages/vue-xlsx/src/render/`
Depends-on: xlsx-composables-split ✅

## 1. Typecheck

```bash
pnpm --filter @extend-ai/vue-xlsx typecheck
```

**Result: ✅ Zero errors.** TypeScript compilation passes cleanly. Build also passes.

## 2. Architecture Coverage

Architecture doc: `docs/xlsx-migration-architecture.md` §2.2 defines xlsx-render as:

```
render/
├── chart-renderer.ts           (~800) 图表渲染主逻辑
├── chart-bar.ts                (~500) 柱状图
├── chart-line.ts               (~500) 折线图
├── chart-pie.ts                (~400) 饼图
├── chart-scatter.ts            (~500) 散点图
├── chart-axis.ts               (~600) 坐标轴
├── chart-legend.ts             (~400) 图例
├── surface-regl.ts            (800) WebGL surface(可超1000)
└── index.ts                    (barrel)
```

**Actual state:**

| File | Lines | Status |
|---|---|---|
| `chart-renderer.tsx` | 3430 | ⚠️ Over limit (~800 target); shared utilities + dispatcher |
| `chart-bar.tsx` | 945 | ✅ Under 1000; bar/waterfall/funnel/box-whisker |
| `chart-line.tsx` | 916 | ✅ Under 1000; line/area/combo/radar/stock |
| `chart-pie.tsx` | 612 | ✅ Under 1000; pie/doughnut/sunburst/treemap |
| `chart-scatter.tsx` | 493 | ✅ Under 1000; scatter/bubble |
| `chart-surface.tsx` | 821 | ℹ️ Extra (not in arch); surface contour/region map |
| `chart-axis.ts` | — | ❌ Missing; axis logic still inside chart-renderer.tsx |
| `chart-legend.ts` | — | ❌ Missing; legend logic still inside chart-renderer.tsx |
| `surface-regl.tsx` | 1178 | ⚠️ Over target but allowed (§4.3 WebGL exception) |
| `index.ts` | 14 | ✅ Barrel; exports MemoChartSvg + MemoSurfaceChartComposite |

**Progress from Review #2:**

- chart-renderer.tsx: 7171 → 3430 lines (52% reduction)
- 4 chart-type sub-files extracted: chart-bar, chart-line, chart-pie, chart-scatter
- 1 extra file: chart-surface (surface chart + region map, which wasn't a separate arch target but is a valid domain split)
- Architecture gaps: chart-axis.ts and chart-legend.ts still unextracted

**Remaining work:**

1. Extract axis rendering (`renderCartesianAxes`, `renderSurfaceAxes`, tick formatting helpers) into `chart-axis.tsx`
2. Extract legend rendering (`renderLegend`, `buildLayout`, `getLegendItems`) into `chart-legend.tsx`
3. chart-renderer.tsx after axis+legend extraction would be ~2430 lines — still above ~800 target, indicating further extractable shared utilities remain

## 3. Import Paths

**Result: ✅ All imports resolve correctly.**

- `@extend-ai/xlsx-core` → workspace symlink ✅
- `@dukelib/sheets-wasm` → external wasm package ✅
- Vue (`vue`) → peer dependency ✅
- d3-* (`d3-scale`, `d3-shape`, `d3-geo`, `d3-hierarchy`) → npm dependencies ✅
- `fflate`, `regl`, `topojson-client` → npm dependencies ✅
- `us-atlas/counties-albers-10m.json`, `world-atlas/countries-50m.json` → resolved ✅
- Internal cross-imports (`./chart-renderer` ← sub-files, `./chart-bar`/`./chart-line`/`./chart-pie`/`./chart-scatter`/`./chart-surface` ← chart-renderer) → correct relative paths ✅
- `src/render/index.ts` barrel → correctly re-exports from `./chart-renderer` and `./surface-regl` ✅
- `src/index.ts` → correctly re-exports `MemoChartSvg`, `MemoSurfaceChartComposite` from `./render` ✅

The hub-and-spoke import pattern (chart-renderer.tsx imports from sub-files, sub-files import shared utilities from chart-renderer.tsx) creates circular dependencies. These resolve correctly because ES module named exports are hoisted and the shared utilities are used at call time (inside render functions), not at module evaluation time.

## 4. Stub/Mock/Fake Residues

**Result: ✅ Render module is clean.** No stub, mock, TODO, FIXME, or placeholder markers found in `src/render/`.

Found elsewhere (out of render scope):

- `src/index.ts`: `XlsxViewer` defineComponent stub ("XLSX Viewer (pending)") — xlsx-components scope
- `src/composables/useXlsxViewerController.ts`: `_setChartSeriesFormula = () => false` is a legitimate wiring placeholder replaced by chart domain during initialization

## Summary

| Check | Result |
|---|---|
| typecheck zero errors | ✅ |
| Architecture coverage | ⚠️ chart-axis/legend missing; chart-renderer 3430 lines (~800 target) |
| Import paths correct | ✅ |
| Stub/mock/fake residues | ✅ |

## Verdict: BLOCKED

Significant progress vs Review #2: chart-renderer.tsx halved from 7171 to 3430 lines, 4 chart-type files + 1 surface file extracted. However, the architecture-defined chart-axis.ts and chart-legend.ts remain inside chart-renderer.tsx, and the main file at 3430 lines still far exceeds the ≤1000-line hard constraint. The surface-regl 1178 lines fall under the §4.3 WebGL exception.

Work required:

1. Extract `renderCartesianAxes` / `renderSurfaceAxes` and tick formatting into `chart-axis.tsx` (~600 lines target)
2. Extract `renderLegend` / `buildLayout` / `getLegendItems` into `chart-legend.tsx` (~400 lines target)
3. Re-evaluate chart-renderer.tsx remaining shared utilities for further extraction toward ~800 target
4. Re-verify typecheck after extraction
