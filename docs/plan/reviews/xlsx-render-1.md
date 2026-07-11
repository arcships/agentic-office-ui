# xlsx-render Review #1

Date: 2026-07-09
Scope: `packages/vue-xlsx/src/render/`
Depends-on: xlsx-composables-split ✅ (composables already split, 9 files, ~5032 lines total)

## 1. Typecheck

```bash
pnpm --filter @arcships/vue-xlsx typecheck
```

**Result: ✅ Zero errors.** TypeScript compilation passes cleanly.

## 2. Architecture Coverage

Architecture doc: `docs/xlsx-migration-architecture.md` section 2.2 defines xlsx-render as:

```
render/
├── chart-renderer.ts           (~800) 图表渲染主逻辑
├── chart-bar.ts                (~500) 柱状图
├── chart-line.ts               (~500) 折线图
├── chart-pie.ts                (~400) 饼图
├── chart-scatter.ts            (~500) 散点图
├── chart-axis.ts               (~600) 坐标轴
├── chart-legend.ts             (~400) 图例
├── surface-regl.ts            (800) WebGL surface
└── index.ts                    (barrel)
```

**Actual state:**

| File | Lines | Status |
|---|---|---|
| `chart-renderer.tsx` | 7171 | ❌ Not split (7x over limit) |
| `surface-regl.tsx` | 1178 | ❌ Not split (over limit) |
| `index.ts` | 14 | ✅ Barrel exists |

The render module contains the old xlsx-004 deliverable: a direct React→Vue port of the upstream monolithic files. The architecture-required split into sub-files by chart type (bar/line/pie/scatter/axis/legend) is **not implemented**. Both files violate the hard constraint of ≤1000 lines per file.

The hard constraint in section 4.3 allows surface-regl to exceed 1000 lines (WebGL shader code is dense), but only as a deliberate exception. The current 1178-line surface-regl.tsx has not been evaluated for what could be extracted.

## 3. Import Paths

**Result: ✅ All imports resolve correctly.**

- `@arcships/xlsx-core` → workspace symlink at `packages/xlsx-core` ✅
- `@dukelib/sheets-wasm` → external wasm package ✅
- Vue, d3-* packages → standard npm dependencies ✅
- Internal cross-imports (`./surface-regl`, `./chart-renderer`) → correct relative paths ✅
- `us-atlas/counties-albers-10m.json` and `world-atlas/countries-50m.json` → present in node_modules ✅

No phantom or broken imports detected.

## 4. Stub/Mock/Fake Residues

**Result: ✅ Render module is clean.** No stub, mock, todo, or placeholder markers found in `src/render/`.

Note: `src/index.ts` contains an `XlsxViewer` stub (`"XLSX Viewer (pending)"`), but that belongs to xlsx-components scope, not xlsx-render.

## Summary

| Check | Result |
|---|---|
| typecheck zero errors | ✅ |
| Architecture coverage | ❌ Files not split per plan (7171/1178 lines vs ≤1000 limit) |
| Import paths correct | ✅ |
| Stub/mock/fake residues | ✅ |

## Verdict: **BLOCKED**

xlsx-render is defined as splitting the monolithic chart-renderer.tsx (7171 lines) into ~7 sub-files. The current code is the unsplit xlsx-004 port. The prerequisite (xlsx-composables-split) is already complete, so the task is unblocked from a dependency standpoint. Work required: split `chart-renderer.tsx` by chart type into `chart-renderer.ts` + `chart-bar.ts` + `chart-line.ts` + `chart-pie.ts` + `chart-scatter.ts` + `chart-axis.ts` + `chart-legend.ts`, and evaluate `surface-regl.tsx` for extractable parts.
