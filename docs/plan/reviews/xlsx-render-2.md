# xlsx-render Review #2

Date: 2026-07-09
Scope: `packages/vue-xlsx/src/render/`
Depends-on: xlsx-composables-split ✅ (composables already split, 9 files)

## 1. Typecheck

```bash
pnpm --filter @extend-ai/vue-xlsx typecheck
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

**Actual state (unchanged from Review #1):**

| File | Lines | Status |
|---|---|---|
| `chart-renderer.tsx` | 7171 | ❌ Not split (7x over limit) |
| `surface-regl.tsx` | 1178 | ❌ Not split (over limit) |
| `index.ts` | 14 | ✅ Barrel exists |

No architecture-required sub-files exist (`chart-bar.ts`, `chart-line.ts`, `chart-pie.ts`, `chart-scatter.ts`, `chart-axis.ts`, `chart-legend.ts`). Both files violate the ≤1000-line hard constraint. The single commit on this module (`ca5f594`) is the pre-existing xlsx-004 port, not a split attempt.

## 3. Import Paths

**Result: ✅ All imports resolve correctly.**

- `@extend-ai/xlsx-core` → workspace symlink ✅
- `@dukelib/sheets-wasm` → external wasm package ✅
- Vue, d3-* packages → standard npm dependencies ✅
- Internal cross-import (`./surface-regl` in `chart-renderer.tsx`) → correct relative path ✅
- `us-atlas/counties-albers-10m.json`, `world-atlas/countries-50m.json`, `topojson-client` → present ✅
- `src/index.ts` barrel correctly re-exports from `./render` ✅

No phantom or broken imports detected.

## 4. Stub/Mock/Fake Residues

**Result: ✅ Render module is clean.** No stub, mock, TODO, FIXME, or placeholder markers in `src/render/`.

Note: `src/index.ts` contains an `XlsxViewer` stub (`"XLSX Viewer (pending)"`), which is xlsx-components scope — out of scope for this review.

## Summary

| Check | Result |
|---|---|
| typecheck zero errors | ✅ |
| Architecture coverage | ❌ Files not split (7171/1178 lines vs ≤1000 limit; 0 of 7 planned sub-files exist) |
| Import paths correct | ✅ |
| Stub/mock/fake residues | ✅ |

## Verdict: BLOCKED

State is identical to Review #1. No progress has been made on splitting `chart-renderer.tsx` (7171 lines) into the 7 architecture-required sub-files (`chart-renderer.ts`, `chart-bar.ts`, `chart-line.ts`, `chart-pie.ts`, `chart-scatter.ts`, `chart-axis.ts`, `chart-legend.ts`) or evaluating `surface-regl.tsx` (1178 lines) for extractable parts. The prerequisite (xlsx-composables-split) is complete, so the task is unblocked from a dependency standpoint.

Work required:

1. Split `chart-renderer.tsx` by chart type into 7 sub-files per architecture §2.2
2. Evaluate `surface-regl.tsx` for extractable non-shader logic (shader code may remain consolidated per §4.3)
3. Re-verify typecheck after split
