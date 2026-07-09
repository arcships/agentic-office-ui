# xlsx-composables-wire — Review 1

**Date:** 2026-07-09
**Status:** ✅ PASS

## Checks

### 1. Typecheck — Zero Errors ✅

```
pnpm --filter @extend-ai/vue-xlsx typecheck
```

`tsc --noEmit` completes with zero diagnostics.

### 2. composables.ts Line Count — 104 ✅

Well under the 1000-line ceiling.

### 3. composables/ File References ✅

All 10 files under `composables/` are transitively reachable from `useXlsxViewerController`:

| File | Imported by `useXlsxViewerController` |
|------|--------------------------------------|
| `useXlsxViewerController.ts` | Self (the main controller) |
| `workbook-state.ts` | ✅ `createHistoryDomain`, `clampZoomScale`, etc. |
| `editing.ts` | ✅ `createEditingDomain` |
| `clipboard.ts` | ✅ `createClipboardDomain` |
| `chart-controller.ts` | ✅ `createChartImageDomain` |
| `navigation.ts` | ✅ `createNavigationDomain` |
| `internal.ts` | ✅ constants & helpers (`DEFAULT_DEFER_LOADING_ABOVE_BYTES`, etc.) |
| `selection.ts` | ✅ `cellAddressToA1`, `rangeToA1` |
| `formatting.ts` | ✅ `XlsxFileSizeLimitExceededError`, `createWorkbookTooLargeError`, etc. |
| `image-assets.ts` | ✅ Transitive — imported by `chart-controller.ts` + `workbook-state.ts` |

No orphan files.

### 4. Dead Code Analysis ⚠️

**No parallel (duplicate) dead code found.** No function is implemented in two places.

**Minor observation:** `composables.ts` re-exports a wide surface (domain creators, utilities, constants, types from all sub-modules), but the official entry point `index.ts` only re-exports `useXlsxViewerController` and `XlsxFileSizeLimitExceededError`. The remaining re-exports in `composables.ts` are unreachable from external consumers via `@extend-ai/vue-xlsx`. These are harmless — typecheck still passes and tree-shaking removes them from the bundle — but the barrel could be trimmed to match the actual public API surface.

## Summary

All hard gates pass. The wiring is clean: `useXlsxViewerController` directly imports from 8 domain files, which cover all 10 files in `composables/`. No orphan code, zero type errors, composables.ts is 104 lines.
