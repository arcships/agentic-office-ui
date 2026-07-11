# docx-viewer Review #1

Date: 2026-07-07
Reviewer: automated
Scope: `packages/docx-core/src/viewer/` (step 4 of migration execution order)
Upstream: `@extend-ai/react-docx` @ commit `6f70b92`

## Findings

### P2 / non-blocking — pretext split uses 3 files instead of 2 per architecture estimate

- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.3 — estimates pretext-layout.ts at ~700 lines and pretext-selection.ts at ~689 lines (split upstream's 1389-line pretext-layout.ts into 2 files).
- **Code:** [pretext-layout.ts](../../packages/docx-core/src/viewer/pretext-layout.ts): 679 lines, [pretext-selection.ts](../../packages/docx-core/src/viewer/pretext-selection.ts): 237 lines, [pretext-items-layout.ts](../../packages/docx-core/src/viewer/pretext-items-layout.ts): 471 lines. Total: 1387 lines.
- **Detail:** The implementation moved `layoutItemsWithPretextAroundExclusions` into a third file (pretext-items-layout.ts) instead of keeping it in pretext-layout.ts. All three files are under the 1000-line limit. The split boundary is a valid refinement — multi-item layout is a distinct concern from single-item layout.
- **Impact:** None. Functional equivalence is maintained. Upstream exports are fully covered across the three files.

### P2 / non-blocking — viewer/index.ts omits 3 modules due to symbol collision

- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.3 — viewer/index.ts is a barrel exporting all 15 modules.
- **Code:** [viewer/index.ts](../../packages/docx-core/src/viewer/index.ts#L3-L7) — explicitly excludes `wasm-source`, `layout-snapshot`, and `pagination-breaks` from the barrel due to symbol collision with engine/ and layout/ barrels. Consumers must import these directly (e.g., `import { … } from "@arcships/docx-core/viewer/wasm-source"`).
- **Detail:** The comment in index.ts documents the reason. These three files are still compiled and typechecked; only the barrel re-export path is missing. The excluded files are available to editor helpers that compose them internally.
- **Impact:** Low. Consumers who need `LayoutSnapshot`, `collectTopLevelExplicitPageBreakStartNodeIndexes`, or wasm source functions must know the direct import path. Consider adding explicit re-exports under renamed identifiers in a future cleanup.

### P3 / non-blocking — thumbnail-raster.ts line count higher than estimate

- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.3 — estimates thumbnail-raster.ts at ~650 lines.
- **Code:** [thumbnail-raster.ts](../../packages/docx-core/src/viewer/thumbnail-raster.ts): 941 lines.
- **Detail:** Still within the ≤1000 hard constraint. The additional lines come from inline type definitions for snapshot elements (ParagraphSnapshot, ImageSnapshot, TableSnapshot, etc.) and the `rasterizeDocxThumbnailSurface` DOM→SVG pipeline implementation. Thumbnail cache was successfully separated into thumbnail-cache.ts (295 lines).
- **Impact:** None at 941 lines, but close to the ceiling. Further refactoring may be needed if future changes push it over 1000.

## Verification Summary

| Check | Result | Detail |
|---|---|---|
| File inventory | ✅ Pass | 16 files: pretext-layout.ts, pretext-selection.ts, pretext-items-layout.ts, thumbnail-raster.ts, thumbnail-cache.ts, docx-import.ts, docx-import-worker.ts, layout-snapshot.ts, section-layout.ts, pagination-breaks.ts, page-count-reconciliation.ts, image-render.ts, content-signature.ts, wasm-source.ts, utif.d.ts, index.ts |
| Upstream file coverage | ✅ Pass | All 15 planned files from architecture §2.3 present. One additional file (pretext-items-layout.ts) is a refinement of the pretext split |
| Line count ≤1000 | ✅ Pass | Max 941 (thumbnail-raster.ts), all under limit |
| Upstream content parity | ✅ Pass | All upstream exports present: layoutTextWithPretextAroundExclusions, layoutItemsWithPretextAroundExclusions, resolveOffsetAtPoint, resolveCaretRectAtOffset, resolveSelectionRects, sliceLayoutToLineRange, importDocxBuffer (dual path), DocxThumbnailSurfaceCache, SerialIdleTaskQueue, renderDocxThumbnailSnapshotSurface, TIFF→PNG conversion, FNV-1a signing, DocumentLayoutMetrics, pagination break detection, page count reconciliation, etc. |
| Import paths | ✅ Pass | All relative (`../engine/*` or `./*`). Zero `@extend-ai/` references in source code |
| Typecheck | ✅ Pass | `pnpm --filter @arcships/docx-core typecheck` exits 0 with no errors |
| Stub/mock/fake | ✅ Pass | No stubs or mocks. "placeholder" in image-render.ts (EMF/WMF) is documented upstream behavior — `imageUsesPlaceholderFallback()` and `unsupportedImageFallbackLabel()` are intentional |
| Circular dependencies | ✅ Pass | Linear internal deps: pretext-selection → pretext-layout, pretext-items-layout → pretext-layout, docx-import → wasm-source, layout-snapshot → pretext-layout + section-layout. No cycles |

### Cross-cutting note: editor-ops.ts in same package

- [editor-ops.ts](../../packages/docx-core/src/editor/editor-ops.ts) is 1329 lines and has not been split per [architecture §2.5](../docx-migration-architecture.md) (planned: paragraph-ops.ts ~480, run-style-ops.ts ~450, table-ops.ts ~399). This is outside the docx-viewer scope (belongs to the `docx-editor-ops` task). The docx-core barrel exports it from `./editor/editor-ops`, and typecheck passes. Flagged for awareness since it's in the same package and exported as "complete."

## Conclusion

**pass** — Viewer layer is complete and aligned. All 15 planned files are present, typecheck passes, import paths are correct, and no stubs are present. The three minor deviations (3-file pretext split, non-re-exported modules, thumbnail-raster at 941 vs ~650) are non-blocking implementation refinements within constraints.
