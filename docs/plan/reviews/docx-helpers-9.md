# docx-helpers Review #9

Date: 2026-07-08
Reviewer: automated
Scope: `packages/docx-core/src/editor/helpers/` (step 5 of migration execution order)
Upstream: `@extend-ai/react-docx` @ commit `6f70b92` (verified: `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`), `react-viewer/src/editor.tsx` lines 1-24953 (pure functions / types / constants region, zero React hooks); upstream file confirmed at 56454 lines total.
Design refs:
- [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — editor/helpers module plan (~21200 lines, 40 content files + barrel)
- [docx-editor-helpers-split-plan.md](../docx-editor-helpers-split-plan.md) — 24-module detailed split
- [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §2.5 — editor.tsx 前半部分 pure-function region; §3 "必须复刻" alignment items #22, #23, #32

Commits under review (cumulative task state):
- `3a4ad69 docx-helpers: docx-core editor/helpers` (batch 1, reviewed in #1)
- `e24a049 docx-helpers: fix #1` (batch 2, reviewed in #2)
- `a909c27 docx-helpers: fix #2` (batch 3, reviewed in #3)
- `62dba91 docx-helpers: fix #6` (batch 4 — added `page-measurement.ts`; reviewed in #7)
- `4326037 docx-helpers: fix #7` (adds `docx-helpers-7.md` only; reviewed in #8)
- `fb017f0 docx-helpers: fix #8` (this cycle — added `selection-helpers.ts`, `selection-restore.ts`, `section-manipulation.ts` + `editor-types.ts` compare/normalize helpers)

HEAD is `fb017f0`; working tree clean. Since review #8, **3 new helper modules landed** (+1577 lines: `selection-helpers.ts` 1000, `selection-restore.ts` 135, `section-manipulation.ts` 370, and +72 lines of range-comparison helpers appended to `editor-types.ts`). The deliverable surface grew from 27 → 30 content modules (~14508 lines excl. barrel). Against the 40-file plan the task is ~75% complete by file count (30 of 40).

## Summary

Review #8 blocked on 14 missing modules with zero progress since #7. This cycle delivered **3 new modules** (`selection-helpers`, `selection-restore`, `section-manipulation`) — the first real progress since batch #6 (`62dba91`). The three modules are well-implemented: they faithfully port upstream `editor.tsx` lines 23538-24639 + 24641-24952, typecheck cleanly, and are acyclic. However they land **disconnected from the barrel** — [index.ts](../../packages/docx-core/src/editor/helpers/index.ts) still lists all three under "Pending" and emits no `export *` for them, so none of their functions are reachable from `@extend-ai/docx-core`. This is a new, blocking integration defect (F1a).

Independently the prior blocking finding F1 persists in reduced form: **11 of 40 planned modules remain absent** (down from 14), including the pagination-plan trio, pretext-build/measure, table-height/line-height-table, tracked-changes pair, style-block-css, and xml-parsing-extra — all on the critical path and several carrying "必须复刻" alignment obligations (#22, #23, #32). `vue-docx` remains a stub (no `composables/` directory), so `docx-composables` still cannot start.

The check matrix is otherwise green: typecheck passes, build (tsup) succeeds, no circular deps (including the new orphan files verified independently), all imports relative, zero `@extend-ai/*` / React imports, zero stub/mock/fake/TODO. The delivered code quality remains high; the task is still incomplete.

**Conclusion: blocked** — F1a (new orphans unwired to barrel) + F1 (11 modules still missing, including all "必须复刻" critical-path modules).

## Findings

### F1a — P1 / blocking — 3 newly delivered modules are not exported from the barrel (NEW this cycle)

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (modules must be aggregated via `editor/helpers/index.ts`); §2.6 (`docx-core/src/index.ts` 聚合导出所有层)
- **Code:** [index.ts:20-28](../../packages/docx-core/src/editor/helpers/index.ts#L20) (Pending comment block still lists `selection-helpers`, `selection-restore`, `section-manipulation` as not-delivered); no `export * from "./selection-helpers"`, `./selection-restore`, or `./section-manipulation` exists anywhere in the barrel ([index.ts:30-56](../../packages/docx-core/src/editor/helpers/index.ts#L30) exports 27 modules, the three new ones absent)
- **Detail:** The three modules this cycle added are correct implementations but are **completely unreferenced**. Grep across `packages/docx-core/src/` finds zero `from "./selection-helpers"`, zero `from "./selection-restore"`, zero `from "./section-manipulation"` imports outside the files themselves. Because [docx-core/src/index.ts:26](../../packages/docx-core/src/index.ts#L26) re-exports only `./editor/helpers` (the barrel), none of `cloneEditorSelection`, `sameEditorSelection`, `normalizeEditorCursorStateForModel`, `shouldReissueDomSelectionRestore`, `updateSectionParagraphTextAtLocation`, `updateSectionImageFloatingAtLocation`, etc. are visible to any consumer of `@extend-ai/docx-core`. The files typecheck only because `tsconfig` `include: ["src"]` compiles every `.ts` file regardless of import graph — so this would not surface as a type error.

  This is a functional/integration defect, not a documentation nit: the modules exist on disk but deliver no reachable capability. The barrel also still says "modular split in progress" and explicitly lists the three as Pending, contradicting the files on disk.

- **Impact:** The selection/section-mutation helpers that downstream `docx-composables` needs (alignment #30 `historyRestoreRequest` nonce-driven restore, #31 `selectionSession`, section header/footer editing) are implemented but unreachable. A consumer importing `@extend-ai/docx-core` cannot use them. Must add the three `export *` lines to the barrel and update the Completed/Pending split.

### F1 — P1 / blocking — 11 of 40 planned helper modules still missing (reduced from 14 in #8)

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (40-file plan, ~21200 lines); [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §3 — "必须复刻" items #22, #23, #32
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 30 content files delivered (excl. barrel), ~14508 lines (note: +1 unplanned leaf `paragraph-toc.ts`)
- **Detail:** Delivered (30, excl. barrel): `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, page-measurement, xml-parsing, style-to-css, paragraph-inspect, text-mutation, numbering, table-utils, table-utils-extra, paragraph-toc, field-helpers, synthetic-textbox, paragraph-geometry, drop-cap, letterhead, line-height, header-footer, paragraph-tracked, selection-helpers, selection-restore, section-manipulation`.

  Missing (11, all within upstream editor.tsx lines 1-24953 — verified absent on disk this review; upstream line refs independently re-confirmed against editor.tsx at commit `6f70b92`):
  - `pretext-build.ts` (~900) — `buildParagraphPretextLayoutSource` (upstream editor.tsx:6586, confirmed present); alignment "必须复刻" #23
  - `pretext-measure.ts` (~900) — pretext measurement integration; alignment "必须复刻" #23
  - `line-height-table.ts` (~730) — table line-height estimation
  - `table-height.ts` (~510) — `estimateTableRowHeightsPx` (upstream editor.tsx:10850, confirmed present); backs the layout-layer callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138) and [page-segmentation-core.ts:682](../../packages/docx-core/src/layout/page-segmentation-core.ts#L682) — callback contract unbacked by a helper
  - `pagination-plan-core.ts` (~850) — `buildRenderColumnSegmentsForPageSection` (upstream editor.tsx:12466, confirmed present); alignment "必须复刻" #22
  - `pagination-plan-iterate.ts` (~850) — measurement-driven iterative pagination; alignment "必须复刻" #22
  - `pagination-plan-stabilize.ts` (~800) — pagination stabilization / oscillation detection
  - `style-block-css.ts` (~580) — `paragraphBlockStyle` (upstream editor.tsx:15433), `wrappedFloatingImageStyle` (upstream editor.tsx:14150), `absoluteFloatingImageStyle` (upstream editor.tsx:14351) — all confirmed present; needed by downstream `docx-render`
  - `xml-parsing-extra.ts` (~560) — XML parsing continuation (upstream 16049-17170 has 37 definitions; current `xml-parsing.ts` exports 18)
  - `tracked-changes.ts` (~550) — `collectTrackedChangesFromModel` (upstream editor.tsx:22770), `collectCommentsFromModel` (upstream editor.tsx:22908) — both confirmed present; alignment "必须复刻" #32
  - `tracked-changes-gutter.ts` (~550) — tracked-change gutter cards

- **Impact:** The pagination-plan trio, pretext-build/measure, table-height, and tracked-changes are the critical path for `docx-composables` (which depends-on `docx-helpers`). The layout layer already declares `estimateTableRowHeightsPx` as a callback parameter and invokes it, but the helpers layer that must supply the real implementation has not delivered it — the callback contract is unbacked. Alignment items #22 (measurement-driven iteration), #23 (pretext variable-width integration), and #32 (tracked-changes from-model derivation) are explicitly "必须复刻" and remain without implementations. `vue-docx` is still a stub (`packages/vue-docx/src/` has only `composables.ts` + `index.ts` + `env.d.ts`; no `composables/` directory), confirming `docx-composables` cannot start.

  Note on the batch model: the workflow doc allows docx-helpers to land in batches, and this is review #9 of an explicitly incremental task. The blocking verdict is scoped to the *task* (`docx-helpers`) not being completable until the remaining 11 modules land and the 3 orphans are wired. Nine successive reviews (#1–#9) have now blocked on missing modules; the deliverable surface grew 20 → 25 → 26 → 26 → 26 → 26 → 27 → 30 across batches #1–#8, with +3 this cycle. Progress resumed this cycle after the #7→#8 stall, but the critical-path modules (pagination-plan, pretext-build/measure, table-height, tracked-changes) remain.

### F2 — P3 / non-blocking — `paragraph-geometry.ts` exceeds the 1000-line hard constraint (carried from #1–#8)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1. The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000. Note `selection-helpers.ts` now sits exactly at the 1000-line cap — any growth there will also breach the constraint.
- **Impact:** Constraint violation; no functional impact.

### F3 — P3 / non-blocking — barrel + README still mark task "in progress" / pending; Pending list now stale and contradictory (carried from #3–#8, worsened)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md:44](../README.md#L44) task table status `pending`
- **Code:** [index.ts:1,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress", Pending comment block lists 17 modules — of which 3 are now on disk but unwired and 14 mix delivered-and-missing); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** The Pending comment block at [index.ts:20-28](../../packages/docx-core/src/editor/helpers/index.ts#L20) is now actively contradictory: it lists `selection-helpers`, `selection-restore`, `section-manipulation` as Pending while the files exist on disk, and also lists 14 modules that are genuinely missing alongside 14 that are delivered and exported (`header-footer`, `page-measurement`, `paragraph-inspect`, `paragraph-geometry`, `paragraph-tracked`, `drop-cap`, `letterhead`, `line-height`, `style-to-css`, `xml-parsing`, `synthetic-textbox`, `field-helpers`, `numbering`, `table-utils`, `table-utils-extra`, `text-mutation`, `state`). Once the 3 orphans are wired (F1a) and the remaining 11 land (F1), the barrel header, the Completed/Pending split, and the README status should all flip to `done`; in the meantime the Pending list should be pruned to the actually-missing 11.
- **Impact:** Scope tracking only.

### F4 — P3 / non-blocking — `style-to-css.ts` duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1–#8)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:10-19](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L10) (local duplicate declaration with comment referencing non-existent `tracked-changes.ts`), used downstream within the file
- **Detail:** Unchanged since batch 1. Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

### F5 — P3 / non-blocking — helpers import from the viewer layer, inverting the documented dependency direction (carried from #1–#8)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §5 — dependency graph states `viewer/{pretext-*,...} ← editor/helpers/*` (viewer depends on editor/helpers)
- **Code:** 13 helpers files import from `../../viewer/section-layout` (twipsToPixels, TWIPS_PER_PIXEL, DocumentLayoutMetrics) and [line-height.ts](../../packages/docx-core/src/editor/helpers/line-height.ts) imports from `../../viewer/pretext-layout`; [paragraph-tracked.ts:22](../../packages/docx-core/src/editor/helpers/paragraph-tracked.ts#L22) and [page-measurement.ts:28](../../packages/docx-core/src/editor/helpers/page-measurement.ts#L28) import from `../../layout/page-segmentation-core`; [section-manipulation.ts:14-17](../../packages/docx-core/src/editor/helpers/section-manipulation.ts#L14) imports from `../editor-ops` (sibling file in `editor/`)
- **Detail:** The documented dependency direction is `viewer → editor/helpers`, but the actual imports flow `editor/helpers → viewer` (and `→ layout`, `→ editor-ops`). This does not create a runtime circular dependency — `viewer` does not import back from `editor/helpers` (verified: zero `from "../../editor"` matches in `packages/docx-core/src/viewer/`), and `editor-ops` does not import from `helpers` (verified), so it is not blocking. The root cause is that `section-layout.ts` (`twipsToPixels` / `TWIPS_PER_PIXEL`) is a foundational unit-conversion utility mis-housed under `viewer/`; it belongs at a lower layer (engine or a shared `units` module). `madge --circular` confirms no cycle.
- **Impact:** None functionally (typecheck green, no cycle). Architectural tidiness only; would resolve naturally if `section-layout` constants were promoted to a lower layer.

## Positive checks (current tree — independently re-verified)

| Check | Result | Detail |
|---|---|---|
| New source since #8 | ✅ 3 modules | `git diff 4326037..HEAD --stat -- packages/docx-core/src/editor/helpers/` = `selection-helpers.ts` +1000, `selection-restore.ts` +135, `section-manipulation.ts` +370, `editor-types.ts` +72 (compare/normalize helpers). Progress resumed after the #7→#8 stall. HEAD `fb017f0`, working tree clean. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core typecheck` exits 0. All `.ts` under `src/` compiled (incl. orphans, since `include: ["src"]`). |
| Build (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core build` (tsup) succeeds — ESM + DTS. |
| Circular deps (barrel) | ✅ Pass | `madge --circular --extensions ts packages/docx-core/src/editor/helpers/index.ts` → "No circular dependency found!" (38 files processed; the 3 orphans not in this graph). |
| Circular deps (orphans) | ✅ Pass | `madge --circular` on the 3 orphan files directly → "No circular dependency found!" (13 files). |
| Import paths | ✅ Pass | All relative (`./` intra-helper, `../../engine/*` / `../../viewer/*` / `../../layout/*` / `../editor-ops` cross-layer). Zero `@extend-ai/*` package refs in source; zero `react` imports. |
| React-type cleanup | ✅ Pass | No `import * as React` / `from "react"` / `@extend-ai/*` in source. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|not.?implemented\|TODO\|FIXME\|XXX\|@ts-ignore\|@ts-expect-error` returns **zero** matches. `placeholder` finds only legitimate uses: form-field `placeholder` property ([text-mutation.ts:128](../../packages/docx-core/src/editor/helpers/text-mutation.ts#L128)) and an explanatory comment at [table-utils.ts:209](../../packages/docx-core/src/editor/helpers/table-utils.ts#L209). |
| Module boundaries | ✅ Pass | Intra-helper graph acyclic (DAG verified); `paragraph-toc` remains a true leaf introduced to break a cycle; helpers never import upward to vue-docx; `editor-ops` does not import back into `helpers`. |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over (F2). `selection-helpers.ts` sits exactly at 1000. |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`. New-module upstream line refs verified present this review: selection helpers @23538-24639 (cloneEditorSelection@24081, sameEditorSelection@24099, normalizeEditorCursorStateForModel@24595, etc.), selection-restore @24148-24259 (shouldReissueDomSelectionRestore@24148, isCollapsedSelectionAtElementStart@24208, isSuspiciousCollapsedSelectionAtElementStart@24247), section-manipulation @24641-24952 (sectionParagraphLocationKey@24641, updateSectionParagraphTextAtLocation@24794, updateSectionImageFloatingAtLocation@24879), and editor-types compare/normalize @3413-3480 (byte-for-byte match to upstream bodies). Missing-module upstream refs re-confirmed: buildParagraphPretextLayoutSource@6586, estimateTableRowHeightsPx@10850, buildRenderColumnSegmentsForPageSection@12466, paragraphBlockStyle@15433, wrappedFloatingImageStyle@14150, absoluteFloatingImageStyle@14351, collectTrackedChangesFromModel@22770, collectCommentsFromModel@22908. |
| Whitespace | ✅ Pass | Working tree clean. |

## Required fixes to unblock (next batch)

1. **F1a (immediate):** Wire the 3 delivered orphans into the barrel — add `export * from "./selection-helpers"`, `export * from "./selection-restore"`, `export * from "./section-manipulation"` to [index.ts](../../packages/docx-core/src/editor/helpers/index.ts); move them out of the Pending comment block into Completed. Verify no export-name collisions with existing barrel exports before/after.
2. **F1:** Deliver the 11 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height` + `line-height-table` (backs the layout-layer `estimateTableRowHeightsPx` callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138)), `tracked-changes` (#32), `style-block-css` (needed by `docx-render`), `xml-parsing-extra`, `tracked-changes-gutter`.
3. **F2** (optional): Split `paragraph-geometry.ts` to ≤1000 lines (cover-image cluster → `paragraph-geometry-cover.ts`); watch `selection-helpers.ts` (at 1000) for growth.
4. **F3** (cleanup alongside): Prune the barrel Pending comment block to the actually-missing modules now; flip barrel header + Completed/Pending list and README status to `done` once F1a + F1 land.
5. **F4** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts` and import the canonical declaration.
6. **F5** (optional, architectural): Promote `section-layout` unit constants (`TWIPS_PER_PIXEL`, `twipsToPixels`) to a lower layer to align with the documented dependency graph.

## Conclusion

**blocked**

F1a (new): the 3 modules delivered this cycle — `selection-helpers.ts`, `selection-restore.ts`, `section-manipulation.ts` — are correctly implemented and upstream-faithful but are **not exported from the barrel**, leaving them unreachable from `@extend-ai/docx-core`. F1 (carried, reduced): 11 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, and tracked-changes that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task (still a stub). The delivered code remains high quality — typecheck green, build succeeds, graph acyclic (incl. orphans), imports relative, zero stubs/mocks/TODOs, no React leakage, new modules byte-for-byte aligned to upstream editor.tsx. Re-verify after the 3 orphans are wired and the remaining 11 modules are delivered.
