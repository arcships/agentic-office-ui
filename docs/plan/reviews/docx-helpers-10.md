# docx-helpers Review #10

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
- `fb017f0 docx-helpers: fix #8` (added `selection-helpers.ts`, `selection-restore.ts`, `section-manipulation.ts` + editor-types compare/normalize helpers; reviewed in #9)
- `207145d docx-helpers: fix #9` (this cycle — barrel wiring only)

HEAD is `207145d`; working tree clean. Since review #9, **only the barrel was edited** — `207145d` changed 2 files: [index.ts](../../packages/docx-core/src/editor/helpers/index.ts) (+`export *` for the 3 #9 orphans, header/Pending comment block updated) and the `docx-helpers-9.md` review doc. **Zero new helper modules landed this cycle.** The deliverable surface remains 30 content modules + 1 barrel (~14527 lines excl. barrel). Against the 40-file plan the task is ~75% complete by file count (30 of 40), unchanged from #9.

## Summary

Review #9 blocked on two P1 findings: F1a (3 newly delivered orphans unwired to the barrel) and F1 (11 of 40 planned modules missing). This cycle resolved **F1a**: the barrel now exports all 30 on-disk modules via `export *`, the "Completed" comment block lists 30, and the "Pending" block is pruned to the actually-missing 11. `selection-helpers`, `selection-restore`, `section-manipulation` are now reachable from `@extend-ai/docx-core`.

F1 persists **with zero progress this cycle**: the same 11 modules remain absent — `pretext-build`, `pretext-measure`, `line-height-table`, `table-height`, `pagination-plan-core`, `pagination-plan-iterate`, `pagination-plan-stabilize`, `style-block-css`, `xml-parsing-extra`, `tracked-changes`, `tracked-changes-gutter`. All were re-confirmed absent on disk. Several carry "必须复刻" alignment obligations (#22 measurement-driven iteration, #23 pretext variable-width integration, #32 tracked-changes from-model derivation) and are on the critical path: the layout layer declares and invokes `estimateTableRowHeightsPx` as a callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138) and [page-segmentation-core.ts:682](../../packages/docx-core/src/layout/page-segmentation-core.ts#L682), but the helpers layer that must supply the real implementation has not delivered it — the callback contract remains unbacked. `vue-docx` remains a stub (`packages/vue-docx/src/` has only `composables.ts` + `index.ts` + `env.d.ts`; no `composables/` directory), so `docx-composables` still cannot start.

The check matrix is otherwise green: typecheck passes, build (tsup) succeeds, no circular deps (barrel processes 42 files), all imports relative, zero `@extend-ai/*` / React imports in source, zero stub/mock/fake/TODO. The delivered code quality remains high; the task is still incomplete.

**Conclusion: blocked** — F1 (11 modules still missing, including all "必须复刻" critical-path modules; no progress this cycle).

## Findings

### F1 — P1 / blocking — 11 of 40 planned helper modules still missing (unchanged from #9)

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

  Note on the batch model: the workflow doc allows docx-helpers to land in batches, and this is review #10 of an explicitly incremental task. The blocking verdict is scoped to the *task* (`docx-helpers`) not being completable until the remaining 11 modules land. Ten successive reviews (#1–#10) have now blocked on missing modules; the deliverable surface grew 20 → 25 → 26 → 26 → 26 → 26 → 27 → 30 → 30 across batches #1–#9, with **+0 modules this cycle** (barrel wiring only). The critical-path modules (pagination-plan, pretext-build/measure, table-height, tracked-changes) remain untouched across all ten cycles.

### F2 — P3 / non-blocking — `paragraph-geometry.ts` exceeds the 1000-line hard constraint (carried from #1–#9)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1. The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000. Note `selection-helpers.ts` sits exactly at the 1000-line cap — any growth there will also breach the constraint.
- **Impact:** Constraint violation; no functional impact.

### F3 — P3 / non-blocking — README still marks task `pending`; barrel header still says "in progress" (carried from #3–#9, partially improved)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md:44](../README.md#L44) task table status `pending`
- **Code:** [index.ts:1](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress"); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** This cycle improved the barrel's internal tracking — the Pending comment block at [index.ts:41-44](../../packages/docx-core/src/editor/helpers/index.ts#L41) now correctly lists only the 11 actually-missing modules (F1a's stale/contradictory state from #9 is resolved), and the Completed block lists 30. However the barrel header still says "modular split in progress" and the README task table still shows `pending`. Once the remaining 11 land (F1), the barrel header, Completed/Pending split, and README status should all flip to `done`.
- **Impact:** Scope tracking only.

### F4 — P3 / non-blocking — `style-to-css.ts` duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1–#9)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:10-19](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L10) (local duplicate declaration with comment referencing non-existent `tracked-changes.ts`), used downstream within the file
- **Detail:** Unchanged since batch 1. Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

### F5 — P3 / non-blocking — helpers import from the viewer layer, inverting the documented dependency direction (carried from #1–#9)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §5 — dependency graph states `viewer/{pretext-*,...} ← editor/helpers/*` (viewer depends on editor/helpers)
- **Code:** 13 helpers files import from `../../viewer/section-layout` (twipsToPixels, TWIPS_PER_PIXEL, DocumentLayoutMetrics) and [line-height.ts](../../packages/docx-core/src/editor/helpers/line-height.ts) imports from `../../viewer/pretext-layout`; [paragraph-tracked.ts:22](../../packages/docx-core/src/editor/helpers/paragraph-tracked.ts#L22) and [page-measurement.ts:28](../../packages/docx-core/src/editor/helpers/page-measurement.ts#L28) import from `../../layout/page-segmentation-core`; [section-manipulation.ts:14-17](../../packages/docx-core/src/editor/helpers/section-manipulation.ts#L14) imports from `../editor-ops` (sibling file in `editor/`)
- **Detail:** The documented dependency direction is `viewer → editor/helpers`, but the actual imports flow `editor/helpers → viewer` (and `→ layout`, `→ editor-ops`). This does not create a runtime circular dependency — `viewer` does not import back from `editor/helpers` (verified: zero `from "../../editor"` matches in `packages/docx-core/src/viewer/`), and `editor-ops` does not import from `helpers` (verified), so it is not blocking. The root cause is that `section-layout.ts` (`twipsToPixels` / `TWIPS_PER_PIXEL`) is a foundational unit-conversion utility mis-housed under `viewer/`; it belongs at a lower layer (engine or a shared `units` module). `madge --circular` confirms no cycle.
- **Impact:** None functionally (typecheck green, no cycle). Architectural tidiness only; would resolve naturally if `section-layout` constants were promoted to a lower layer.

## Positive checks (current tree — independently re-verified)

| Check | Result | Detail |
|---|---|---|
| F1a resolution (this cycle) | ✅ Fixed | The 3 #9 orphans (`selection-helpers`, `selection-restore`, `section-manipulation`) are now exported via `export *` at [index.ts:73-75](../../packages/docx-core/src/editor/helpers/index.ts#L73); Completed block lists 30, Pending block pruned to 11. Barrel's stale/contradictory state resolved. |
| New source since #9 | ⚠️ Barrel-only | `git diff fb017f0..HEAD --stat` = `index.ts` +38/-19 (barrel wiring + comment update) + `docx-helpers-9.md` review doc. Zero new helper modules. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core typecheck` exits 0. All `.ts` under `src/` compiled. |
| Build (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core build` (tsup) succeeds — ESM (483.77 KB) + DTS (161.83 KB). |
| Circular deps (barrel) | ✅ Pass | `madge --circular --extensions ts packages/docx-core/src/editor/helpers/index.ts` → "No circular dependency found!" (42 files processed). |
| Import paths | ✅ Pass | All relative (`./` intra-helper, `../../engine/*` / `../../viewer/*` / `../../layout/*` / `../editor-ops` cross-layer). Zero `@extend-ai/*` package refs in source (the one `@extend-ai/react-docx` match is a header *comment* in [index.ts:3](../../packages/docx-core/src/editor/helpers/index.ts#L3), not an import); zero `react` imports. |
| React-type cleanup | ✅ Pass | No `import * as React` / `from "react"` / `@extend-ai/*` in source. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|not.?implemented\|TODO\|FIXME\|XXX\|@ts-ignore\|@ts-expect-error` returns **zero** matches. `placeholder` finds only legitimate uses: form-field `placeholder` property ([text-mutation.ts:128](../../packages/docx-core/src/editor/helpers/text-mutation.ts#L128)) and an explanatory comment at [table-utils.ts:209](../../packages/docx-core/src/editor/helpers/table-utils.ts#L209). |
| Module boundaries | ✅ Pass | Intra-helper graph acyclic (DAG verified, 42 files); `paragraph-toc` remains a true leaf introduced to break a cycle; helpers never import upward to vue-docx; `editor-ops` does not import back into `helpers`. |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over (F2). `selection-helpers.ts` sits exactly at 1000. |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`. Missing-module upstream refs re-confirmed present this review: buildParagraphPretextLayoutSource@6586, estimateTableRowHeightsPx@10850, buildRenderColumnSegmentsForPageSection@12466, paragraphBlockStyle@15433, wrappedFloatingImageStyle@14150, absoluteFloatingImageStyle@14351, collectTrackedChangesFromModel@22770, collectCommentsFromModel@22908. |
| Whitespace | ✅ Pass | Working tree clean. |

## Required fixes to unblock (next batch)

1. **F1:** Deliver the 11 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height` + `line-height-table` (backs the layout-layer `estimateTableRowHeightsPx` callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138)), `tracked-changes` (#32), `style-block-css` (needed by `docx-render`), `xml-parsing-extra`, `tracked-changes-gutter`.
2. **F2** (optional): Split `paragraph-geometry.ts` to ≤1000 lines (cover-image cluster → `paragraph-geometry-cover.ts`); watch `selection-helpers.ts` (at 1000) for growth.
3. **F3** (cleanup alongside): Flip barrel header ("modular split in progress" → done) and [plan/README.md:44](../README.md#L44) status to `done` once F1 lands.
4. **F4** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts` and import the canonical declaration.
5. **F5** (optional, architectural): Promote `section-layout` unit constants (`TWIPS_PER_PIXEL`, `twipsToPixels`) to a lower layer to align with the documented dependency graph.

## Conclusion

**blocked**

This cycle resolved F1a (the 3 #9 orphans are now wired to the barrel and reachable from `@extend-ai/docx-core`), but delivered **zero new helper modules**. F1 persists unchanged: 11 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, and tracked-changes that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task (still a stub). The layout layer's `estimateTableRowHeightsPx` callback contract remains unbacked. The delivered code remains high quality — typecheck green, build succeeds, graph acyclic (42 files), imports relative, zero stubs/mocks/TODOs, no React leakage, upstream refs re-confirmed. Re-verify after the remaining 11 modules are delivered.
