# docx-helpers Review #5

Date: 2026-07-08
Reviewer: automated
Scope: `packages/docx-core/src/editor/helpers/` (step 5 of migration execution order)
Upstream: `@extend-ai/react-docx` @ commit `6f70b92` (verified: `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`), `react-viewer/src/editor.tsx` lines 1-24953 (pure functions / types / constants region, zero React hooks)
Design refs:
- [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — editor/helpers module plan (~21200 lines, 40 content files + barrel)
- [docx-editor-helpers-split-plan.md](../docx-editor-helpers-split-plan.md) — 24-module detailed split
- [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §2.5 — editor.tsx 前半部分 pure-function region; §3 "必须复刻" alignment items #22, #23, #32

Commits under review (cumulative task state):
- `3a4ad69 docx-helpers: docx-core editor/helpers` (batch 1, reviewed in #1)
- `e24a049 docx-helpers: fix #1` (batch 2, reviewed in #2)
- `a909c27 docx-helpers: fix #2` (batch 3, reviewed in #3 — last source commit)
- `7cf6675 docx-helpers: fix #3` (doc-only, review #3)
- `2659ea0 docx-helpers: fix #4` (doc-only, review #4)

Working tree clean; HEAD `2659ea0`.

## Summary

No helpers source has landed since `a909c27` (review #3). The two intervening commits `7cf6675` and `2659ea0` are doc-only (review markdown files). The deliverable surface is therefore identical to reviews #3 and #4: 26 content modules + barrel, 12241 lines (~12186 excl. barrel), against a plan of 40 content files / ~21200 lines.

The blocking finding F1 is unchanged across reviews #1–#5: 15 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, line-height-table, tracked-changes/-gutter, selection-helpers/restore, section-manipulation, page-measurement, style-block-css, and xml-parsing-extra. Several carry "必须复刻" alignment obligations (#22 measurement-driven iteration, #23 pretext integration, #32 tracked-changes from-model) and are prerequisites for the downstream `docx-composables` task (still a stub — `vue-docx/src/` has no `composables/` directory).

This review independently re-ran the full check matrix against the current tree (not copied from #4): typecheck passes for `@extend-ai/docx-core` and the full 6-package workspace; `madge --circular` reports no cycles (37 files processed); imports are all relative with zero `@extend-ai/*` package refs in source; no React imports; no residual stub/mock/fake/TODO/FIXME (grep for `stub|mock|fake|not.?implemented|TODO|FIXME|XXX` returns zero matches; the only `placeholder` hits are the legitimate form-field `placeholder` property and an explanatory code comment). The structural concern `paragraph-geometry.ts` at 1027 lines (27 over the ≤1000 hard constraint) is also unchanged.

The quality of what *is* delivered remains high and the dependency graph is clean, but the task is still ~65% complete by file count (26 of 40) and the missing modules are on the critical path. No progress was made on the blocking finding across the last two review cycles (#4 and #5).

**Conclusion: blocked** — identical blocking state to reviews #3 and #4; F1 unchanged.

## Findings

### F1 — P1 / blocking — 15 of 40 planned helper modules still missing (carried from #3/#4, unchanged)

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (40-file plan, ~21200 lines); [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §3 — "必须复刻" items #22, #23, #32
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 26 modules delivered (excl. barrel), ~12186 lines
- **Detail:** Delivered (26): `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, xml-parsing, style-to-css, paragraph-inspect, paragraph-toc, text-mutation, numbering, table-utils, table-utils-extra, field-helpers, synthetic-textbox, paragraph-geometry, drop-cap, letterhead, line-height, header-footer, paragraph-tracked`.

  Missing (15, all within upstream editor.tsx lines 1-24953 — verified absent on disk; upstream line refs confirmed present in editor.tsx this review):
  - `page-measurement.ts` (~740) — `resolveMeasuredPageContentHeightPx` (upstream editor.tsx:2571)
  - `pretext-build.ts` (~900) — `buildParagraphPretextLayoutSource` (upstream editor.tsx:6586); alignment "必须复刻" #23
  - `pretext-measure.ts` (~900) — pretext measurement integration; alignment #23
  - `line-height-table.ts` (~730) — table line-height estimation
  - `table-height.ts` (~510) — `estimateTableRowHeightsPx` (upstream editor.tsx:10850)
  - `pagination-plan-core.ts` (~850) — `buildRenderColumnSegmentsForPageSection` (upstream editor.tsx:12466)
  - `pagination-plan-iterate.ts` (~850) — measurement-driven iterative pagination; alignment "必须复刻" #22
  - `pagination-plan-stabilize.ts` (~800) — pagination stabilization / oscillation detection
  - `style-block-css.ts` (~580) — `paragraphBlockStyle` etc. (upstream editor.tsx:15433; needed by `docx-render`)
  - `xml-parsing-extra.ts` (~560) — XML parsing continuation
  - `tracked-changes.ts` (~550) — `collectTrackedChangesFromModel` (upstream editor.tsx:22770; alignment "必须复刻" #32)
  - `tracked-changes-gutter.ts` (~550) — tracked-change gutter cards
  - `selection-helpers.ts` (~605) — selection / caret helpers
  - `selection-restore.ts` (~605) — DOM selection restore
  - `section-manipulation.ts` (~310) — section operations

- **Impact:** The pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers are the critical path for `docx-composables` (which depends-on `docx-helpers`). The layout layer already declares `estimateTableRowHeightsPx` as a callback parameter and invokes it ([page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138), [page-segmentation-core.ts:682](../../packages/docx-core/src/layout/page-segmentation-core.ts#L682)), but the helpers layer that must supply the real implementation has not delivered it — the callback contract is unbacked by a helper. Alignment items #22 (measurement-driven iteration), #23 (pretext variable-width integration), and #32 (tracked-changes from-model derivation) are explicitly "必须复刻" and remain without implementations. `vue-docx` is still a stub (`packages/vue-docx/src/` has no `composables/` directory; only `composables.ts`/`index.ts`), confirming `docx-composables` cannot start. The barrel [index.ts:20-28](../../packages/docx-core/src/editor/helpers/index.ts#L20) still says "modular split in progress" and lists these 15 as Pending — honest, but confirms the task is not complete.

  Note on the batch model: the workflow doc allows docx-helpers to land in batches, and this is review #5 of an explicitly incremental task. The blocking verdict is scoped to the *task* (`docx-helpers`) not being completable until the remaining 15 modules land. Five successive reviews (#1–#5) have now blocked on this same F1; the deliverable surface has grown 20 → 25 → 26 → 26 → 26 across batches, with **zero growth across #3, #4, and #5** (the intervening commits were doc-only). The completion rate has stalled; the remaining 15 modules must land to close the task.

### F2 — P3 / non-blocking — `paragraph-geometry.ts` still exceeds the 1000-line hard constraint (carried from #1–#4)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1. No helpers file crossed the ceiling this batch (no source landed at all). The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000.
- **Impact:** Constraint violation; no functional impact.

### F3 — P3 / non-blocking — barrel + README still mark task "in progress" / pending (carried from #3/#4)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md:44](../README.md#L44) task table status `pending`
- **Code:** [index.ts:1,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress", 15 modules listed Pending); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** Honest documentation of partial state, consistent with the batch model. Once the remaining 15 modules land, the barrel header, the Completed/Pending split, and the README status should all flip to `done`.
- **Impact:** Scope tracking only.

### F4 — P3 / non-blocking — `style-to-css.ts` still duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1–#4)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:13](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L13) (local duplicate declaration), used at [style-to-css.ts:227](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L227)
- **Detail:** Unchanged since batch 1. Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

## Positive checks (this batch / current tree — independently re-verified)

| Check | Result | Detail |
|---|---|---|
| No new source since #3 | ✅ confirmed | `git log --oneline -- packages/docx-core/src/editor/helpers/` shows last source commit is `a909c27`; `7cf6675` and `2659ea0` touch only review markdown. Re-verification run against current HEAD `2659ea0`, working tree clean. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core typecheck` exits 0. |
| Typecheck (full workspace) | ✅ Pass | `pnpm typecheck` (6 packages: docx-core, xlsx-core, vue-extend, vue-docx, vue-xlsx, +root) exits 0 — no ripple. |
| Circular deps | ✅ Pass | `madge --circular --extensions ts packages/docx-core/src/editor/helpers/index.ts` → "No circular dependency found!" (37 files processed). Intra-helper import graph is acyclic. |
| Import paths | ✅ Pass | All relative (`./` intra-helper, `../../engine/*` / `../../viewer/*` cross-layer). Zero `@extend-ai/*` package refs in source. |
| React-type cleanup | ✅ Pass | No `import * as React` / `from "react"` / `@extend-ai/*` in source. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|not.?implemented\|TODO\|FIXME\|XXX` returns **zero** matches. `placeholder` finds only legitimate uses: form-field `placeholder` property ([text-mutation.ts:128](../../packages/docx-core/src/editor/helpers/text-mutation.ts#L128)) and an explanatory comment at [table-utils.ts:209](../../packages/docx-core/src/editor/helpers/table-utils.ts#L209). |
| Module boundaries | ✅ Pass | Layer DAG respected (helpers import only engine + viewer, never upward to vue-docx); intra-helper graph acyclic; `paragraph-toc` remains a true leaf. |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over (F2). |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`; missing-module upstream line refs verified present in editor.tsx this review (`resolveMeasuredPageContentHeightPx`@2571, `buildParagraphPretextLayoutSource`@6586, `estimateTableRowHeightsPx`@10850, `buildRenderColumnSegmentsForPageSection`@12466, `paragraphBlockStyle`@15433, `collectTrackedChangesFromModel`@22770). |
| Whitespace | ✅ Pass | `git diff --check` clean (working tree clean). |

## Required fixes to unblock (next batch)

1. **F1**: Deliver the 15 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height` + `line-height-table` (backs the layout-layer `estimateTableRowHeightsPx` callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138)), `tracked-changes` (#32), `selection-helpers`/`selection-restore`, `style-block-css` (needed by `docx-render`), `section-manipulation`, `page-measurement`, `xml-parsing-extra`, `tracked-changes-gutter`.
2. **F2** (optional): Split `paragraph-geometry.ts` to ≤1000 lines (cover-image cluster → `paragraph-geometry-cover.ts`).
3. **F3** (cleanup alongside): Flip barrel header + Completed/Pending list and README status to `done` once F1 lands.
4. **F4** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts` and import the canonical declaration.

## Conclusion

**blocked**

Identical blocking state to reviews #3 and #4: no helpers source has landed since `a909c27` (the two intervening commits `7cf6675` and `2659ea0` were doc-only). The blocking finding F1 is unchanged across five reviews — 15 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task (still a stub). The delivered code remains high quality — typecheck green (6 packages), graph acyclic, imports relative, zero stubs/mocks/TODOs, no React leakage — but the task is ~65% complete by file count with zero progress across the last two review cycles. Re-verify after the remaining 15 modules are delivered.
