# docx-helpers Review #4

Date: 2026-07-07
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
- `a909c27 docx-helpers: fix #2` (batch 3, reviewed in #3)
- `7cf6675 docx-helpers: fix #3` (this batch — adds review #3 doc only, **no helpers source changes**)

## Summary

No new helpers source landed since review #3. The only task-scoped commit since `a909c27` is `7cf6675`, which adds the review #3 markdown file and touches zero files under `packages/docx-core/src/editor/helpers/`. The deliverable surface is therefore identical to review #3: 26 content modules + barrel, 12241 lines (~12186 excl. barrel), against a plan of 40 content files / ~21200 lines.

The blocking finding F1 from reviews #1–#3 is unchanged: 15 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, line-height-table, tracked-changes/-gutter, selection-helpers/restore, section-manipulation, page-measurement, style-block-css, and xml-parsing-extra. Several of these carry "必须复刻" alignment obligations (#22 measurement-driven iteration, #23 pretext integration, #32 tracked-changes from-model) and are prerequisites for the downstream `docx-composables` task (still a stub — `vue-docx/src/` has no `composables/` directory).

This review re-verified the full check matrix against the current tree: typecheck passes for `@extend-ai/docx-core` and the full 6-package workspace; `madge --circular` reports no cycles; imports are all relative with zero `@extend-ai/*` package refs in source; no React imports; no residual stub/mock/fake (the only `placeholder` hits are the legitimate form-field `placeholder` property and an explanatory code comment about upstream generator behavior). The one structural concern carried from earlier reviews — `paragraph-geometry.ts` at 1027 lines (27 over the ≤1000 hard constraint) — is also unchanged.

The quality of what *is* delivered remains high and the graph is clean, but the task is still ~65% complete by file count (26 of 40) and the missing modules are on the critical path. No progress was made on the blocking finding between review #3 and #4.

**Conclusion: blocked** — identical blocking state to review #3; F1 unchanged.

## Findings

### F1 — P1 / blocking — 15 of 40 planned helper modules still missing (carried from #3, unchanged)

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (40-file plan, ~21200 lines); [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §3 — "必须复刻" items #22, #23, #32
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 26 modules delivered (excl. barrel), ~12186 lines
- **Detail:** Delivered (26): `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, xml-parsing, style-to-css, paragraph-inspect, paragraph-toc, text-mutation, numbering, table-utils, table-utils-extra, field-helpers, synthetic-textbox, paragraph-geometry, drop-cap, letterhead, line-height, header-footer, paragraph-tracked`.

  Missing (15, all within upstream editor.tsx lines 1-24953 — verified absent on disk; upstream line refs confirmed present in editor.tsx):
  - `page-measurement.ts` (~740) — `resolveMeasuredPageContentHeightPx` (upstream editor.tsx:2571)
  - `pretext-build.ts` (~900) — `buildParagraphPretextLayoutSource` (upstream editor.tsx:6586); alignment "必须复刻" #23
  - `pretext-measure.ts` (~900) — pretext measurement integration; alignment #23
  - `line-height-table.ts` (~730) — table line-height estimation
  - `table-height.ts` (~510) — `estimateTableRowHeightsPx` (upstream editor.tsx:10850)
  - `pagination-plan-core.ts` (~850) — `buildRenderColumnSegmentsForPageSection` (upstream editor.tsx:12466)
  - `pagination-plan-iterate.ts` (~850) — measurement-driven iterative pagination; alignment "必须复刻" #22
  - `pagination-plan-stabilize.ts` (~800) — pagination stabilization / oscillation detection
  - `style-block-css.ts` (~580) — `paragraphBlockStyle` etc. (needed by `docx-render`)
  - `xml-parsing-extra.ts` (~560) — XML parsing continuation
  - `tracked-changes.ts` (~550) — `collectTrackedChangesFromModel` (alignment "必须复刻" #32)
  - `tracked-changes-gutter.ts` (~550) — tracked-change gutter cards
  - `selection-helpers.ts` (~605) — selection / caret helpers
  - `selection-restore.ts` (~605) — DOM selection restore
  - `section-manipulation.ts` (~310) — section operations

- **Impact:** The pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers are the critical path for `docx-composables` (which depends-on `docx-helpers`). The layout layer already declares `estimateTableRowHeightsPx` as a callback parameter ([page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138), [page-segmentation-core.ts:682](../../packages/docx-core/src/layout/page-segmentation-core.ts#L682)) and invokes it, but the helpers layer that must supply the real implementation has not delivered it — so the callback contract is unbacked by a helper. Alignment items #22 (measurement-driven iteration), #23 (pretext variable-width integration), and #32 (tracked-changes from-model derivation) are explicitly "必须复刻" and remain without implementations. `vue-docx` is still a stub (`packages/vue-docx/src/` has no `composables/` directory; only `composables.ts`/`index.ts`), confirming `docx-composables` cannot start. The barrel [index.ts:20-28](../../packages/docx-core/src/editor/helpers/index.ts#L20) still says "modular split in progress" and lists these 15 as Pending — honest, but confirms the task is not complete.

  Note on the batch model: the workflow doc allows docx-helpers to land in batches, and this is review #4 of an explicitly incremental task. The blocking verdict is scoped to the *task* (`docx-helpers`) not being completable until the remaining 15 modules land. Four successive reviews (#1–#4) have now blocked on this same F1; the deliverable surface has grown 20 → 25 → 26 → 26 across batches, with **zero growth between #3 and #4** (the intervening commit was doc-only). The completion rate needs to accelerate materially to close the task.

### F2 — P3 / non-blocking — `paragraph-geometry.ts` still exceeds the 1000-line hard constraint (carried from #1/#2/#3)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1 (review #1 F4 / #2 F3 / #3 F2). No helpers file crossed the ceiling this batch (no source landed at all). The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000.
- **Impact:** Constraint violation; no functional impact.

### F3 — P3 / non-blocking — barrel + README still mark task "in progress" / pending (carried from #3)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md:44](../README.md#L44) task table status `pending`
- **Code:** [index.ts:1,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress", 15 modules listed Pending); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** Honest documentation of partial state, consistent with the batch model. Once the remaining 15 modules land, the barrel header, the Completed/Pending split, and the README status should all flip to `done`.
- **Impact:** Scope tracking only.

### F4 — P3 / non-blocking — `style-to-css.ts` still duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1/#2/#3)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:10-19](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L10-L19)
- **Detail:** Unchanged since batch 1 (review #1 F6 / #2 F5 / #3 F4). Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

## Positive checks (this batch / current tree)

| Check | Result | Detail |
|---|---|---|
| No new source since #3 | ✅ confirmed | `git log -- packages/docx-core/src/editor/helpers/` shows last source commit is `a909c27`; `7cf6675` touches only `docs/plan/reviews/docx-helpers-3.md`. Re-verification run against current HEAD `7cf6675`. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core typecheck` exits 0. |
| Typecheck (full workspace) | ✅ Pass | `pnpm typecheck` (6 packages: docx-core, xlsx-core, vue-extend, vue-docx, vue-xlsx, +root) exits 0 — no ripple. |
| Circular deps | ✅ Pass | `madge --circular --extensions ts packages/docx-core/src/editor/helpers/index.ts` → "No circular dependency found!". Intra-helper import graph (top hubs: constants×13, editor-types×10, paragraph-inspect×9) is acyclic. |
| Import paths | ✅ Pass | All relative (`./` intra-helper, `../../engine/*` / `../../viewer/*` cross-layer). Zero `@extend-ai/*` package refs in source. |
| React-type cleanup | ✅ Pass | No `import * as React` / `from "react"` in source; `React.CSSProperties` only mentioned in comments documenting the replacement. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|placeholder\|not implemented\|forward declaration` finds only legitimate uses: form-field `placeholder` property ([text-mutation.ts:128](../../packages/docx-core/src/editor/helpers/text-mutation.ts#L128)) and an explanatory comment at [table-utils.ts:209](../../packages/docx-core/src/editor/helpers/table-utils.ts#L209). All `return 0` hits are empty-input / no-border edge cases, not stubs. |
| Module boundaries | ✅ Pass | Layer DAG respected (helpers import only engine + viewer, never upward to vue-docx); intra-helper graph acyclic; `paragraph-toc` remains a true leaf. |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over (F2). |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`; missing-module upstream line refs verified present in editor.tsx (e.g. `resolveMeasuredPageContentHeightPx`@2571, `buildParagraphPretextLayoutSource`@6586, `estimateTableRowHeightsPx`@10850, `buildRenderColumnSegmentsForPageSection`@12466). |
| Whitespace | ✅ Pass | `git diff --check` clean (exit 0). |

## Required fixes to unblock (next batch)

1. **F1**: Deliver the 15 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height` + `line-height-table` (backs the layout-layer `estimateTableRowHeightsPx` callback), `tracked-changes` (#32), `selection-helpers`/`selection-restore`, `style-block-css` (needed by `docx-render`), `section-manipulation`, `page-measurement`, `xml-parsing-extra`, `tracked-changes-gutter`.
2. **F2** (optional): Split `paragraph-geometry.ts` to ≤1000 lines (cover-image cluster → `paragraph-geometry-cover.ts`).
3. **F3** (cleanup alongside): Flip barrel header + Completed/Pending list and README status to `done` once F1 lands.
4. **F4** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts` and import the canonical declaration.

## Conclusion

**blocked**

Identical blocking state to review #3: no helpers source landed since `a909c27` (the intervening commit `7cf6675` was doc-only). The blocking finding F1 is unchanged across four reviews — 15 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task (still a stub). The delivered code remains high quality — typecheck green (6 packages), graph acyclic, imports relative, no stubs, no React leakage — but the task is ~65% complete by file count with zero progress this batch. Re-verify after the remaining 15 modules are delivered.
