# docx-helpers Review #3

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
- `a909c27 docx-helpers: fix #2` (batch 3 — this review's delta)

## Summary

Batch 3 (`a909c27`) correctly resolves the non-blocking circular-dependency regression flagged in review #2 (F2): the TOC predicates and PAGE/NUMPAGES/StyleRef field-instruction parsers are extracted into a new acyclic leaf module [paragraph-toc.ts](../../packages/docx-core/src/editor/helpers/paragraph-toc.ts), and the three former cycle participants (`field-helpers`, `line-height`, `synthetic-textbox`) now import them from that leaf. The extraction is byte-equivalent to upstream (editor.tsx:15864-15895 for TOC predicates, 17653-17676 for field-instruction parsers), and [field-helpers.ts](../../packages/docx-core/src/editor/helpers/field-helpers.ts) re-exports the symbols so existing barrel consumers are unaffected (no TS2308/TS2440 duplicate-export conflict — `tsc` dedupes the `export *` of identical symbols). Batch 3 also backfills a previously-missing real implementation `paragraphContainsExplicitLineBreakText` in [line-height.ts:257](../../packages/docx-core/src/editor/helpers/line-height.ts#L257), byte-equivalent to upstream editor.tsx:9157.

`madge --circular` on the helpers barrel now reports **No circular dependency found** (was 2 cycles in #2). typecheck passes for both `@extend-ai/docx-core` and the full 6-package workspace. All imports are relative; zero `@extend-ai/*` package references in source (one comment in the barrel documents the migration source). No residual stub/mock/fake — every `return 0` hit is a legitimate empty-input / no-border edge case, not a placeholder. 26 content modules now exist (25 from batch 2 + `paragraph-toc`), ~11057 lines.

**However, the blocking finding from review #2 (F1) is unchanged**: the 15 planned modules that carry "必须复刻" alignment obligations and are prerequisites for the downstream `docx-composables` task are still absent. Batch 3 delivered 1 new structural module (the `paragraph-toc` leaf) + 1 backfilled helper function, not the pagination-plan trio, pretext-build/measure, table-height, tracked-changes, or selection helpers. The barrel [index.ts](../../packages/docx-core/src/editor/helpers/index.ts) still says "modular split in progress" and still lists these 15 as Pending.

**Conclusion: blocked** — the cycle fix is clean and the new code is high quality, but the task remains ~65% complete by file count (26 of 40) and the missing 15 modules carry "必须复刻" obligations (#22 measurement-driven iteration, #23 pretext integration, #32 tracked-changes) and block `docx-composables`/`docx-render`.

## Findings

### F1 — P1 / blocking — 15 of 40 planned helper modules still missing (carried from #2, unchanged in batch 3)

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (40-file plan, ~21200 lines); [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §3 — "必须复刻" items #22, #23, #32
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 26 modules delivered (excl. barrel), ~11057 lines
- **Detail:** Delivered (26): `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, xml-parsing, style-to-css, paragraph-inspect, paragraph-toc, text-mutation, numbering, table-utils, table-utils-extra, field-helpers, synthetic-textbox, paragraph-geometry, drop-cap, letterhead, line-height, header-footer, paragraph-tracked`.

  Missing (15, all within upstream editor.tsx lines 1-24953, i.e. in-scope for this task — verified absent on disk):
  - `page-measurement.ts` (~740) — `resolveMeasuredPageContentHeightPx` (upstream line 2571)
  - `pretext-build.ts` (~900) — `buildParagraphPretextLayoutSource` (upstream line 6586); alignment "必须复刻" #23
  - `pretext-measure.ts` (~900) — pretext measurement integration; alignment #23
  - `line-height-table.ts` (~730) — table line-height estimation
  - `table-height.ts` (~510) — `estimateTableRowHeightsPx` (upstream line 10850)
  - `pagination-plan-core.ts` (~850) — `buildRenderColumnSegmentsForPageSection` (upstream line 12466)
  - `pagination-plan-iterate.ts` (~850) — measurement-driven iterative pagination; alignment "必须复刻" #22
  - `pagination-plan-stabilize.ts` (~800) — pagination stabilization / oscillation detection
  - `style-block-css.ts` (~580) — `paragraphBlockStyle` etc. (needed by `docx-render`)
  - `xml-parsing-extra.ts` (~560) — XML parsing continuation
  - `tracked-changes.ts` (~550) — `collectTrackedChangesFromModel` (alignment "必须复刻" #32)
  - `tracked-changes-gutter.ts` (~550) — tracked-change gutter cards
  - `selection-helpers.ts` (~605) — selection / caret helpers
  - `selection-restore.ts` (~605) — DOM selection restore
  - `section-manipulation.ts` (~310) — section operations

- **Impact:** The pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers are the critical path for `docx-composables` (which depends-on `docx-helpers`). Alignment items #22 (measurement-driven iteration), #23 (pretext variable-width integration), and #32 (tracked-changes from-model derivation) are explicitly "必须复刻" and remain without implementations. The barrel [index.ts:20-28](../../packages/docx-core/src/editor/helpers/index.ts#L20) still says "modular split in progress" and lists these 15 as Pending — honest, but confirms the task is not complete.

  Note on batch model: the workflow doc ([plan/workflow.md](../workflow.md)) does say docx-helpers is delivered "分批 5-6 个", and review #1/#2 explicitly allowed renegotiating into `docx-helpers-1..N` batches. Batch 3 is itself clean and is consistent with that model. The blocking verdict here is scoped to the *task* (`docx-helpers`) not being completable until the remaining 15 modules land. Three successive reviews (#1, #2, #3) have now blocked on this same F1; the deliverable surface has grown only 20 → 25 → 26 across batches, so the completion rate needs to accelerate materially to close the task.

### F2 — P3 / non-blocking — `paragraph-geometry.ts` still exceeds the 1000-line hard constraint (carried from #1/#2)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1 (review #1 F4 / #2 F3). No file in batch 3 crosses the ceiling (`paragraph-toc.ts` is 85 lines). The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000.
- **Impact:** Constraint violation; no functional impact.

### F3 — P3 / non-blocking — barrel + README still mark task "in progress" / pending

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md](../README.md) task table status `pending`
- **Code:** [index.ts:1,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress", 15 modules listed Pending); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** Honest documentation of partial state, consistent with the batch model. The barrel Pending list was not updated to reflect `paragraph-toc.ts` landing (it is re-exported at [index.ts:47](../../packages/docx-core/src/editor/helpers/index.ts#L47) but not listed under Completed). Once the remaining 15 modules land, the barrel header, the Completed/Pending split, and the README status should all flip to `done`.
- **Impact:** Scope tracking only.

### F4 — P3 / non-blocking — `style-to-css.ts` still duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1/#2)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:10-19](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L10-L19)
- **Detail:** Unchanged since batch 1 (review #1 F6 / #2 F5). Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

## Positive checks (this batch)

| Check | Result | Detail |
|---|---|---|
| Circular deps resolved (#2 F2) | ✅ Pass | `madge --circular` on [index.ts](../../packages/docx-core/src/editor/helpers/index.ts) → "No circular dependency found!". The 2 cycles from #2 are gone: `line-height.ts:49` and `synthetic-textbox.ts:18` now import `isTableOfContentsParagraph`/`instructionTextToPageFieldKind`/`instructionTextToStyleRefTarget` from the new `paragraph-toc` leaf, not from `field-helpers`. |
| paragraph-toc extraction upstream-faithful | ✅ Pass | `isTableOfContentsStyle`/`tableOfContentsLevel`/`isTableOfContentsParagraph` byte-equivalent to upstream editor.tsx:15864-15895; `instructionTextToPageFieldKind`/`instructionTextToStyleRefTarget` byte-equivalent to upstream 17653-17676. |
| Backward-compat re-exports | ✅ Pass | [field-helpers.ts:25-32](../../packages/docx-core/src/editor/helpers/field-helpers.ts#L25) re-exports the 5 symbols + `PageFieldKind` type from `paragraph-toc`; `paragraphPageFieldSequence` and the `PageFieldValueToken`/`StyleRefFieldValueToken` interfaces remain in field-helpers (not lost in extraction). No duplicate-export conflict — tsc dedupes identical `export *` symbols. |
| Backfilled real impl | ✅ Pass | `paragraphContainsExplicitLineBreakText` in [line-height.ts:257](../../packages/docx-core/src/editor/helpers/line-height.ts#L257) byte-equivalent to upstream editor.tsx:9157 (previously absent). |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core typecheck` exits 0. |
| Typecheck (full workspace) | ✅ Pass | `pnpm typecheck` (6 packages) exits 0 — no ripple in vue-docx/xlsx. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|placeholder\|forward declaration\|not implemented` finds only legitimate uses (form-field `placeholder` property at text-mutation.ts:128). All `return 0` hits are empty-input / no-border edge cases (drop-cap empty text, line-height empty text, paragraph-tracked empty sections, table-utils none/nil border), not stubs. |
| Import paths | ✅ Pass | All relative; zero `@extend-ai/*` package refs in source (only the barrel comment at index.ts:3 documents the migration source). Cross-layer imports only to `../../engine/*`, `../../viewer/*` (respects §5 DAG across layers). |
| Module boundaries | ✅ Pass | Layer DAG respected; intra-helper graph now acyclic. `paragraph-toc` is a true leaf (imports only `../../engine/types` + `./xml-parsing`). |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over; batch-3 files are well under (`paragraph-toc.ts` 85). |
| React-type cleanup | ✅ Pass | No `React.CSSProperties` / `import * as React` / `from "react"` in source. |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92`; module header line-number references match upstream editor.tsx. |
| Whitespace | ✅ Pass | `git diff --check` clean (exit 0). |

## Required fixes to unblock (next batch)

1. **F1**: Deliver the 15 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height`, `line-height-table`, `tracked-changes` (#32), `selection-helpers`/`selection-restore`, `style-block-css` (needed by `docx-render`), `section-manipulation`, `page-measurement`, `xml-parsing-extra`, `tracked-changes-gutter`.
2. **F2** (optional): Split `paragraph-geometry.ts` to ≤1000 lines (cover-image cluster → `paragraph-geometry-cover.ts`).
3. **F3** (cleanup alongside): Update the barrel Completed/Pending list to reflect `paragraph-toc.ts`; flip barrel header + README status to `done` once F1 lands.
4. **F4** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts` and import the canonical declaration.

## Conclusion

**blocked**

Batch 3 cleanly resolves the circular-dependency regression from review #2 (F2) via a byte-equivalent `paragraph-toc.ts` leaf extraction with backward-compatible re-exports, backfills a real missing helper, and keeps typecheck green, imports relative, and the graph acyclic. But the blocking finding F1 is unchanged across three reviews: 15 of 40 planned modules remain absent — including the pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task. Re-verify after the remaining 15 modules are delivered.
