# docx-helpers Review #7

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
- `a909c27 docx-helpers: fix #2` (batch 3, reviewed in #3)
- `62dba91 docx-helpers: fix #6` (batch 4 — added `page-measurement.ts`, 689 lines; reviewed this cycle)

HEAD is `62dba91`; working tree clean. Since review #6, one new helpers source file landed (`page-measurement.ts`), reducing missing modules from 15 → 14. The deliverable surface is now 27 content modules + barrel, ~13620 lines (excl. barrel), against a plan of 40 content files / ~21200 lines.

## Summary

Review #6 blocked on 15 missing modules. This cycle added `page-measurement.ts` (689 lines, 15 exports including the upstream `resolveMeasuredPageContentHeightPx`@2571), closing one module and advancing the task from 26 → 27 of 40 content modules. The blocking finding F1 persists: 14 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task (still a stub).

Independently re-ran the full check matrix against the current tree: typecheck passes for `@extend-ai/docx-core`; build (tsup) succeeds; `madge --circular` reports no cycles (38 files processed); all imports relative with zero `@extend-ai/*` package refs; no React imports; no residual stub/mock/fake/TODO/FIXME. The quality of what *is* delivered remains high and the dependency graph is clean, but the task is ~68% complete by file count (27 of 40) and the missing modules are on the critical path.

**Conclusion: blocked** — F1 persists (14 of 40 modules missing); one module closed since #6.

## Findings

### F1 — P1 / blocking — 14 of 40 planned helper modules still missing (carried from #3–#6, reduced from 15)

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (40-file plan, ~21200 lines); [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §3 — "必须复刻" items #22, #23, #32
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 27 modules delivered (excl. barrel), ~13620 lines
- **Detail:** Delivered (27): `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, page-measurement, xml-parsing, style-to-css, paragraph-inspect, text-mutation, numbering, table-utils, table-utils-extra, paragraph-toc, field-helpers, synthetic-textbox, paragraph-geometry, drop-cap, letterhead, line-height, header-footer, paragraph-tracked`.

  Missing (14, all within upstream editor.tsx lines 1-24953 — verified absent on disk this review; upstream line refs independently re-confirmed against editor.tsx at commit `6f70b92`):
  - `pretext-build.ts` (~900) — `buildParagraphPretextLayoutSource` (upstream editor.tsx:6586, confirmed present); alignment "必须复刻" #23
  - `pretext-measure.ts` (~900) — pretext measurement integration; alignment #23
  - `line-height-table.ts` (~730) — table line-height estimation
  - `table-height.ts` (~510) — `estimateTableRowHeightsPx` (upstream editor.tsx:10850, confirmed present); backs the layout-layer callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138) and [page-segmentation-core.ts:682](../../packages/docx-core/src/layout/page-segmentation-core.ts#L682) — callback contract unbacked by a helper
  - `pagination-plan-core.ts` (~850) — `buildRenderColumnSegmentsForPageSection` (upstream editor.tsx:12466, confirmed present); alignment "必须复刻" #22
  - `pagination-plan-iterate.ts` (~850) — measurement-driven iterative pagination; alignment "必须复刻" #22
  - `pagination-plan-stabilize.ts` (~800) — pagination stabilization / oscillation detection
  - `style-block-css.ts` (~580) — `paragraphBlockStyle` (upstream editor.tsx:15433), `wrappedFloatingImageStyle` (upstream editor.tsx:14150), `absoluteFloatingImageStyle` (upstream editor.tsx:14351) — all confirmed present; needed by downstream `docx-render`
  - `xml-parsing-extra.ts` (~560) — XML parsing continuation (upstream 16049-17170 has ~29 definitions; current `xml-parsing.ts` exports 18)
  - `tracked-changes.ts` (~550) — `collectTrackedChangesFromModel` (upstream editor.tsx:22770), `collectCommentsFromModel` (upstream editor.tsx:22908) — both confirmed present; alignment "必须复刻" #32
  - `tracked-changes-gutter.ts` (~550) — tracked-change gutter cards
  - `selection-helpers.ts` (~605) — selection / caret helpers (`cloneEditorSelection`@24081, `sameEditorSelection`@24099, `shouldReissueDomSelectionRestore`@24148, `selectionFallbackParagraphLocation`@24366, `selectionFromTextRangeLocation`@24385, `normalizeSelectionForModel`@24577, `normalizeEditorCursorStateForModel`@24595 — all confirmed present, all <24953)
  - `selection-restore.ts` (~605) — DOM selection restore (`isCollapsedSelectionAtElementStart`@24208, `isSuspiciousCollapsedSelectionAtElementStart`@24247 — confirmed present)
  - `section-manipulation.ts` (~310) — section document mutation (`updateSectionParagraphTextAtLocation`@24794, `updateSectionImageFloatingAtLocation`@24879 — confirmed present, both <24953)

- **Impact:** The pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers are the critical path for `docx-composables` (which depends-on `docx-helpers`). The layout layer already declares `estimateTableRowHeightsPx` as a callback parameter and invokes it, but the helpers layer that must supply the real implementation has not delivered it — the callback contract is unbacked by a helper. Alignment items #22 (measurement-driven iteration), #23 (pretext variable-width integration), and #32 (tracked-changes from-model derivation) are explicitly "必须复刻" and remain without implementations. `vue-docx` is still a stub (`packages/vue-docx/src/` has only `composables.ts` + `index.ts` + `env.d.ts`; no `composables/` directory), confirming `docx-composables` cannot start. The barrel [index.ts:1](../../packages/docx-core/src/editor/helpers/index.ts#L1) still says "modular split in progress" and lists the missing modules as Pending — honest, but confirms the task is not complete.

  Note on the batch model: the workflow doc allows docx-helpers to land in batches, and this is review #7 of an explicitly incremental task. The blocking verdict is scoped to the *task* (`docx-helpers`) not being completable until the remaining 14 modules land. Seven successive reviews (#1–#7) have now blocked on this same F1; the deliverable surface has grown 20 → 25 → 26 → 26 → 26 → 26 → 27 across batches, with +1 module landed this cycle (`page-measurement`). Progress is resuming but the critical-path modules (pagination-plan, pretext-build/measure, table-height, tracked-changes, selection) remain.

### F2 — P3 / non-blocking — `paragraph-geometry.ts` exceeds the 1000-line hard constraint (carried from #1–#6)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1. The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000.
- **Impact:** Constraint violation; no functional impact.

### F3 — P3 / non-blocking — barrel + README still mark task "in progress" / pending (carried from #3–#6)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md:44](../README.md#L44) task table status `pending`
- **Code:** [index.ts:1,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress", modules listed Pending — note: `page-measurement` was added to the export list at [index.ts:40](../../packages/docx-core/src/editor/helpers/index.ts#L40) but the Pending comment block at lines 20-28 still lists it as pending); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** Honest documentation of partial state, consistent with the batch model. The Pending comment list is now stale — it lists `page-measurement` as Pending even though the file exists and is exported. Once the remaining 14 modules land, the barrel header, the Completed/Pending split, and the README status should all flip to `done`.
- **Impact:** Scope tracking only.

### F4 — P3 / non-blocking — `style-to-css.ts` duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1–#6)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:13](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L13) (local duplicate declaration with comment referencing non-existent `tracked-changes.ts`), used at [style-to-css.ts:227](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L227)
- **Detail:** Unchanged since batch 1. Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

### F5 — P3 / non-blocking — helpers import from the viewer layer, inverting the documented dependency direction (carried, newly formalized)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §5 — dependency graph states `viewer/{pretext-*,...} ← editor/helpers/*` (viewer depends on editor/helpers)
- **Code:** 10 helpers files import from `../../viewer/section-layout` (twipsToPixels, TWIPS_PER_PIXEL, DocumentLayoutMetrics) and [line-height.ts:21](../../packages/docx-core/src/editor/helpers/line-height.ts#L21) imports from `../../viewer/pretext-layout`; [paragraph-tracked.ts:22](../../packages/docx-core/src/editor/helpers/paragraph-tracked.ts#L22) imports from `../../layout/page-segmentation-core`; [page-measurement.ts:28](../../packages/docx-core/src/editor/helpers/page-measurement.ts#L28) imports from `../../layout/page-segmentation-core`
- **Detail:** The documented dependency direction is `viewer → editor/helpers`, but the actual imports flow `editor/helpers → viewer` (and `→ layout`). This does not create a runtime circular dependency — `viewer` does not import back from `editor/helpers` (verified: zero `from "../../editor"` matches in `packages/docx-core/src/viewer/`) — so it is not blocking. The root cause is that `section-layout.ts` (`twipsToPixels` / `TWIPS_PER_PIXEL`) is a foundational unit-conversion utility mis-housed under `viewer/`; it belongs at a lower layer (engine or a shared `units` module). `madge --circular` confirms no cycle.
- **Impact:** None functionally (typecheck green, no cycle). Architectural tidiness only; would resolve naturally if `section-layout` constants were promoted to a lower layer.

## Positive checks (current tree — independently re-verified)

| Check | Result | Detail |
|---|---|---|
| New source since #6 | ✅ confirmed | `git log --oneline -- packages/docx-core/src/editor/helpers/` shows `62dba91 docx-helpers: fix #6` added `page-measurement.ts` (689 lines). HEAD `62dba91`, working tree clean. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core typecheck` exits 0. |
| Build (docx-core) | ✅ Pass | `pnpm --filter @extend-ai/docx-core build` (tsup) succeeds — ESM + DTS. |
| Circular deps | ✅ Pass | `madge --circular --extensions ts packages/docx-core/src/editor/helpers/index.ts` → "No circular dependency found!" (38 files processed). Intra-helper import graph is acyclic. |
| Import paths | ✅ Pass | All relative (`./` intra-helper, `../../engine/*` / `../../viewer/*` / `../../layout/*` cross-layer). Zero `@extend-ai/*` package refs in source. |
| React-type cleanup | ✅ Pass | No `import * as React` / `from "react"` / `@extend-ai/*` in source. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|not.?implemented\|TODO\|FIXME\|XXX` returns **zero** matches. `placeholder` finds only legitimate uses: form-field `placeholder` property ([text-mutation.ts:128](../../packages/docx-core/src/editor/helpers/text-mutation.ts#L128)) and an explanatory comment at [table-utils.ts:209](../../packages/docx-core/src/editor/helpers/table-utils.ts#L209). No `@ts-ignore` / `@ts-expect-error`. |
| Module boundaries | ✅ Pass | Intra-helper graph acyclic (DAG verified); `paragraph-toc` remains a true leaf introduced to break a cycle; helpers never import upward to vue-docx. |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over (F2). |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`; missing-module upstream line refs independently re-verified present in editor.tsx this review (`buildParagraphPretextLayoutSource`@6586, `estimateTableRowHeightsPx`@10850, `buildRenderColumnSegmentsForPageSection`@12466, `paragraphBlockStyle`@15433, `wrappedFloatingImageStyle`@14150, `absoluteFloatingImageStyle`@14351, `collectTrackedChangesFromModel`@22770, `collectCommentsFromModel`@22908, selection helpers @24081-24595). Delivered `page-measurement.ts` verified to contain `resolveMeasuredPageContentHeightPx`@2571 upstream equivalent. |
| Whitespace | ✅ Pass | Working tree clean. |

## Required fixes to unblock (next batch)

1. **F1**: Deliver the 14 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height` + `line-height-table` (backs the layout-layer `estimateTableRowHeightsPx` callback at [page-segmentation-table.ts:138](../../packages/docx-core/src/layout/page-segmentation-table.ts#L138)), `tracked-changes` (#32), `selection-helpers`/`selection-restore`, `style-block-css` (needed by `docx-render`), `section-manipulation`, `xml-parsing-extra`, `tracked-changes-gutter`.
2. **F2** (optional): Split `paragraph-geometry.ts` to ≤1000 lines (cover-image cluster → `paragraph-geometry-cover.ts`).
3. **F3** (cleanup alongside): Flip barrel header + Completed/Pending list and README status to `done` once F1 lands; update the stale Pending comment block (remove `page-measurement`).
4. **F4** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts` and import the canonical declaration.
5. **F5** (optional, architectural): Promote `section-layout` unit constants (`TWIPS_PER_PIXEL`, `twipsToPixels`) to a lower layer to align with the documented dependency graph.

## Conclusion

**blocked**

F1 persists: 14 of 40 planned modules remain absent, including the pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task (still a stub). Since review #6, one module was closed (`page-measurement.ts`, 689 lines, 15 exports) — progress is resuming but the critical-path modules remain. The delivered code remains high quality — typecheck green, build succeeds, graph acyclic, imports relative, zero stubs/mocks/TODOs, no React leakage. Re-verify after the remaining 14 modules are delivered.
