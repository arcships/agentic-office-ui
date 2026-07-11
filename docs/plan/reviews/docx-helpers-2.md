# docx-helpers Review #2

Date: 2026-07-07
Reviewer: automated
Scope: `packages/docx-core/src/editor/helpers/` (step 5 of migration execution order)
Upstream: `@extend-ai/react-docx` @ commit `6f70b92` (verified: `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`), `react-viewer/src/editor.tsx` lines 1-24953 (pure functions / types / constants region, zero React hooks)
Design refs:
- [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — editor/helpers module plan (~21200 lines, 40 content files + barrel)
- [docx-editor-helpers-split-plan.md](../docx-editor-helpers-split-plan.md) — 24-module detailed split
- [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §2.5 — editor.tsx 前半部分 pure-function region; §3 "必须复刻" alignment items

Commits under review:
- `3a4ad69 docx-helpers: docx-core editor/helpers` (batch 1, reviewed in #1)
- `e24a049 docx-helpers: fix #1` (batch 2 — this review's delta)

## Summary

Batch 2 (`e24a049`) correctly resolves the two runtime-functional stubs flagged as blocking in review #1:

- **F2 fixed** — `estimateTextAdvanceWidthPx` is now a real per-character advance-width estimator with cache + tab/space/CJK/digit heuristics in [drop-cap.ts:223](../../packages/docx-core/src/editor/helpers/drop-cap.ts#L223), imported by [synthetic-textbox.ts:21](../../packages/docx-core/src/editor/helpers/synthetic-textbox.ts#L21). Byte-equivalent to upstream editor.tsx:7747.
- **F3 fixed** — `resolveParagraphFirstLineLeftTabStopsPx` is now the canonical implementation in [line-height.ts:423](../../packages/docx-core/src/editor/helpers/line-height.ts#L423) (filters `alignment !== "right"`, applies `firstLineOriginPx` offset via `resolveParagraphFirstLineOriginPx`), imported by [field-helpers.ts:12](../../packages/docx-core/src/editor/helpers/field-helpers.ts#L12). Matches upstream editor.tsx:9310.

Batch 2 also delivers 5 genuinely new modules (`drop-cap`, `letterhead`, `line-height`, `header-footer`, `paragraph-tracked`), all real ports of upstream (spot-checked `resolveHeaderPaginationReservePx` against upstream line 2217 — faithful). 25 of 40 planned modules now exist (≈63% by file count, ~12198 of ~21200 lines ≈58%). typecheck passes for both `@arcships/docx-core` and the full 6-package workspace; no residual stub/mock/fake; no React-type leakage; cross-layer imports are relative and DAG-respecting (`../../engine/*`, `../../layout/*`, `../../viewer/*` only).

Two issues remain:

1. **Batch 2 introduced 2 circular module dependencies** that were not present in batch 1 (review #1 explicitly passed "acyclic"). Both are runtime-safe (ESM live bindings, all cross-refs used inside function bodies, not at module top level), so typecheck passes and no runtime crash occurs — but they violate the architecture's "clean DAG" principle (§5) and are a regression from #1.
2. **15 of 40 planned modules are still missing**, including the pagination-plan trio (alignment "必须复刻" #22 measurement-driven iteration, #23 pretext integration), `tracked-changes.ts` ("必须复刻" #32), `table-height`, `pretext-build`/`pretext-measure`, `selection-helpers`/`selection-restore`, `section-manipulation`, `style-block-css`, `page-measurement`, `line-height-table`, `xml-parsing-extra`. These are the critical path for the downstream `docx-composables` task and cannot be omitted.

**Conclusion: blocked** — the two blocking stubs from #1 are resolved and the new code is clean, but the task scope is still ~58% complete and the remaining 15 modules carry "必须复刻" alignment obligations and are prerequisites for `docx-composables`. The 2 new circular dependencies are a non-blocking regression to fix alongside the next batch.

## Findings

### F1 — P1 / blocking — 15 of 40 planned helper modules still missing

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (40-file plan, ~21200 lines); [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §3 — "必须复刻" items #22, #23, #32
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 25 modules delivered (excl. barrel), 12198 lines
- **Detail:** Delivered (25): `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, xml-parsing, style-to-css, paragraph-inspect, text-mutation, numbering, table-utils, table-utils-extra, field-helpers, synthetic-textbox, paragraph-geometry, drop-cap, letterhead, line-height, header-footer, paragraph-tracked`.

  Missing (15, all within upstream editor.tsx lines 1-24953, i.e. in-scope for this task):
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

- **Impact:** The pagination-plan trio, pretext-build/measure, table-height, tracked-changes, and selection helpers are the critical path for `docx-composables` (which depends-on `docx-helpers`). Alignment items #22 (measurement-driven iteration), #23 (pretext variable-width integration), and #32 (tracked-changes from-model derivation) are explicitly "必须复刻" and remain without implementations. The barrel [index.ts:1](../../packages/docx-core/src/editor/helpers/index.ts#L1) still says "modular split in progress" and lists these as Pending — honest, but confirms the task is not complete.

  Note on batch model: the workflow doc ([plan/workflow.md](../workflow.md)) does say docx-helpers is delivered "分批 5-6 个", and review #1's required-fix #1 explicitly allowed renegotiating into `docx-helpers-1..N` batches. Batch 2 is consistent with that model and is itself clean. The blocking verdict here is scoped to the *task* (`docx-helpers`) not being completable until the remaining 15 modules land in subsequent batches.

### F2 — P2 / non-blocking — batch 2 introduced 2 circular module dependencies (regression from #1)

- **Severity:** P2
- **Blocking:** non-blocking (runtime-safe via ESM live bindings; typecheck passes)
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §5 — "依赖关系图" specifies a clean DAG (engine ← layout ← viewer ← editor/helpers); review #1 Positive-check table explicitly passed "Helper-to-helper import graph is acyclic"
- **Code:** detected by `madge --circular` on [packages/docx-core/src/editor/helpers/index.ts](../../packages/docx-core/src/editor/helpers/index.ts):
  ```
  1) field-helpers.ts > line-height.ts
  2) field-helpers.ts > line-height.ts > paragraph-geometry.ts > synthetic-textbox.ts
  ```
  - [field-helpers.ts:12](../../packages/docx-core/src/editor/helpers/field-helpers.ts#L12) imports `resolveParagraphFirstLineLeftTabStopsPx` (value) from `./line-height`
  - [line-height.ts:49](../../packages/docx-core/src/editor/helpers/line-height.ts#L49) imports `isTableOfContentsParagraph` (value) from `./field-helpers`
  - [synthetic-textbox.ts:18](../../packages/docx-core/src/editor/helpers/synthetic-textbox.ts#L18) imports `instructionTextToPageFieldKind` / `instructionTextToStyleRefTarget` (value) from `./field-helpers`, closing the longer cycle via `paragraph-geometry.ts:35 → synthetic-textbox`
- **Detail:** Confirmed these cycles did NOT exist in batch 1 (`line-height.ts` was absent at commit `3a4ad69`); they were introduced by `e24a049`. Runtime safety verified: every cross-cycle reference is consumed inside a function body (e.g. `isTableOfContentsParagraph` is called only at [line-height.ts:697](../../packages/docx-core/src/editor/helpers/line-height.ts#L697), never at module top level), so ESM live bindings resolve to the defined function by call time and no TDZ/`undefined` crash occurs. `tsc --noEmit` passes.
- **Impact:** No functional break today. Risk: a future maintainer adding a top-level call into the cycle would get a silent `undefined`. Also violates the architecture's stated DAG and regresses a check #1 had passed. Fix is low-cost: break the cycle by extracting the shared leaf predicate (e.g. move `isTableOfContentsParagraph` / `tableOfContentsLevel` into `paragraph-inspect.ts` or a new `paragraph-toc.ts` leaf, or convert the `line-height → field-helpers` edge to a callback parameter mirroring the `header-footer.ts` callback-decoupling pattern already used in this codebase).

### F3 — P3 / non-blocking — `paragraph-geometry.ts` still exceeds the 1000-line hard constraint (carried from #1)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates `paragraph-geometry.ts` at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines (27 over)
- **Detail:** Unchanged since batch 1 (review #1 F4). The cover-image detection cluster (`isLikelyFullPageCover*` / `fullPageCover*`) could be split into a `paragraph-geometry-cover.ts` to bring it under 1000. No new file in batch 2 crosses the ceiling.
- **Impact:** Constraint violation; no functional impact.

### F4 — P3 / non-blocking — barrel + README still mark task "in progress" / pending

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full module list); [plan/README.md](../README.md) task table status `pending`
- **Code:** [index.ts:1,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1) ("modular split in progress", 15 modules listed Pending); [plan/README.md:44](../README.md#L44) (`docx-helpers … pending`)
- **Detail:** Honest documentation of partial state, consistent with the batch model. No code impact. Once the remaining 15 modules land, the barrel header and README status should be updated to `done`.
- **Impact:** Scope tracking only.

### F5 — P3 / non-blocking — `style-to-css.ts` still duplicates the `ParagraphTrackedInlineChange` forward type (carried from #1)

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:10-19](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L10-L19)
- **Detail:** Unchanged since batch 1 (review #1 F6). Local duplicate of the type with a comment promising `tracked-changes.ts` will re-export the canonical declaration. Correct interim measure; cleanup owed when F1's `tracked-changes.ts` lands.
- **Impact:** None currently.

## Positive checks (this batch)

| Check | Result | Detail |
|---|---|---|
| F2 stub removed | ✅ Pass | `estimateTextAdvanceWidthPx` real impl in [drop-cap.ts:223](../../packages/docx-core/src/editor/helpers/drop-cap.ts#L223); `synthetic-textbox.ts` now imports it (`./drop-cap`), no `return 0` stub remains. |
| F3 divergence fixed | ✅ Pass | `resolveParagraphFirstLineLeftTabStopsPx` canonical in [line-height.ts:423](../../packages/docx-core/src/editor/helpers/line-height.ts#L423) (filters `!== "right"`, applies `firstLineOriginPx`); `field-helpers.ts` imports it, local divergent copy removed. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @arcships/docx-core typecheck` exits 0. |
| Typecheck (full workspace) | ✅ Pass | `pnpm typecheck` (6 packages) exits 0 — no ripple in vue-docx/xlsx. |
| Residual stubs/mocks/fakes | ✅ Pass | Grep for `stub\|mock\|fake\|placeholder\|forward declaration\|not implemented` finds only legitimate uses (form-field `placeholder` property; a comment about generator "placeholder uniform grid"). No `return 0` / throw-not-implemented stubs. |
| Import paths | ✅ Pass | All relative; zero `@extend-ai/*` package refs in source; cross-layer imports only to `../../engine/*`, `../../layout/*`, `../../viewer/*` (respects §5 DAG across layers). |
| Module boundaries | ⚠️ Partial | Layer DAG respected, but 2 intra-helper cycles introduced (F2 above) — runtime-safe but a DAG regression within the helpers layer. |
| Single-file ≤1000 | ⚠️ Partial | Only `paragraph-geometry.ts` (1027) over; all 5 new batch-2 modules are ≤959 (`numbering.ts` is the largest new-era file at 959). |
| Real implementations | ✅ Pass | Spot-checked `estimateTextAdvanceWidthPx`, `resolveParagraphFirstLineLeftTabStopsPx`+`resolveParagraphFirstLineOriginPx`, `resolveHeaderPaginationReservePx`, `paragraph-tracked.ts` predicates — all genuine ports of upstream, not stubs. |
| React-type cleanup | ✅ Pass | No `React.CSSProperties` / `import * as React` / `from "react"` in source (only comment mentions documenting the migration). |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92`; module header line-number references match upstream editor.tsx. |
| Callback-decoupling | ✅ Pass | `header-footer.ts` reserve entry points accept an `estimateDocNodeHeightPx` callback to stay decoupled from the not-yet-migrated table-height/line-height cluster — mirrors the layout layer's `PageSegmentationCallbacks` contract. Good modular design. |

## Required fixes to unblock (next batch)

1. **F1**: Deliver the 15 remaining modules. Priority for the critical path: `pagination-plan-core`/`-iterate`/`-stabilize` (alignment #22, #23), `pretext-build`/`pretext-measure` (#23), `table-height`, `line-height-table`, `tracked-changes` (#32), `selection-helpers`/`selection-restore`, `style-block-css` (needed by `docx-render`), `section-manipulation`, `page-measurement`, `xml-parsing-extra`, `tracked-changes-gutter`.
2. **F2** (non-blocking but fix alongside): Break the `field-helpers ↔ line-height` and `field-helpers → line-height → paragraph-geometry → synthetic-textbox` cycles — extract the shared TOC predicate to a leaf module or convert the `line-height → field-helpers` edge to a callback.
3. **F3** (optional): Split `paragraph-geometry.ts` to ≤1000 lines.
4. **F5** (cleanup when `tracked-changes.ts` lands): Remove the duplicated `ParagraphTrackedInlineChange` from `style-to-css.ts`.

## Conclusion

**blocked**

Batch 2 cleanly resolves both blocking stubs from review #1 (F2 `estimateTextAdvanceWidthPx`, F3 `resolveParagraphFirstLineLeftTabStopsPx`) and the 5 new modules are faithful upstream ports with correct relative imports, passing typecheck, and no residual stubs. However, 15 of 40 planned modules remain absent — including the pagination-plan trio, pretext-build/measure, tracked-changes, table-height, and selection helpers that carry "必须复刻" alignment obligations (#22, #23, #32) and are prerequisites for the downstream `docx-composables` task. Batch 2 also introduced 2 runtime-safe but architecture-violating circular dependencies (a regression from #1's acyclic graph). Re-verify after the remaining 15 modules are delivered and the cycles are broken.
