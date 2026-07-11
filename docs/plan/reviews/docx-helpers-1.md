# docx-helpers Review #1

Date: 2026-07-07
Reviewer: automated
Scope: `packages/docx-core/src/editor/helpers/` (step 5 of migration execution order)
Upstream: `@extend-ai/react-docx` @ commit `6f70b92`, `react-viewer/src/editor.tsx` lines 1-24953 (pure functions / types / constants region, zero React hooks)
Design refs:
- [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — editor/helpers 39-module plan (~21200 lines)
- [docx-editor-helpers-split-plan.md](../docx-editor-helpers-split-plan.md) — 24-module detailed split
- [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §2.5 — editor.tsx 前半部分 pure-function region

Commit under review: `3a4ad69 docx-helpers: docx-core editor/helpers` (branch `task/docx-remigration`)

## Summary

The delivery implements **20 of 39** planned modules (~9364 of ~21200 planned lines, ≈44%). What was delivered is correct in itself — relative imports, zero circular dependencies, typecheck passes, and the implemented functions are real ports of upstream (not stubs). But the task is materially **incomplete**: the pagination-plan subsystem, pretext-build/measure, drop-cap, line-height(+table), table-height, header-footer, page-measurement, tracked-changes(+gutter), selection-helpers(+restore), section-manipulation, and xml-parsing-extra modules are all absent, leaving many of the 26 upstream alignment points (items #23 pagination-plan, #41 thumbnail measurement inputs, drop cap, line-height estimation, table-height estimation, tracked-changes derivation #32, selection/caret helpers) without their helper implementations.

The barrel [index.ts](../../packages/docx-core/src/editor/helpers/index.ts) honestly documents this state ("modular split in progress", lists 30 pending modules), and forward declarations were used to keep the delivered subset self-consistent. Two of those forward declarations are functional stubs that return wrong values at runtime (one returns 0, one deviates from upstream filtering).

**Conclusion: blocked** — the scope is far short of the architecture's file list; two residual stubs violate the "no stub" rule (the architecture does not mark these as explicitly-deferred), and the missing modules block downstream tasks (docx-composables needs line-height / pagination-plan / tracked-changes / selection helpers).

## Findings

### F1 — P1 / blocking — 19 of 39 planned helper modules missing

- **Severity:** P1
- **Blocking:** blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (39-file plan, ~21200 lines)
- **Code:** [packages/docx-core/src/editor/helpers/](../../packages/docx-core/src/editor/helpers/) — 20 modules delivered (excluding barrel), 9364 lines
- **Detail:** Implemented: `constants, performance, cache-utils, editor-types, editor-types-extra, ooxml-helpers, zoom-utils, dom-helpers, default-model, state, xml-parsing, style-to-css, paragraph-inspect, text-mutation, numbering, table-utils, table-utils-extra, field-helpers, synthetic-textbox, paragraph-geometry`.

  Missing (all within upstream editor.tsx lines 1-24953, i.e. in-scope for this task):
  - `header-footer.ts` (~625) — `resolveHeaderPaginationReservePx` (upstream line 2217)
  - `page-measurement.ts` (~740) — `resolveMeasuredPageContentHeightPx` (upstream line 2571)
  - `letterhead.ts` (~215)
  - `paragraph-tracked.ts` (~750) — structural tracked-change analysis (upstream 7936-8640); referenced by paragraph-inspect.ts:7 as "lives in paragraph-tracked.ts"
  - `pretext-build.ts` (~900) — `buildParagraphPretextLayoutSource` (upstream line 6586)
  - `pretext-measure.ts` (~900) — pretext measurement integration
  - `drop-cap.ts` (~450) — `resolveDropCapFontSizePx` (7695), `estimateTextAdvanceWidthPx` (7747)
  - `line-height.ts` (~740) — `estimateParagraphLineHeightPx` (10316), `paragraphLineCountWithinWidth` (9982), `resolveParagraphFirstLineLeftTabStopsPx` (9310)
  - `line-height-table.ts` (~730) — table line-height estimation
  - `table-height.ts` (~510) — `estimateTableRowHeightsPx` (upstream line 10850)
  - `pagination-plan-core.ts` (~850) — `buildRenderColumnSegmentsForPageSection` (upstream line 12466) and pagination-plan orchestration
  - `pagination-plan-iterate.ts` (~850) — measurement-driven iterative pagination (alignment item #22)
  - `pagination-plan-stabilize.ts` (~800) — pagination stabilization / oscillation detection
  - `style-block-css.ts` (~580) — `paragraphBlockStyle` etc.
  - `xml-parsing-extra.ts` (~560) — XML parsing continuation
  - `tracked-changes.ts` (~550) — `collectTrackedChangesFromModel` (upstream line 22770, alignment item #32)
  - `tracked-changes-gutter.ts` (~550) — tracked-change gutter cards
  - `selection-helpers.ts` (~605) — selection / caret helpers
  - `selection-restore.ts` (~605) — DOM selection restore
  - `section-manipulation.ts` (~310) — section operations

- **Impact:** Downstream tasks cannot proceed correctly. `docx-composables` needs line-height / pagination-plan / tracked-changes / selection helpers; `docx-render` needs paragraphBlockStyle / run-style CSS continuation. The pagination-plan trio is the core of alignment items #22 (measurement-driven iteration) and #23 (pretext variable-width layout integration) — both flagged "必须复刻" in the alignment doc.

### F2 — P1 / blocking — residual stub `estimateTextAdvanceWidthPx` returns 0

- **Severity:** P1
- **Blocking:** blocking (residual stub not marked as explicitly-deferred by the architecture)
- **Design ref:** [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §2.5 — drop-cap / line-height helpers "可直接复制"; architecture §2.5 lists `drop-cap.ts` and `line-height.ts` as deliverables (not deferred)
- **Code:** [synthetic-textbox.ts:22-29](../../packages/docx-core/src/editor/helpers/synthetic-textbox.ts#L22-L29)
  ```ts
  // Forward declaration: estimateTextAdvanceWidthPx lives in drop-cap.ts (not
  // yet migrated). Placeholder until that module lands.
  function estimateTextAdvanceWidthPx(_text, _style): number {
    return 0;
  }
  ```
- **Detail:** Upstream implementation (editor.tsx:7747) is a per-character advance-width estimator with cache, tab/space/punctuation/digit width heuristics — ~50 lines of real logic. The stub unconditionally returns 0. It is consumed at [synthetic-textbox.ts:500](../../packages/docx-core/src/editor/helpers/synthetic-textbox.ts#L500) and [:594](../../packages/docx-core/src/editor/helpers/synthetic-textbox.ts#L594) to compute `fitScale = safeWidth / estimatedTextWidth`; because callers wrap with `Math.max(1, ...)`, `fitScale` collapses to 1 and the synthetic-text-box fit-to-width scaling never triggers. Upstream also uses this function for form-field width sizing (editor.tsx:7889-7932).
- **Impact:** Synthetic text boxes (page-anchored floating text frames, used for page-number/letterhead art) render at wrong widths; form-field intrinsic width estimation is wrong. This is a runtime-functional stub, not a type placeholder.

### F3 — P2 / blocking — `resolveParagraphFirstLineLeftTabStopsPx` diverges from upstream despite "placeholder" comment

- **Severity:** P2
- **Blocking:** blocking (functional divergence from upstream while commented as a temporary placeholder)
- **Design ref:** [upstream-docx-feature-alignment.md](../upstream-docx-feature-alignment.md) §2.5 — line-height helpers "可直接复制"; architecture §2.5 `line-height.ts` is a deliverable
- **Code:** [field-helpers.ts:13-27](../../packages/docx-core/src/editor/helpers/field-helpers.ts#L13-L27)
- **Detail:** The comment says "Forward declaration … Replaced with a placeholder until that module lands", but the body is a real (different) implementation. Upstream (editor.tsx:9310) filters tab stops by `alignment !== "right"` and offsets by `firstLineOriginPx` via `resolveParagraphFirstLineOriginPx`, then maps `Math.round(value - firstLineOriginPx)`. The ported version filters by `alignment === "left"` only and ignores the first-line origin entirely. The two produce different tab-stop positions whenever a paragraph has a first-line indent or non-left tabs.
- **Impact:** Tab-stop-based form-field alignment (TOC dot leaders, page-number fields) computes wrong first-line tab positions. Either migrate `line-height.ts` with the canonical implementation and import it, or make the local copy byte-equivalent to upstream.

### F4 — P3 / non-blocking — `paragraph-geometry.ts` exceeds the 1000-line hard constraint

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) 硬约束 — "单文件 ≤ 1000 行"; §2.5 estimates paragraph-geometry.ts at ~800 lines
- **Code:** [paragraph-geometry.ts](../../packages/docx-core/src/editor/helpers/paragraph-geometry.ts) — 1027 lines
- **Detail:** 27 lines over the ceiling. The overflow is modest and the module is cohesive (floating-image detection + render-mode predicates), but it technically violates the stated hard constraint. Splitting the cover-image detection cluster (`isLikelyFullPageCover*`, `fullPageCover*`) into a `paragraph-geometry-cover.ts` would bring it under 1000.
- **Impact:** Constraint violation; no functional impact.

### F5 — P3 / non-blocking — barrel marks task "in progress" rather than complete

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 (full 39-module list)
- **Code:** [index.ts:1-2,20-28](../../packages/docx-core/src/editor/helpers/index.ts#L1)
- **Detail:** The barrel header says "modular split in progress" and lists 30 modules as "Pending". This is honest documentation, but it confirms the task was committed in a partial state. If the intent was a batched delivery, the task brief ("你的任务是 docx-helpers … 按 objective 完成") and architecture §2.5 imply the full set; a partial batch should be split into `docx-helpers-1`, `docx-helpers-2`, etc. rather than committed as `docx-helpers`.
- **Impact:** Process / scope tracking; no code impact.

### F6 — P3 / non-blocking — `style-to-css.ts` duplicates a forward type to break a cycle

- **Severity:** P3
- **Blocking:** non-blocking
- **Design ref:** [docx-migration-architecture.md](../docx-migration-architecture.md) §2.5 — `tracked-changes.ts` is a planned module
- **Code:** [style-to-css.ts:10-19](../../packages/docx-core/src/editor/helpers/style-to-css.ts#L10-L19)
- **Detail:** `ParagraphTrackedInlineChange` is declared locally with a comment "Duplicated here to avoid a circular module dependency; tracked-changes.ts re-exports the canonical declaration." This is a reasonable interim measure while `tracked-changes.ts` is absent, but it creates a second source of truth for the type. Once `tracked-changes.ts` lands, the local duplicate should be removed and imported from there.
- **Impact:** None currently; cleanup owed when F1's tracked-changes.ts is delivered.

## Positive checks (delivered subset)

| Check | Result | Detail |
|---|---|---|
| Import paths | ✅ Pass | All relative (`../../engine/*`, `../../viewer/*`, `./*`). Zero `@extend-ai/react-docx-*` package references in source. |
| Typecheck (docx-core) | ✅ Pass | `pnpm --filter @arcships/docx-core typecheck` exits 0. |
| Typecheck (full workspace) | ✅ Pass | `pnpm typecheck` (6 packages) exits 0 — no ripple errors in vue-docx/xlsx. |
| Module boundaries | ✅ Pass | Helper-to-helper import graph is acyclic. Leaf modules (constants, editor-types, state, text-mutation, zoom-utils, default-model, dom-helpers, performance, cache-utils, ooxml-helpers) have no intra-helper deps; the field-helpers ↔ numbering ↔ text-mutation cluster and paragraph-geometry → {paragraph-inspect, synthetic-textbox} edges form a clean DAG. |
| Real implementations | ✅ Pass | Spot-checked `buildParagraphNumberingLabels`, `mutateParagraphTextStyleInRange`, `runStyleToCss`, `defaultStarterModel`, paragraph-geometry predicates — all are genuine ports of upstream logic, not stubs (modulo F2/F3). |
| React-type cleanup | ✅ Pass | No `React.CSSProperties` / `import * as React` residual in helpers; types replaced with plain equivalents per alignment doc Phase 2 sed plan. |
| Upstream commit match | ✅ Pass | Upstream confirmed at `6f70b92`; line-number references in module headers match. |

## Required fixes to unblock

1. **F1**: Deliver the 19 missing modules (or explicitly renegotiate the task into batched sub-tasks `docx-helpers-1..N` and update [plan/README.md](../README.md) status). The pagination-plan trio, line-height, table-height, tracked-changes, and selection helpers are the critical path for `docx-composables`.
2. **F2**: Replace the `return 0` stub with the real `estimateTextAdvanceWidthPx` (port upstream editor.tsx:7747-7840) — either inline or by delivering `drop-cap.ts`.
3. **F3**: Port the canonical `resolveParagraphFirstLineLeftTabStopsPx` (upstream editor.tsx:9310-9326, including `resolveParagraphFirstLineOriginPx` and the `firstLineOriginPx` offset) — either inline or by delivering `line-height.ts`.
4. **F4** (optional for unblock): Split `paragraph-geometry.ts` to ≤1000 lines.

## Conclusion

**blocked**

The delivered 20 modules are high-quality and correctly integrated, but the task is ~44% complete against the architecture's 39-module file list, and two residual runtime-functional stubs (F2, F3) violate the "no unmarked stub" rule. The missing modules (F1) are prerequisites for the downstream `docx-composables` and `docx-render` tasks. Re-verify after the missing modules are delivered and F2/F3 stubs are replaced with upstream-equivalent implementations.
