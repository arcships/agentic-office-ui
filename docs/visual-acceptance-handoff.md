# Visual Acceptance Handoff — Upstream Comparison Plan

## Current state

The repository has completed the first engineering verification pass:

- Real DOCX / XLSX / PDF / PNG / JSON fixtures exist under `apps/demo/public/samples`.
- Demo routes exist for Home, DOCX Viewer, DOCX Editor, XLSX Viewer, PDF Viewer, and Components.
- Material verification passes with `scripts/verify_test_materials.py`.
- Workspace `pnpm typecheck` and `pnpm build` pass after core packages have been built.
- Initial browser smoke checks found and fixed several data-loading, error-state, editing, upload, and responsive overflow issues.
- Upstream Extend UI attribution is documented in `README.md` and `docs/upstream-extend-ui.md`.

The next phase is a user-visible visual acceptance pass. This phase should treat the shared browser as the source of truth and should not rely on code inspection alone.

Because this project is restoring/porting Extend UI behavior, the upstream React playgrounds are now the primary comparison baseline. Generic shadcn/ui standards remain important, but upstream React rendering and interactions are the best reference for parity decisions.

## Goal for this phase

Perform a full visual and interaction acceptance review of the Vue Extend UI demo by comparing each local route against the closest upstream React playground state, then fixing local Vue-side differences until the experience matches upstream intent or an intentional difference is explicitly documented.

The completion target is:

> Every route and component looks and behaves correctly in the shared browser across desktop, laptop, tablet, and mobile viewports, with upstream Extend UI visual/interaction parity where available, shadcn-style consistency where upstream is indirect, real user interaction coverage, and explicit user confirmation per page.

## Revised acceptance contract

The acceptance target is not a smoke test and not a documentation exercise. A route only passes when a user can operate the visible UI in the shared browser and the behavior is close enough to the upstream Extend UI intent to be used as a believable Vue port.

A page is **not accepted** if any of the following are true:

- The page visually looks like raw browser DOM controls rather than a cohesive shadcn/upstream-style product UI. Native `button`, `select`, `input`, `table`, and file controls must either be wrapped/styled consistently or explicitly documented as temporary non-accepted implementation.
- Switching demo fixtures from dropdowns leaves stale state, incorrect page/sheet counts, broken rendering, missing content, impossible-to-read content, or misleading success status.
- A toolbar button or advertised feature is visible but does not work, unless the UI clearly marks it as unavailable and the gap is documented as an intentional non-goal for the current acceptance slice.
- The component passes only by DOM/text presence while the rendered visual surface is wrong in the shared browser.
- The implementation relies on claims like “API compatible” while the visible interaction differs materially from upstream behavior.
- There are known high/medium user-visible issues that have not been fixed or explicitly re-scoped before the user review.

Acceptance evidence must distinguish three statuses:

- **Accepted** — browser-tested, user-visible behavior works, user explicitly accepted the page.
- **Known intentional gap** — browser-tested gap remains, documented with user/project agreement that it is outside the current acceptance slice.
- **Rejected / needs work** — broken, misleading, visually raw, or not close enough to upstream/shadcn expectations.

The current user review rejected the previous pass. The following observed problems are acceptance blockers until fixed or explicitly re-scoped:

1. Many pages still use raw DOM-looking controls despite a shadcn-style shell.
2. Dropdown fixture switching causes broken or misleading render/state in multiple viewers.
3. DOCX Viewer pagination/page display is incorrect for paginated content.
4. DOCX Editor editing is not upstream-equivalent; typing can mutate the whole line/block unexpectedly.
5. XLSX Viewer lacks core spreadsheet interactions such as multi-cell selection and drag selection; many advertised controls do not work.
6. PDF Viewer has broken or incomplete pagination, search highlight/scroll behavior, and PDF frame sizing.

## Source documents

Use these documents as supporting references:

- `docs/shadcn-components-browser-acceptance.md` — detailed shadcn/components acceptance standard and checklists.
- `docs/component-browser-verification-plan.md` — original browser verification scope.
- `docs/component-browser-verification-evidence.md` — engineering verification evidence from the previous pass.
- `docs/upstream-extend-ui.md` — upstream attribution and sync policy.
- `docs/upstream-visual-comparison-plan.md` — compatibility redirect; this document is now the canonical merged plan.
- `docs/upstream-parity-gap-audit.md` — strict source-level DOCX/XLSX upstream parity gap audit. This is a required blocker list for any future full-parity claim.

## Upstream reference snapshots

The upstream repositories are cloned outside this repository so their source code is available for reference without vendoring it into the Vue workspace.

| Upstream | Local path | Checked commit | Notes |
|---|---|---:|---|
| `@extend-ai/react-docx` | `/Users/eric8810/Code/extend-ui-upstream/react-docx` | `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764` | React DOCX monorepo and playground |
| `@extend-ai/react-xlsx` | `/Users/eric8810/Code/extend-ui-upstream/react-xlsx` | `f285a1c1a2a02e441a2e1f56e2fa480a0a979502` | React XLSX monorepo and playground |

Do not copy upstream source files into this repository unless a future task explicitly decides to vendor a snapshot with license and commit metadata. Use these clones for comparison, behavior study, screenshots, and browser observations only.

## Demo and reference environments

### Local Vue demo

Project root:

```bash
/Users/eric8810/Code/agentic-office-ui
```

Run demo:

```bash
pnpm dev
```

Default browser URL:

```text
http://localhost:5000
```

Routes:

```text
http://localhost:5000/#/
http://localhost:5000/#/docx-viewer
http://localhost:5000/#/docx-editor
http://localhost:5000/#/xlsx-viewer
http://localhost:5000/#/pdf-viewer
http://localhost:5000/#/components
```

### Upstream React DOCX playground

Project root:

```bash
/Users/eric8810/Code/extend-ui-upstream/react-docx
```

Useful commands:

```bash
pnpm install --frozen-lockfile
pnpm dev
pnpm test:visual
pnpm test:visual:update
```

Default upstream dev command runs the DOCX playground via:

```bash
pnpm --filter @extend-ai/react-docx-playground dev
```

The upstream repository also has Playwright visual tests configured in `playwright.config.ts`; the visual test web server uses `http://127.0.0.1:4173`.

### Upstream React XLSX playground

Project root:

```bash
/Users/eric8810/Code/extend-ui-upstream/react-xlsx
```

Useful commands:

```bash
pnpm install --frozen-lockfile
pnpm dev
pnpm build:playground
```

Default upstream dev command runs the XLSX playground via:

```bash
pnpm --filter @react-xlsx/playground dev
```

## Acceptance principles

1. Use the shared browser for all acceptance decisions.
2. Compare local Vue routes against upstream React playground states wherever a direct upstream reference exists.
3. Use shadcn/ui standards for areas without direct upstream coverage, especially PDF and generic shared components.
4. Real interactions beat static inspection: cover mouse, keyboard, upload, drag, scroll, hover, focus, disabled, selected, readonly, dirty, loading, empty, error, and success states.
5. If upstream and local results disagree visually, the upstream browser result wins unless the difference is intentional and documented.
6. Fix issues immediately where practical, then retest in the shared browser.
7. Ask for user confirmation after each page before moving to the next page.

## Source-level upstream parity audit gate

The current strict audit found that the Vue DOCX and XLSX implementations are not yet full upstream-equivalent ports. The complete blocker list is maintained in `docs/upstream-parity-gap-audit.md` and must be treated as part of this handoff.

A DOCX or XLSX page cannot be accepted as upstream-parity complete while any High or Medium gap in `docs/upstream-parity-gap-audit.md` remains unresolved, unless that exact gap is explicitly re-scoped as an intentional non-goal with user/project agreement. Browser smoke evidence, passing build/typecheck, or a stable demo route is not enough to override source-level no-ops, placeholder exports, or missing upstream features.

Current top-level gap categories from the audit:

### DOCX unresolved categories

- Public API surface mismatch: missing form-field, image-wrap, context-menu, richer controller, and several upstream extension APIs.
- Editor/controller incompleteness: range selection, partial-run editing, links, lists, table cell editing, table structure editing, image insertion, comments, tracked changes, form fields, and context menus are missing or incomplete.
- Formatting shells/no-ops: font family, font size, text color, highlight, superscript/subscript, paragraph styles, line spacing, and borders are not fully implemented.
- Rendering/model fidelity gaps: complex tables, image fallback/crop/wrap/floating, headers/footers, sections, columns, notes, hyperlinks, pagination precision, dark/night-reader mode, and large-document virtualization are incomplete.
- Engine/export blockers: DOCX WASM bridge and serialization/export path are not production-equivalent to upstream; fallback parsing drops many OOXML structures.
- Demo coverage gap: local DOCX Editor exposes only a small toolbar subset and cannot validate the upstream playground feature set.

### XLSX unresolved categories

- Controller/workbook engine mismatch: local controller lacks a real upstream-equivalent workbook/worksheet engine and returns many no-op mutation methods.
- Worker/deferred/large-file behavior gaps: upstream worker-backed parsing, deferred loading, read-only threshold behavior, and snapshots are not fully implemented.
- Workbook model fidelity gaps: hidden rows/columns, merges, frozen panes, styles, tables, validations, conditional formats, sparklines, charts/images/shapes/form controls, and chartsheets are missing or incomplete.
- Editing/persistence gaps: formulas/recalculation, style/format persistence, merge/unmerge, copy/paste/fill, row/column resize, table sorting, undo/redo coverage, and export/reimport are not fully workbook-backed.
- Rendering gaps: canvas/virtualized large-grid renderer, custom `getCellStyle`, image/chart/shape overlays, table header menus, thumbnails, and frozen panes are incomplete.
- Export/download blocker: local `download()` writes a JSON summary with `.xlsx` extension and `exportXlsx()` delegates to it, so true XLSX export parity is not present.
- Demo coverage gap: local XLSX demo labels controller/viewer partial/scoped and does not yet cover the upstream playground feature set.

Before final acceptance, each listed category must be decomposed to the concrete IDs in `docs/upstream-parity-gap-audit.md`, then either implemented and verified or explicitly accepted as an intentional gap.

## Feature acceptance gates

These gates override the more general page checklist below. A page cannot be accepted unless every applicable gate passes in the shared browser and the source-level upstream parity audit gate above is satisfied.

### Shared UI gate

- All visible buttons, selects, inputs, file uploads, tabs, toolbars, cards, tables, badges, and error states use a consistent shadcn/upstream visual language.
- Focus-visible, hover, active, disabled, selected, loading, empty, success, and error states are visibly distinct.
- Controls that do nothing are removed, disabled with explanation, or implemented.
- Keyboard Tab order follows visual order and always shows focus.

### Fixture-switching gate

For every dropdown-driven fixture route:

- Switching every valid fixture fully resets stale viewer/editor state.
- The loaded filename/status, page/sheet count, selection, zoom, error state, and toolbar disabled state match the new fixture.
- Corrupted fixtures show a clear error state and do not retain stale content from the previous valid fixture.
- Re-switching from corrupted back to valid fixture recovers without page reload.

### DOCX Viewer gate

- Multi-page DOCX files render as multiple visible pages or an explicitly accepted equivalent page navigation model.
- Page count/current page display matches the rendered pages.
- Scrolling through pages updates visible page/navigation state where such state is shown.
- Tables, images, headings, CJK text, and page margins remain readable after fixture switching.
- No footer/status/pagination UI reports impossible or stale values.
- Source-level DOCX viewer gaps from `docs/upstream-parity-gap-audit.md` are closed or agreed intentional before parity acceptance, including headers/footers, hyperlinks, complex tables, images/fallbacks, sections/columns, notes/comments/revisions visibility, pagination fidelity, dark/night-reader behavior, import depth, and virtualization.

### DOCX Editor gate

- Editing a paragraph changes only the intended insertion/selection range; typing must not rewrite, duplicate, reorder, or restyle the whole line/block unexpectedly.
- Bold/italic/underline and heading controls affect only the intended selected/current text block according to the documented editor model.
- Undo/redo restore the actual edited content and selection-adjacent state, not only toolbar availability.
- Theme toggling changes the editor surface without losing text, selection, or undo history.
- Every visible editor toolbar/control is backed by real behavior, undo/redo participation, and export/reimport behavior where applicable; otherwise it is hidden/disabled with explanation.
- Source-level DOCX editor gaps from `docs/upstream-parity-gap-audit.md` are closed or agreed intentional before parity acceptance, including range selection, font family/size/color/highlight/subscript/superscript, links, lists, paragraph styles, line spacing, borders, table editing, image insertion/wrap, comments, tracked changes, form fields, context menus, thumbnails, WASM import/export, and serialization.
- If full upstream editor parity is not implemented, the route must be renamed/scoped honestly (for example “DOCX editing demo”) and upstream-missing behaviors must be documented before acceptance.

### XLSX Viewer gate

- Workbook switching works for all fixtures and clears stale sheet/selection/edit/error state.
- Sheet tabs switch sheets and update grid content.
- Single-cell selection, range selection, Shift/Cmd-style extension where implemented, and drag-to-select visibly work.
- Double-click edit, Enter commit, Escape cancel, Tab commit-and-move, undo, redo, readonly blocking, zoom, upload, and download are either implemented or hidden/disabled with a clear reason.
- Grid headers, gridlines, active cell, selected range, edited cell, and readonly state are visually clear.
- Download/export controls generate truthful outputs: source download returns the original workbook source where applicable, XLSX export returns a real openable workbook, and CSV export reflects the active sheet. JSON summaries must not be downloaded with `.xlsx` extensions.
- Source-level XLSX gaps from `docs/upstream-parity-gap-audit.md` are closed or agreed intentional before parity acceptance, including real workbook/controller engine behavior, worker/deferred loading, formulas/recalculation, persisted styles, merges, frozen panes, clipboard/paste/fill, row/column resizing, tables/sorting, conditional formats, data validation, sparklines, images/charts/chartsheets/shapes/form controls, canvas/virtualized rendering, custom render hooks, thumbnails, URL/file load replacement, and export/reimport.

### PDF Viewer gate

- PDF frame has usable height and does not collapse at desktop, laptop, tablet, or mobile widths.
- Page next/previous and thumbnail clicks actually change the displayed page and the page indicator.
- Zoom and rotate visibly affect the displayed PDF without breaking layout.
- Search provides visible result feedback, highlights or otherwise locates matches, and scrolls/navigates to the active result when feasible.
- Corrupted PDFs show a clear error and do not leave stale thumbnails/page counts/content.

## Page mapping

| Local Vue route | Upstream reference | Comparison target |
|---|---|---|
| `/#/` Home | Upstream DOCX/XLSX playground shell and shadcn component patterns | Overall visual language, nav/card/table density, page hierarchy |
| `/#/docx-viewer` | `react-docx/apps/playground/src/App.tsx`, `DocxEditorViewer`, page thumbnails APIs, visual tests where relevant | DOCX page surface, document scale, pagination/thumbnail conventions, loading/error states |
| `/#/docx-editor` | React DOCX playground editor toolbar and `DocxEditorViewer` | Toolbar grouping, editing surface, selection/focus, undo/redo, theme, formatting controls |
| `/#/xlsx-viewer` | `react-xlsx/apps/playground/src/App.tsx`, `XlsxViewer`, `XlsxViewerProvider` | Spreadsheet shell, ribbon/toolbar density, gridlines, headers, selection/editing, sheet tabs, zoom/thumbnail behavior where applicable |
| `/#/pdf-viewer` | No direct upstream React DOCX/XLSX equivalent | Use upstream shadcn shell/component patterns as indirect reference |
| `/#/components` | Upstream playground `components/ui/*` and shared UI patterns | Buttons, inputs, tooltip, spinner, cards, upload-like surfaces, empty/error state style |

## Viewport matrix

| Viewport | Purpose | Acceptance focus |
|---|---|---|
| `1440×900` | Desktop default | Layout rhythm, toolbar density, page shell alignment, full interaction path |
| `1280×720` | Laptop short viewport | Vertical density, sticky/scroll areas, toolbar wrapping |
| `768×1024` | Tablet | Internal scroll containers, stacked controls, touch target size |
| `390×844` | Mobile | No page-level horizontal scroll, readable hierarchy, usable controls |

Responsive pass criteria:

- Page-level horizontal overflow is absent.
- Wide document/grid content uses internal scrolling with clear affordance.
- Toolbar/control rows wrap or collapse without clipping.
- Touch targets remain usable.
- Critical state text remains readable.

## Recommended loop per page

1. Open the upstream reference in the shared browser.
2. Set viewport to `1440×900`.
3. Record the upstream URL, commit, viewport, and exact interaction state.
4. Capture notes or screenshots for the upstream baseline.
5. Open the matching local Vue route.
6. Reproduce the same viewport and interaction state.
7. Compare layout, density, surface, controls, rendering fidelity, interaction states, and keyboard behavior.
8. Repeat at `1280×720`, `768×1024`, and `390×844`.
9. Record every mismatch using the issue template below.
10. Fix local issues immediately where practical.
11. Reopen/retest the failed state.
12. Ask the user to confirm the page before moving on.

## Comparison categories

| Category | Questions |
|---|---|
| Shell | Does page width, nav/header, background, and card hierarchy match upstream intent? |
| Controls | Are buttons, selects, inputs, toggles, labels, disabled states, and focus rings aligned? |
| Density | Are padding, gaps, row heights, toolbar grouping, and text sizes close to upstream? |
| Surface | Do document pages, grids, thumbnails, sidebars, and viewer frames feel like upstream surfaces? |
| Interaction | Do hover, active, focus-visible, selected, dirty, readonly, error, empty, loading, and success states match? |
| Keyboard | Is Tab order logical and focus styling comparable? |
| Responsive | Does local behavior preserve upstream intent while preventing page-level overflow? |
| Fidelity | For DOCX/XLSX, are rendered content, pagination/gridlines, selection, and editing behavior acceptably close? |

## shadcn/components visual standards

Check each page for:

- Clear title/subtitle hierarchy.
- Consistent spacing between header, controls, status, viewer/content, and tables.
- Consistent border radius on cards, buttons, selects, inputs, upload zones, and viewer shells.
- Border and muted background usage matching shadcn visual language and upstream playground patterns.
- Button variants that communicate primary, secondary, destructive, icon, disabled, hover, active, and focus states.
- Inputs/selects with visible focus rings and stable sizing.
- Tables and grids with readable headers, cell states, selected states, and overflow handling.
- Empty/loading/error states that are visually stable, readable, and specific.
- Tooltip and popover-like surfaces that feel intentional and do not clip unexpectedly.
- Keyboard focus visibility throughout the route.

## Fix order

Fix in this order to avoid churn:

1. Global tokens and base shell: CSS variables, focus rings, body/background, nav layout.
2. Shared component primitives: button, input/select, card, table, tooltip, spinner, upload zone.
3. Viewer shells: border/radius/background/overflow and internal scroll boundaries.
4. DOCX-specific rendering and editor toolbar behavior.
5. XLSX-specific grid, tabs, selection, editing, zoom, readonly behavior.
6. PDF route shadcn consistency, using upstream shell/component patterns as indirect reference.
7. Responsive polish and keyboard/a11y pass.

Prefer small Vue-side fixes with browser retest after each meaningful change. Avoid broad rewrites unless an upstream comparison shows a systemic mismatch.

## Issue record template

Use this format for every upstream/local mismatch or acceptance issue:

```text
ID:
Local page/component:
Upstream reference:
Viewport:
State/interaction:
Steps:
Upstream observed:
Local observed:
Expected local behavior:
Type: visual / interaction / responsive / keyboard / a11y / state / data-fidelity / intentional-difference
Severity: blocker / high / medium / low
Likely local files:
Files changed:
Fix summary:
Retest result:
User confirmation: pending / accepted
```

## Page-by-page checklist

### 1. Home

Route: `/#/`

Upstream references:

- Upstream DOCX/XLSX playground page shell.
- Upstream `components/ui/card.tsx`, `components/ui/table.tsx`, `components/ui/badge.tsx`, `components/ui/button.tsx`.

Visual review:

- Main title and subtitle align and read well.
- Navigation cards have upstream-like spacing, border, radius, hover, and status badges.
- API map table has shadcn-like table styling and readable density.
- Nav spacing and active state feel intentional.
- Page remains balanced across all viewports.

Interactions:

- Hover each card.
- Tab through every card link.
- Activate each card with keyboard and mouse.
- Return to Home after route navigation.

Expected user confirmation:

> Home navigation and API map visual experience accepted.

### 2. DOCX Viewer

Route: `/#/docx-viewer`

Upstream references:

- `react-docx/apps/playground/src/App.tsx`, especially `DocxEditorViewer`, thumbnails, page layout, theme, and toolbar usage.
- Upstream visual tests under `react-docx/tests/visual` where relevant.

Visual review:

- Controls area looks like an upstream/shadcn control row.
- Status area clearly communicates loaded file and expected coverage.
- Viewer shell has clear border, radius, background, and internal scrolling.
- Rendered DOCX page background, scale, shadow, and margins are close to upstream intent.
- Rendered DOCX page does not break the app layout.
- Error state for corrupted DOCX is specific and visually integrated.

Interactions:

- Load each sample: `demo.docx`, `legal-contract.docx`, `invoice-table.docx`, `report-with-image.docx`, `chinese-mixed.docx`, `corrupted.docx`.
- Upload a valid DOCX fixture.
- Scroll the document.
- Check table/image/chinese content visibility.
- Tab through select, button, and file input.
- Verify mobile/tablet internal scroll behavior.

Expected user confirmation:

> DOCX Viewer visual parity, loading, upload, scroll, and error-state experience accepted.

### 3. DOCX Editor

Route: `/#/docx-editor`

Upstream references:

- Upstream DOCX playground editor toolbar and document surface.
- `DocxEditorViewer`, `useDocxEditor`, formatting/border/theme hooks in upstream `App.tsx`.

Visual review:

- Editor toolbar feels like a cohesive upstream/shadcn toolbar.
- Toolbar grouping, density, and icon/button variants align with upstream intent.
- Formatting buttons have visible hover/focus/active/disabled states.
- Contenteditable page surface has clear document affordance.
- Focus ring on editable text is visible and tasteful.
- Undo/redo/theme status feedback is understandable.

Interactions:

- Click editable paragraph.
- Type Chinese + English + numbers.
- Toggle Bold.
- Undo and redo.
- Toggle theme.
- Use Tab to traverse toolbar and editor controls.
- Verify Escape/Tab/focus behavior remains stable.
- Verify mobile/tablet internal scroll behavior.

Expected user confirmation:

> DOCX Editor editing, toolbar, focus, undo/redo, theme, and upstream parity experience accepted.

### 4. XLSX Viewer

Route: `/#/xlsx-viewer`

Upstream references:

- `react-xlsx/apps/playground/src/App.tsx`.
- `XlsxViewer`, `XlsxViewerProvider`, `useXlsxViewerSelection`, `useXlsxViewerEditing`, `useXlsxViewerZoom`, and thumbnail behavior.

Visual review:

- Toolbar/ribbon density and grouping align with upstream intent.
- Sheet tabs, grid headers, and grid cells feel consistent.
- Gridline color, header style, row height, and column width are acceptably close to upstream.
- Selection outline is visible and not visually noisy.
- Edit input is aligned with the cell and readable.
- Readonly and corrupted workbook states are clear.
- Wide sheets use internal scrolling instead of page-level overflow.
- Upstream latest commit restored gridlines in XLSX viewer thumbnails, so gridline visibility deserves special attention.

Interactions:

- Load each workbook: `financial-model.xlsx`, `sales-table.xlsx`, `charts-images.xlsx`, `large-grid.xlsx`, `corrupted.xlsx`.
- Switch sheets.
- Click cells.
- Double-click a cell and edit.
- Press Enter to commit.
- Press Escape to cancel.
- Press Tab to commit and move.
- Use undo/redo.
- Zoom in/out.
- Toggle readonly and verify editing is blocked.
- Upload an XLSX fixture.
- Trigger download.

Expected user confirmation:

> XLSX Viewer grid, tabs, editing, keyboard, zoom, readonly, upload, error-state, and upstream parity experience accepted.

### 5. PDF Viewer

Route: `/#/pdf-viewer`

Upstream reference:

- No direct upstream React DOCX/XLSX equivalent in the cloned sources.
- Use upstream shadcn shell/components as indirect references.

Visual review:

- Toolbar, thumbnails, PDF frame, status, and error surfaces feel unified.
- Current page state is readable.
- Thumbnail rail does not crowd the page.
- Search feedback is visible and understandable.
- Corrupted PDF error state is clear.
- Mobile/tablet layout remains usable.

Interactions:

- Load each PDF: `sample.pdf`, `scanned-invoice.pdf`, `rotated-pages.pdf`, `large-contract.pdf`, `corrupted.pdf`.
- Upload a valid PDF fixture.
- Use next/previous page controls.
- Click thumbnails.
- Zoom in/out.
- Rotate.
- Search.
- Download.
- Tab through toolbar controls.

Expected user confirmation:

> PDF Viewer frame, thumbnails, pagination, zoom, search, upload, download, and error-state experience accepted.

### 6. Components

Route: `/#/components`

Upstream references:

- `react-docx/apps/playground/src/components/ui/*`
- `react-xlsx/apps/playground/src/components/ui/*`

Visual review:

- Component demo cards have upstream-like spacing and hierarchy.
- Controls and status text are aligned and easy to scan.
- Button/input/select/card/table/tooltip/spinner visual styling is close to upstream shadcn components.
- Error messages look intentional.
- Image thumbnails and overlay components do not overflow or distort.
- Tooltip and spinner styling matches the rest of the UI.

Interactions:

- SignaturePad: draw, check empty, clear, check empty again.
- FileUpload: select/drag a valid file, invalid type, too-large file, disabled state.
- FileThumbnail: inspect PDF, DOCX, XLSX, PNG, and unknown type thumbnails.
- BoundingBoxCitations: click fields and verify selected field state.
- LayoutBlocks: click OCR blocks and verify selected block state; verify empty state.
- Spinner: inspect size/color/placement.
- Tooltip: hover and focus trigger; verify hide behavior.
- Tab through interactive controls.

Expected user confirmation:

> Components visual consistency and all component interaction states accepted.

## Known areas needing special attention

- The previous phase verified engineering behavior, but upstream visual parity still needs to happen page by page.
- DOCX and XLSX naturally contain wide document/grid content. The acceptance standard is internal scrolling with clear affordance, not page-level horizontal overflow.
- PDF viewer currently uses browser iframe rendering. Acceptance should focus on visible controls, state, fallback/error handling, and user-perceived behavior.
- File upload and download paths should be exercised in the browser, not only checked by code.
- Keyboard focus styling should be reviewed visually; existing smoke checks did not fully evaluate focus aesthetics.
- Hover/focus/disabled/active states need actual browser inspection.
- DOCX page surface, toolbar grouping, and XLSX gridlines/selection are likely the highest-value upstream comparison areas.
- The strict upstream parity gap audit in `docs/upstream-parity-gap-audit.md` supersedes any earlier evidence that implied only user confirmation remained. Treat all High/Medium DOCX/XLSX gaps in that audit as unresolved blockers until implemented, retested, or explicitly re-scoped with agreement.

## Evidence to update during execution

During this phase, update or create evidence alongside fixes:

- `docs/component-browser-verification-evidence.md` — append browser evidence after retesting fixed behavior.
- `docs/upstream-visual-comparison-evidence.md` — side-by-side upstream/local observations and screenshots.
- `docs/upstream-parity-gap-audit.md` — required source-level DOCX/XLSX parity gap ledger. Update each affected gap ID when functionality is implemented, verified, intentionally re-scoped, or newly discovered.

## Commands to run after fixes

Run after each meaningful code change or at the end of a page batch:

```bash
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py
```

If `pnpm typecheck` fails because dependent workspace package `dist` types are missing, build the relevant core packages first, then rerun the full check.

## Final completion criteria

This phase is complete only when all of the following are proven by current evidence:

- Upstream React DOCX and XLSX playgrounds have been opened locally and used as visual/interaction references.
- All six local routes have been opened and operated in the shared browser, not just inspected by code or snapshots.
- All page-specific checklists and all feature acceptance gates above pass.
- All four viewports pass for every route.
- Fixture switching has been tested for every listed DOCX/XLSX/PDF sample, including corrupted fixtures and recovery back to valid fixtures.
- All visible controls either work, are disabled with a clear explanation, or are removed from the accepted UI.
- All high/medium upstream mismatches and all user-reported blockers are fixed or explicitly documented as intentional gaps with user agreement.
- Every High/Medium gap ID in `docs/upstream-parity-gap-audit.md` is implemented and verified, or explicitly marked as an intentional non-goal with user/project agreement.
- No page-level horizontal overflow remains.
- No browser page errors or business console errors remain after a clean reload of each route.
- All discovered visual and interaction issues are fixed and retested in the shared browser.
- Evidence documents record accepted, known-gap, and rejected statuses honestly; rejected prior evidence is not reused as proof of completion.
- The user explicitly confirms each page after seeing the working shared browser state.
- Final `pnpm typecheck`, `pnpm build`, and material verification pass.

Do not mark this phase complete if the only proof is passing typecheck/build, text presence in the DOM, no horizontal overflow, or a checklist filled without successful browser operation.
