# Upstream Visual Comparison Evidence

> Status update (2026-07-05): The prior visual acceptance pass is **rejected / needs work** after user review. This file is retained as historical evidence only and must not be used as proof that `docs/visual-acceptance-handoff.md` is complete. Current blockers include raw DOM-looking controls, broken fixture switching/render state, incorrect DOCX pagination/editor behavior, missing XLSX multi-select/drag selection and broken advertised controls, and broken PDF pagination/search/frame sizing.


## Scope

This evidence records the first upstream-comparison implementation pass for the merged plan in `docs/visual-acceptance-handoff.md`.

References used:

- Local Vue demo: `http://127.0.0.1:5002`
- Upstream DOCX playground: `http://127.0.0.1:5174`
- Upstream XLSX playground: `http://127.0.0.1:5175`
- Upstream DOCX clone: `/Users/eric8810/Code/extend-ui-upstream/react-docx` @ `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`
- Upstream XLSX clone: `/Users/eric8810/Code/extend-ui-upstream/react-xlsx` @ `f285a1c1a2a02e441a2e1f56e2fa480a0a979502`

## Upstream observations

### DOCX playground

Browser snapshot confirmed the upstream DOCX playground uses:

- compact app chrome with `Ready` status,
- grouped shadcn-like toolbar clusters,
- disabled undo/redo states,
- paragraph style, font family/size, line spacing,
- bold/italic/underline/strike/superscript/subscript,
- text/highlight color controls,
- alignment, list, border, image/table, zoom, import/download controls,
- switches for edits, comments, and read-only.

### XLSX playground

Browser snapshot confirmed the upstream XLSX playground uses:

- workbook title bar,
- ribbon tabs (`Home`, `Insert`, `Page Layout`, `Formulas`, `Data`, `View`),
- grouped ribbon controls with disabled states,
- formula/name box area,
- empty workbook state,
- shadcn-like dense controls and tokenized focus/disabled styling.

## Issues addressed in this pass

### UVC-001 — Global shadcn token coverage incomplete

- Type: visual / shared styling
- Severity: high
- Local files changed:
  - `apps/demo/index.html`

Fix:

- Expanded the demo token set to include upstream-style `--card`, `--popover`, `--secondary`, `--accent`, `--destructive`, `--input`, `--ring`, radius scales, and shadows.
- Added Inter/system font stack and global focus-visible ring behavior.
- Added global disabled cursor/opacity and transition baseline for controls.

Retest:

- `pnpm typecheck` passed.
- `pnpm build` passed.
- Browser snapshots show updated nav labels and toolbar semantics after rebuild.

### UVC-002 — App nav did not match upstream density and focus/active affordance

- Type: visual / shell
- Severity: medium
- Local files changed:
  - `apps/demo/src/App.vue`

Fix:

- Converted nav from emoji-heavy links to compact text links.
- Added sticky, translucent, token-based shell styling.
- Added upstream-like compact active/hover states.

Retest:

- Browser snapshot confirmed compact text nav on local routes.

### UVC-003 — DOCX editor toolbar too primitive visually

- Type: visual / interaction
- Severity: high
- Local files changed:
  - `packages/vue-docx/src/DocxEditor.vue`

Fix:

- Reworked toolbar markup into accessible `role="toolbar"` groups.
- Added shadcn-like toolbar clusters, compact 28px controls, hover/active/disabled states, and page count styling.
- Changed heading option label from `Normal` to `Body`, matching upstream language more closely.

Retest:

- Browser snapshot for `/#/docx-editor` shows grouped DOCX editor toolbar with Undo/Redo, heading select, Bold/Italic/Underline, theme, and page count.
- `pnpm typecheck` passed.
- `pnpm build` passed.

Known remaining gap:

- Upstream DOCX editor has many more controls and a much richer editing model. This pass improves visual grouping and density but does not implement full upstream editor parity.

### UVC-004 — DOCX page surface/canvas less upstream-like and block geometry double-applied margins

- Type: visual / data-fidelity
- Severity: high
- Local files changed:
  - `packages/vue-docx/src/DocxViewer.vue`
  - `packages/vue-docx/src/DocxEditor.vue`

Fix:

- Added muted canvas padding/background to DOCX viewer/editor content.
- Added upstream-like page border and stronger page shadow.
- Removed per-block `left/top` visual offsets in Vue viewer/editor rendering and kept blocks in flow inside the padded page, reducing double-margin distortion.
- Stopped keying editor blocks by text content to avoid remounting editable paragraphs on every text change.

Retest:

- `pnpm typecheck` passed.
- `pnpm build` passed.

Known remaining gap:

- Header/footer, rich run semantics, table fidelity, and true selection/caret preservation are still materially less complete than upstream.

### UVC-005 — XLSX toolbar did not resemble upstream ribbon

- Type: visual / interaction
- Severity: high
- Local files changed:
  - `packages/vue-xlsx/src/XlsxViewer.vue`

Fix:

- Replaced the primitive single-row toolbar with a compact workbook title area, ribbon tab list, and grouped ribbon toolbar.
- Added History, Zoom, and File groups with labels and disabled states.
- Added shadcn-like button styling and scrollable ribbon groups.

Retest:

- Browser snapshot for `/#/xlsx-viewer` shows `Workbook ribbon tabs`, `XLSX viewer toolbar`, grouped History/Zoom/File controls, and updated disabled states.
- `pnpm typecheck` passed.
- `pnpm build` passed.

Known remaining gap:

- This is a visual approximation. Full upstream formula bar, bottom tabs with thumbnail previews, virtualized grid, style fidelity, and keyboard grid semantics remain future work.

### UVC-006 — SignaturePad overflow risk on mobile

- Type: responsive
- Severity: medium
- Local files changed:
  - `packages/vue-extend/src/components/SignaturePad.vue`

Fix:

- Made the canvas wrapper max-width aware.
- Changed rendered canvas CSS width to `100%` with `max-width` based on the configured width.

Retest:

- Browser mobile viewport check on `/#/components` at `390×844` reported no page-level horizontal overflow.

## Browser checks performed

- Opened upstream DOCX playground and recorded toolbar structure through shared browser snapshot.
- Opened upstream XLSX playground and recorded title/ribbon/formula/empty-state structure through shared browser snapshot.
- Opened local `/#/docx-editor`; verified updated accessible grouped toolbar in browser snapshot.
- Opened local `/#/xlsx-viewer`; verified updated ribbon-like toolbar in browser snapshot.
- Opened local `/#/components` at `390×844`; `documentElement.scrollWidth === clientWidth`, so no page-level horizontal overflow was observed for that route/viewport.

## Command verification

Passed:

```bash
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py
```

Notes:

- A combined command `pnpm typecheck && pnpm build && uv run ...` was terminated with exit code 137 after typecheck/build had completed; the output showed typecheck and build success before termination. The material verification command was then rerun separately and passed.
- `pnpm build` still reports the known Vite dynamic/static import chunking warnings for `docx-core` and `xlsx-core`; build exits successfully.

## Remaining high/medium parity gaps

The following are not fully resolved by this pass and should be treated as intentional staged follow-up unless the acceptance bar requires full upstream parity immediately:

1. DOCX editor does not implement upstream's full toolbar/control surface, import/export, drag-drop import, read-only mode, comments/track-changes UI, thumbnails sheet, or rich editor selection model.
2. DOCX viewer/editor still lack upstream-level header/footer rendering, rich run rendering, link/highlight/script handling, and table styling fidelity.
3. XLSX viewer still lacks full upstream workbook shell including formula bar, bottom sheet tabs, thumbnail previews, virtualization, style/theme fidelity, copy/paste, and arrow-key grid navigation.
4. XLSX export still does not match upstream real workbook export behavior.
5. Tooltip/FileUpload/BoundingBox/LayoutBlocks selected-state parity needs a dedicated shared-components pass.
6. Full page-by-page user confirmation is still pending.

## Execution pass — 2026-07-05 full-route sweep

### Environment actually used

- Local Vue demo: `http://localhost:5000` (`pnpm --filter demo dev`, Vite dev server)
- Upstream DOCX reference: `http://127.0.0.1:5174`, clone commit `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`
- Upstream XLSX reference: `http://127.0.0.1:5175`, clone commit `f285a1c1a2a02e441a2e1f56e2fa480a0a979502`
- Browser automation evidence files generated under `/tmp/agentic-visual-evidence/`:
  - `metrics-pass.json` — all-route viewport metrics
  - `interaction-pass.json` — route interaction results

### Viewport matrix result

All six routes were opened at all required viewports: `1440×900`, `1280×720`, `768×1024`, `390×844`.

Result from `/tmp/agentic-visual-evidence/metrics-pass.json`:

| Route | 1440×900 | 1280×720 | 768×1024 | 390×844 |
|---|---:|---:|---:|---:|
| Home `/#/` | overflow=false | overflow=false | overflow=false | overflow=false |
| DOCX Viewer `/#/docx-viewer` | overflow=false | overflow=false | overflow=false | overflow=false |
| DOCX Editor `/#/docx-editor` | overflow=false | overflow=false | overflow=false | overflow=false |
| XLSX Viewer `/#/xlsx-viewer` | overflow=false | overflow=false | overflow=false | overflow=false |
| PDF Viewer `/#/pdf-viewer` | overflow=false | overflow=false | overflow=false | overflow=false |
| Components `/#/components` | overflow=false | overflow=false | overflow=false | overflow=false |

Responsive notes:

- Wide DOCX pages and XLSX grids remain inside internal scroll containers.
- XLSX ribbon and sheet tabs use horizontal internal scrolling on narrow widths.
- Components route no longer page-overflows at mobile width after `SignaturePad` and shared component responsive fixes.

### Interaction sweep result

Evidence source: `/tmp/agentic-visual-evidence/interaction-pass.json` plus shared browser follow-up for DOCX editor and PDF viewer.

| Page | Interaction evidence |
|---|---|
| Home | Tab traversal completed; DOCX Viewer card activation navigated away and return to Home succeeded. |
| DOCX Viewer | Loaded all DOCX samples: `demo.docx`, `legal-contract.docx`, `invoice-table.docx`, `report-with-image.docx`, `chinese-mixed.docx`, `corrupted.docx`; each reported loaded sample text/error-state evidence; valid DOCX upload path completed; keyboard traversal completed. |
| DOCX Editor | Focused editable text and inserted `中英123` via browser DOM/editing command; toolbar controls for Bold/Undo/Redo/theme were present and operated; keyboard/Escape traversal completed. |
| XLSX Viewer | Loaded all workbook samples including corrupted workbook; grid presence confirmed; double-click edit committed value `42`; readonly checkbox blocked edit mode; workbook upload path completed; keyboard traversal completed; visible sheet tabs `Assumptions`, `P&L`, `Notes` confirmed. |
| PDF Viewer | Shared browser follow-up loaded all PDFs: `sample.pdf`, `scanned-invoice.pdf`, `rotated-pages.pdf`, `large-contract.pdf`, `corrupted.pdf`; valid and corrupted states produced page/error evidence; controls are present and disabled appropriately for corrupted state; no page overflow. |
| Components | SignaturePad draw/clear/check path exercised; FileUpload valid upload completed; BoundingBox/LayoutBlocks click paths exercised; Tooltip hover/focus triggers present; keyboard traversal completed. |

### Additional issues fixed during this pass

#### UVC-007 — Shared component Tailwind utility classes were not backed by compiled CSS

- Type: visual / interaction
- Severity: high
- Files changed:
  - `packages/vue-extend/src/components/FileUpload.vue`
  - `packages/vue-extend/src/components/Tooltip.vue`
  - `packages/vue-extend/src/components/BoundingBoxCitations.vue`
  - `packages/vue-extend/src/components/LayoutBlocks.vue`

Fix:

- Replaced Tailwind utility class dependency inside shared Vue components with scoped CSS so demo visuals do not depend on unavailable Tailwind generation.
- Added keyboard-operable FileUpload drop zone (`role="button"`, Enter/Space activation), explicit disabled/dragging/focus states, and shadcn-like border/radius/background styling.
- Reworked Tooltip teleport positioning to use viewport-fixed coordinates computed from trigger/tooltip geometry, avoiding body-origin absolute positioning drift and viewport clipping.
- Added explicit selected visual states to BoundingBoxCitations and LayoutBlocks overlay/list items.

Retest:

- `pnpm --filter @extend-ai/vue-extend typecheck` passed.
- Full `pnpm typecheck` and `pnpm build` passed.
- Components route metrics show overflow=false at all four viewports.

#### UVC-008 — XLSX ribbon tabs and sheet tab placement were still less upstream-like

- Type: visual / interaction
- Severity: medium
- Files changed:
  - `packages/vue-xlsx/src/XlsxViewer.vue`

Fix:

- Expanded workbook ribbon tabs to match upstream labels more closely: `Home`, `Insert`, `Page Layout`, `Formulas`, `Data`, `View`.
- Moved sheet tabs below the grid surface, matching upstream spreadsheet shell convention more closely.

Retest:

- Browser text evidence includes `HomeInsertPage LayoutFormulasDataView` for XLSX route.
- Active workbook sheet tabs are visible at bottom and include `Assumptions`, `P&L`, `Notes`.
- Full `pnpm typecheck` and `pnpm build` passed.

#### UVC-009 — Home card status styling selector did not apply because class value included emoji/space

- Type: visual
- Severity: medium
- Files changed:
  - `apps/demo/src/pages/HomePage.vue`

Fix:

- Split status text from CSS class with `statusClass: "ready"`.
- Removed large emoji card icons to better match the compact upstream/shadcn card language.

Retest:

- Browser metrics show Home cards with text `Ready` and no page overflow at all four viewports.

### Browser errors and console status

- Shared browser `browser.errors` after final checks: no page errors.
- Vite dev console contained transient HMR messages while files were being edited, including one temporary reload failure for `XlsxViewerPage.vue`; subsequent typecheck/build passed and the route loaded successfully. This is not considered a remaining business/runtime console error.

### Final command gates — 2026-07-05

Passed:

```bash
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py
```

`pnpm build` still emits the known Vite chunking warnings for mixed dynamic/static imports of `docx-core` and `xlsx-core`; the command exits successfully.

## User confirmation status — updated

The browser/command evidence is now available for all pages, but explicit user confirmation is still required by `docs/visual-acceptance-handoff.md` before the phase can be called complete.

- Home: pending user confirmation
- DOCX Viewer: pending user confirmation
- DOCX Editor: pending user confirmation
- XLSX Viewer: pending user confirmation
- PDF Viewer: pending user confirmation
- Components: pending user confirmation

## Current acceptance rerun — 2026-07-06

Status: **in progress, not complete**. The previous rejected pass remains historical only; this section records the newer fixes and evidence gathered after the stricter acceptance contract was added.

### Fixes completed in this rerun

- DOCX layout/options:
  - `packages/docx-core/src/types.ts` extends `LayoutOptions` with `marginTop`, `marginBottom`, `marginLeft`, and `marginRight`.
  - `packages/docx-core/src/layout.ts` uses four-sided margins consistently, honors page-break-before, explicit page breaks, and section breaks, and avoids trailing empty pages.
  - `packages/vue-docx/src/DocxViewer.vue` now merges `layoutOptions` into resolved page metrics and passes the same options to `layoutDocument()`.
- DOCX editor text update:
  - `packages/docx-core/src/model.ts` changed `updateParagraphText()` from replacing the whole paragraph with a single text run to a run-preserving diff update.
  - `packages/vue-docx/src/DocxEditor.vue` displays page count from actual computed pages and restores caret offset after input.
  - `apps/demo/src/pages/DocxEditorPage.vue` now includes mixed styled runs so browser editing can prove style preservation.
- XLSX viewer:
  - `packages/vue-xlsx/src/XlsxViewer.vue` supports drag range selection, Shift-click extension, normalized reverse selections, `@dblclick.stop`, Enter commit, Tab commit-and-move, and read-only edit blocking.
  - `packages/vue-xlsx/src/composables.ts` clears selection/edit/history state on workbook load/error to prevent stale fixture state.
- PDF viewer:
  - `packages/vue-extend/src/components/PdfViewer.vue` no longer uses a tiny collapsed iframe. The frame has usable height at desktop and mobile sizes.
  - Page navigation updates the page indicator and iframe hash.
  - Search gives visible result feedback and opens the estimated hit page when feasible.
  - Corrupted PDFs clear valid content and recover when switching back to valid fixtures.
- Demo styling:
  - `apps/demo/index.html` adds shadcn-like global styling for page-level controls, file inputs, checkboxes, focus rings, and tables.
  - `apps/demo/src/main.ts` imports built package CSS for docx/extend/xlsx components so scoped component CSS is actually visible in the demo.

### Shared-browser evidence gathered

Evidence files:

- `/tmp/agentic-visual-evidence/route-viewport-browser-cli.jsonl` — route/viewport metrics gathered through `agent-browser` against `http://127.0.0.1:5002`.

Viewport matrix summary from current rerun:

| Route | 1440×900 | 1280×720 | 768×1024 | 390×844 |
|---|---:|---:|---:|---:|
| Home `#/` | overflow=false | overflow=false | overflow=false | overflow=false |
| DOCX Viewer `#/docx-viewer` | overflow=false, 6 pages | overflow=false, 6 pages | overflow=false, 6 pages | overflow=false, 6 pages |
| DOCX Editor `#/docx-editor` | overflow=false, 1 page | overflow=false, 1 page | overflow=false, 1 page | overflow=false, 1 page |
| XLSX Viewer `#/xlsx-viewer` | overflow=false, 220 grid cells visible | overflow=false, 220 grid cells visible | overflow=false, 220 grid cells visible | overflow=false, 220 grid cells visible |
| PDF Viewer `#/pdf-viewer` | overflow=false, `1 / 4`, frame 635px | overflow=false, `1 / 4`, frame 540px | overflow=false, `1 / 4`, frame 690.188px | overflow=false, `1 / 4`, frame 540px |
| Components `#/components` | overflow=false | overflow=false | overflow=false | overflow=false |

Targeted interaction evidence from shared browser:

- DOCX Viewer: `/docx-viewer` sample renders 6 `.docx-page` elements, first pages have 1056px height, visible text starts with `MASTER SERVICES AGREEMENT`.
- DOCX Editor: appending text to a paragraph with a bold middle run preserved the existing bold `paragraph` span and appended new text to the ordinary trailing run; Undo became enabled after Vue update.
- XLSX Viewer: workbook `financial-model.xlsx` loads with tabs `Assumptions`, `P&L`, `Notes`; drag selection visibly highlights an active cell and selected range; double-click edit can set a visible cell to `999`; read-only blocks edit input.
- PDF Viewer: corrupted fixture shows `Unable to load PDF document.` and `1 / —`; switching back to `large-contract.pdf` recovers to `1 / 31` with 31 thumbnails; search `page` shows `34 matches; opened page 1`; zoom-in updates hash to `zoom=125`; rotate produces a 90-degree transform matrix.

### Command gates in current rerun

Passed:

```bash
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py
git diff --check
```

`pnpm build` still emits the known Vite chunking warnings for dynamic/static imports of `docx-core` and `xlsx-core`, but exits successfully.

### Current upstream playground refresh — 2026-07-06

Evidence file: `/tmp/agentic-visual-evidence/upstream-refresh.jsonl`.

The upstream clones were verified at their documented commits:

- DOCX upstream: `/Users/eric8810/Code/extend-ui-upstream/react-docx` @ `6f70b92b8d32dcf352130bc8ad0f2f15a87a6764`.
- XLSX upstream: `/Users/eric8810/Code/extend-ui-upstream/react-xlsx` @ `f285a1c1a2a02e441a2e1f56e2fa480a0a979502`.

Shared-browser observations at `1440×900`:

| Page | URL | Current observation | Local implication |
|---|---|---|---|
| Upstream DOCX playground | `http://localhost:5173/` | No page-level overflow. Toolbar exposes grouped document/theme/editing controls: Document, Theme, Undo/Redo, Body, Calibri, font size, line spacing, Bold/Italic/Underline/Strike, script controls, text/highlight colors, link, alignment, lists, columns, pages, border, image/table, zoom, import/download, show edits/comments, read-only. | Local DOCX Editor now has an honest smaller accepted surface: grouped undo/redo, heading, B/I/U, theme, actual page count, and run-preserving edits. Missing upstream-rich controls remain intentional gaps unless user requests full parity. |
| Upstream XLSX playground | `http://localhost:5174/` | No page-level overflow. Workbook shell exposes title/status area, Download, theme/customize controls, ribbon tabs Home/Insert/Page Layout/Formulas/Data/View, grouped ribbon actions, formula/name input area, and empty workbook upload state. | Local XLSX Viewer now follows this intent with workbook title/status, ribbon-like tabs/groups, grid headers/gridlines, selection/editing, read-only, zoom, upload/download, and explicit fixture states. |
| Local DOCX Viewer | `http://127.0.0.1:5002/#/docx-viewer` | No page-level overflow; default sample renders 6 `.docx-page` surfaces and readable legal-contract content. | Addresses rejected pagination/page-surface blocker. |
| Local DOCX Editor | `http://127.0.0.1:5002/#/docx-editor` | No page-level overflow; toolbar and editor page are styled; default page count is `Page 1 / 1`; mixed styled run edit behavior was browser-tested separately. | Addresses rejected whole-line rewrite/style-loss blocker for the current scoped editor demo. |
| Local XLSX Viewer | `http://127.0.0.1:5002/#/xlsx-viewer` | No page-level overflow; workbook ribbon, grid headers, sheet tabs, 220+ visible cells, zoom/read-only/upload controls present. | Addresses rejected missing selection/edit/read-only fixture-state blockers for current scope. |

### Still not complete

The current rerun still cannot be marked complete because these requirements remain pending:

- Final command gates need to be rerun after the latest evidence-document updates.
- User page-by-page confirmation is still required by `docs/visual-acceptance-handoff.md`.

### Final command gate rerun — 2026-07-06

After the current code and evidence-document updates, the following combined gate passed:

```bash
pnpm typecheck && pnpm build && uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && git diff --check
```

Notes:

- `pnpm typecheck` passed for workspace packages.
- `pnpm build` passed for core packages, Vue packages, and demo app.
- The material verifier passed and reported expected valid DOCX/PDF/XLSX/PNG/JSON fixtures plus corrupted negative fixtures.
- `git diff --check` passed.
- Vite still emitted the known dynamic/static import chunking warnings for `docx-core` and `xlsx-core`; these are warnings and the build exited successfully.

### Localhost shared-browser rerun — 2026-07-06

Because the shared preview panel could not display the IPv4-only `127.0.0.1:5002` server, a current local demo server was started with:

```bash
pnpm dev -- --host 0.0.0.0 --port 5000
```

The shared browser successfully displayed `http://localhost:5000/#/`, and the route/viewport matrix was rerun through `agent-browser` against that same URL. Evidence file: `/tmp/agentic-visual-evidence/route-viewport-localhost5000.jsonl`.

Results:

| Route | 1440×900 | 1280×720 | 768×1024 | 390×844 |
|---|---:|---:|---:|---:|
| Home `#/` | overflow=false | overflow=false | overflow=false | overflow=false |
| DOCX Viewer `#/docx-viewer` | overflow=false, 6 pages | overflow=false, 6 pages | overflow=false, 6 pages | overflow=false, 6 pages |
| DOCX Editor `#/docx-editor` | overflow=false, 1 page | overflow=false, 1 page | overflow=false, 1 page | overflow=false, 1 page |
| XLSX Viewer `#/xlsx-viewer` | overflow=false, 232 visible grid/table cells | overflow=false, 232 visible grid/table cells | overflow=false, 232 visible grid/table cells | overflow=false, 232 visible grid/table cells |
| PDF Viewer `#/pdf-viewer` | overflow=false, `1 / 4`, frame 635px | overflow=false, `1 / 4`, frame 540px | overflow=false, `1 / 4`, frame 690.188px | overflow=false, `1 / 4`, frame 540px |
| Components `#/components` | overflow=false | overflow=false | overflow=false | overflow=false |

Fixture-switching evidence file: `/tmp/agentic-visual-evidence/fixture-switching-localhost5000.jsonl`.

Fixture switching summary at `1440×900`:

- DOCX valid fixtures all loaded without page-level overflow: demo/legal contract 6 pages; invoice/report/chinese samples 1 page each; corrupted DOCX cleared pages and showed `invalid zip data`.
- XLSX valid fixtures all loaded without page-level overflow: financial model, sales table, charts/images, and large grid all rendered cells; corrupted workbook produced a negative fixture state without page-level overflow.
- PDF fixtures all loaded without page-level overflow: sample `1 / 4`, scanned invoice `1 / 3`, rotated pages `1 / 2`, large contract `1 / 31`; corrupted PDF showed `Unable to load PDF document.` and `1 / —`.

### Shared-browser clean-route and interaction spot checks — 2026-07-06

Using the shared browser preview at `http://localhost:5000`, each route was opened after the latest rebuild/evidence updates and checked with the browser `errors` and `console` buffers:

- `/#/`: no page errors. Console output only contained DimCode inspect preload messages during the first inspection pass; after buffer clear and route reload there were no business console errors.
- `/#/docx-viewer`: no page errors and no console output.
- `/#/docx-editor`: no page errors and no console output.
- `/#/xlsx-viewer`: no page errors and no console output.
- `/#/pdf-viewer`: no page errors and no console output.
- `/#/components`: no page errors and no console output.

Additional shared-browser manual interaction spot checks:

- DOCX Editor: inserted `中英123` into an editable paragraph with `document.execCommand('insertText')`; the new text appeared, Undo became enabled after Vue updated, the existing bold `paragraph` run remained bold, Undo removed the inserted text, and Redo restored it.
- XLSX Viewer: CDP mouse events selected a visible cell; the cell background changed to active blue with a blue outline. Double-clicking a cell opened an inline editor, entering `999` and pressing Enter committed the value. Dragging across cells produced visible selected/active range styling across the grid.
- PDF Viewer: Next page changed the indicator to `2 / 4`; search for `page` reported `7 matches; opened page 1`; Zoom In updated the zoom combobox/iframe hash to `125%`; Rotate applied a 90-degree iframe transform matrix.

These checks replace the failed `/tmp/agentic-visual-evidence/interaction-gates-localhost5000.sh` attempt, which had JavaScript quoting errors and is not used as evidence.

### Final acceptance audit status — 2026-07-06

Current evidence now proves the technical/browser portions that can be independently verified by the agent:

- Upstream DOCX/XLSX playgrounds were opened and compared against local DOCX/XLSX routes.
- All six local routes were opened in the shared browser/runtime and retested at all required viewports.
- Fixture switching was tested for all listed DOCX/XLSX/PDF samples, including corrupted fixtures and recovery states where applicable.
- No page-level horizontal overflow was observed in the current route/viewport matrix.
- Clean shared-browser route checks reported no page errors and no business console errors.
- Key interaction gates were manually verified in the shared browser for DOCX editor editing/undo/redo, XLSX selection/edit/range styling, and PDF pagination/search/zoom/rotate.
- Final command gates passed after evidence updates: `pnpm typecheck`, `pnpm build`, material verifier, and `git diff --check`.

Remaining non-agent-verifiable requirement:

- `docs/visual-acceptance-handoff.md` requires explicit user confirmation for each page after seeing the working shared browser state. This remains pending and prevents marking the phase as Accepted/complete.

Acceptance status by page before user confirmation:

| Page | Agent technical/browser status | User confirmation |
|---|---|---|
| Home | Ready for review | Pending |
| DOCX Viewer | Ready for review | Pending |
| DOCX Editor | Ready for review | Pending |
| XLSX Viewer | Ready for review | Pending |
| PDF Viewer | Ready for review | Pending |
| Components | Ready for review | Pending |

### User-reported blocker fixes — 2026-07-06

Follow-up from user visual review fixed three accepted-blocking issues:

1. DOCX Viewer `report-with-image.docx` image was missing.
   - `packages/docx-core/src/wasm.ts` now parses `word/_rels/document.xml.rels`, resolves drawing `r:embed` image relationships, converts binary assets to data URLs, and creates `ImageRunNode` entries with OOXML extent-derived dimensions.
   - Shared browser verification on `http://localhost:5000/#/docx-viewer`: selecting `Report with image and rich text` renders one image with `naturalWidth=1000`, `naturalHeight=1330`, visible size about `499×664`, and 2 rendered pages.
2. XLSX Viewer ribbon tabs appeared clickable but did nothing.
   - `packages/vue-xlsx/src/XlsxViewer.vue` now tracks the active ribbon tab. Home shows real actions; Insert/Page Layout/Formulas/Data/View switch to explicit explanatory ribbon content instead of silently doing nothing.
   - Shared browser verification: clicking `Insert` sets `aria-selected=true`, applies active styling, and displays `Use the upload control above to insert/open a workbook fixture.`
3. XLSX grid/table cells and table surface should not inherit rounded shadcn table styling.
   - `packages/vue-xlsx/src/XlsxViewer.vue` now forces the grid table, headers, cells, and grid-scoped controls to square spreadsheet styling.
   - Shared browser verification: `.xlsx-viewer-table` border radius is `0px` and visible grid `td` border radius is `0px`.
4. Components annotation/floating marker visuals should not look like shadcn cards/buttons.
   - `packages/vue-extend/src/components/BoundingBoxCitations.vue` and `packages/vue-extend/src/components/LayoutBlocks.vue` now use square, crosshair-style overlay boxes and flat annotation-list rows instead of rounded shadcn component surfaces.
   - Shared browser verification: `.citation-box`, `.citation-list-item`, `.layout-block-box`, and `.layout-block-card` all report `border-radius: 0px`; overlay cursors are `crosshair`.

Regression gates after these fixes passed:

```bash
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py
git diff --check
```

### XLSX upstream ribbon parity implementation rerun — 2026-07-06

After the user rejected explanation-only / disabled ribbon behavior as non-compliant with the 1:1 upstream objective, the XLSX viewer was reworked again against `react-xlsx/apps/playground/src/App.tsx`.

Implemented in `packages/vue-xlsx/src/XlsxViewer.vue`:

- Upstream tab structure remains: `Home`, `Insert`, `Page Layout`, `Formulas`, `Data`, `View`.
- Home now exposes real groups matching upstream intent: `Clipboard`, `Font`, `Alignment`, `Number`, `Styles`, `Cells`.
- Insert now exposes real groups: `Workbook`, `Open`, `Source`.
- Page Layout now exposes real groups: `Theme`, `Export`.
- Formulas now exposes real groups: `Calculation`, `Defined Names`.
- Data now exposes real groups: `Tables`, `Refresh`, `Workbook`.
- View now exposes real groups: `Zoom`, `Sheets`, `Display`.
- Added upstream-like name box and formula bar (`fx`) with Enter commit / Escape cancel behavior.
- Added local style application for font family/size, B/I/U/S, text/fill color, alignment, number format, presets, borders, merge/unmerge, highlight, and document-dark view state.

Implemented in `packages/vue-xlsx/src/composables.ts`:

- Added `loadWorkbookFromBuffer()` controller API.
- Converted display filename to reactive state.
- Implemented CSV export instead of leaving it as noop.
- Implemented active sheet deletion.
- Implemented cell display/formula getters.
- Implemented cell value / selected formula mutation and basic undo/redo over cell edits.

Shared-browser verification at `http://localhost:5000/#/xlsx-viewer`:

- Ribbon tabs rendered exactly as `Home`, `Insert`, `Page Layout`, `Formulas`, `Data`, `View`.
- Home groups rendered as `Clipboard`, `Font`, `Alignment`, `Number`, `Styles`, `Cells`.
- Selecting cell `A2`, editing through the formula bar to `=1+2`, and committing with Enter updated the grid and showed status `Updated A2`.
- Clicking each non-Home tab updates `aria-selected=true` and renders the expected group labels:
  - Insert: `Workbook`, `Open`, `Source`.
  - Page Layout: `Theme`, `Export`.
  - Formulas: `Calculation`, `Defined Names`.
  - Data: `Tables`, `Refresh`, `Workbook`.
  - View: `Zoom`, `Sheets`, `Display`.
- Insert `+ Sheet` changed sheet tabs from `Assumptions`, `P&L`, `Notes` to `Assumptions`, `P&L`, `Notes`, `Sheet4` and showed status `Sheet added`.
- XLSX table/cell radius remained square (`0px`).
- Shared browser reported no page errors; console output only contained Vite/DimCode inspection logs during dev-server hot reload.

This supersedes the earlier explanation-only ribbon evidence, which should be treated as rejected.

### Shared-browser acceptance rerun — 2026-07-06 current worktree

This rerun uses the current worktree after the XLSX ribbon/formula work and the DOCX editor undo render fix. It supersedes older historical/rejected evidence only for the specific checks below; explicit user page confirmation is still pending and remains required before final acceptance.

Viewport matrix in shared browser via CDP device metrics:

| Viewport | Result |
|---|---|
| 1440×900 | Home, DOCX Viewer, DOCX Editor, XLSX Viewer, PDF Viewer, Components all opened; page-level horizontal overflow=false; no visible `.error`/`role=alert` on valid default state. |
| 1280×720 | Same six routes opened; page-level horizontal overflow=false; no valid-state route errors. |
| 768×1024 | Same six routes opened; page-level horizontal overflow=false; wide document/grid surfaces stayed inside internal scroll areas. |
| 390×844 | Same six routes opened; page-level horizontal overflow=false; controls/content remained reachable via wrapping/internal scroll. |

Fixture switching / corrupted recovery rerun:

| Route | Fixtures tested | Browser result |
|---|---|---|
| DOCX Viewer | `demo.docx`, `legal-contract.docx`, `invoice-table.docx`, `report-with-image.docx`, `chinese-mixed.docx`, `corrupted.docx`, then `demo.docx` again | Valid fixtures loaded with correct `Loaded:` state. `demo`/`legal-contract` rendered 6 `.docx-page` surfaces; `invoice-table` rendered table content; `report-with-image` rendered an image; corrupted fixture showed `invalid zip data` with 0 pages; recovery back to `demo.docx` restored 6 pages. |
| XLSX Viewer | `financial-model.xlsx`, `sales-table.xlsx`, `charts-images.xlsx`, `large-grid.xlsx`, `corrupted.xlsx`, then `financial-model.xlsx` again | Valid fixtures loaded with expected sheet tabs/cells: financial model sheets `Assumptions`, `P&L`, `Notes`; sales data single sheet; dashboard single sheet; large grid rendered 30,500 visible cells in current viewport/query. Corrupted workbook showed `Failed to load workbook: invalid zip data`, cleared sheet tabs/cells, and recovery back to `financial-model.xlsx` restored grid/sheets. |
| PDF Viewer | `sample.pdf`, `scanned-invoice.pdf`, `rotated-pages.pdf`, `large-contract.pdf`, `corrupted.pdf`, then `sample.pdf` again | Valid fixtures loaded iframe and thumbnail counts 4/3/2/31 respectively. Corrupted PDF cleared iframe/thumbnails. Recovery back to `sample.pdf` restored iframe and 4 thumbnails. |

Focused interaction checks:

- XLSX formula bar: selected `A2`, entered `=1+2`, committed with Enter; grid updated and status showed `Updated A2`.
- XLSX double-click edit / Enter / undo / redo: double-clicked value cell `420`, inline editor opened; entered `12345`; Enter committed; Undo removed the edit; Redo restored it.
- XLSX Insert `+ Sheet`: changed sheet tabs from `Assumptions`, `P&L`, `Notes` to include `Sheet4` and status `Sheet added`.
- XLSX read-only: page read-only checkbox blocked inline editor opening on double click.
- XLSX tab groups: non-Home tabs rendered expected groups: Insert=`Workbook/Open/Source`; Page Layout=`Theme/Export`; Formulas=`Calculation/Defined Names`; Data=`Tables/Refresh/Workbook`; View=`Zoom/Sheets/Display`.
- PDF search: searching `Page` on `sample.pdf` produced `7 matches; opened page 1` and a visible locator `Search result opened on page 1`.
- PDF frame sizing: default sample iframe measured about 986×635 at 1440×900, satisfying non-collapsed usable frame check.
- DOCX Editor typing scope and undo/redo: inserted `中英123` into the second editable paragraph; only that paragraph changed. Undo restored the previous paragraph text and enabled redo; redo restored the edit. This required adding a paragraph text keyed render in `packages/vue-docx/src/DocxEditor.vue` so controller undo/redo state is reflected in contenteditable DOM.

Current known final blocker:

- User page-by-page visual confirmation is still pending for Home, DOCX Viewer, DOCX Editor, XLSX Viewer, PDF Viewer, and Components. Do not mark final completion until the user confirms the shared-browser state.

### Clean console/page-error rerun — 2026-07-06

After clearing the browser console and page-error buffers, the six local routes were opened sequentially in the shared browser at `1440×900`:

- `/#/`
- `/#/docx-viewer`
- `/#/docx-editor`
- `/#/xlsx-viewer`
- `/#/pdf-viewer`
- `/#/components`

Result:

- Each route displayed the expected page title.
- No valid-state visible `.error` / `role=alert` messages were present.
- Page-level horizontal overflow remained false for each route.
- `browser.errors` returned no page errors.
- `browser.console` returned no console output after the clean route pass.

Remaining non-technical gate:

- The handoff still requires explicit user confirmation for each page. Current technical evidence is ready for user review, but final acceptance/goal completion remains pending until the user confirms the six pages in the shared browser.

### Components focused interaction rerun — 2026-07-06

The Components route was re-opened in the shared browser and operated beyond static DOM checks:

- SignaturePad: pointer-event drawing over the canvas changed status to `Status: Has signature`; Clear returned it to `Status: Empty`; Check empty confirmed empty state.
- FileUpload: a browser-created `tiny.pdf` file was selected through the enabled file input and the page showed `Uploaded: tiny.pdf`.
- FileUpload invalid type: selecting `bad.exe` produced the visible validation message `"bad.exe" is not an accepted file type`.
- Disabled FileUpload: disabled `.xlsx` upload input remained disabled in DOM.
- BoundingBoxCitations: clicking `Vendor` selected field data and surfaced `Northwind Industrial Supplies` in the selected-field panel.
- LayoutBlocks: clicking the `MASTER SERVICES AGREEMENT` block selected the OCR/layout block state.
- Tooltip: hovering `Hover top` produced visible tooltip text `Tooltip appears on hover after a short delay.`.
- Focus styling: key component controls expose visible focus styling via component-level focus-visible styles / focus box shadows. The browser automation environment did not advance focus with synthetic Tab reliably, so this evidence is limited to programmatic focus/style inspection and visible component focus CSS rules, not a full human keyboard traversal acceptance.

This strengthens the Components page evidence but does not replace the final required user visual confirmation.

### DOCX Editor toolbar parity extension — 2026-07-06

To reduce the remaining upstream editor-toolbar gap without adding fake controls, the DOCX editor toolbar was extended with controller-backed actions only:

- Added Strikethrough (`S`) via `editor.toggleStrike()`.
- Added paragraph alignment controls: left, center, right, justify via `editor.setParagraphAlignment(...)`.
- Added paragraph operations: insert paragraph, duplicate paragraph, remove paragraph via existing controller methods.
- Added `⬇ DOCX` download using `editor.exportDocx()` and a real browser download anchor.
- Propagated paragraph alignment through the core layout model by adding `LayoutBlock.align` and setting it from `paragraph.style?.align` in `paragraphToLayout()`.

Shared-browser retest on `/#/docx-editor`:

- Toolbar now showed: `↶`, `↷`, `B`, `I`, `U`, `S`, `L`, `C`, `R`, `J`, `+ ¶`, `Copy ¶`, `Del ¶`, `☾`, `⬇ DOCX`.
- Clicking a paragraph and pressing center alignment changed that paragraph's computed `text-align` to `center` while other paragraphs stayed `left`.
- Duplicate paragraph increased editable paragraph count from 3 to 4.
- Remove paragraph reduced editable paragraph count back from 4 to 3.

These changes do not claim full upstream DOCX editor parity for comments/track-changes/rich selection/table/image editing, but they remove several obvious visible toolbar-action gaps using real implemented behavior.

### Final gate rerun — 2026-07-06 current runtime spot-check

After re-reading `docs/visual-acceptance-handoff.md`, the current shared-browser runtime was spot-checked again at `1440×900` across all six accepted routes:

- Home: opened `/#/`; title/navigation/API map present; page-level horizontal overflow=false; no valid-state alerts.
- DOCX Viewer: opened `/#/docx-viewer`; default `demo.docx` loaded; 6 DOCX page surfaces detected; page-level horizontal overflow=false; no valid-state alerts.
- DOCX Editor: opened `/#/docx-editor`; toolbar actions visible (`B/I/U/S`, alignment, paragraph actions, theme, export); 3 editable paragraphs detected; page-level horizontal overflow=false; no valid-state alerts.
- XLSX Viewer: opened `/#/xlsx-viewer`; default financial model loaded with ribbon tabs, sheet tabs, and 200+ visible cells; page-level horizontal overflow=false; no valid-state alerts.
- PDF Viewer: opened `/#/pdf-viewer`; default sample PDF iframe measured about 986×635 with page controls and thumbnails/page buttons; page-level horizontal overflow=false; no valid-state alerts.
- Components: opened `/#/components`; SignaturePad/FileUpload/FileThumbnail/BoundingBox/LayoutBlocks/Spinner/Tooltip demos present; page-level horizontal overflow=false; no valid-state alerts.

Current final command gate was rerun from repository root and passed:

```bash
pnpm typecheck && pnpm build && uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && git diff --check
```

Result: `pnpm typecheck` passed, `pnpm build` passed, material verification passed for DOCX/PDF/XLSX/PNG/JSON/corrupted fixtures, and `git diff --check` passed. Vite still reports known mixed dynamic/static import warnings during demo build, but the build exits successfully.

The only known remaining final-completion requirement not satisfied by agent-verifiable evidence is explicit user page-by-page visual confirmation for Home, DOCX Viewer, DOCX Editor, XLSX Viewer, PDF Viewer, and Components.

### Final criteria / feature-gate cross-check — 2026-07-06

This table maps the handoff final criteria to current evidence so the final state is not inferred from a narrow smoke check. Historical rejected sections above remain historical only; use the current 2026-07-06 rerun sections for acceptance evidence.

| Handoff criterion / gate | Current status | Current evidence | Remaining requirement |
|---|---|---|---|
| Upstream React DOCX/XLSX references opened and used | Technical ready | Upstream refresh/comparison notes and local parity reruns in this document and `docs/upstream-visual-comparison-evidence.md`. | User must still accept local result. |
| Six local routes opened and operated in shared browser | Technical ready | Viewport matrix, fixture recovery, focused DOCX/XLSX/PDF/Components interaction reruns, and current runtime spot-check. | User must still visually confirm each page. |
| Shared UI gate: styled controls, visible states, no fake controls | Technical ready | Raw-control blockers were replaced with shadcn/upstream-style shells; visible non-working controls were either implemented or disabled with clear context; current runtime check found no valid-state alerts/overflow. | User visual judgment still required. |
| Fixture-switching gate | Technical ready | DOCX/XLSX/PDF listed valid fixtures plus corrupted fixtures and recovery back to valid state were browser-tested. | None known. |
| DOCX Viewer gate | Technical ready | Default and fixture-rerun evidence shows multi-page surfaces, correct loaded status, table/image/CJK coverage, corrupted error and recovery. | User confirmation pending. |
| DOCX Editor gate | Scoped technical ready | Current page is now labeled as a scoped DOCX Editing Demo; toolbar actions are backed by controller behavior; typing scope, undo/redo, theme, alignment, paragraph operations, and export were browser-tested. Full comments/track-changes UI, rich range selection, and table/image editing are documented as not claimed in the accepted surface. | User/project agreement required through page confirmation. |
| XLSX Viewer gate | Technical ready | Sheet tabs, grid cells, selection/editing, Enter/Escape/Tab/undo/redo, readonly, zoom, upload/download-capable controls, fixture switching, corrupted recovery, and ribbon grouping were browser-tested/documented. | User confirmation pending. |
| PDF Viewer gate | Technical ready | Usable iframe sizing, next/previous, thumbnails/page buttons, zoom/rotate, search locator feedback, download, corrupted clearing/recovery documented. | User confirmation pending. |
| Four viewport matrix | Technical ready | `1440×900`, `1280×720`, `768×1024`, and `390×844` route matrix documented with page-level overflow=false. | User confirmation pending. |
| No page-level overflow / no page errors / no business console errors | Technical ready | Current clean console/page-error rerun and current runtime spot-check document no page errors and no valid-state alerts/overflow. Vite HMR logs during development are not business errors. | None known. |
| Evidence honesty | Technical ready | Prior rejected evidence and failed scripts are explicitly marked historical/rejected and not reused as final proof; scoped DOCX Editor gaps are documented. | User confirmation pending. |
| Final commands | Passed | `pnpm typecheck && pnpm build && uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && git diff --check` passed on current worktree. | Rerun after any further code change. |
| User explicit page-by-page confirmation | Pending / blocker | Evidence tables list all six pages as ready for review. | Must receive explicit user confirmation before goal completion. |

### Clean shared-browser reload after scope-note change — 2026-07-06

After the DOCX Editor route was honestly scoped as `DOCX Editing Demo`, the shared browser console and page-error buffers were cleared and all six routes were opened again at `1440×900`:

| Route | Result |
|---|---|
| `/#/` | Home title/navigation rendered; page-level horizontal overflow=false; no valid-state alerts. |
| `/#/docx-viewer` | DOCX Viewer rendered default `demo.docx` with 6 DOCX page surfaces; page-level horizontal overflow=false; no valid-state alerts. |
| `/#/docx-editor` | `DOCX Editing Demo — Verification` heading and scope note rendered; 3 editable paragraphs present; page-level horizontal overflow=false; no valid-state alerts. |
| `/#/xlsx-viewer` | XLSX Viewer rendered default financial model with 232 queried cells and sheet tabs `Assumptions`, `P&L`, `Notes`; page-level horizontal overflow=false; no valid-state alerts. |
| `/#/pdf-viewer` | PDF Viewer rendered `sample.pdf` iframe at about 986×635; page-level horizontal overflow=false; no valid-state alerts. |
| `/#/components` | Components page rendered SignaturePad canvas and component controls; page-level horizontal overflow=false; no valid-state alerts. |

After the route pass, `browser.errors` returned no page errors and `browser.console` returned no console output. The remaining final-completion blocker is still explicit user page-by-page visual confirmation.

## Strict upstream parity re-audit — 2026-07-06

After a stricter review prompted by the question “is everything really upstream-identical and feature-complete?”, the answer is **no**. The current Vue demo is a usable/scoped port for the visual acceptance surface, but it must not be represented as full 1:1 upstream feature parity. The following gaps are current and intentional only if the project/user accepts the scoped surface:

| Area | Upstream capability observed | Current Vue state | Acceptance implication |
|---|---|---|---|
| DOCX Editor toolbar | React playground exposes paragraph style previews, font family/size, line spacing, bold/italic/underline/strike, superscript/subscript, text color, highlight, link, alignment, lists, columns status, page thumbnails, border presets, image insert, table insert/edit, zoom, import/export, tracked changes/comments toggles, read-only mode, context menus, form-field editing. | Vue demo implements a much smaller controller-backed toolbar: undo/redo, heading/body, B/I/U/S, paragraph alignment, paragraph insert/duplicate/delete, theme, export. Many controller methods are API shells/no-ops. | Not full upstream parity. Page is now labeled `DOCX Editing Demo` and scope note says full comments/track-changes/rich selection/table/image editing are not claimed. |
| DOCX comments / tracked changes | React playground renders/toggles actual comments and tracked-change cards. | Vue `useDocxComments` / `useDocxTrackChanges` expose toggle state but comments/trackedChanges arrays are empty and no authoring/rendered cards exist. | Must be documented as shell/scoped, not “1:1”. |
| DOCX paragraph styles / line spacing / borders / thumbnails | React has real selectors, hover previews, line spacing changes, border preset application, independent page thumbnail rendering. | Vue has API shape but `setParagraphStyle`, `setLineSpacing`, `applyBorderPreset`, thumbnail painting/rendering are placeholders/no-ops. | Not complete upstream parity. |
| DOCX Viewer rendering fidelity | React viewer supports richer editor/viewer rendering surfaces, virtualization/thumbnails/context hooks. | Vue viewer renders parsed model pages/tables/images/text but styling/layout fidelity is simplified. | Acceptable only as scoped visual demo, not full upstream viewer equivalence. |
| XLSX Viewer advanced editing | React playground has real controller hooks for style application, defined names, merge/unmerge, export XLSX/CSV, thumbnails, table sort menus, worker/canvas/highlight/resize options. | Vue surface implements visible ribbon/local editing/selection/zoom and some local style/merge/status behavior; underlying controller still has multiple no-ops (defined ranges, fill, recalc, merge/unmerge, resize, paste structured data, chart/image edits, style persistence). `exportXlsx()` currently delegates to JSON-like `download()` rather than producing a real workbook. | Not full upstream parity. Visible demo should be treated as scoped; advanced controls must be implemented, hidden, or clearly scoped before strict 1:1 acceptance. |
| Home/API claims | Previous page text claimed “1:1 API compatible” and many rows said `✅ 1:1`. | Updated to say parity is partial/tracked; API map now marks scoped/partial/shell/placeholder statuses. | Evidence and UI no longer overclaim full parity. |

Conclusion: current technical gates may pass, but full objective completion under a strict “1:1 upstream, no missing features” interpretation is **not proven** and would require additional implementation, especially in DOCX Editor and advanced XLSX APIs.
