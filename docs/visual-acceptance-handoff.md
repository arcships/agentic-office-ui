# Visual Acceptance Handoff

## Current state

The repository has completed the first engineering verification pass:

- Real DOCX / XLSX / PDF / PNG / JSON fixtures exist under `apps/demo/public/samples`.
- Demo routes exist for Home, DOCX Viewer, DOCX Editor, XLSX Viewer, PDF Viewer, and Components.
- Material verification passes with `scripts/verify_test_materials.py`.
- Workspace `pnpm typecheck` and `pnpm build` pass.
- Initial browser smoke checks found and fixed several data-loading, error-state, editing, upload, and responsive overflow issues.
- Upstream Extend UI attribution is documented in `README.md` and `docs/upstream-extend-ui.md`.

The next phase is a user-visible shadcn/components visual acceptance pass. This phase should treat the shared browser as the source of truth and should not rely on code inspection alone.

## Goal for the next phase

Perform a full visual and interaction acceptance review of the Vue Extend UI demo using shadcn/ui + components standards. Open every page in the shared browser, operate every visible control, check every important state, fix visual or interaction issues immediately, and repeat until the user confirms each page.

The completion target is:

> Every route and component looks and behaves correctly in the shared browser across desktop, laptop, tablet, and mobile viewports, with shadcn-style visual consistency, real user interaction coverage, and explicit user confirmation per page.

## Source documents

Use these documents as references:

- `docs/shadcn-components-browser-acceptance.md` — detailed acceptance standard and checklists.
- `docs/component-browser-verification-plan.md` — original browser verification scope.
- `docs/component-browser-verification-evidence.md` — engineering verification evidence from the previous pass.
- `docs/upstream-extend-ui.md` — upstream attribution and sync policy.

## Demo environment

Project root:

```text
D:\code\ui\vue-extend
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

## Acceptance approach

### Core rule

Use the shared browser for all acceptance decisions. The user needs to see the actual interface and actual interactions. Code review can support diagnosis, but it cannot replace visual acceptance.

### Recommended loop per page

1. Open the route in the shared browser.
2. Set viewport to `1440×900`.
3. Inspect the page visually against shadcn-style standards.
4. Perform the page-specific interactions listed below.
5. Check keyboard/focus behavior.
6. Check error, empty, loading, success, disabled, selected, readonly, and dirty states where applicable.
7. Repeat at `1280×720`, `768×1024`, and `390×844`.
8. Record issues in this document or a follow-up evidence document.
9. Fix issues immediately.
10. Reopen the page and repeat the interaction that failed.
11. Ask the user to confirm the page before moving on.

### Viewport matrix

| Viewport | Purpose | Acceptance focus |
|---|---|---|
| 1440×900 | Desktop default | Layout rhythm, page density, shell alignment, full interaction path |
| 1280×720 | Laptop short viewport | Vertical density, sticky/scroll areas, toolbar wrapping |
| 768×1024 | Tablet | Internal scroll containers, stacked controls, touch target size |
| 390×844 | Mobile | No page-level horizontal scroll, readable hierarchy, usable controls |

## shadcn/components visual standards

Check each page for:

- Clear title/subtitle hierarchy.
- Consistent spacing between header, controls, status, viewer/content, and tables.
- Consistent border radius on cards, buttons, selects, inputs, upload zones, and viewer shells.
- Border and muted background usage matching shadcn visual language.
- Button variants that communicate primary, secondary, destructive, icon, disabled, hover, active, and focus states.
- Inputs/selects with visible focus rings and stable sizing.
- Tables and grids with readable headers, cell states, selected states, and overflow handling.
- Empty/loading/error states that are visually stable, readable, and specific.
- Tooltip and popover-like surfaces that feel intentional and do not clip unexpectedly.
- Keyboard focus visibility throughout the route.

## Page-by-page handoff checklist

### 1. Home

Route: `/#/`

Visual review:

- Main title and subtitle align and read well.
- Navigation cards have consistent spacing, border, radius, hover, and status badges.
- API map table has shadcn-like table styling and readable density.
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

Visual review:

- Controls area looks like a shadcn form/control row.
- Status area clearly communicates loaded file and expected coverage.
- Viewer shell has clear border, radius, background, and internal scrolling.
- Rendered DOCX page feels like a document surface and does not break the app layout.
- Error state for corrupted DOCX is specific and visually integrated.

Interactions:

- Load each sample: `demo.docx`, `legal-contract.docx`, `invoice-table.docx`, `report-with-image.docx`, `chinese-mixed.docx`, `corrupted.docx`.
- Upload a valid DOCX fixture.
- Scroll the document.
- Check table/image/chinese content visibility.
- Tab through select, button, and file input.

Expected user confirmation:

> DOCX Viewer visual, loading, upload, scroll, and error-state experience accepted.

### 3. DOCX Editor

Route: `/#/docx-editor`

Visual review:

- Editor toolbar feels like a cohesive shadcn toolbar.
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
- Verify mobile/tablet internal scroll behavior.

Expected user confirmation:

> DOCX Editor editing, toolbar, focus, undo/redo, and theme experience accepted.

### 4. XLSX Viewer

Route: `/#/xlsx-viewer`

Visual review:

- Toolbar, sheet tabs, grid headers, and grid cells feel consistent.
- Selection outline is visible and not visually noisy.
- Edit input is aligned with the cell and readable.
- Readonly and corrupted workbook states are clear.
- Wide sheets use internal scrolling instead of page-level overflow.

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

> XLSX Viewer grid, tabs, editing, keyboard, zoom, readonly, upload, and error-state experience accepted.

### 5. PDF Viewer

Route: `/#/pdf-viewer`

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

Visual review:

- Component demo cards have consistent spacing and hierarchy.
- Controls and status text are aligned and easy to scan.
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

## Issue record template

Use this format for every issue:

```text
ID:
Page/component:
Viewport:
Steps:
Actual:
Expected:
Type: visual / interaction / responsive / keyboard / a11y / state / data
Severity: blocker / high / medium / low
Files changed:
Fix summary:
Retest result:
User confirmation: pending / accepted
```

## Known areas needing special attention

- The previous phase verified engineering behavior, but user-visible shadcn acceptance still needs to happen page by page.
- DOCX and XLSX naturally contain wide document/grid content. The acceptance standard is internal scrolling with clear affordance, not page-level horizontal overflow.
- PDF viewer currently uses browser iframe rendering. Acceptance should focus on visible controls, state, fallback/error handling, and user-perceived behavior.
- File upload and download paths should be exercised in the browser, not only checked by code.
- Keyboard focus styling should be reviewed visually; existing smoke checks did not fully evaluate focus aesthetics.
- Hover/focus/disabled/active states need actual browser inspection.

## Commands to run after fixes

Run after each meaningful code change or at the end of a page batch:

```bash
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts\verify_test_materials.py
```

## Final completion criteria

The next phase is complete when:

- All six routes have been opened and operated in the shared browser.
- All page-specific checklists above are completed.
- All four viewports pass for every route.
- No page-level horizontal overflow remains.
- No browser page errors or business console errors remain.
- All discovered visual and interaction issues are fixed and retested.
- The user explicitly confirms each page.
- Final `pnpm typecheck`, `pnpm build`, and material verification pass.
