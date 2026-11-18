# Changelog

All notable changes to the `@arcships` package family are documented in this file.

## [0.7.0] - 2026-11-18

### DOCX fixture and acceptance coverage (#53)

- Add three DOCX fixtures that carry the structures the reported defects live in: `review-threads.docx` (comment threads with replies and a resolved thread, two comments on one paragraph, a range crossing three paragraphs, comments anchored in a table cell and in the header, a deleted table row, a row-height `trPrChange`, two footnotes and two endnotes), `table-structures.docx` (`vMerge` restart/continuation cells, `gridSpan`, a nested table, multi-paragraph cells and a `tblHeader` row that repeats across pages), and `sections-columns.docx` (portrait/landscape/two-column sections with their own headers and footers, a text box and a behind-text watermark).
- `scripts/generate_test_materials.py` gained `--only`, so a new fixture can be regenerated without rebuilding the unrelated binaries with a different toolchain; `scripts/verify_test_materials.py` now fails when a regenerated fixture loses its comments, threading, merges, header rows or sections.
- Add `tests/unit/docx-fixture-fidelity.test.mjs` and `tests/component/docx-fixture-fidelity.test.mjs` (registered in `scripts/ci/run-suite.mjs`): they assert comment threads and anchors, cross-paragraph ranges, one gutter card per comment per page with card numbers matching their inline markers, cards placed within one line of their anchor, deleted rows marked and listed once, row-height revisions anchored to their own row, merged/nested table column counts and cell text, repeating header rows, per-section headers/footers, note numbering, and an import → edit → export → re-import field diff that changes only the edited paragraph.
- Add the “Acceptance coverage map” to `docs/visual-acceptance-handoff.md`: every acceptance row carries an id and maps either to a real CI-registered test or to the browser suite that still has to accept it; the unit suite verifies the mapping so an unverified claim cannot read as accepted.
- `tests/blackbox/docx_annotations.py` and `tests/blackbox/fidelity_content.py` now check gutter placement against the anchor markers, thread visibility, cross-paragraph/cell/header anchors, merged/nested tables and per-section headers instead of only asserting that a card or a notes section exists. The insertion revision renders inline and is not ballooned, so the review fixture asserts two gutter cards instead of three, and the outer table's column count comes from its own `colgroup` instead of also counting the nested table's columns.

### XLSX page setup and printing (#20)

- Parse `pageSetup`, `printOptions`, `pageMargins`, `rowBreaks`/`colBreaks`, `sheetPr/pageSetUpPr` and the `_xlnm.Print_Area`/`_xlnm.Print_Titles` defined names into `XlsxSheetData`, with an engine fallback for `skipXmlParsing` sources.
- Add pure pagination to `@arcships/xlsx-core` (`buildWorksheetPrintLayout`, `buildPrintPageGrid`): paper size/orientation, usable margins, fit-to-page scaling, manual breaks, repeated title rows/columns, merged-cell spans and `pageOrder` produce a flat page list of cell ranges with the paper size.
- Add an Excel-style dashed page-break preview overlay to the grid plus page numbers, and a paginated `XlsxPrintPreview` with paper-sized page blocks, `@media print` output and `window.print()`.
- Expose `printLayout` / `printPages` / `getPrintLayout()`, `setPageBreakPreview()`, `openPrintPreview()` / `closePrintPreview()` and `print()` on the viewer controller; the toolbar gains page-break and print-preview buttons.
- Export writes `pageMargins`, `printOptions`, `pageSetUpPr`, manual breaks and print-area/title defined names back into the saved workbook (the engine drops several of them), so export → reload keeps the same pagination; unmodified workbooks still export byte-for-byte unchanged.

### XLSX worksheet rename (#9)

- Add `renameSheet(index, name)` and `renameActiveSheet(name)` to the controller, plus `getSheetNameError(name, index?)` for live validation. Names follow Excel's rules: trimmed and non-empty, at most 31 characters, no `[ ] : * ? / \\`, no leading or trailing apostrophe, and unique case-insensitively across worksheets *and* chartsheets; a rejected name leaves the workbook, the tab strip and the undo stack untouched. A rename keeps worksheet order, so `activeSheetIndex` and `activeTabIndex` do not move, and a case-only rename of the same sheet is allowed.
- `@dukelib/sheets-wasm` 0.1.17 exposes no rename setter (`Workbook` has `addSheet`/`removeSheet`/`sheetNames` but no name write), so the viewer renames through its own saved package: the workbook part's `<sheet name>` and every reference to that sheet are rewritten (worksheet formulas, conditional-format and data-validation formulas, table calculated-column and totals formulas, internal hyperlink locations, defined names, chart series references, pivot cache worksheet sources and `formula` attributes), and the result is parsed back into a fresh workbook model. History snapshots are byte-based, so the rename is still one undo step, redo works, and export/reload keep the new name and its cross-sheet formulas.
- The tab strip gains inline rename: double-click a worksheet tab to edit it, `Enter` commits, `Esc` cancels and blur commits; an empty or duplicate name keeps the editor open with a visible message. A tab context menu (and the `ContextMenu`/`Shift+F10` keys) offers the same Rename command, chartsheet tabs and read-only viewers expose no rename affordance, and the sheet's protection state does not block a rename, matching Excel.
- Export `validateWorksheetName`, `renameSheetInFormula`, `applyWorkbookSheetRename`, `XLSX_MAX_SHEET_NAME_LENGTH` and the `XlsxSheetNameIssue` type from `@arcships/xlsx-core`, so hosts can validate a name or apply the same rename to their own saved bytes without the viewer.

### XLSX chart type editing (#139)

- Add `changeWorkbookChartType()`: classic chart families rebuild the `c:plotArea` while carrying over series names, colours, data labels, point styles and axis formatting, and chartEx families (treemap, sunburst, histogram, pareto, box & whisker, waterfall, funnel, filled map) write a `cx:chartSpace` part with the `chartStyle`/`chartColorStyle` parts, relationships and content-type overrides Excel expects.
- Migrate classic ↔ chartEx in place: the drawing relationship type, the `cx:chart`/`c:chart` reference, `a:graphicData@uri`, the `mc:AlternateContent`/`mc:Fallback` wrapper and the chart content type are all kept consistent, and the chart id follows the migrated kind.
- Read chartEx parts from the package archive because the WASM workbook model does not expose them, so authored and exported modern charts survive re-parse, export and reload; histogram/pareto binning is no longer applied twice, and the declared series stay recoverable when a point-series layout collapses them for rendering.
- Expose the gallery to hosts and the viewer: `chartTypeOptions` (57 classic and modern entries), `chartTypeFamilies()`, `resolveChartTypeOption()`, `matchesChartType()` and `chartTypeUnavailable()`, plus `XlsxViewerController.changeChartType(id, type)`; the XLSX context menu shows a family/type submenu when a chart is right-clicked, and the change is undoable and included in exported XLSX bytes.

### XLSX structure editing, paste special and Name Manager

- Add `insertRows`/`deleteRows`/`insertColumns`/`deleteColumns` to the controller and both header surfaces (ribbon, header menus, row/column header context menus). The WASM engine exposes no structural API, so the viewer moves cell values, formulas, styles, row/column sizes and merges itself, rewrites A1 references (absolute references included, `#REF!` when a reference is deleted) and then shifts hyperlinks, comments, conditional-format ranges and formula thresholds, data validations, table ranges, frozen panes, defined names and image/chart/shape/form control anchors. Structural edits are one snapshot undo step and survive export/reload.
- Add `shiftWorkbookSheetStructure`, the pure range/cell/anchor/conditional-format/data-validation/sparkline/table remappers, `resolveWorksheetMergedRegions` and `applyWorkbookDefinedNames`/`readWorkbookDefinedNames` to `@arcships/xlsx-core`. Absolute relationship targets in worksheet rels are normalized so table parts with `/xl/tables/...` targets update and delete correctly.
- Add Paste Special (all / values only / formats only / transpose) to the context menu, ribbon and a dedicated dialog, plus `cutSelection`, `pasteSpecialFromClipboard` and `pasteSpecialStructuredData`. A cut moves the source range on the first paste and is one undo step.
- Add a Format Cells dialog with decimal places, thousands separator, currency/percent/date/text categories and a custom Excel format code (`setSelectionNumberFormat`).
- Add a Name Manager dialog backed by `getNamedRanges`/`setNamedRange`/`deleteNamedRange`; deleted names are written out of `xl/workbook.xml` on export because the engine cannot remove a defined name.
- Upgrade the cell context menu to Excel's item set (cut/copy/paste/paste special submenu, edit cell, clear contents, insert/delete row and column, merge & center/across/cells/unmerge, sort, select row/column/all) with shortcut hints, checked selection state and `↑`/`↓`/`Home`/`End`/type-ahead/`←`/`→`/`Esc` keyboard semantics. Row and column headers now open their own menu, and hovering a header exposes a column/row menu trigger in the grid.
- Follow-up audit fixes: loaded defined names (from `workbook.xml`) now shift even though the engine rejects redefining them, the Name Manager overlay carries the new reference through export/reload; unmodeled conditional-format rules (`expression`, `cellIs`, `top10`, …) get their `sqref` and `<formula>` rewritten on structure edits; formulas on other sheets that reference the edited sheet are rewritten (unqualified references on those sheets stay untouched); and both new structure test files are registered in `scripts/ci/run-suite.mjs`.

### XLSX authoring surface (#102)

- Write conditional-format and data-bar colours as 8-digit ARGB (`FFRRGGBB`) instead of the CSS `#rrggbb` form kept in the model; LibreOffice and Excel read the latter as black. Theme, indexed and automatic colours keep their `theme`/`indexed`/`auto` attributes, and the ARGB/hex input forms are both accepted.

- Add a ribbon group, context-menu entries and an `XlsxAuthoringDialog` for charts, comments, hyperlinks, tables, conditional formatting, data validation and freeze panes; `XlsxViewer` renders the dialog and hosts can mount it directly.
- Public controller commands mutate the workbook and are undoable: `insertChart`, `getCellComment`/`setCellComment`, `getCellHyperlink`/`setCellHyperlink`, `getConditionalFormatRules`/`addConditionalFormatRule`/`updateConditionalFormatRule`/`deleteConditionalFormatRule`, `createTable`/`removeTable`/`getTableAtCell`, `setSelectionDataValidation`, `freezePanes`/`setFreezePanes`/`toggleFreezePanesAtActiveCell`.
- Authoring writes OOXML on the source archive (worksheet `pane`, `hyperlinks`, `conditionalFormatting`, `dataValidations`, `legacyDrawing`, `tableParts`, plus comments/VML/table/chart/drawing parts) because the WASM model exposes no setters and its save pass drops most of them.
- `mergeWorkbookSheetAuthoringAssets()` merges those fragments into every saved workbook (export and undo snapshots) in schema order, keeps conditional-format rules the editor cannot model, and copies comments, tables and inserted charts; `insertWorkbookChart()` generates bar/line/pie chart parts plus their drawing and relationships, and chart type now reads from the chart XML instead of defaulting to column.
- Fix `mergeWorkbookImageAssets` crashing with `node.remove is not a function` when the engine re-emits a drawing, by using `parentNode.removeChild` in the image and chart export paths.

### DOCX measure-driven pagination iteration (#99)

- Run the full estimate → render → measure → repaginate loop in the viewer: rendered paragraph heights and table row heights feed back into `buildDocumentPageNodeSegments`, bounded by three passes with a tolerance-based convergence check and a signature history that stops oscillation.
- Replace the ad-hoc viewer feedback counter with `PaginationMeasurementLoop`, which reports `repaginating`, `converged`, `oscillating` or `exhausted` plus the iteration count.
- Treat `documentPageCount` as an advisory target (`LayoutSnapshot.pageCountTarget`, `enforced: false`); drop the default 1.02–1.3 virtual-page scaling candidates and keep the deprecated reconciliation/merge helpers for `0.x` compatibility only.
- Keep page surfaces on unchanged pages by reusing their segment/column arrays and render entries when only other pages repaginate, and restore the reader's page offset after a measurement pass moves page positions.
- Add the measurement loop, page reuse, scroll anchor and legal/invoice fixture acceptance tests (nine pages for `legal-contract.docx`, one for `invoice-table.docx`, every page exactly one paper tall).

### DOCX clipped page measurement (#135)

- Port the clipped body-rect reader from upstream (`createClippedBodyRectReader`): every measured element intersects the box of each ancestor whose overflow clips it (`hidden`, `clip`, `auto`, `scroll`, `overlay`), results are cached per element for the measurement pass, and screen-space rects are converted back to unscaled layout pixels so the values match `offsetHeight` at any zoom.
- Add `collectPageContentMeasurementDescendants` (skips `data-docx-pagination-ignore` subtrees), `resolveClippedBodyRenderedBottomPx` (the first caller of `resolveMeasuredBodyRenderedBottomPx`) and `resolveClippedMeasuredHeightPx` (clipped bottom first, box height as the fallback) to the root `@arcships/docx-core` entry.
- Use the clipped bottom for `DocxPageSurface`'s page measurement (`onPageMeasure`), `DocxPageWrapper`'s measured height and the `DocxPageBody` content measurement, so a table cell or fixed-height container that clips overflowing paragraphs no longer reports a taller page; `offsetHeight` stays the fallback when nothing can be measured and a page never drops below one paper.
- Mark the page's cover, border, floating-image and form-field overlay layers as `data-docx-pagination-ignore`, so page chrome can never extend the measured content height.
- Add the clipped-cell fixture and the reader/overlay unit tests plus viewer, wrapper and page-body measurement tests (measured height equals the visible height and the next page starts one paper plus the page-break gap below).

### XLSX formatting ribbon (#120)

- Add the missing ribbon formatting controls: a border preset dropdown (All / Outside / Left / Right / Top / Bottom / None) with a border colour picker, while the existing All-borders button stays as a shortcut. Border presets are drawn through the existing style command, so each edge is written to `styles.xml` and survives export and reload; `outside` patches only the perimeter cells, so the inside of the selection keeps its own borders and colours.
- Add `setSelectionBorders({ preset, style, color })` and the `XlsxBorderPreset` / `XlsxBorderPresetInput` types, so hosts can draw a preset (and set the line colour for every edge it touches) without dropping to `setRangeStyle`.
- Add `mergeAndCenterSelection()`, the default Merge action: it merges the selection and writes `horizontal: center` + `vertical: center` in one undo step. A single-cell selection is only centred, because Excel never writes a one-cell merge region. `mergeSelection()` keeps merging without touching alignment for hosts that want it.
- Fix undo of a style edit on a cell that had no explicit style: the parsed workbook default is OOXML-shaped (`{ theme }` / `{ rgb }` / `{ indexed }` colours), which `setCellStyleAt` rejects, and it only named the groups the source style carried, so a removal (bold, border, fill) stayed applied. The fallback style now translates colours to the engine input and spells out every group and property.

### DOCX annotation gutter

- Position tracked-change and comment cards against their text anchors in page space instead of rendering them as an unpositioned list at a fixed `top: 72px`, with stable document-order ties, collision-free stacking, and measured card heights.
- Counter-scale the annotation pane so cards keep a fixed screen size at every zoom level, and reserve a constant screen-space footprint beside the paper.
- Number every comment and draw a numbered marker plus leader line at its anchor, linking the highlighted range to its card; clicking a card, marker, or highlighted range activates the pair.
- Stop switching the pane to `position: static` below 1100px, which inflated the measured page height and broke multi-page scroll positions.
- Store `w:commentRangeStart` / `w:commentRangeEnd` on the paragraph model as UTF-16 character offsets, so a comment highlights exactly the commented characters instead of a whole run, and so highlights and gutter cards survive the edits that clear `sourceXml`.
- Remap those offsets on every text edit (typing, delete, find/replace, paragraph split and join) with an `anchorText` fallback when an edit cannot be mapped, and re-emit the range markers plus `w:commentReference` when serializing an edited paragraph.
- Turn the pane into a toggleable overlay drawer below 1024px instead of a column that leaves the paper, so the paper keeps the full viewport width on laptop screens.

### DOCX character typography

- Preserve OOXML underline/strike variants (`double`, `dotted`, `dash`, `wave`, `dstrike`) and expose them as CSS text-decoration style/thickness instead of flattening every run to a single solid line.
- Map run, paragraph and table borders through one OOXML-to-CSS line-style table so `wave`, `triple`, `thick` and compound gap styles stay visible instead of producing invalid CSS declarations.
- Model and render `w:caps`, `w:smallCaps`, `w:spacing`, `w:w`, `w:bdr`, text shadow/outline, run RTL and paragraph `w:bidi`/extended `w:jc` values.
- Keep ASCII, high-ANSI, eastAsia and complex-script fonts separately, build a per-glyph font stack and resolve major/minor theme fonts from `theme1.xml` during import.
- Read `w:lvlJc`/`w:suff` from `numbering.xml`, align list markers accordingly, and add decimalZero, Chinese, ideographic, enclosed and text numbering formats.

### XLSX drawing objects

- Honour picture `zIndex` inside per-pane stacking contexts and render rotation, flip and `srcRect` crop.
- Add drag, eight-way resize and arrow-key nudge for images and charts; pointer gestures preview locally and commit one history entry.
- Add the same move/resize/nudge interaction to shapes, plus Delete for images, charts and shapes backed by stable non-visual object identities.
- Expose programmable form-control CRUD (`addFormControl`/`updateFormControl`/`removeFormControl`) backed by worksheet controls, ctrlProps and content-type rewrites.
- Cull anchored objects by visible pane bounds and cache content-space rects by sheet and zoom.
- Render shape text from paragraphs and runs with alignment, insets, vertical alignment and flips.
- Make checkbox and radio controls interactive and write `TRUE`/`FALSE` to their `linkedCell`.
- Use one visible-column axis model for grid rendering and anchor write-back, including hidden rows and columns.
- Memoise chart SVG trees so scrolling does not rebuild chart DOM.

### XLSX worker robustness

- Treat cell-snapshot and row-batch cancellation as per-request aborts instead of terminating the shared Worker, and recover cleanly from Worker failures.
- Report non-cloneable Worker results as `WORKER_FAILED` immediately instead of waiting for the 30 s parse timeout.
- Fall back to main-thread parsing when a Worker load fails for a non-limit reason, and apply `maxParseMs` to the main-thread path as well.
- Keep charts refreshed after an edit from being overwritten by a late Worker chart-hydration result.
- Render Worker-mode sheet thumbnails from the transferred sheet snapshot instead of an unavailable main-thread workbook.
- Inflate only `xl/media/*` for image budget validation, avoid redundant source-buffer copies, and bound the Worker cell-snapshot cache as an LRU.

### PPTX archive limits, single-pass parsing and media policy

- Pass the full PPTX archive budget subset through `toOfficeLimits` (`maxCompressionRatio`, `maxArchivePathLength`, `maxXmlBytes`, `maxSingleXmlBytes`, `maxXmlDepth`, `maxXmlAttributeBytes`, `maxTextNodeBytes`, `maxRelationships`) and expose it in `DEFAULT_PPTX_PREVIEW_LIMITS`, so the shared runtime limits are no longer silently disabled for PPTX.
- Decompress each deck once: `loadPptxArchive` reads only the central directory, lets the renderer parse the container, and validates the decoded parts with `validateOfficeArchiveParts` instead of inflating the archive a second time.
- Use one `slideId` for PPTX object fingerprints, locators and resolution, so hosts that supply `slideIds` keep resolving their own references as `exact`.
- Route external media through the shared URL policy via `externalMediaUrlPolicy`; `externalMedia: "allowed"` alone no longer loads arbitrary `http(s)` targets.
- Read byte sources in place instead of copying them twice and hand decoded media bytes to `Blob` without an extra copy.

### Cross-format Office references

- Prepare the nine public packages as the `0.6.0` candidate while keeping `0.5.4` as the published stable line.
- Add the framework-neutral `@arcships/office-interaction` package with serializable references, runtime validation, candidate navigation, geometry helpers, and transient selection state.
- Add stage-one reference adapters for DOCX, XLSX, PPTX, and PDF, including precise text, page or sheet regions, spreadsheet rows and columns, and stable PPTX objects.
- Add one controlled reference-selection contract to all four minimal Surfaces without making them own the host's reference collection or Agent workflow.
- Add optional object-outline and region-selection primitives, demo coverage, external package-consumer checks, and release-gate coverage.

### XLSX chart data

- Render value, series-name, category, and percent data labels on 2D and 3D bar, column, line, and area charts, including label position and number format.
- Read combo-chart data labels independently from each column, line, or area type group, falling back to chart-level labels only when a group does not define its own.
- Normalize 100 % stacked charts with `Σ|v|`, so mixed positive and negative categories render at their real percentages.
- Honour logarithmic value axes, smooth line series, marker symbols, shared combo-chart value axes, and date category axes with elapsed-time spacing.
- Exclude hidden rows and columns from chart series unless the chart explicitly opts out with `plotVisOnly=0`.
- Keep every level of multi-level category caches instead of only the deepest label.
- Refresh chart series from an edited workbook through `refreshChartDataFromWorkbook` and the Worker `refreshCharts` request, so worker-backed charts do not stay on the source file's cached XML values after an in-worker edit.

### PDF selection and annotation hit layer

- Allow PDF text selection to cross page boundaries: drag focus is resolved from the client position, each touched page becomes its own text slice, and copy joins the pages in document order.
- Extend `PdfSelectionState` with per-page `segments` while keeping `range`, `text`, and `rects` compatible for single-page selections.
- Add optional `PdfRenderRuntime.getPageAnnotations()` and `PdfRenderAnnotation`, exposing link, widget, and annotation rectangles in canonical page space with sanitized link targets and form-field identity.
- Add an invisible hit layer to `PdfSurface` that emits `linkActivate` and `annotationActivate`; external links never navigate automatically, and a host `urlPolicy` blocks unapproved targets.
- Stop the region interaction layer from bubbling Escape to the Surface handler so hosts receive one `selectionCancel` per key press.
- Write AcroForm values end to end: `PdfRenderRuntime.setFormFieldValue()` maps text, checkbox and choice widgets through `resolvePdfFormFieldValue()`, and `PdfSurface` renders native controls that commit through the runtime, refresh the page bitmap and emit `formFieldChange`/`formFieldError`.
- Edit annotation appearance through `PdfRenderRuntime.updateAnnotationAppearance()` (`contents`, `color`, `strokeColor`, `opacity`, `strokeWidth`) and `PdfSurface.updateAnnotationAppearance()`, which regenerates the appearance stream, refreshes the bitmap and emits `annotationChange`/`annotationError`.
- Add single-page text selection to `PdfViewer`: drag selection, double-click word selection, clipboard copy, `Escape` clear, and a `selectionChange` event fed by the same canonical transform and glyph geometry as `PdfSurface`.
- Restore a real PDF text layer: `PdfRenderRuntime.getPageTextLayer()` batches PDFium geometry and per-character slices into `PdfPageTextLayer`, and `PdfSurface`/`PdfViewer` render one transparent, non-`aria-hidden` span per glyph positioned from its glyph box. A single CSS matrix per page applies zoom and 0/90/180/270 rotation, so no browser re-shaping or glyph drift is involved; unmapped-Unicode glyphs are marked `aria-hidden` and unselectable. The layer keeps `pointer-events: none`, so cross-page selection, annotation hits and region selection are unchanged, and screen readers, browser find and CJK extraction now see the document text.

### DOCX tracked-change decisions (#41)

- Add `acceptTrackedChange` / `rejectTrackedChange`, row variants and `acceptAllTrackedChanges` / `rejectAllTrackedChanges` to the DOCX editor controller, backed by the pure `applyTrackedChangeDecision` model rewrite in `@arcships/docx-core`. Every decision is one undoable history step.
- Accepting or rejecting a revision now rewrites the paragraph `sourceXml` envelope instead of dropping it, so remaining revisions stay visible and exportable, rejected deletions come back as ordinary runs, and accepting an insertion no longer needs a second edit pass.
- Wire per-card Accept/Reject buttons into the revision gutter and Accept all / Reject all into the editor toolbar.
- Keep the document editable while revisions are shown: the `Edits` toggle is a display choice again, for body paragraphs and table cells alike. Committed text now skips inline deleted/moved-from runs (`readEditableText`), so an edit cannot resurrect rejected content by accident, and edits that still clear revision markup report it on the status line instead of silently dropping the cards.
- Own each gutter card on exactly one page using `trackedChangeBelongsToPageSegments` and the new `groupGutterAnnotationsByPage`, so a change in row 3 of a table spanning pages (or in a paragraph split across pages) is no longer repeated.
- Stop duplicating insertions in the gutter (`trackedChangeUsesGutterBalloon`), and stop collecting revisions while they are hidden: `trackedChanges` stays empty and `trackedChangeCount` reports `null` until the markup is shown.
- Model limitation to be explicit about: run format revisions revert their previous run style and paragraph format revisions revert an explicit alignment, but table/row/cell property revisions only have their markup cleared — the model stores no previous table formatting. Those ids come back in `clearedOnlyChangeIds` and the status line says so.

### Boundaries

- Keep screenshot bitmap capture and annotation editing outside the built-in Surface implementation; unsupported capture rejects with a structured error.
- Keep prompts, Agent calls, document mutation, multi-reference roles, and accept/reject workflows owned by the host application.

## [0.5.4] - 2026-07-13

### Surface search and result positioning

- Added one shared search-session contract and `OfficeFindBar` for DOCX, XLSX, PPTX, and PDF surfaces, with default fuzzy matching, result counts, next/previous navigation, activation, and stale-task suppression.
- Added precise DOCX text-range highlights, XLSX sheet/cell activation, PPTX shape-level activation, and PDF search rectangles using each format's stable semantic anchors.
- Kept built-in viewers and host-owned Surface toolbars on the same search state instead of maintaining duplicate result lists.

### PDF glyph selection

- Replaced browser-reflowed transparent text spans with PDFium character indices, glyph geometry, and overlay rectangles for search, drag selection, and copy.
- Added visible-page geometry prefetch with request deduplication, a drag threshold, retained down/up points while geometry loads, run-scoped glyph-gap snapping, animation-frame coalescing, and final pointer-up flushing.
- Added native double-click word selection using coarse PDF boundaries, per-character text slices, an explicit PDF-index-to-UTF-16 map, and `Intl.Segmenter`.
- Kept selection single-page and did not add OCR, sentence/line gestures, edge scrolling, touch handles, or new public selection APIs.

### Validation and compatibility

- Added unit coverage for coordinate transforms, run-scoped hit testing, English/numeric/Chinese word boundaries, and multi-code-unit mapping.
- Added component coverage for drag completion after delayed geometry and native double-click selection, plus real Chromium interaction verification.
- Kept existing public entry points and package relationships compatible with `0.5.3`.

## [0.5.3] - 2026-07-13

### Unified Surface gesture zoom

- Added controlled `zoom` and `update:zoom` contracts to `DocxDocumentSurface`, `PdfSurface`, `PptxStage`, and `XlsxSheetSurface`; `1` means 100%.
- Added `enableGestureZoom`, `Ctrl+wheel`, trackpad pinch, and WebKit gesture support without consuming ordinary wheel scrolling.
- Unified the supported Surface range at 50%–200% and kept existing zoom APIs compatible when the controlled prop is absent.
- Preserved the pointed page or slide during DOCX, PDF, and PPTX zoom; XLSX preserves the pointed cell and handles frozen rows, frozen columns, and headers per axis.
- Updated built-in viewers and all four Surface demos to use host-owned zoom state.

### Expanded Office input formats

### @arcships/xlsx-core / @arcships/vue-xlsx

- Added XLSB, XLSM, XLTX and XLTM input support; macro content is ignored and never executed.
- Added CSV source detection for local `.csv` file names and remote `.csv` URLs or `text/csv` responses.
- CSV input opens as one editable sheet while preserving the original source bytes for download and reusing XLSX validation, Worker, rendering, and export paths.
- Added UTF-8 and BOM-marked UTF-16 decoding, plus CSV support to the XLSX Viewer and Surface demo upload controls.

### @arcships/vue-pptx

- Added PPTM, PPSX, PPSM, POTX and POTM input support with package validation and ignored-macro warnings.

### @arcships/docx-core / @arcships/vue-docx

- Added read-only DOCM, DOTX and DOTM input support; VBA is never executed and original-format macro editing is not implied.

### Internal and compatibility changes

- Added shared source-format detection and Surface zoom math to the private `@arcships/office-runtime` package; public packages continue to ship without a runtime dependency on that private package.
- Added an explicit PPTX list scroll container so renderer virtualization and Surface anchor restoration use the same owner.
- Kept PDF page slots geometrically stable while zoomed and exposed XLSX grid anchor capture/restoration to its Surface.
- Added focused unit and component coverage for zoom bounds and wheel-event consumption, while retaining the existing release regression suites.

## [0.5.2] - 2026-07-13

### @arcships/vue-pptx

- Changed browse and Surface rendering to a vertically scrollable, windowed slide list; present mode remains a single-slide playback stage.
- Added runtime `selectionChange`, `objectClick`, and contextual `contextMenu` events with slide index, object key, and viewport/container coordinates.
- Kept the visible slide, toolbar counter, search navigation, and host navigation synchronized while scrolling.
- Reserved left/right keyboard navigation for present mode and use up/down or PageUp/PageDown in browse mode.

### @arcships/pptx-core

- Added explicit `renderMode: "list" | "slide"` and list-rendering options to browser sessions.
- Static preview sessions default to list mode; document/playback sessions default to slide mode.

## [0.5.1] - 2026-07-13

### Release

- Bumped the eight public packages together for the search API release candidate.

## [0.5.0] - 2026-07-13

### Unified Surface Components

- New `DocxDocumentSurface` (DOCX), `XlsxSheetSurface` (XLSX), `PdfSurface` (PDF) — minimal embeddable renderers for third-party use.
- `PptxStage` (PPTX) already existed; usage documented.
- Unified CSS variables `--xxx-surface-bg` for background color customization.
- `fitWidth` prop on DOCX/PDF for container-width adaptive zoom.
- Unified event model: `contextMenu`, `selectionChange`, `objectClick` across all four surfaces.
- Container-relative coordinates (`containerX`/`containerY`) in all `contextMenu` events.

### @arcships/vue-docx

- `DocxDocumentSurface` exported as stable public component.
- `fitWidth` prop: ResizeObserver-based auto zoom.
- `contextMenu` event with page index + container coords.
- `selectionChange` event via browser Selection API.
- Fixed: double zoom scaling from `getBoundingClientRect` + `zoomFactor`.
- Fixed: annotation gutter horizontal layout (replaced hardcoded 48px with actual 16+240px).

### @arcships/vue-xlsx

- `XlsxSheetSurface` exported as stable public component.
- `selectionChange` event via `controller.selection` watch.
- `contextMenu` event with cell address + sheet name + container coords.
- Fixed: single-character edit bug (grid keydown re-entering edit mode).
- Fixed: `allowResizeInReadOnly` default changed to `true`.
- Fixed: right-click no longer cancels existing cell selection.

### @arcships/vue-pptx

- `PptxStage` added `--pptx-surface-bg` CSS variable.
- `contextMenu`, `selectionChange`, `objectClick` event declarations; `0.5.2` completes their runtime behavior and structured payloads.
- HomePage now includes all three PPTX demo page cards.

### @arcships/vue-pdf

- `PdfSurface` with vertical scroll through all pages (replaces single-page navigation).
- `fitWidth` prop: auto zoom based on widest page vs container.
- `contextMenu` event with page index + container coords.
- `--pdf-surface-bg` CSS variable.
- Known limitation: text selection not available pending PDF engine CMap issues.

### Demo

- Four new surface demo pages: `/docx-surface`, `/xlsx-surface`, `/pptx-surface`, `/pdf-surface`.
- Each demo shows host-owned toolbar, event listeners, and live status display.
- Dev server port changed to 5173.

## [0.4.0] - Unreleased (candidate)

### @arcships/docx-core

- New instance-level `createDocxRuntime` API; Worker, WASM, source rules, resource limits, tasks, and diagnostics belong to the instance.
- Public entry points: `./runtime`, `./wasm-url`, `./worker`, and `./assets/docx_wasm_bg.wasm`.
- Unified rendering surface for Viewer and Editor.
- Instance-scoped limits for input, archive, XML, relationships, images, model nodes, layout pages, parse time, history, and cache — with structured errors (`actual` / `allowed`).

### @arcships/vue-docx

- `DocxViewer` and `DocxEditor` use the unified document rendering surface.
- New `useDocxPageThumbnails` composable (replaces `useDocxViewerThumbnails`).
- History bounded by both entry count (default 100) and estimated bytes (default 32 MiB).
- Stable style entry: `@arcships/vue-docx/style.css`.

### @arcships/xlsx-core

- New instance-level `createXlsxRuntime`; unified Worker/WASM config, cancellation, timeout, limits, and diagnostics.
- Worker failures no longer silently fall back to main thread.
- Instance-scoped limits for input, archive, XML, relationships, images, parse time, sheet XML, shared strings, rows/columns, sheet count, and formula count.
- Public entry points: `./runtime`, `./wasm-url`, `./worker`, and `./assets/duke_sheets_wasm_bg.wasm`.

### @arcships/vue-xlsx

- Optional entries: `@arcships/vue-xlsx/chart`, `./map`, and `./webgl` — plain tables no longer download optional renderers.
- Large-sheet rendering uses real scroll ranges, accumulated offsets, and in-frame merged draws.
- Workbook child objects released per-instance lifecycle.
- History, images, object URLs, and thumbnails bounded by entry count and bytes.
- Stable style entry: `@arcships/vue-xlsx/style.css`.

### @arcships/vue-pdf

- `PdfViewer` uses a controlled PDF engine to display real documents.
- PDF Runtime, Worker, page image URLs, and tasks belong to specific instances; released on source switch or unmount.
- Single public file-size rejection limit: `maxFileSize`, default `50 MiB`, host-adjustable. Returns `PDF_TOO_LARGE` with `actual` and `allowed`.
- Public entry points: `./assets/pdfium.wasm` and `./style.css`.

### @arcships/vue-ui

- Upload, signature, thumbnail, citation, and layout components retain stable public events and error types.
- Public entry point: `./style.css`.

## [0.1.0] - Initial workspace release

- DOCX/XLSX monorepo with core and Vue packages.
- PDF viewer and utility components.
