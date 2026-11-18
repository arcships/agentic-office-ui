# Changelog

All notable changes to the `@arcships` package family are documented in this file.

## [0.7.0] - 2026-11-18

### XLSX number display and input follow the viewer locale (#137)

- Add the upstream locale primitives to `vue-xlsx`: `localizeCellNumber(formatted, formatCode, locale)` rewrites the engine's en-US decimal/group separators for the requested locale while masking quoted, escaped and currency literals, and `parseLocaleNumber(text, locale)` parses grouped/decimal input (`1.234,5` → 1234.5 in every locale; a lone `1.234` is 1.234 in `en-US` and 1234 in `de-DE`). `parseDisplayedNumber` accepts the same locale so conditional-format data bars, color scales, icon sets and sparklines read localized cell text correctly.
- `XlsxViewer` and `XlsxSheetSurface` now treat their `locale` prop as the default number/date locale as well as the UI locale; the new `documentLocale` prop overrides it, and a default `"en"` leaves a controller created with its own `locale` untouched. Switching the prop re-localizes existing formats and changes how typed values are parsed. The controller option `locale`, the `controller.locale` getter and `setDocumentLocale()` keep working as before.
- Formatting and coercion now route through those helpers: displayed values, the formula bar, copy text, a11y labels, conditional formats and sparklines use `localizeCellNumber`, while typed and pasted text uses `parseLocaleNumber` before coercion. Explicit `@` text formats keep the typed string, `[$-407]`-style format sections still win over the viewer locale, date/time formats keep their punctuation, and format literals such as `"kg"#,##0.00` are preserved.
- Add `tests/unit/xlsx-number-locale.test.mjs` and extend `tests/component/xlsx-number-locale.test.mjs` with the `locale="de-DE"` acceptance (typing `1.234,5` → 1234.5 and `1.234,50` display, `en-US` keeping `1,234.50`, `@` text and `documentLocale` override), register the unit file in `scripts/ci/run-suite.mjs`, and document the props in the XLSX guide, custom-surface guide and public API contract (including #104's locale acceptance for numbers and dates).

### DOCX host render hooks (#151)

- Add `renderCommentCard` and `renderTrackedChangeCard` to `DocxEditor` and `DocxViewer` (forwarded through `DocxDocumentSurface`, `DocxViewerRoot` and `DocxPageSurface`), and `renderContextMenu` / `renderTableContextMenu` to `DocxEditor` and `DocxContextMenu`. They use the render-prop shapes already published by `@arcships/docx-core`, so a host can replace the body of comment and revision cards and the whole right-click menu without forking the components. Without a renderer every card and menu keeps its previous DOM, classes and behavior, and the gutter still owns anchoring, activation, drawer mode and height measurement.
- `DocxCommentCardRenderProps` now carries `activate()` and `setResolved(resolved)`, plus `canModify`; `DocxTrackedChangeCardRenderProps` carries `activate()`, `accept()`, `reject()` and `canDecide`. The callbacks reuse the existing gutter commands (`resolveComment`, `acceptTrackedChange`, `rejectTrackedChange`), so a custom card gets the same undo entries as the built-in buttons. Read-only surfaces report `canModify`/`canDecide` as `false` and the callbacks become no-ops there.
- Custom menus receive the localized built-in `actions`, `runAction(id)`, `closeMenu()`, the pointer `position`, `documentTheme` and `readOnly`. `runAction()` accepts the public `DocxContextMenuActionId` ids (table commands, image layer commands, clipboard and formatting), plus `highlight-<color>` / `text-color-<css color>` entries for the color rows; `renderTableContextMenu` wins over `renderContextMenu` for table targets.
- Read-only editing sessions now open the context menu instead of showing nothing: the built-in menu renders a single Copy item (`data-docx-context-menu-read-only`, `docx-context-menu-copy`), so selected text stays copyable when `editable` is false or revisions lock editing. `DocxTableContextMenuActionId` gains `select-table` and `clear-table-cells` to match the actions the built-in table menu already offers.
- Add `tests/component/docx-host-render-hooks.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering the default appearance, custom comment cards resolving a thread, custom revision cards accepting a change, the read-only no-op guards, `DocxDocumentSurface`/`DocxEditor` forwarding, custom text/table menus with `runAction`/`closeMenu`, the read-only action filtering and the read-only Copy menu; document the hooks in `docs/custom-components/docx.md` and `docs/guide/docx.md`.

### XLSX cell cut and range move (#130)

- Add the missing cut workflow to the XLSX viewer. `Ctrl/Cmd+X`, the cell context menu (“Cut”, `Ctrl+X` hint) and the ribbon call `cutSelection()`, which writes the range payload to the internal/system clipboard and arms a pending move instead of clearing the source immediately. The grid draws Excel's dashed marching-ants marquee around the source range (`drawCutAntsBorder()`, phase advanced by `scheduleCutAnts()`), exposes it as `data-cut-range`, and `XlsxSelectionOverlay` announces “Cut K1:K2, press Escape to cancel” (`selection.cutPending`, English and zh-CN).
- The next `pasteFromClipboard()` writes values, formulas (shifted by the paste distance) and styles, then clears the source cells — all inside one undo entry, so one undo restores both ranges and one redo re-applies the whole move. A formats-only `pasteSpecialFromClipboard({ mode: "formats" })`, a fresh copy and a workbook reload each end the cut state without touching the source; `Escape` in the grid (and in the viewer for hosts whose focus is elsewhere) cancels a pending cut.
- Add `XlsxClipboardCutState` (`{ range, sheetIndex }`), `XlsxViewerController.cutState` and `cancelCut()` to `@arcships/xlsx-core`; `cutSelection()` reports whether a cut is pending and `cancelCut()` whether one was cancelled. A cut made on one sheet keeps its marquee there and still clears its source when pasted on another sheet, and a rebuilt workbook (undo, rename, reload) drops a stale payload instead of pasting it into the new document.
- Add `tests/component/xlsx-cut.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering the cut → paste → one-undo/one-redo acceptance, the grid's Ctrl/Cmd+X and Escape shortcuts, the read-only refusal, the copy and formats-only endings, cross-sheet moves, the context-menu item with its shortcut hint, the overlay announcement and the marquee paint/animation; document the workflow in `docs/guide/xlsx.md`, `docs/custom-components/xlsx.md` and the public API contract.

### DOCX drawing objects, charts and equations (#33)

- The model can now represent the content a Word document carries beyond paragraphs, tables, text, pictures and form fields. Add `DrawingRunNode` (`type: "drawing"`) with `kind` `textBox` / `shape` / `wordArt` / `group` / `smartArt` / `chart` / `oleObject` / `model3d`, plus `MathRunNode` (`type: "math"`); both join `ParagraphChildNode`. A drawing node carries the frame (`widthPx` / `heightPx`), the anchor and paint order (`floating`, mirroring `wp:anchor` — offsets, wrap, `behindDocument`, `relativeHeight`-derived `zIndex`), the stored `sourceXml`, and the content the importer can read: `chart` (plot type, title, categories and cached series values from the chart part), `shape` (preset, fill, stroke), `ole.progId`, a 3D poster `src`, and `textBox` paragraphs. A math node carries the linearized equation (`a/b+x^2`) next to the original OMML.
- `buildDocModel()` now enriches the WASM-built model: `paragraphDrawingObjects()` reads each paragraph's own markup, classifies the non-picture drawings (`classifyDrawingXml()`) and equations, resolves chart parts through the story relationship part (`parseChartData()`), poster images through `binaryAssets`, and attaches `DrawingRunNode` / `MathRunNode` children at the text offset the object sits at. Charts and shapes the engine rendered as a flat SVG become drawing nodes that keep that preview; text boxes keep the existing synthetic image representation and additionally expose their parsed paragraphs on `ImageRunNode.textBox` (plus `textBoxSourceXml` for the export).
- The renderer paints the objects instead of dropping them: charts draw an SVG preview from their cached data, 3D models draw their poster frame, text boxes draw their model paragraphs, equations render their linear text, and shapes / SmartArt / embedded OLE objects draw a labelled placeholder at the anchored position (`data-docx-drawing`, `data-docx-drawing-kind`, `data-docx-drawing-placeholder`, `data-docx-drawing-preview`, `data-docx-drawing-textbox`, `data-docx-drawing-behind`, `data-docx-math`). `drawingObjectPlaceholderStyle()`, `drawingObjectFrameStyle()`, `drawingObjectSizePx()`, `chartPlaceholderText()` and `chartPreviewDataUri()` are the shared, framework-free rules, and `messages.docx.drawing` carries the localized labels (English and zh-CN).
- Export no longer loses these objects. The engine rebuilds an edited paragraph from model fields and has no slot for shapes, charts, embedded objects or equations, so a serializer patch (`applyPreservedDrawingObjectsToDocumentXml()` / `...ToPartContent()` / `...ToPackage()`) writes each object's stored markup back at its text offset for body, header, footer, footnote and endnote stories. Untouched paragraphs still pass through verbatim; a part whose paragraph count disagrees with the model is left unchanged instead of patched at a shifted offset; `w:drawing` / `w:object` / `w:pict` are wrapped in a run, `m:oMath` stays paragraph content. Editing a text box writes the new text into `w:txbxContent` (both `mc:Choice` and `mc:Fallback`), so a text box survives the edit and its export. Duplicated paragraphs renumber drawing ids in the preserved markup as well.
- Drop caps are drawn: a paragraph with `w:framePr/@w:dropCap` renders its first letter as a floated span sized by `resolveDropCapFontSizePx()` / `resolveDropCapVisualHeightPx()` (`data-docx-drop-cap`, `data-docx-drop-cap-type`, `data-docx-drop-cap-lines`) with the rest of the text flowing around it; `renderParagraphRuns()` accepts `renderDropCaps: false` for hosts that paint their own frames.
- Add `tests/unit/docx-drawing-objects.test.mjs` and `tests/component/docx-drawing-objects.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering the model nodes for text boxes, shapes, groups, charts, SmartArt, OLE objects, 3D models and equations, chart-part parsing and preview, edit → export → reopen for every object class, verbatim unedited paragraphs, text box text edits, header story parts, the package serializer, the WASM boundary, placeholder styles, drop-cap metrics and the rendered chart/placeholder/text-box/math/drop-cap output; document the behavior in `docs/guide/docx.md`, `docs/custom-components/docx.md` and the public API contract.
### DOCX font metrics, kerning and script fonts (#123)

- Add the DOCX font subsystem to `@arcships/docx-core`: `font-metrics` (revision plus `document.fonts` and rendering-profile invalidation), `font-advances` (canvas/DOM advance cache keyed by family, weight, style, size, kerning and OpenType feature string), `font-kerning` (`w:kern` threshold resolution), `font-features` (`w14:ligatures` / `w14:cntxtAlts` to `font-feature-settings`, plus the shared measurement canvas), `font-table` (`w:font` scanning that does not leak a self-closing entry's children), `script-fonts` (OOXML ascii/hAnsi/eastAsia/cs slot selection with UTF-16 segment offsets) and `document-fonts` (body/header/footer/note/comment/numbering family collection, embedded `FontFace` installation and `document.fonts.ready` awaiting). `invalidateFontMetrics()`, `getFontMetricsRevision()`, `registerFontMetricCache()`, `subscribeFontMetrics()` and the `fontMetrics` facade expose the cache lifecycle; `ensureDocumentFontInvalidation()` now also watches DPR/visualViewport changes.
- Text measurement is script-aware: `measureTextWidthPx()` measures one advance per resolved font slot with the run's kerning mode and feature settings, falling back to a hidden DOM span (document-only hosts) and then to deterministic glyph advances (SSR/worker). `resolveMeasureFontForText()` / `resolveMeasureFontForSegment()` name the same font the renderer uses, `layout-engine` uses the platform-neutral font vertical metrics (0.9 em ascent + 0.25 em descent) instead of `fontSizePt * 1.6`, and the real page-segmentation path uses Canvas font boxes with the same fallback.
- Runs render per script slot: `DocxParagraphHost` emits one span per Latin/CJK/complex-script segment, each carrying the slot family, `font-kerning` and `font-feature-settings`; hyperlinks split the same way. The pretext layout source is split by the same slots, so pagination and the DOM agree on line breaks for mixed-script paragraphs.
- Add `w:kern`, `w14:ligatures`, `w14:cntxtAlts`, `w:rFonts/@w:ascii|hAnsi|hint`, `w:szCs`, `w:cs` and `w:lang` parsing to imported run styles, and expose `TextLigatures` plus the new `TextStyle` fields.
- `useDocxModel` and the editor import path now load the document's fonts (including embedded font-table faces) before layout, and `DocxViewerRoot` re-paginates when the font-metric revision changes.
- Add `tests/unit/docx-font-metrics.test.mjs` and `tests/component/docx-font-rendering.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering script slots, kerning thresholds, ligature settings, OOXML parsing, the advance/DOM/headless fallbacks, `document.fonts.ready` invalidation, Latin/CJK/RTL line-break agreement within 1px, page-count stability across metric invalidation, font-table isolation, document font requests and per-slot rendering.
### DOCX table move handle and floating tables (#150)

- Add a table move handle to `DocxTableHost`: a small square (`data-docx-table-move-handle`, excluded from thumbnails) sits at the table's top-left corner when the table is hovered or has an active cell, and dragging it draws a dashed ghost (`data-docx-table-drag-ghost`) that follows the pointer. Pressing the handle selects the whole table (a `cellRange` covering every cell), a click without travel only selects, and a drag of more than 3px drops the table as a floating table. The handle is only rendered for editable body tables on their first page fragment — nested tables, header/footer tables and read-only hosts have none — and hosts can opt out with `:show-table-move-handle="false"`.
- Add `controller.setTableFloating(tableIndex, update)` to `DocxEditorController`, plus the `DocxTableFloatingUpdate` / `DocxTableFloatingDrop` types and the `resolveTableFloatingDropPatch()` / `mergeTableFloatingStyle()` helpers. The drop writes `TableStyle.floating` (position in twips, `margin`/`text` anchors and the text-wrap distances), keeps the previous wrap distances on a second drag, refuses values `w:tblpPr` cannot represent, and records exactly one undo entry; `setTableFloating(tableIndex, null)` puts the table back into the text flow. Undo restores the inline layout in one step.
- The export now writes `w:tblPr/w:tblpPr` from `TableStyle.floating` (drop position `w:tblpX`/`w:tblpY`, `w:vertAnchor`/`w:horzAnchor`, `w:tblpXSpec`/`w:tblpYSpec` and the `*FromText` distances) as the first `tblPr` child. The bundled engine rebuilds an edited table from the model and ignored `floating`, so a cell edit used to turn a floating table back into an inline one; the new serializer patch restores the marker for body, header and footer tables, leaves tables that still carry their stored markup verbatim, aligns nested tables depth-first, and `modelToDocumentXml`, `serializeDocModel` and `serializeDocx` all use it. Reopening the exported file parses the same `floating` back.
- `tableWrapperStyle()` keeps rendering floating tables with the existing wrapper and now folds `tblpX`/`tblpY` into the wrapper's margins, so a dropped table lands where it was released and the surrounding text still wraps on the other side; negative offsets (outdenting a table) are preserved.
- New localized status messages (`floatedTable`, `inlinedTable`, `cannotFloatTableNotFound`) and the `moveTableHandle` tooltip ship in English and zh-CN.
- Add `tests/unit/docx-table-floating.test.mjs` and `tests/component/docx-table-floating.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering the drop math and its page clamp, patch merging and refusal, the wrapper offsets, `tblpPr` element order and attributes, export → reopen, an edited floating table keeping its position, clearing back to inline, nested-table alignment, header story parts, stored markup left verbatim, the handle/ghost rendering, whole-table selection, drag → floating → undo, click-without-drag, the controller refusals and the read-only/nested/header/opt-out cases; document the affordance in `docs/guide/docx.md`, `docs/custom-components/docx.md` and the public API contract.

### DOCX bookmark navigation (#148)

- Add `createDocxBookmarkIndex(model)` and `normalizeDocxBookmarkName(name)` to `@arcships/docx-core`. The index maps every `w:bookmarkStart` name to the body node — and table cell — that carries it, reusing the `BOOKMARK_START_XML_PATTERN` scan the REF/PAGEREF field resolvers already use. Lookups accept `#Name`, any case and URL-escaped names; `_GoBack` and duplicate names follow Word's behavior. REF/PAGEREF now resolve table-cell bookmarks through the same index.
- Internal `#bookmark` links now navigate. Clicking a table-of-contents entry or cross-reference in a paragraph, table cell, header/footer or note scrolls to the bookmark's page and paragraph, briefly highlights the target and places the caret on editable surfaces. `DocxViewer`, `DocxEditor` and `DocxDocumentSurface` expose `revealBookmark(name)` and return `DocxBookmarkRevealResult` (`{ status: "revealed", name, nodeIndex, pageIndex?, text }` or `{ status: "missing", name }`); the viewer emits `requestPageReveal(pageIndex)` when a virtualized target page has to mount, and missing bookmarks show a localized non-blocking notice plus a `bookmarkMissing(name)` event instead of throwing. The link hover card's open action follows the same path while still emitting `link-open`.
- Add `tests/unit/docx-bookmarks.test.mjs` and `tests/component/docx-bookmarks.test.mjs` (registered in `scripts/ci/run-suite.mjs`), covering the import path (`w:anchor` → `#name`, `w:bookmarkStart` → target), normalization, table-cell targets, page/paragraph scrolling, target highlight, caret placement, missing-bookmark notices, read-only paragraph/table hosts and the `DocxViewer`/`DocxEditor` APIs; document the behavior in `docs/guide/docx.md`, `docs/custom-components/docx.md` and the public API contract.

### DOCX revision display modes (#149)

- Add `DocxRevisionDisplay = "inline" | "balloons" | "none"` and `reviewDisplayModel(model, display)` to `@arcships/docx-core`. `inline` keeps insertions and deletions in the text flow, `balloons` moves deleted/moved-from runs (and the deleted paragraph mark) out of the body onto gutter cards while insertions stay inline, and `none` shows the final text with no inline markup or revision cards. `reviewDisplayParagraph()` derives the same model for one paragraph, and `hiddenParagraphChildIndexes()`, `hiddenParagraphTextRanges()`, `displayTextOffsetForSourceOffset()` and `sourceTextOffsetForDisplayOffset()` expose the display offsets caret/selection code needs.
- The display model is read-only: the source model keeps every revision, accept/reject still finds them, and export still writes `<w:ins>`/`<w:del>`. A display paragraph is cloned only when it actually hides content; `w:delText` runs never become model children, so clean paragraphs and inline display share the input.
- `renderParagraphRuns()` accepts `revisionDisplay` (default `inline` for direct callers) and computes `data-docx-text-start`/`data-docx-text-end` against the display text. `DocxParagraphHost`, `DocxTableHost`, `DocxPageBody`, `DocxPageSurface`, `DocxPageHeader`, `DocxPageFooter`, `DocxDocumentSurface` and `DocxViewerRoot` pass the mode through; `DocxViewerRoot` paginates and renders the display model while comments, tracked-change collection, editing and export keep the source model.
- `DocxEditorController` gains `revisionDisplay` and `setRevisionDisplay()`, `UseDocxEditorOptions` gains `initialRevisionDisplay`, and `useDocxTrackChanges()` exposes both. The editor and viewer toolbars gain a localized “Display for review” selector (`All markup` / `Balloons` / `No markup`); `DocxViewer` accepts `revisionDisplay`/`defaultRevisionDisplay` and emits `update:revisionDisplay`. Inline mode hides tracked-change cards but keeps comment cards; balloons is the default so the old “inline + gutter” duplication is gone.
- Add `tests/unit/docx-revision-display.test.mjs` and `tests/component/docx-revision-display.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering the display model, text-flow switching, comment-range and caret offset remapping, table/header stories, toolbar switching, gutter visibility, and export/revision preservation; document the modes in `docs/guide/docx.md`.
- Selection integration follows the display model: DOM selections and carets are translated from display offsets to source offsets before they reach the controller, and source ranges are translated back to display offsets for DOM restore and context-menu highlighting. The reusable paragraph-level helpers cover body, table, header/footer and note paragraphs, so a caret after deleted text lands on the same source offset Word's final text uses and survives undo/redo restore.

### DOCX table range selection (#146)

- Add a rectangular table selection to the DOCX editor. `DocxTableCellSelectionRange` now carries `gridRange: { startColumnIndex, endColumnIndex }`: the physical table-grid columns covered by the rectangle after `gridSpan` and `vMerge` expansion. `selectTableCellRange()`, drag selection and Shift+click store the normalized range, and `selectedTableCellRange` exposes it on the controller.
- Add `selectWholeTable(tableIndex?)` and `collapseTableCellSelection()` to `DocxEditorController`; `Esc` collapses a multi-cell range to its anchor cell. Highlighting is physical-grid based (`isCellWithinTableSelectionRange()` / `isCellSelected()`), so a merged cell that only partially intersects the rectangle is highlighted as a whole; `DocxTableHost` renders the stronger `docx-table-cell--range-selected` and `docx-table-cell--range-focus` classes.
- Character and paragraph formatting, copy and clear now operate on every cell in the range. `applySelectedStyleChange()` formats all paragraphs under the rectangle when no text range exists; `selectedParagraphLocations()` returns every selected cell's paragraphs for alignment/heading/list commands; copy writes a tab/line separated matrix plus an HTML `<table>`; `clearTableSelection()` reuses `clearTableCellContents(tableIndex, cells[])` so one undo entry clears the whole rectangle (also wired to `Delete`/`Backspace` and cut). The toolbar and context menu gain localised “Select table” and “Clear contents” actions, and toolbar Merge/Clear enabled states follow the range.
- Add `isCellWithinTableSelectionRange()`, `tableCellSelectionRangeCells()`, `tableCellSelectionRangeParagraphEntries()`, `tableCellSelectionRangeHasMultipleCells()`, `tableCellSelectionAnchor()` and `tableSelectionGridRectangle()` to `@arcships/docx-core`, and add `tests/unit/docx-table-selection.test.mjs` (registered in `scripts/ci/run-suite.mjs`). Extend `tests/component/docx-table-editing.test.mjs` with a 2×3 drag, merged-cell whole-table highlight, range formatting/copy/clear/undo and Escape/Delete coverage; document the API in `docs/guide/docx.md` and `docs/custom-components/docx.md`.

### DOCX image layer order (#145)

- Add `bringToFront(location)`, `bringForward(location)`, `sendBackward(location)` and `sendToBack(location)` to `DocxEditorController`. They rewrite the floating image's `floating.zIndex` — Word's `wp:anchor/@relativeHeight` — with the same clamping as Word (`0` … `251658240`, one `65536` step per "forward"/"backward"), clear the stale `sourceXml` of the image, paragraph and table, and record exactly one undo entry. Non-floating images are refused because inline images have no paint order.
- The image context menu gets a localized "Layer" group with Bring to Front, Bring Forward, Send Backward and Send to Back. Items are enabled per Word's rules: Bring actions are disabled at the top of the stack, Send actions at the bottom, and all four are disabled for inline images. `resolveDocxImageLayerOrderState()`, `resolveDocxImageLayerOrderPatch()` and `resolveDocxImageLayerOrderMenuActions()` expose the same rules (and consume `DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS`) to custom menu renderers.
- Rendering follows the model order: `DocxImageLayer` now applies the `relativeHeight`-derived CSS `z-index` per floating image, so two overlapping images paint in the order Word stores, and the same ordering reaches the exported `relativeHeight` attribute. Right-clicking an image or the floating-image layer opens the image menu instead of starting a drag.
- Add `tests/unit/docx-image-layer-order.test.mjs` and `tests/component/docx-image-layer-order.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering clamping and disabled states, menu-driven reordering of two overlapping images, layer paint order, export/re-import XML, stale-`sourceXml` invalidation, one-step undo, right-click context resolution and the localized menu; document the commands in `docs/guide/docx.md`, `docs/custom-components/docx.md` and the public API contract.

### DOCX export pre-serialization (#134)

- Add `prepareModelForSerialization(model)` to `@arcships/docx-core`. The serializer now prepares the model before the WASM bridge: it recurses into body nodes, header/footer sections (including the per-section lists), footnotes and endnotes, then projects each collapsed table grid onto its physical columns. `modelToDocumentXml`, `serializeDocModel` and `serializeDocx` all use the prepared model, so an edited table in a header or a note gets the same preparation as a body table.
- Add `projectCollapsedTableGrid(table, { resolvedBorderSpacingPx })` and `hasCollapsedTableGrid(table)`. A collapsed grid keeps zero-width logical columns in `TableStyle.columnWidthsTwips` and names them in the new `TableStyle.collapsedGridColumns`; projection rewrites the cell spans and `columnWidthsTwips` to the physical grid without mutating the imported model. Import records the logical grid (and `TableStyle.styleId` / `sourceStyleId`) from the table's stored `w:tblGrid`, so a document that Word wrote with a collapsed column is projected after the next edit. Unedited tables keep their stored markup and are written verbatim.
- Add the pre-serialization invariants: a collapsed grid must not have a changed table style (`sourceStyleId` vs `styleId`) or visible table borders, and its marker, widths and cell spans must agree. Illegal models throw the new `DocxSerializationError` (`COLLAPSED_TABLE_GRID_STYLE_CHANGED`, `COLLAPSED_TABLE_GRID_BORDER_SPACING`, `INVALID_COLLAPSED_TABLE_GRID` with the projection `reason`) instead of writing a package Word cannot open.
- Add the `INVALID_MODEL` code to `DocxExportError`. `exportDocxBuffer` (Worker and inline paths) maps `DocxSerializationError` to it, so hosts can distinguish an unserializable model from a generic `EXPORT_FAILED`.
- Keep vertical merges through an edit. The bundled engine build writes `w:gridSpan` from the model but not `w:vMerge`; a table rebuilt after a cell edit lost its row spans. The serializer now re-emits `w:vMerge w:val="restart"` / `w:vMerge` from `rowSpan` / `vMergeContinuation` for body, header, footer, footnote and endnote parts, including tables nested inside table cells.
- Add `tests/unit/docx-export-prepare-model.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering projection boundaries/rejections, story recursion, the style/border invariants and typed errors, the merged-table corpus fixture edited through `updateTableCellParagraphText` / `mergeTableCells` and reopened, a nested table edited through the recursive cell path and reopened, an edited collapsed grid (imported and constructed), a collapsed-grid column insert, a merged table inside a header through both serializers, and the `INVALID_MODEL` export mapping.

### DOCX programmatic export (#147)

- Add `getDocxBytes()` and `getDocxBlob()` to `DocxEditorController`. They return the serialized DOCX without touching the DOM download path, so hosts can upload, store, diff or post-process the edited document instead of intercepting a browser download. `getDocxBlob()` uses the same container MIME type as the current file name (`.docx`, `.docm`, `.dotx`, `.dotm`).
- Both methods reuse the full export pipeline: `flushTextDrafts()`, the pending export model transformer, the imported `basePackage`, required-form-field validation, and the Worker-backed `exportDocxBuffer` with its cancellation, timeout and structured `DocxExportError` failures. A document replaced mid-serialization or an unmounted editor rejects the pending promise with `ABORTED` instead of returning stale or partial bytes; an `exportDocx()` request is also cancelled when a newer export starts.
- `exportDocx()` now calls `getDocxBytes()` internally and only builds the `<a download>` link after serialization succeeds, so programmatic reads and downloads share one serialization path and one set of MIME/file-name rules. Rapid exports still let only the latest request download.
- Worker export messages are now clone-safe: before `postMessage`, the model is deep-copied with `cloneDocModel(toRaw(model))` and the base package is rebuilt as plain Maps/objects. `getDocxBytes()`, `getDocxBlob()` and `exportDocx()` therefore work when the controller exposes Vue reactive Proxies instead of throwing `DataCloneError` in a real browser.
- Add `tests/component/docx-export-bytes.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering Worker-backed and headless main-thread serialization, Word-openable OOXML parts and re-import, Blob MIME types, no-download behavior, model-change/unmount cancellation with Worker termination, structured failures, required-field refusal and in-flight imports; its fake Worker uses `structuredClone` so proxy payloads fail like a real Worker. Add `tests/blackbox/docx_export_bytes.py` (registered in the blackbox suite) to assert `getDocxBytes()`/`getDocxBlob()` and a real `<a download>` export in Chromium. Document the API in `docs/composables/docx.md`, `docs/custom-components/docx.md` and the public API contract.

### DOCX and XLSX find & replace (#113)

- Add `replaceMatch(target, text)` and `replaceAll(query, text, options?)` to `DocxEditorController`. `replaceMatch` rewrites exactly one search result (the target is produced by `docxReplaceTargetFromMatch(match)`); `replaceAll` rewrites every body occurrence with the search session's `caseSensitive`/`wholeWord`/`regex` options and returns `{ replaced, skipped }`. Header, footer, footnote and endnote results stay searchable but are skipped: headers and footers are edited through their click-to-edit page bands (#97) and notes are read-only, so find & replace never rewrites those stories behind the user's back.
- Replacing is run-style preserving: `replaceTextRangeInParagraph()` splits the text runs only at the match boundaries, keeps images/form fields and links in place, gives the inserted text the style of the character it replaced, remaps comment-range offsets and drops the stale `sourceXml` so export re-emits the paragraph. Matches inside nested tables are addressed through their full nesting path (`replaceTextInParagraphLocation`, `replaceTextInDocxParagraphs`).
- Both commands record one undo entry per call: replacing all matches in a document is one undo step, and redo re-applies the whole batch.
- The DOCX editor toolbar gains a find & replace bar (search field with case/whole-word/regex toggles, previous/next, replace field, Replace and Replace All). `Ctrl/Cmd + F` opens it and `Escape` closes it, in read-only editors too (the replace half is disabled there). `useDocxFindReplace()` and the `find` slot of `DocxToolbar` expose the same wiring to hosts.
- Add `replaceInSheet(query, replacement, { matchCase, wholeCell, formula, sheetIndex, range, occurrence })` to `XlsxViewerController`. It covers both values and formulas (formula cells are edited through their formula text, never through the cached result), coerces replaced values like typed input, can be scoped to a range or to one search occurrence, activates the target sheet first and records the whole sweep as a single undo entry. Read-only, protected-sheet and Worker read-only loads return `changed: false`.
- The XLSX search panel gains a replace field, Replace/Replace All buttons and “Match entire cell” / “Include formulas” switches; the result list refreshes after every replace.
- `OfficeFindBar` accepts optional `showReplace`, `replacement`, `replaceDisabled`, `replaceAllDisabled` and `replaceStatus` props, plus `update:replacement`, `replace` and `replaceAll` events; new `common.search` strings (`replacePlaceholder`, `replace`, `replaceAll`, `replaced`, `wholeCell`, `formulas`) ship in English and zh-CN.
- Add `tests/unit/docx-replace.test.mjs`, `tests/component/docx-search-replace.test.mjs` and `tests/component/xlsx-replace.test.mjs` (registered in `scripts/ci/run-suite.mjs`), covering run/style preservation, nested-table and table-cell targets, skipped read-only stories, case/whole-word options, export → reopen, undo/redo, value and formula replacement, whole-cell and range scoping, single-occurrence replace, read-only workbooks and the toolbar/search-panel wiring.

### DOCX section page numbering and parity blank pages (#122)

- Add `createSectionPageNumberTracker()` to `@arcships/docx-core`. It resolves what every physical page shows from the section's `w:pgNumType`: `w:start` restarts the count, `w:fmt` picks the format (`decimal`, `decimalZero`, `lowerRoman`, `upperRoman`, `lowerLetter`, `upperLetter`, `none`, and the other `ST_NumberFormat` names Word writes), and a section without `w:start` continues the previous section's numbering. `normalizeDocxPageNumberFormat()` returns the canonical spelling of a raw `w:fmt` value.
- The tracker exposes one `DocxSectionPageRange` per section (`startPageIndex`/`endPageIndex`, `startNumber`/`endNumber`, `format`, `formatXml`, `restarts`, `requiredParity`, `parityBlankPageIndexes`, plus the `nodeRange` `DocumentPageRange` the pages cover) and the per-page accessors the renderer needs: `pageNumberForPage()` for `PAGE`, `pageNumberFormatForPage()`, `pageLabelForPage()` (`i`, `12`, …), `sectionPagesForPage()` for `SECTIONPAGES` and `fieldsForPage()`. `NUMPAGES` stays the document's physical page count, like Word.
- Add `createSectionParityBlankPage()` / `isSectionParityBlankPage()`: a `w:type="oddPage"/"evenPage"` break inserts a blank page, and that page belongs to the section that ends there — it keeps that section's header/footer, continues its numbering and counts in its `SECTIONPAGES`, while the next section still starts on the required parity with its own `w:start` number. `layoutDocument()` now attributes the blank page the same way the pagination engine does.
- The viewer, the header/footer stories and the thumbnail rail all read the tracker: `DocxPageSurface` exposes `data-docx-page-label` / `data-docx-page-parity-blank`, `DocxThumbnailPanel` labels each page with `data-thumbnail-page-label` and flags `data-thumbnail-parity-blank`, and `PAGE`/`NUMPAGES`/`SECTIONPAGES` resolve from the same numbers, so no page can show a label that disagrees with its own footer field.
- Add `tests/unit/docx-section-page-numbers.test.mjs` and `tests/component/docx-section-page-numbers.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering the acceptance fixture (decimal `1 … n`, then `lowerRoman` restarting at `1` on an `oddPage` start, with and without the parity blank page), the format table, continuation numbering, the field values and the rail labels, and document the numbering and parity rules in `docs/guide/docx.md`.

### DOCX document blocks, words and highlight overlays (#121)

- Add `extractDocumentBlocks` (plus the same-tick `extractDocumentBlocksSync`), `flattenDocumentBlocks` and `collectDuplicateDocModelBlockIds` to `@arcships/docx-core`. The block tree covers body paragraphs, resolved headings, title/subtitle (`section_heading`) and caption (`legend`) styles, numbered list metadata, form-control (`key_value`) and page-number paragraphs, tables with repeating header rows and addressable cells, inline figures, header/footer containers and note paragraphs. Every block carries a deterministic source id, a part/story reference and the model path used by the unified reference contract; empty paragraphs/cells follow the upstream rule (empty paragraph wrappers promote their figures, empty cells are omitted).
- Add word segmentation with `includeWords`. Whitespace separates words, punctuation and symbols are retained, CJK text is split per code point, and every `DocumentWord` carries UTF-16 `sourceRefs` offsets and a `<blockId>/word/{start}-{end}` id.
- Add DOM geometry resolution: `resolveDocxSourceGeometry`, `resolveDocxBlockGeometries`, `resolveDocxWordGeometries`, `normalizedDocxBoundingBox`, `toDocxPageBlocks` and `revealDocxSource`. Boxes are measured from the rendered viewer DOM and projected into page CSS pixels by dividing the page surface's client rect by its CSS size, so 100%/200% zoom, landscape/rotated paper and clipped page slices all share one page coordinate space; each mounted page contributes its own fragment and a resized/repaginated page is re-measured.
- Add the `DocxBlockHighlights` and `DocxWordHighlights` Vue overlays. Both accept a viewer `root`, `pageIndex`, colors and an `onGeometry` callback; blocks can be selected with `block`/`blocks`/`blockIds` and words with `word`/`words`/`wordRanges` (`{ blockId, start, end }`). Highlight boxes are teleported into the matching page surface, so the page zoom transform applies without extra math, and carry `data-docx-block-highlight`, `data-docx-word-highlight`, `data-docx-block-type` and `data-docx-highlight-active`.
- Expose blocks and words through the unified reference contract. `createDocxBlockReferenceDraft` and `createDocxWordReferenceDraft` build `paragraph`/`heading`/`list-item`/`table`/`table-cell`/`image`/`section` and `text-range` drafts, and `resolveDocxReference` now resolves structure paths (paragraph, table row/cell, header/footer/footnote containers) and text ranges in non-body stories. `DocxDocumentSurface` gains `getDocumentBlocks()`, `confirmBlockReference()` and `confirmWordReference()` so `reference-confirm` can carry block and word identity.
- The viewer now writes source ids and UTF-16 text offsets into the DOM (`data-docx-source-id`, `data-docx-text-start`/`data-docx-text-end`, `data-docx-document-id`) for paragraph hosts, table rows/cells/cell paragraphs, images, header/footer containers and note paragraphs.
- Add `tests/unit/docx-document-blocks.test.mjs` and `tests/component/docx-block-highlights.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering extraction from a model and an imported fixture, word ranges, 100%/200% zoom, landscape page rotation, pagination slices, block 3–4 overlays, `pageIndex`/`blockIds`/`wordRanges`, and `referenceConfirm` round-trips; document the layer in `docs/guide/docx.md` and `docs/custom-components/docx.md`.

### DOCX editor affordances (#101)

- Add `DocxFormFieldDialog`: double-clicking a form field in `DocxEditor` opens a properties dialog that edits the field's value, display type, name, placeholder, required flag, maximum length and dropdown options. Saving goes through the new `controller.updateFormField(location, patch)`, which validates the value against the option list it is saved with and writes the whole patch as **one** undo entry; `DocxFormFieldDialog` is exported for custom surfaces and can be disabled with `:show-form-field-dialog="false"`.
- Add `DocxLinkHoverCard`: hovering a rendered hyperlink in an editable document shows open/edit/remove actions. Editable body paragraphs and table cells now emit `data-docx-link-start`/`data-docx-link-end`, so edit and remove first select exactly the linked characters and then call the existing `controller.setLink()`. The card opens external links in a new tab and labels internal bookmarks, and `:show-link-card="false"` removes it.
- Add table row resize handles to `DocxTableHost` (`data-docx-table-row-handle`, next to the column handles). Dragging a row boundary previews the new height and commits it through `controller.setTableRowHeight(tableIndex, rowIndex, heightPx)`, which writes `w:trHeight` with the `atLeast` rule in one undo entry. Handles appear when the table is hovered or has an active cell, only for editable, non-nested tables, and hosts can opt out with `:show-row-resize-handles="false"` (the same hover wiring also reveals the existing column handles).
- Add a paragraph-style hover preview: hovering the toolbar's style picker opens `DocxStylePreviewPopover`, which renders a sample line with the hovered style's own run/paragraph properties (`docxStylePreviewCss`, also exported) and applies the style when clicked. The native select stays the keyboard path, and `:show-style-preview="false"` disables the popover.
- Add drag-to-reorder for body paragraphs and tables. Hovering a body node shows a handle (`useDocxNodeReorder`, also exported with its pure insertion-slot helpers); dragging it resolves the drop slot from the rendered node frames, shows a drop indicator, and commits one `controller.moveBodyNode(fromIndex, targetNodeIndex)` call. The handle also reorders with `↑`/`↓`, `Escape` cancels a drag, a drop on the node's own slot is a no-op, and section boundaries stay attached to their nodes. `:show-node-reorder="false"` removes the affordance.
- `useDocxFormFields()` gains `updateFormField`/`updateSelectedFormField`, and `DocxEditorController` gains `updateFormField`, `setTableRowHeight` and `moveBodyNode` (with the localized `updatedFormField`, `resizedTableRow` and `movedBodyNode` status messages).
- Add `tests/component/docx-editor-affordances.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering all five affordances end to end: dialog open/save/undo, link card open/edit/remove through `setLink`, row drag → `w:trHeight`, style hover preview → apply, drag reorder plus keyboard moves and the no-op drop, and every opt-out prop; document the affordances in `docs/guide/docx.md` and `docs/custom-components/docx.md`.

### DOCX SVG pictures (#124)

- Add an SVG picture layer to `@arcships/docx-core`. `isSvgPicture()` recognises an `image/svg+xml` part from the content type the engine resolves through `[Content_Types].xml`/relationships, from the `.svg` part name, or from the media bytes themselves; `resolveRenderableImageSource()` now returns a sanitised `data:image/svg+xml` source for those pictures instead of passing the raw engine bytes to `<img>`. The source stays vector at the `wp:extent` anchored size, so 200% zoom re-rasterises at display resolution instead of stretching a bitmap.
- Sanitisation (`sanitizeSvgPictureMarkup()`, `inspectSvgPicture()`) strips `<script>`, `on*` handlers, `javascript:` URLs, `@import`, external `href`/`src`/`xlink:href` and external CSS `url(...)` references, while internal `#id` references and embedded `data:image/` paint survive. DOCTYPE/entity prologs and comments are dropped. Inspections are cached per media-bytes array or source string, so cloned models and re-renders do not re-parse a picture.
- Pictures using `<foreignObject>` or filters, malformed markup and markup over `MAX_SVG_PICTURE_MARKUP_CHARS` (4 MiB) fall back to the placeholder (`SVG`, or `s` for small icons). `svgPictureUnsupportedDiagnostic()` returns the structured `IMAGE_UNSUPPORTED_FEATURE` diagnostic (`reason`, `partName`); the inline runner reports it through the new `onImageUnsupported` / `imageUnsupportedText` render options, with the localized `docx.diagnostics.unsupportedImageFeature` message, and `DocxImageLayer` labels floating fallbacks from the same helper instead of showing only alt text.
- Original bytes are never replaced: editing keeps the model's `data`/`src` and the base package's media parts, so export → reopen preserves every SVG part byte-for-byte. Thumbnails keep SVG sources vector (`thumbnailImageSourceQualifiesForDownscale()` skips them) and rasterise the page surface once, and an unsupported picture can no longer leave a broken `<img>` in the clone.
- Add `test-data/corpus/docx-svg-pictures.docx` (generated by `scripts/fixtures/generate-docx-svg-pictures.py`) with a script/external-reference icon, a vector chart, a floating SVG, a Word-style `asvg:svgBlip` + PNG fallback, and `foreignObject`/filter pictures, plus `tests/unit/docx-svg-pictures.test.mjs` and `tests/component/docx-svg-pictures.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering detection, sanitisation, anchored sizes, fallback diagnostics, the 200% zoom vector path, thumbnails and edit/export/reimport fidelity. Document the layer in `docs/guide/docx.md`.

### DOCX page setup and header/footer editing (#97)

- Add `setPageMargins(sectionIndex, margins)`, `setPageSize(sectionIndex, { widthTwips, heightTwips, orientation })` and `addHeaderFooter(sectionIndex, kind, { referenceType })` to the editor controller. Each command is a single undoable step and writes the section's `w:sectPr` in place, so the header/footer references, `w:type`, columns, page borders and page numbering of a section survive a page-setup edit. Margins are in twips (`DocxPageMargins`); paper size uses `DocxPageSize`, and a landscape/portrait contradiction swaps the edges the way Word stores it (portrait omits `w:orient`).
- The commands keep every representation of a section's properties in step: `metadata.sections[i].sectionPropertiesXml`, the paragraph-level `sectionBreak` marker of a mid-document section (the #56 contract) and `metadata.sectionPropertiesXml` for the final section. The same ops are exported from `@arcships/docx-core` (`setPageMargins`, `setPageSize`, `addHeaderFooter`, `resolvePageSetup`, `resolveSectionPropertiesXml`, plus the `w:pgSz`/`w:pgMar` read/write helpers and `DOCX_PAGE_SIZE_PRESETS`).
- Export now also writes the package wrapper a newly created part needs: `word/_rels/document.xml.rels` gains the `header`/`footer` relationship (renaming the `r:id` and rewriting the reference if the id is already taken) and `[Content_Types].xml` gains the matching `<Override>`. A document that already declares its parts is untouched.
- UI: the toolbar gains a page-setup dialog (`DocxPageSetupDialog`, opened from `editor-page-setup`) with paper size presets, custom width/height, orientation and the four margins in inches, plus a section picker for multi-section documents. `useDocxPageSetup` resolves every section's geometry, follows the caret's section, and applies paper size and margins as one undo entry.
- UI: `DocxPageSurface` renders a click-to-edit header/footer band. Activating it creates the variant the page is actually looking for (`first` on a title page, `even` on even pages, otherwise `default`) when the section has none, then makes the header/footer paragraphs content-editable; text commits go through `commitSectionParagraphText`, so one edit updates every page of the section and is undone as one step. Header/footer stories no longer touch the body selection or the body structural commands.
- Tests: `tests/unit/docx-page-setup.test.mjs` (pgMar/pgSz writes, per-section isolation, export → reopen round-trip for margins, orientation, a mid-document section and a created header with its relationship/content-type, comments and a new header sharing one export wrapper, unchanged-wrapper guard) and `tests/component/docx-page-setup.test.mjs` (toolbar dialog apply/cancel with one undo entry, header creation, multi-page header update, undo/redo, localized chrome).

### XLSX sheet tab surface for host chrome (#7)

- Add `showSheetTabs` (default `true`) to `XlsxViewer`. Setting it to `false` removes only the bottom strip: the grid, ribbon, formula bar and the controller's active tab stay intact, so sheets can still be switched by host UI through `setActiveTabIndex()` (and by `Ctrl/Cmd+PageUp`/`PageDown`).
- Add a controlled `activeTabIndex` prop with the matching `update:activeTabIndex` event. `undefined` keeps the controller authoritative, so existing hosts are unaffected; a number is mirrored onto the controller, including a restored index that arrives before the workbook is parsed (applied with the first tabs), and out-of-range values are ignored exactly like `setActiveTabIndex()`.
- Add `tabChange(tab)` with the full `XlsxWorkbookTab` (including `kind`). It fires exactly once per committed switch and never for the workbook's opening tab, a no-op activation, renames, visibility changes, reorders or plain re-renders, and a host-driven `activeTabIndex` write is not echoed. Tab indices are `tabs` indices that interleave worksheets and chartsheets, so `kind`/`sheetIndex`/`chartsheetIndex` stay readable instead of being derived from `sheets`.
- Reordering the active tab changes its index without switching sheets: `update:activeTabIndex` follows the index while `tabChange` stays quiet.
- Add `tests/component/xlsx-viewer-tab-surface.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering strip visibility, the controller-owned default, controller-restored tabs, the controlled index before/after load, the `v-model` round-trip, out-of-range values, chartsheet payloads and rename/reorder rebuilds, and document the surface in `docs/guide/xlsx.md`, `docs/custom-components/xlsx.md` and the package README.

### XLSX sheet tab strip navigation and semantics (#11)

- Keep the active sheet tab in view. `XlsxSheetTabs` scrolls the newly active tab with `scrollIntoView({ block: "nearest", inline: "nearest" })` and `scroll-behavior: smooth` on the strip, so a switch from the arrows, `setActiveTabIndex()`, a restored document or undo reveals its tab without scrolling the page or the grid viewport.
- `Ctrl/Cmd+PageUp` / `Ctrl/Cmd+PageDown` switch to the previous/next tab, clamped at the ends and `preventDefault()`ed; read-only viewers keep the shortcut, while an open inline cell editor owns it. The same editor guard now keeps `Ctrl/Cmd+Z`/`Ctrl/Cmd+Y` from reaching the workbook while the cell editor is open.
- The tab strip is a real tab widget: `role="tablist"` with a label, roving `tabindex` (`0` on the focused/active tab, `-1` elsewhere), `ArrowLeft`/`ArrowRight`/`Home`/`End` move focus, `Enter`/`Space` activate, focus follows the active tab after a strip switch, and each tab carries an `id` plus `aria-controls` pointing at the viewer's grid. All tabs keep the `data-testid="xlsx-sheet-tab"` contract.
- Add `tests/component/xlsx-sheet-tabs.test.mjs` (registered in `scripts/ci/run-suite.mjs`) and extend `tests/component/xlsx-sheet-order.test.mjs` with the active/visible tab state assertions.

### DOCX comments CRUD (#96)

- Add `addDocumentComment(model, { range, text, author, initials, date })`, `replyToDocumentComment(model, parentId, text, options)`, `resolveDocumentComment(model, commentId, resolved)` and `deleteDocumentComment(model, commentId)` to `@arcships/docx-core`. They return immutable revisions: a comment is one definition in `metadata.comments` plus one `commentRanges` offset pair per paragraph the selection touches, so highlights, gutter cards, edits, undo and export all keep working from the same model data. A selection spanning paragraphs or table cells anchors every paragraph it covers, and `resolveDocumentCommentAnchors` exposes that split.
- Serialization writes `word/comments.xml` and `word/commentsExtended.xml` (replies as `w15:paraIdParent`, resolved threads as `w15:done="1"`), adds the `comments`/`commentsExtended` relationships and content-type overrides, and removes both parts when the last comment is deleted. Deleted comments also lose their stale `w:commentRangeStart`/`w:commentRangeEnd`/`w:commentReference` anchors in `document.xml` and in header/footer parts, while authored comments get their anchors re-emitted at the stored offsets. Imported documents whose comments were not edited keep their original comment parts byte-for-byte.
- `DocxEditorController` gains `addComment(range, text)`, `replyToComment(id, text)`, `resolveComment(id, resolved)` and `deleteComment(id)` (one undo entry each), plus `beginComment` / `beginCommentReply` / `submitCommentComposer` / `cancelCommentComposer` / `commentComposer` for the inline editor. `useDocxEditor` accepts `commentAuthor`, `commentInitials` and `commentClock`; the default author comes from `docx.editor.commentAuthor`.
- The DOCX toolbar gains a New comment action next to Insert image/table, and the annotation gutter's comment cards gain Reply, Resolve/Reopen (thread root only) and Delete. The inline editor renders inside the replied-to card for replies and at the top of the pane for a new comment on the selection.
- Add `test-data/corpus/docx-comments.docx` (generated by `scripts/fixtures/generate-docx-comments.py`) with a plain thread with two replies and a resolved thread with one reply, plus `tests/unit/docx-comments.test.mjs` and `tests/component/docx-comments-crud.test.mjs` (registered in `scripts/ci/run-suite.mjs`) covering create/reply/resolve/delete, thread inheritance, table-cell and cross-paragraph anchors, export/reopen with resolved state, part removal, undo and the toolbar/gutter/composer UI.

### Viewer keyboard navigation (#116)

- Add `useViewerKeys()` and the pure `resolveViewerKeyCommand()` to `@arcships/vue-ui`, with a documented map for page navigation (`PageUp`/`PageDown`, `Home`/`End`), zoom (`Ctrl/Cmd` + `+`/`=`/`-`/`0`), search (`Ctrl/Cmd + F`) and `Escape`. The composable refuses inputs, textareas, selects and `contenteditable` targets, keeps `Shift`/`Alt` navigation native, and only calls `preventDefault()` when the host actually handled the command.
- `DocxViewer` and `PdfViewer` now flip pages, jump to the first/last page, zoom, focus their search field and cancel the active search from the keyboard; `Escape` in the PDF viewer clears search before it clears the text selection. The PDF print shortcut (`Ctrl/Cmd + P`) is preserved outside the shared map, but now follows the same editable-target and `defaultPrevented` guards, so it cannot fire from the search or page fields.
- `PdfSurface` gains one-page scrolling (`PageUp`/`PageDown`), first/last jumps, keyboard zoom through the controlled `update:zoom` path, and `Escape` for selection/region cancellation. `DocxViewerRoot` now falls back to assigning `scrollTop` when the host element has no `scrollTo`, and `scrollToPage` tolerates a missing `scrollIntoView`.
- `PptxViewer` no longer hijacks arrows and page keys in browse mode: `PageUp`/`PageDown`, `↑`/`↓`, `Home`/`End` and `Space` scroll the slide list, while slide navigation stays in present mode. Zoom works in both modes, and `Escape` closes browse search or leaves present fullscreen. `Space` on a focused playback button remains a button activation.
- All four viewer roots and `PdfSurface` show a `:focus-visible` outline for keyboard users.
- Add `tests/unit/viewer-keys.test.mjs` and `tests/component/viewer-keyboard.test.mjs` (registered in `scripts/ci/run-suite.mjs`), extend `tests/component/pptx-preview.test.mjs`, and document the shared map in `docs/guide/viewer-keyboard.md` plus the DOCX/PDF/PPTX guides.

### DOCX search coverage and options (#49)

- Add a shared `SurfaceSearchOptions` (`caseSensitive`, `wholeWord`, optional `regex`) to `@arcships/office-runtime`; `createSurfaceSearchExpression` builds the Unicode-aware literal/whole-word/regex matcher, and both DOCX and XLSX searches consume it. DOCX and XLSX toolbars now expose match-case, whole-word and regex toggles.
- DOCX search walks tables recursively: text inside nested tables is matched with a `nested-table-cell` location (`path` + the viewer's stable `data-docx-table-nesting-path`), while direct cell/body ranges keep the existing `DocxTextRange` shape. Header, footer, footnote and endnote stories are searched too, tagged with `region` plus `partName`/`noteId`, using the section-level header/footer parts the viewer actually renders (rendered part wins over the editing fallback, and shared parts are searched once).
- Search navigation resolves header/footer parts and note bodies to the page that renders them (`DocxViewerRoot.searchTargetPage`, `findDocxSearchTargetPage`) and scrolls there; highlights are scoped to the header/footer/note DOM story so identical text in another story cannot be highlighted.
- The search session gained `refreshSearch()`: editing a paragraph (a clone-on-write model replacement) re-runs the active query against the new model, updates match offsets, keeps the result list and preserves/clamps the active result index instead of clearing the search. `DocxDocumentSurface` exposes the options-aware `search()` plus `refreshSearch()`.
- Tests: nested-table, header/footer/note region, rendered-section preference, case/whole-word/regex, session refresh (offsets + active index), page lookup, surface model-refresh, toolbar toggles and XLSX per-call options.

### DOCX fixture and acceptance coverage (#53)

- Add three DOCX fixtures that carry the structures the reported defects live in: `review-threads.docx` (comment threads with replies and a resolved thread, two comments on one paragraph, a range crossing three paragraphs, comments anchored in a table cell and in the header, a deleted table row, a row-height `trPrChange`, two footnotes and two endnotes), `table-structures.docx` (`vMerge` restart/continuation cells, `gridSpan`, a nested table, multi-paragraph cells, a `tblHeader` row that repeats across pages and a two-row table inside the page header), and `sections-columns.docx` (portrait/landscape/two-column sections with their own headers and footers, a text box and a behind-text watermark).
- `scripts/generate_test_materials.py` gained `--only`, so a new fixture can be regenerated without rebuilding the unrelated binaries with a different toolchain; `scripts/verify_test_materials.py` now fails when a regenerated fixture loses its comments, threading, merges, header rows or sections.
- Add `tests/unit/docx-fixture-fidelity.test.mjs` and `tests/component/docx-fixture-fidelity.test.mjs` (registered in `scripts/ci/run-suite.mjs`): they assert comment threads and anchors, cross-paragraph ranges, one gutter card per comment per page with card numbers matching their inline markers, cards placed within one line of their anchor, deleted rows marked and listed once, row-height revisions anchored to their own row, merged/nested table column counts and cell text, repeating header rows, a table inside the page header rendered through the header's table host on every page, per-section headers/footers, note numbering, and an import → edit → export → re-import field diff that changes only the edited paragraph.
- Add the “Acceptance coverage map” to `docs/visual-acceptance-handoff.md`: every acceptance row carries an id and maps either to a real CI-registered test or to the browser suite that still has to accept it; the unit suite verifies the mapping so an unverified claim cannot read as accepted.
- `tests/blackbox/docx_annotations.py` and `tests/blackbox/fidelity_content.py` now check gutter placement against the anchor markers, thread visibility, cross-paragraph/cell/header anchors, merged/nested tables, a table rendered inside the page header on every page, and per-section headers instead of only asserting that a card or a notes section exists. The insertion revision renders inline and is not ballooned, so the review fixture asserts two gutter cards instead of three, and the outer table's column count comes from its own `colgroup` instead of also counting the nested table's columns.

### XLSX page setup and printing (#20)

- Parse `pageSetup`, `printOptions`, `pageMargins`, `rowBreaks`/`colBreaks`, `sheetPr/pageSetUpPr` and the `_xlnm.Print_Area`/`_xlnm.Print_Titles` defined names into `XlsxSheetData`, with an engine fallback for `skipXmlParsing` sources.
- Add pure pagination to `@arcships/xlsx-core` (`buildWorksheetPrintLayout`, `buildPrintPageGrid`): paper size/orientation, usable margins, fit-to-page scaling, manual breaks, repeated title rows/columns, merged-cell spans and `pageOrder` produce a flat page list of cell ranges with the paper size.
- Add an Excel-style dashed page-break preview overlay to the grid plus page numbers, and a paginated `XlsxPrintPreview` with paper-sized page blocks, `@media print` output and `window.print()`.
- Expose `printLayout` / `printPages` / `getPrintLayout()`, `setPageBreakPreview()`, `openPrintPreview()` / `closePrintPreview()` and `print()` on the viewer controller; the toolbar gains page-break and print-preview buttons.
- Export writes `pageMargins`, `printOptions`, `pageSetUpPr`, manual breaks and print-area/title defined names back into the saved workbook (the engine drops several of them), so export → reload keeps the same pagination; unmodified workbooks still export byte-for-byte unchanged.

### XLSX worksheet reorder (#10)

- Add `moveSheet(fromIndex, toIndex)` to the controller. Both arguments are tab indices — positions in `tabs`, which interleave worksheets and chartsheets — not `sheets`/`chartsheets` indices. The resulting strip is `tabs.splice(toIndex, 0, ...tabs.splice(fromIndex, 1))`, the moved tab becomes active and stays selected, dropping it on its own slot is a no-op, and read-only viewers reject the move. Sheet protection does not block a move, matching Excel.
- `@dukelib/sheets-wasm` 0.1.17 exposes no reorder setter, so the viewer rewrites the `<sheets>` order in its own saved package and parses the result back into a fresh workbook. `applyWorkbookSheetOrder` also remaps the positions stored beside the list (`localSheetId` defined names, workbook-view `activeTab`/`firstSheet`) so scoped names and the file's active tab keep pointing at the same sheet.
- Every per-sheet array is re-read from the reordered package rather than spliced by hand: worksheet state (`sheetStatesByWorkbookSheetIndex`), images, shapes, chart/drawing origins, `sheetOrigins` and the visible-index map all follow their sheet, so `sheetOriginsRef`, `imagesByWorkbookSheetIndex` and `shapesByWorkbookSheetIndex` cannot drift onto a neighbour. Image assets are now indexed by worksheet rather than by workbook.xml tab slot, so a chartsheet placed before a worksheet no longer shifts its images, shapes or sheet state; chartsheet drawings are exposed separately as `chartsheetOrigins`.
- A move is one byte-based undo entry (redo restores the new order), and the new order survives `exportXlsx()` and reload, including workbooks with chartsheets mixed into the tab strip.
- The tab strip gains drag-to-reorder: drag a tab to show an insertion indicator between tabs, `Esc` cancels, a same-slot drop changes nothing, and the moved tab is scrolled back into view after the drop. Read-only viewers expose no drag affordance, and a rejected move shows a visible message instead of leaving the strip stale.
- Export `applyWorkbookSheetOrder` from `@arcships/xlsx-core` so hosts can apply the same slot reorder to their own saved bytes.

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

### DOCX rich clipboard (#115)

- Copy now writes rich clipboard flavours instead of plain text: the internal structured fragment (`application/x-arcships-docx-fragment+json`, carrying the source `DocxTextRange` and the selected runs), `text/html` with inline run styles (bold/italic/underline/strike, colour, highlight, size, family, super/subscript) plus `<a>`/`<ul>`/`<ol>` markup, `text/rtf` (colour table, `\uN?` escapes, `HYPERLINK` fields, list markers) for editors that only take RTF, and `text/plain`. A rejected composite write is retried with the standard flavours and finally falls back to `writeText`, so a copy is never lost.
- Paste prefers the internal payload, then `text/html`, then `text/rtf`, then `text/plain`. HTML from Word is mapped back onto runs (`b/strong/i/em/u/s/strike/del/a/font/span` and inline declarations, `h1`–`h6`, `ul`/`ol`/`li`, Word's `mso-list` paragraphs), and Word RTF is read for character formatting, colour tables, hyperlink fields and list markers, so bold, links and lists survive a cross-application paste; cells keep their existing single-paragraph behaviour by joining blocks with `\n`.
- Add the pure `@arcships/docx-core` clipboard module (`createDocxClipboardPayload`, `serializeDocxClipboardPayload`, `parseDocxClipboardPayload`, `renderDocxClipboardHtml`, `parseDocxClipboardHtml`, `renderDocxClipboardRtf`, `parseDocxClipboardRtf`, `docxClipboardParagraphsForRange`, `insertDocxClipboardParagraphs`) and expose `writeSelectionToClipboardData` / `pasteClipboardData` on the editor controller. Rich pastes stay one undo step: an expanded cross-paragraph selection is collapsed and the styled runs inserted inside the same transaction.
- Sanitize the clipboard as external input: script/style/iframe/object/embed/svg content is dropped, only whitelisted inline declarations and colour/font-size/font-family shapes are read, link targets must pass the existing Office URL policy (`sanitizeOfficeLink`), and payload/HTML/RTF/paragraph/run caps (512 KiB each, 2000 paragraphs, 20000 runs) bound the work done on a pasted clipboard.
- Wire copy/cut on the viewer root and paste on the paragraph host through the same controller API, so `Ctrl+C`/`Ctrl+X`/`Ctrl+V` use the rich flavours instead of the reduced DOM copy.

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
