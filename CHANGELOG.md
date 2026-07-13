# Changelog

All notable changes to the `@arcships` package family are documented in this file.

## [0.5.2] - 2026-07-13

### @arcships/xlsx-core / @arcships/vue-xlsx

- Added XLSB, XLSM, XLTX and XLTM input support; macro content is ignored and never executed.
- Added CSV source detection for local `.csv` file names and remote `.csv` URLs or `text/csv` responses.
- CSV input opens as one editable sheet while preserving the original source bytes for download and reusing XLSX validation, Worker, rendering, and export paths.
- Added UTF-8 and BOM-marked UTF-16 decoding, plus CSV support to the XLSX Viewer and Surface demo upload controls.

### @arcships/vue-pptx

- Added PPTM, PPSX, PPSM, POTX and POTM input support with package validation and ignored-macro warnings.
- Changed browse and Surface rendering to a vertically scrollable, windowed slide list; present mode remains a single-slide playback stage.
- Added runtime `selectionChange`, `objectClick`, and contextual `contextMenu` events with slide index, object key, and viewport/container coordinates.
- Kept the visible slide, toolbar counter, search navigation, and host navigation synchronized while scrolling.
- Reserved left/right keyboard navigation for present mode and use up/down or PageUp/PageDown in browse mode.

### @arcships/pptx-core

- Added explicit `renderMode: "list" | "slide"` and list-rendering options to browser sessions.
- Static preview sessions default to list mode; document/playback sessions default to slide mode.

### @arcships/docx-core / @arcships/vue-docx

- Added read-only DOCM, DOTX and DOTM input support; VBA is never executed and original-format macro editing is not implied.

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
