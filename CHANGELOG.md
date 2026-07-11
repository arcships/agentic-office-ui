# Changelog

All notable changes to the `@arcships` package family are documented in this file.

## [0.2.0] - Unreleased (candidate)

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
