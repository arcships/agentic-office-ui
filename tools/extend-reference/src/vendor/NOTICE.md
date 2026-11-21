# Vendored Extend UI notices

The files in this directory (`ui/`, `extend/`) and the vendored stylesheets in
`../styles/extend-*.css` are copied or adapted from
[`extend-hq/ui`](https://github.com/extend-hq/ui) at commit
`70fe70c6f1e1eee53954c6923f3d9f583557a57e` (2026-11-17).

Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.

Licensed under the MIT License. The full upstream license text is kept at
[`docs/upstream/licenses/extend-ui-MIT.txt`](../../../../docs/upstream/licenses/extend-ui-MIT.txt)
and is reproduced in [`THIRD_PARTY_NOTICES.md`](../../../../THIRD_PARTY_NOTICES.md).

## MIT License

Copyright (c) 2026 CrowdView Inc, dba Extend
Portions Copyright (c) 2023 shadcn

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Files and what changed

| File | Upstream source | Change |
|---|---|---|
| `ui/button.tsx` | `apps/v4/components/ui/button.tsx` | Base UI `useRender`/`mergeProps` replaced with a plain `<button>`; class strings unchanged |
| `ui/spinner.tsx` | `apps/v4/registry/new-york-v4/ui/spinner.tsx` | `IconPlaceholder` resolved to an inline Lucide Loader2 SVG (matches `OfficeSpinner`) |
| `ui/card.tsx` | `apps/v4/registry/new-york-v4/ui/card.tsx` | `useRender` replaced with `<div>`; only `Card` copied |
| `ui/separator.tsx` | `apps/v4/registry/new-york-v4/ui/separator.tsx` | Base UI Separator replaced with a `<div>`; classes unchanged |
| `ui/select.tsx` | `apps/v4/registry/new-york-v4/ui/select.tsx` | Static trigger/value only (`selectTriggerVariants` copied verbatim); chevron is an inline Lucide SVG; no popup |
| `ui/input.tsx` | `apps/v4/registry/new-york-v4/ui/input.tsx` | Base UI Input replaced by a native `<input>`; wrapper/input class strings unchanged |
| `ui/dialog.tsx` | `apps/v4/registry/new-york-v4/ui/dialog.tsx` | Base UI dialog/portal/backdrop replaced by static divs with the same class strings so the open surface can be measured |
| `ui/tooltip.tsx` | `apps/v4/registry/new-york-v4/ui/tooltip.tsx` | Base UI tooltip popup replaced by static markup with the same class strings and `data-parity-tooltip-content` |
| `ui/toggle.tsx` | `apps/v4/components/ui/toggle.tsx` | Base UI `Toggle`/`ToggleGroup` replaced by static `<button>`/`<div>` markup with the same class strings and `data-pressed`/`data-state` attributes; `pressed`/`disabled` are props (#170) |
| `ui/group.tsx` | `apps/v4/components/ui/group.tsx` | `useRender`/`mergeProps` replaced by plain elements; class strings unchanged; used by the toggle group's `spacing="none"` branch (#170) |
| `ui/color-picker.tsx` | `apps/v4/components/ui/color-picker.tsx` | Swatch trigger only (`ColorPickerTrigger`): ghost `icon-sm` button, `size-4` glyph and the 16x4 colour bar; popover panel, editors, eyedropper and `lib/color-picker-utils` dropped (#170) |
| `ui/registry-icon-props.ts` | `apps/v4/lib/registry-icon-props.ts` | Verbatim type (#170) |
| `extend/file-thumbnail.tsx` | `apps/v4/components/extend/file-thumbnail.tsx` | Rendered markup and classes unchanged |
| `extend/file-upload.tsx` | `apps/v4/components/extend/file-upload.tsx` | `border-beam` dependency dropped (`BorderBeam` only paints while dragging); icons resolved directly; vendored `Card` |
| `extend/document-viewer-sidebar.tsx` | `apps/v4/components/extend/document-viewer-sidebar.tsx` | Rendered markup and classes unchanged |
| `extend/viewer-states.tsx` | `apps/v4/components/extend/pdf-editor-shared.tsx` (`PdfEditorEmptyState`), `docx-viewer.tsx`/`xlsx-viewer.tsx` (`ViewerLoadingSurface`), `xlsx-viewer.tsx` (`loadError` card), `file-system.tsx`/`xlsx-viewer.tsx`/`pdf-editor-panels.tsx` (footer strips) | Class strings copied verbatim; the empty-state glyph is a prop instead of the upstream icon set, the error-card action drops its upload glyph, and the status footer composes the pinned footer classes with the page/status/file content |
| `extend/docx-toolbar.tsx` | `apps/v4/components/extend/docx-viewer.tsx` | Toolbar/page/zoom/actions subset: tooltip/dropdown popups dropped; static trigger classes and icon sizes kept |
| `extend/pdf-toolbar.tsx` | `apps/v4/components/extend/pdf-viewer.tsx` | Toolbar subset: tooltip/popover/select popups dropped, static trigger classes kept |
| `extend/pptx-toolbar.tsx` | `apps/v4/components/extend/pptx-viewer.tsx` | Toolbar subset: tooltip/select/dropdown popups dropped; static trigger classes, slide counter, zoom cluster and actions trigger kept |
| `extend/pptx-slide-card.tsx` | `apps/v4/components/extend/pptx-viewer.tsx` + `file-thumbnail.tsx` | Thumbnail-rail slide card: rendered markup/classes unchanged; the engine canvas is replaced by the shared fixture PNG |
| `extend/bounding-box-citations.tsx` | `apps/v4/components/extend/bounding-box-citations.tsx` | `HumanReviewPanel` Form tab + scalar field cards; pierre/diffs JSON tab, glide-data-grid array/object editors and tooltips dropped; loading/error states are harness extensions |
| `extend/layout-blocks-panel.tsx` | `apps/v4/components/extend/layout-blocks-panel.tsx` | `@tanstack/react-virtual` and the Base UI ScrollArea replaced by flow/scroll divs with the same classes; loading/error/markdown states are harness extensions |
| `extend/signature-fields.tsx` | `apps/v4/registry/new-york-v4/blocks/e-signature.tsx` | `SignatureFieldsPanel` only; PDF viewer and `signature_pad` dialog dropped; `showPreview`, loading and error are harness extensions |
| `extend/document-splits.tsx` | `apps/v4/components/extend/document-splits.tsx` | Static split cards and thumbnails; `@dnd-kit` wrappers dropped (drag/drop is covered by Vue interaction tests); ScrollArea replaced; loading/error/empty are harness extensions |
| `extend/xlsx-toolbar.tsx` | `apps/v4/components/extend/xlsx-viewer.tsx` | `WorkbookToolbar` + tooltip/search/actions triggers only; engine hooks become props; popups dropped (#160) |
| `extend/xlsx-sheet-tabs.tsx` | `apps/v4/components/extend/xlsx-viewer.tsx` + registry `tabs.tsx` / `components/ui/scroll-area.tsx` | `WorkbookSheetTabsInner` without the hover thumbnail portal; Base UI Tabs/ScrollArea replaced by static elements with the same classes (#160) |
| `extend/xlsx-status-bar.tsx` | `apps/v4/components/extend/file-system.tsx` | the `aria-live` status strip; the pinned XLSX viewer has no status bar (#160) |
| `extend/xlsx-ribbon.tsx` | `apps/v4/components/extend/xlsx-editor.tsx` (`EditorToolbar` format row, ~L1840-L2270) | Static resting-state format row: undo/redo, font family + size, B/I/U/S, text/fill colour, alignment, wrap, number format, borders, merge. Private `useXlsxViewer*` hooks become props; select/dropdown/popover popups dropped; ColorPicker swatch trigger only; pivot, sheet selector, zoom and formula row belong to other surfaces (#170) |
| `../styles/extend-globals.css` | `apps/v4/app/globals.css` | Only the `@import` graph narrowed; the rest is byte-identical (sha256 of the body is verified by `scripts/ci/extract-extend-tokens.mjs`) |
| `../styles/extend-shadcn-tailwind.css` | `apps/v4/app/shadcn-tailwind.css` | Provenance header only |
| `../lib/cn.ts` | `apps/v4/lib/utils.ts` | `siteConfig`/`absoluteUrl` dropped |
| `../upstream/components/extend/pdf-viewer.tsx` | `apps/v4/components/extend/pdf-viewer.tsx` | Byte-identical body (attribution header only). L2 route reference (#182), rendered through `route.html` with the real EmbedPDF plugins |
| `../upstream/components/extend/document-viewer-sidebar.tsx` | `apps/v4/components/extend/document-viewer-sidebar.tsx` | Byte-identical body (attribution header only) |
| `../upstream/components/{icon-placeholder.tsx,ui/*.tsx}` | `apps/v4/components/{icon-placeholder.tsx,ui/*.tsx}` | Byte-identical bodies (attribution header only); these re-export the registry primitives below |
| `../upstream/registry/new-york-v4/ui/*.tsx` | `apps/v4/registry/new-york-v4/ui/*.tsx` | Byte-identical bodies (attribution header only) |
| `../upstream/lib/{utils,config}.ts` | `apps/v4/lib/{utils,config}.ts` | Byte-identical bodies (attribution header only) |
| `../upstream/lib/pdf-thumbnail-utils.ts` | `apps/v4/lib/pdf-thumbnail-utils.ts` | Body unchanged except `PDFIUM_WASM_URL`: upstream fetches `@embedpdf/pdfium@2.15.0` from jsDelivr, this app serves the same file from its own origin (`scripts/copy-pdfium.mjs`) so a capture never needs the network |

Hugeicons (`@hugeicons/react`, `@hugeicons/core-free-icons`) is the icon set the
pinned reference resolves by default (`apps/v4/components/icon-placeholder.tsx`).
