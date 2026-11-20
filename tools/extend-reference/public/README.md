# DOCX parity page image

The DOCX sidebar reference no longer ships a page bitmap here. The shared
fixture lives at `tools/parity-fixtures/assets/docx-thumbnail.png`: a
deterministic 184×276 (DPR 2) capture of the real `DocxThumbnailPanel` canvas
for the parity fixture document (12 paragraphs, 200×300pt page, 20pt margin).

The Extend reference `DocxSidebarSurface` draws that bitmap into a canvas at
1:1 device pixels, which is what upstream's DOCX sidebar does
(`previewContent={<canvas/>}`), so the reference and the Vue subject rasterize
the same 92×138 card.

Regenerate with the parity fixture: render `tools/vue-reference` with
`?surface=docx-sidebar`, then export the first `[data-parity-cell='open']
canvas` with `canvas.toDataURL('image/png')` (see
`docs/components/docx-ui-stories.md`).
