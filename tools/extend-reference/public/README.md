# DOCX sidebar parity page image

`docx-sidebar-page-1.png` is a deterministic 280×420 (DPR 2) capture of the real
`DocxThumbnailPanel` canvas for the parity fixture document (12 paragraphs,
200×300pt page, 20pt margin). It is used by the Extend reference
`DocxSidebarSurface` so both sides compare the same DOCX page imagery instead of
two different file-thumbnail cards.

Regenerate by running the DOCX story capture or by rendering the same model in
`tools/vue-reference` and exporting the canvas at 2×.
