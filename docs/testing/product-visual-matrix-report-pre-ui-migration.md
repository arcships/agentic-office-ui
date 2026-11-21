# Product visual matrix

- Mode: `report-only`
- Generated: 2026-11-19T19:15:05.515406+00:00
- Git HEAD: `3e0f13ddb2ad1bd5d0e1499dbc54b855ccad57b3`
- Browser: 149.0.7827.55
- Baseline: `tests/baseline/visual`
- Pixel policy: non-blocking, per-pixel threshold 8, max differing ratio 0.0020

## Summary

| Metric | Value |
|---|---:|
| Matrix scenarios | 12 |
| Evidence scenarios | 17 |
| Structural failures | 0 |
| Pixel diffs | 0 |
| Missing baselines | 0 |

## Route matrix


### `/docx-viewer` — 3/3 structural

| Viewport | Theme | Zoom | Screenshot | Baseline | Structural | Pixel | Diff ratio |
|---|---|---:|---|---|---|---|---:|
| desktop-1440x900 | light | 100% | `tests/baseline/visual/screenshots/docx-viewer/desktop-1440x900/light-zoom100.webp` | `tests/baseline/visual/screenshots/docx-viewer/desktop-1440x900/light-zoom100.webp` | PASS | BASELINE | — |
| desktop-1440x900 | dark | 100% | `tests/baseline/visual/screenshots/docx-viewer/desktop-1440x900/dark-zoom100.webp` | `tests/baseline/visual/screenshots/docx-viewer/desktop-1440x900/dark-zoom100.webp` | PASS | BASELINE | — |
| mobile-390x844 | light | 100% | `tests/baseline/visual/screenshots/docx-viewer/mobile-390x844/light-zoom100.webp` | `tests/baseline/visual/screenshots/docx-viewer/mobile-390x844/light-zoom100.webp` | PASS | BASELINE | — |


### `/xlsx-viewer` — 3/3 structural

| Viewport | Theme | Zoom | Screenshot | Baseline | Structural | Pixel | Diff ratio |
|---|---|---:|---|---|---|---|---:|
| desktop-1440x900 | light | 100% | `tests/baseline/visual/screenshots/xlsx-viewer/desktop-1440x900/light-zoom100.webp` | `tests/baseline/visual/screenshots/xlsx-viewer/desktop-1440x900/light-zoom100.webp` | PASS | BASELINE | — |
| desktop-1440x900 | dark | 100% | `tests/baseline/visual/screenshots/xlsx-viewer/desktop-1440x900/dark-zoom100.webp` | `tests/baseline/visual/screenshots/xlsx-viewer/desktop-1440x900/dark-zoom100.webp` | PASS | BASELINE | — |
| mobile-390x844 | light | 100% | `tests/baseline/visual/screenshots/xlsx-viewer/mobile-390x844/light-zoom100.webp` | `tests/baseline/visual/screenshots/xlsx-viewer/mobile-390x844/light-zoom100.webp` | PASS | BASELINE | — |


### `/pdf-viewer` — 3/3 structural

| Viewport | Theme | Zoom | Screenshot | Baseline | Structural | Pixel | Diff ratio |
|---|---|---:|---|---|---|---|---:|
| desktop-1440x900 | light | 100% | `tests/baseline/visual/screenshots/pdf-viewer/desktop-1440x900/light-zoom100.webp` | `tests/baseline/visual/screenshots/pdf-viewer/desktop-1440x900/light-zoom100.webp` | PASS | BASELINE | — |
| desktop-1440x900 | dark | 100% | `tests/baseline/visual/screenshots/pdf-viewer/desktop-1440x900/dark-zoom100.webp` | `tests/baseline/visual/screenshots/pdf-viewer/desktop-1440x900/dark-zoom100.webp` | PASS | BASELINE | — |
| mobile-390x844 | light | 100% | `tests/baseline/visual/screenshots/pdf-viewer/mobile-390x844/light-zoom100.webp` | `tests/baseline/visual/screenshots/pdf-viewer/mobile-390x844/light-zoom100.webp` | PASS | BASELINE | — |


### `/pptx-viewer` — 3/3 structural

| Viewport | Theme | Zoom | Screenshot | Baseline | Structural | Pixel | Diff ratio |
|---|---|---:|---|---|---|---|---:|
| desktop-1440x900 | light | 100% | `tests/baseline/visual/screenshots/pptx-viewer/desktop-1440x900/light-zoom100.webp` | `tests/baseline/visual/screenshots/pptx-viewer/desktop-1440x900/light-zoom100.webp` | PASS | BASELINE | — |
| desktop-1440x900 | dark | 100% | `tests/baseline/visual/screenshots/pptx-viewer/desktop-1440x900/dark-zoom100.webp` | `tests/baseline/visual/screenshots/pptx-viewer/desktop-1440x900/dark-zoom100.webp` | PASS | BASELINE | — |
| mobile-390x844 | light | 100% | `tests/baseline/visual/screenshots/pptx-viewer/mobile-390x844/light-zoom100.webp` | `tests/baseline/visual/screenshots/pptx-viewer/mobile-390x844/light-zoom100.webp` | PASS | BASELINE | — |

## Issue evidence scenarios

| Scenario | Route | Screenshot | Structural | Pixel | Notes |
|---|---|---|---|---|---|
| `docx-editor-default` | `/docx-editor` | `tests/baseline/visual/evidence/docx-editor-default.webp` | PASS | BASELINE | Editor shell, toolbar and editable surface (#28) |
| `docx-viewer-comments-zoom50` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-comments-zoom50.webp` | PASS | BASELINE | Comment gutter at 50% zoom (#29) |
| `docx-viewer-comments` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-comments.webp` | PASS | BASELINE | Threaded comments, tracked changes, gutter anchoring (#29/#30) |
| `docx-viewer-thumbnails` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-thumbnails.webp` | PASS | BASELINE | Thumbnail panel with an image-bearing document (#44/#48) |
| `docx-viewer-page-sizes` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-page-sizes.webp` | PASS | BASELINE | Mixed portrait/landscape pages and per-page thumbnails (#27) |
| `docx-viewer-tables` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-tables.webp` | PASS | BASELINE | Merged/nested tables with repeating headers (#30) |
| `docx-viewer-embedded` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-embedded.webp` | PASS | BASELINE | Embedded text boxes, shapes and chart/OLE objects (#36) |
| `docx-viewer-comments-zoom200` | `/docx-viewer` | `tests/baseline/visual/evidence/docx-viewer-comments-zoom200.webp` | PASS | BASELINE | Comment gutter at 200% zoom (#29) |
| `pdf-viewer-rotated` | `/pdf-viewer` | `tests/baseline/visual/evidence/pdf-viewer-rotated.webp` | PASS | BASELINE | PDF page rendering after fixture switching |
| `pptx-surface-stack` | `/pptx-surface` | `tests/baseline/visual/evidence/pptx-surface-stack.webp` | PASS | BASELINE | PPTX continuous surface with a real deck |
| `pptx-viewer-deck` | `/pptx-viewer` | `tests/baseline/visual/evidence/pptx-viewer-deck.webp` | PASS | BASELINE | PPTX viewer with a real deck |
| `xlsx-surface-freeze-merge` | `/xlsx-surface` | `tests/baseline/visual/evidence/xlsx-surface-freeze-merge.webp` | PASS | BASELINE | Surface grid freeze panes and merged regions (#14/#15) |
| `xlsx-viewer-freeze-merge` | `/xlsx-viewer` | `tests/baseline/visual/evidence/xlsx-viewer-freeze-merge.webp` | PASS | BASELINE | Two-axis freeze panes and a scrolled-out merge anchor (#14/#15) |
| `xlsx-viewer-databar-icons` | `/xlsx-viewer` | `tests/baseline/visual/evidence/xlsx-viewer-databar-icons.webp` | PASS | BASELINE | Negative data bars and 4/5-bucket icon sets (#15) |
| `xlsx-viewer-iconsets` | `/xlsx-viewer` | `tests/baseline/visual/evidence/xlsx-viewer-iconsets.webp` | PASS | BASELINE | 4- and 5-bucket icon sets (#15) |
| `xlsx-viewer-charts` | `/xlsx-viewer` | `tests/baseline/visual/evidence/xlsx-viewer-charts.webp` | PASS | BASELINE | Chart families, combo axes, tick labels and legends (#67/#70) |
| `xlsx-viewer-dpr2` | `/xlsx-viewer` | `tests/baseline/visual/evidence/xlsx-viewer-dpr2.webp` | PASS | BASELINE | Backing-store resolution at DPR 2 (#15) |

## Confirmed / refuted / needs-another-fixture

| Issue | Finding | Verdict | Evidence | Observed | Remaining gap |
|---:|---|---|---|---|---|
| #13 | XLSX grid per-paint conditional-format stats, worksheet handles and sheet thumbnails | **needs-another-fixture** | `evidence/xlsx-viewer-databar-icons.webp`, `evidence/xlsx-viewer-charts.webp` | The data-bar/icon workbook and chart workbook render without visible layout or paint failures; conditional-format output is present. | This is a per-frame performance finding and cannot be confirmed or refuted from screenshots. It needs a CDP trace or the component perf suite. |
| #14 | XLSX frozen panes ignore the freeze offset in headers, hit-testing, repaint and the editor | **refuted** | `evidence/xlsx-viewer-freeze-merge.webp`, `evidence/xlsx-surface-freeze-merge.webp` | charts-images.xlsx reports freezeRow=2/freezeColumn=2. After scrolling to column K / row 13, A–B and rows 1–2 stay pinned and painted; the grid host and surface agree. | Pointer hit-testing, column-header resize and the inline editor position were not exercised by a screenshot scenario. |
| #15 | XLSX merged regions, anchored objects, text overflow, icon sets, DPR and canvas limits | **refuted** | `evidence/xlsx-viewer-freeze-merge.webp`, `evidence/xlsx-viewer-databar-icons.webp`, `evidence/xlsx-viewer-iconsets.webp`, `evidence/xlsx-viewer-dpr2.webp` | The merged region is still drawn after its anchor is scrolled away; negative data bars and 4/5-bucket icon glyphs render; at DPR 2 every visible grid canvas backing store follows the 2× ratio. | Hidden-row/column anchor drift and text overflow into empty neighbours need a dedicated fixture. Canvas pixel limits need an oversized-worksheet fixture. |
| #27 | DOCX multi-page shell page-height cache, thumbnails, gutter inflation, fit-width and drag outline | **refuted** | `evidence/docx-viewer-page-sizes.webp`, `evidence/docx-viewer-thumbnails.webp` | page-sizes.docx renders five page surfaces and five thumbnails with distinct heights (198/99/181 px at a fixed 140 px width), so portrait/landscape pages are not forced into one aspect ratio. thumbnails render text and images. | The stale per-index measured-height cache and thumbnail cache thrash require a scripted document switch; this snapshot only verifies the visible shell. |
| #28 | DOCX editor global keydown scope, tracked-changes toolbar, dead context menu, fake paste, dirty tracking | **needs-another-fixture** | `evidence/docx-editor-default.webp` | The editor route mounts the toolbar, editable page surface and status line (status ready, editable state, undo/redo indicators). | Global Ctrl+B/Ctrl+S scope, context-menu actions, paste insertion and dirty tracking are interaction behaviours. They need a keyboard/pointer fixture and are already covered by the behavioural suite rather than a screenshot. |
| #29 | DOCX comment cards are not anchored to text and scale with page zoom | **refuted** | `evidence/docx-viewer-comments.webp`, `evidence/docx-viewer-comments-zoom50.webp`, `evidence/docx-viewer-comments-zoom200.webp` | review-threads.docx renders 11 cards, 9 inline markers and 9 comment highlights. Cards are ordered by anchor and the 100% layout matches anchor + 4 px collision stacking to 0 px error. At 50% and 200% viewer zoom the card heights stay 71/57 px (no pane scaling). | Card click → anchor reveal and highlight click → card focus were not scripted. The 50/200% screenshots put the gutter outside the narrow viewport, so the height evidence comes from the DOM snapshot. |
| #30 | DOCX comment collection loses replies, header/footer comments, edited comments and repeats table anchors | **refuted** | `evidence/docx-viewer-comments.webp`, `evidence/docx-viewer-tables.webp` | The review-threads fixture yields a header comment, a thread with two replies (Alex/Morgan), a table-cell comment, a cross-paragraph comment and unique card ids. table-structures.docx renders merged, nested and repeating-header tables without duplicate cards. | Editing a commented paragraph and a table that spans pages were not exercised; those need scripted edits. |
| #36 | DOCX floating images page origin, behind-text z-order, rotation, wrap and stale drag hit box | **needs-another-fixture** | `evidence/docx-viewer-embedded.webp` | docx-embedded-objects.docx loads and renders its text box, shape and SmartArt through the file input. The snapshot reports 0 floating-image layers. | None of the committed fixtures contains a floating image with wrap/exclusion settings, so the reported floating-image defects cannot be observed. A floating-image fixture is required. |
| #44 | DOCX thumbnail signature, resolution cache, stale unavailable paint and O(N²) invalidation | **refuted** | `evidence/docx-viewer-thumbnails.webp`, `evidence/docx-viewer-page-sizes.webp` | The thumbnail panel renders real page content, keeps images and headers, and gives mixed-size pages distinct aspect ratios. No blank or 'unavailable' thumbnail is visible after load. | Style-table signature coverage and cache hits after a DPR/resolution change need a scripted edit/DPR fixture; this snapshot only proves the rendered output. |
| #48 | DOCX thumbnails lose scoped CSS, bake in editor chrome and have no image-decode timeout | **refuted** | `evidence/docx-viewer-thumbnails.webp` | The report-with-image thumbnails show the styled heading, table borders and embedded image; no active-paragraph tint or editor handles are baked into the thumbnail. | A stalled image decoder needs a fault-injection fixture to exercise the timeout path; that cannot be observed from a normal screenshot. |
| #67 | XLSX charts: data labels, signed 100% denominator, log axis, smooth/markers, combo axes, date axis | **needs-another-fixture** | `evidence/xlsx-viewer-charts.webp` | xlsx-chart-families.xlsx renders two visible chart items with a column chart, a negative bar, value ticks -10…30, Q1–Q4 category labels and a Revenue legend. No rendering crash or clearly wrong geometry is visible. | Data labels, log axes, smooth/marker options, 100% stacking with mixed signs, date axes and the one-axis/two-axis combo distinction need dedicated fixtures. The current fixture only covers basic families and combo sheets. |
| #70 | XLSX chart axes and layout: gridlines, label position/skip, tick number formats, legend wrap, plot insets, pie angles | **needs-another-fixture** | `evidence/xlsx-viewer-charts.webp` | Basic value ticks, category labels and the legend are drawn as SVG text; the first two chart families are not clipped or overlapping at 1440 px. | Percent/currency/date tick formats, 12 long category labels, legend wrapping, gridline/label-skip configuration and BarOfPie angles are not present in the committed fixture. They need another workbook. |

## Reproduction

```bash
python3 scripts/ci/capture-visual-matrix.py --update-baseline
python3 scripts/ci/capture-visual-matrix.py --check
```

Structural assertions are blocking. Pixel diffs are reported but do not fail the command until the gate owned by #165 enables `--fail-on-pixel`.
