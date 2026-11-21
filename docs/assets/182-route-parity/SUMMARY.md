# L2 route parity report

- harness: `extend-ui-route-parity`
- upstream pin: `70fe70c6f1e1eee53954c6923f3d9f583557a57e`
- device scale factor: 2
- thresholds: geometry 0px, pixels 0.10%, channel delta 2

## Route matrix

| route | gate | status | cases | worst pixels | worst chrome | worst geometry |
|---|---|---|---|---|---|---|
| `pdf-viewer` | report-only | fail | 8 | 92.388% | 92.647% | 13.00px |

## Cases

### `pdf-viewer`

| viewport | theme | status | pixels | chrome | geometry | failures |
|---|---|---|---|---|---|---|
| 1440x900 | light | fail | 56.484% | 92.063% | 1.73px | pixels, chrome-pixels |
| 1440x900 | dark | fail | 92.388% | 92.436% | 1.73px | pixels, chrome-pixels |
| 390x844 | light | fail | 38.586% | 79.567% | 13.00px | pixels, chrome-pixels, geometry |
| 390x844 | dark | fail | 78.993% | 81.076% | 13.00px | pixels, chrome-pixels, geometry |
| 1280x720 | light | fail | 56.484% | 92.063% | 1.73px | pixels, chrome-pixels |
| 1280x720 | dark | fail | 92.388% | 92.436% | 1.73px | pixels, chrome-pixels |
| 768x1024 | light | fail | 43.319% | 92.050% | 13.00px | pixels, chrome-pixels, geometry |
| 768x1024 | dark | fail | 91.287% | 92.647% | 13.00px | pixels, chrome-pixels, geometry |

Dominant chrome colours (background evidence):

- `1440x900/light`: reference `#fcfcfc` 91.43% vs subject `#e5e7eb` 91.33%
- `1440x900/dark`: reference `#0d0d0d` 91.42% vs subject `#020617` 91.33%
- `390x844/light`: reference `#fcfcfc` 75.25% vs subject `#e5e7eb` 75.21%
- `390x844/dark`: reference `#0d0d0d` 75.21% vs subject `#020617` 75.21%
- `1280x720/light`: reference `#fcfcfc` 91.43% vs subject `#e5e7eb` 91.33%
- `1280x720/dark`: reference `#0d0d0d` 91.42% vs subject `#020617` 91.33%
- `768x1024/light`: reference `#fcfcfc` 90.07% vs subject `#e5e7eb` 90.05%
- `768x1024/dark`: reference `#0d0d0d` 90.05% vs subject `#020617` 90.05%

Classified differences (product block, scaffolding excluded):

| id | category | blocking | status |
|---|---|---|---|
| `sidebar-default-open` | product-component | yes | open |
| `optional-status-bar` | product-component | yes | open |
| `canvas-background` | product-component | yes | open |
| `dark-page-inversion` | product-component | yes | open |
| `page-raster-softening` | product-component | yes | open |
| `mobile-canvas-padding` | product-component | yes | open |
| `page-raster-size-rounding` | product-component | no | open |

Excluded scaffolding (never compared):

- `reference` `[data-slot="code"]` — The docs page renders a usage code block under the preview; the product region ends at [data-slot="pdf-viewer"].
- `reference` `[data-slot="component-preview"]` — Docs preview card (rounded border, page margins) wraps the product component; the harness reproduces only its box metrics so both regions receive the same width.
- `subject` `[data-testid="pdf-viewer-page"] > h2, [data-testid="pdf-viewer-page"] > .desc, [data-testid="pdf-viewer-page"] > .controls, [data-testid="pdf-viewer-page"] > .policy-panel, [data-testid="pdf-viewer-page"] > .status-grid, [data-testid="pdf-viewer-page"] > .diagnostics` — Demo-only page header, sample selector, runtime-policy panel, status grid and diagnostics list; the upstream docs page has none of them.
- `subject` `[data-route-region="viewer"]` — Demo frame (1px border, radius, overflow) that wraps the viewer; only its box metrics are reproduced on the reference side.

## Landmark geometry

Deltas are `subject - reference` for each landmark box relative to its region origin.

| route | case | landmark | Δx | Δy | Δw | Δh | tolerance | status |
|---|---|---|---|---|---|---|---|---|
| `pdf-viewer` | 1440x900/light | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1440x900/light | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1440x900/light | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1440x900/light | page1 | +0.14 | +0.00 | -0.27 | -0.59 | 1px | pass |
| `pdf-viewer` | 1440x900/light | page1Raster | -0.86 | -1.00 | +1.73 | +1.41 | 2px | pass |
| `pdf-viewer` | 1440x900/dark | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1440x900/dark | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1440x900/dark | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1440x900/dark | page1 | +0.14 | +0.00 | -0.27 | -0.59 | 1px | pass |
| `pdf-viewer` | 1440x900/dark | page1Raster | -0.86 | -1.00 | +1.73 | +1.41 | 2px | pass |
| `pdf-viewer` | 390x844/light | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/light | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/light | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/light | page1 | -12.00 | -12.00 | -0.27 | -0.59 | 1px | fail |
| `pdf-viewer` | 390x844/light | page1Raster | -13.00 | -13.00 | +1.73 | +1.41 | 2px | fail |
| `pdf-viewer` | 390x844/dark | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/dark | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/dark | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/dark | page1 | -12.00 | -12.00 | -0.27 | -0.59 | 1px | fail |
| `pdf-viewer` | 390x844/dark | page1Raster | -13.00 | -13.00 | +1.73 | +1.41 | 2px | fail |
| `pdf-viewer` | 1280x720/light | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1280x720/light | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1280x720/light | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1280x720/light | page1 | +0.14 | +0.00 | -0.27 | -0.59 | 1px | pass |
| `pdf-viewer` | 1280x720/light | page1Raster | -0.86 | -1.00 | +1.73 | +1.41 | 2px | pass |
| `pdf-viewer` | 1280x720/dark | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1280x720/dark | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1280x720/dark | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 1280x720/dark | page1 | +0.14 | +0.00 | -0.27 | -0.59 | 1px | pass |
| `pdf-viewer` | 1280x720/dark | page1Raster | -0.86 | -1.00 | +1.73 | +1.41 | 2px | pass |
| `pdf-viewer` | 768x1024/light | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 768x1024/light | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 768x1024/light | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 768x1024/light | page1 | +0.14 | -12.00 | -0.27 | -0.59 | 1px | fail |
| `pdf-viewer` | 768x1024/light | page1Raster | -0.86 | -13.00 | +1.73 | +1.41 | 2px | fail |
| `pdf-viewer` | 768x1024/dark | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 768x1024/dark | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 768x1024/dark | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 768x1024/dark | page1 | +0.14 | -12.00 | -0.27 | -0.59 | 1px | fail |
| `pdf-viewer` | 768x1024/dark | page1Raster | -0.86 | -13.00 | +1.73 | +1.41 | 2px | fail |

