# L2 route parity report

- harness: `extend-ui-route-parity`
- upstream pin: `70fe70c6f1e1eee53954c6923f3d9f583557a57e`
- device scale factor: 2
- thresholds: geometry 0px, pixels 0.10%, channel delta 2

## Route matrix

| route | gate | status | cases | worst pixels | worst chrome | worst geometry |
|---|---|---|---|---|---|---|
| `pdf-viewer` | blocking | pass | 4 | 26.017% | 0.000% | 1.73px |

## Route gates

- `pdf-viewer` (blocking): blocking checks chrome-only pixels, landmarks, console errors
  full-page pixels are reported as diagnostic because the page bitmap comes from a different pinned renderer version than the reference

## Cases

### `pdf-viewer`

| viewport | theme | status | pixels | chrome | geometry | failures |
|---|---|---|---|---|---|---|
| 1440x900 | light | pass-with-diagnostic | 13.584% | 0.000% | 1.73px | — (diagnostic: pixels) |
| 1440x900 | dark | pass-with-diagnostic | 15.205% | 0.000% | 1.73px | — (diagnostic: pixels) |
| 390x844 | light | pass-with-diagnostic | 23.478% | 0.000% | 1.73px | — (diagnostic: pixels) |
| 390x844 | dark | pass-with-diagnostic | 26.017% | 0.000% | 1.73px | — (diagnostic: pixels) |

Dominant chrome colours (background evidence):

- `1440x900/light`: reference `#fcfcfc` 91.43% vs subject `#fcfcfc` 91.43%
- `1440x900/dark`: reference `#0d0d0d` 91.42% vs subject `#0d0d0d` 91.42%
- `390x844/light`: reference `#fcfcfc` 75.25% vs subject `#fcfcfc` 75.25%
- `390x844/dark`: reference `#0d0d0d` 75.21% vs subject `#0d0d0d` 75.21%

Classified differences (product block, scaffolding excluded):

| id | category | blocking | status |
|---|---|---|---|
| `sidebar-default-open` | product-component | yes | fixed |
| `optional-status-bar` | product-component | yes | fixed |
| `canvas-background` | product-component | yes | fixed |
| `dark-page-inversion` | product-component | yes | fixed |
| `page-raster-softening` | product-component | yes | fixed |
| `mobile-canvas-padding` | product-component | yes | fixed |
| `page-raster-size-rounding` | product-component | no | open |
| `page-raster-renderer-version` | environment | no | accepted |

Excluded scaffolding (never compared):

- `reference` `[data-slot="code"]` — The docs page renders a usage code block under the preview; the product region ends at [data-slot="pdf-viewer"].
- `reference` `[data-slot="component-preview"]` — Docs preview card (rounded border, page margins) wraps the product component; the harness reproduces only its box metrics so both regions receive the same width.
- `subject` `[data-testid="pdf-viewer-page"] > h2, [data-testid="pdf-viewer-page"] > .desc, [data-testid="pdf-viewer-page"] > .controls, [data-testid="pdf-viewer-page"] > .policy-panel, [data-testid="pdf-viewer-page"] > .status-grid, [data-testid="pdf-viewer-page"] > .diagnostics` — Demo-only page header, sample selector, runtime-policy panel, status grid and diagnostics list; the upstream docs page has none of them.
- `subject` `.nav` — Global demo navigation above the page. The capture hides it so the product region sits in the reference host shell (24px page padding + 1px frame) instead of below a sticky nav; it is never part of the region.
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
| `pdf-viewer` | 390x844/light | page1 | +0.00 | +0.00 | -0.27 | -0.59 | 1px | pass |
| `pdf-viewer` | 390x844/light | page1Raster | -1.00 | -1.00 | +1.73 | +1.41 | 2px | pass |
| `pdf-viewer` | 390x844/dark | root | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/dark | toolbar | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/dark | viewport | +0.00 | +0.00 | +0.00 | +0.00 | 0px | pass |
| `pdf-viewer` | 390x844/dark | page1 | +0.00 | +0.00 | -0.27 | -0.59 | 1px | pass |
| `pdf-viewer` | 390x844/dark | page1Raster | -1.00 | -1.00 | +1.73 | +1.41 | 2px | pass |

