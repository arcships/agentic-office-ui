# Product visual matrix baselines

This directory owns the **product-route** visual matrix (issue #74). It is
independent from the upstream Extend reference harness
(`tests/blackbox/ux_parity.py`, #155): that harness compares local routes to an
upstream React checkout, while this matrix only captures the local demo's
product routes and their structural invariants.

## Current status: pre-UI-migration snapshot

The repository is in the middle of the 0.8 UI migration (#157 primitives,
#158–#164 page parity). The full `route × viewport × theme × zoom` matrix will
change underneath those branches, so the committed baseline is deliberately
small:

- `docx-viewer`, `xlsx-viewer`, `pdf-viewer`, `pptx-viewer`
- `desktop-1440x900` in light and dark
- `mobile-390x844` in light
- 100% zoom

That is **12 matrix baselines** plus **17 issue-evidence baselines**
(`evidence/`). The capture script still reproduces the full matrix (10 routes ×
4 viewports × 2 themes × 3 zooms = 240 scenarios). Recapture and commit the full
matrix after #159–#164 land:

```bash
python3 scripts/ci/capture-visual-matrix.py --update-baseline \
  --baseline-label post-ui-migration
```

See `manifest.json` for the exact files, SHA-256 hashes and captured Git HEAD.

## Commands

Preview must not already be running on the selected port (default `4186`). The
script builds `demo` and its workspace dependencies unless `--skip-build` is
passed, starts `vite preview`, and always stops the preview process group in a
`finally` block.

```bash
# Full matrix + evidence into output/visual-matrix/<timestamp>/
python3 scripts/ci/capture-visual-matrix.py

# Rewrite this baseline directory
python3 scripts/ci/capture-visual-matrix.py --update-baseline

# Capture again and compare; structural assertions are blocking,
# pixel diffs are reported but do not fail unless --fail-on-pixel is passed
python3 scripts/ci/capture-visual-matrix.py --check

# One route/viewport/theme/zoom during development
python3 scripts/ci/capture-visual-matrix.py --skip-build \
  --filter docx-viewer --viewport desktop-1440x900 --theme dark --zoom 200 \
  --no-evidence --output /tmp/visual-matrix

# Rebuild a Markdown table from the committed manifest without a browser
python3 scripts/ci/capture-visual-matrix.py --report-only \
  --output /tmp/visual-matrix-report
```

Local Python needs Playwright and Pillow (`/tmp/bb-venv/bin/python` in this
workspace). CI installs `requirements-ci.txt`.

## Naming scheme

```text
tests/baseline/visual/
├── config.json
├── manifest.json
├── screenshots/<route>/<viewport>/<theme>-zoom<050|100|200>.webp
├── structure/<route>/<viewport>/<theme>-zoom<050|100|200>.json
└── evidence/<scenario>.webp
    evidence/<scenario>.json
```

Screenshots are stored as **lossless WebP** to keep the repository small while
preserving exact pixels. The capture run keeps PNGs under `output/` for human
review; the pixel runner decodes both formats with Pillow.

`<route>` is the demo route id (`docx-viewer`, `xlsx-surface`, …), `<viewport>`
is one of `desktop-1440x900`, `laptop-1280x720`, `tablet-768x1024`,
`mobile-390x844`, and `<theme>` is `light` or `dark`.

## Determinism

Each scenario gets a fresh browser context with:

- Chromium (Playwright bundle, system Chromium/Chrome fallback)
- DPR 1, sRGB, hidden scrollbars, `--font-render-hinting=none`,
  `--disable-lcd-text`
- `locale=en-US`, `timezone_id=UTC`, `reducedMotion=reduce`
- URL theme: `?theme=light` / `?theme=dark` (the demo reads it before mount)
- CSS `zoom` = 0.5 / 1 / 2 for the zoom axis
- fixed font stack `Liberation Sans → DejaVu Sans → Noto Sans CJK SC → Arial`
- animations and transitions disabled, caret hidden
- `document.fonts.ready`, images complete, two `requestAnimationFrame` ticks and
  a 300 ms settle before the screenshot
- route-specific readiness predicates (page status, grid rows, PDF pages,
  PPTX slides, parity `PASS`)

The structure dump records only stable data: tags, selected classes/testids,
selected `data-*` state, rounded rectangles, counts and short text. It omits
random ids, timestamps and computed style strings.

## Blocking policy

| Check | Blocking? | Owner |
|---|---|---|
| Route readiness, root visible, no visible error slots | yes | this script |
| Theme, zoom, no horizontal overflow at 100% | yes | this script |
| DOCX page surfaces, parity `PASS` | yes | this script |
| XLSX grid rows/columns/canvas, dark canvas sample | yes | this script |
| PDF pages, PPTX slides | yes | this script |
| Evidence-specific assertions (`evidence_assertions`) | yes | this script |
| Pixel diff (`threshold=8`, `maxDiffRatio=0.002`) | **no** | #57 / #165 later gate |

Pixel diffs are nightly-only and non-blocking: the nightly job uploads the
report and diff images as artifacts, but only structural failures fail the job.
`--fail-on-pixel` is reserved for the eventual blocking gate in #165.

## Issue evidence

`evidence/` contains the screenshots and DOM snapshots used by the verdict table
in [`verdicts.json`](./verdicts.json). Each evidence scenario is defined in
`EVIDENCE_SCENARIOS` in the capture script and targets one of the historical
rendering findings:

| Scenario | Finding |
|---|---|
| `docx-viewer-comments*` | #29 / #30 comment anchoring, ordering, zoom, replies, header/table comments |
| `docx-viewer-thumbnails` | #44 / #48 thumbnail panel content and rendering |
| `docx-viewer-page-sizes` | #27 mixed page sizes and per-page thumbnail aspect ratios |
| `docx-viewer-tables` | #30 merged/nested tables and repeating headers |
| `docx-viewer-embedded` | #36 embedded objects; the fixture does **not** contain a floating image |
| `docx-editor-default` | #28 editor shell (interaction behaviour needs another fixture) |
| `xlsx-viewer-freeze-merge`, `xlsx-surface-freeze-merge` | #14 / #15 freeze panes and merged regions after scrolling |
| `xlsx-viewer-databar-icons`, `xlsx-viewer-iconsets` | #15 data bars and icon sets |
| `xlsx-viewer-dpr2` | #15 DPR 2 canvas backing store |
| `xlsx-viewer-charts` | #67 / #70 chart families, ticks, legend, combo charts |
| `pdf-viewer-rotated`, `pptx-viewer-deck`, `pptx-surface-stack` | general route coverage |

## Known gaps

- #13 is a performance finding; screenshots cannot observe per-frame allocations.
- #28's keydown scope, context menu, paste and dirty tracking are interaction
  behaviours; a screenshot only confirms the editor shell.
- #36 needs a DOCX fixture with an actual floating image (none of the current
  samples contains one).
- #67/#70 need percent/date axis and long-legend fixtures to observe every
  sub-claim. The current `xlsx-chart-families.xlsx` covers families, combo
  charts, categories and basic value ticks.
- #27's stale page-height cache and thumbnail cache thrash need a scripted
  document switch, not a single snapshot.
- Cross-OS font rendering still differs; that is why pixel diffs remain
  non-blocking until the dedicated gate is calibrated on CI.
