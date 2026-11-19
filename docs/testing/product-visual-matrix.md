# Product visual matrix (#74)

`scripts/ci/capture-visual-matrix.py` turns the demo's product routes into a
deterministic screenshot + DOM-structure matrix:

- **10 routes**: `docx-viewer`, `docx-editor`, `docx-surface`, `docx-parity`,
  `xlsx-viewer`, `xlsx-surface`, `pdf-viewer`, `pdf-surface`, `pptx-viewer`,
  `pptx-surface`
- **4 viewports**: 1440×900, 1280×720, 768×1024, 390×844
- **2 themes**: light / dark (URL `?theme=`)
- **3 zooms**: 50% / 100% / 200% (CSS `zoom`)

The full matrix is 240 scenarios. It also captures 17 issue-evidence scenarios
(comments, thumbnails, frozen panes/merges, data bars/icons, DPR 2, charts,
embedded objects) used by the verdict table for #13–#15, #27–#30, #36, #44,
#48, #67 and #70.

## Current status (pre-UI-migration)

The 0.8 UI migration (#157, #158–#164) is still landing, so the committed
baseline under [`tests/baseline/visual/`](../../tests/baseline/visual) is a
deliberately small snapshot:

| Routes | Viewports | Themes | Zoom | Baselines |
|---|---|---|---|---:|
| docx-viewer, xlsx-viewer, pdf-viewer, pptx-viewer | 1440×900, 390×844 | light, dark (1440 only) | 100% | 12 |
| evidence scenarios | 1440×900 (DPR 2 for one case) | light | 100% | 17 |

The full matrix was captured locally while developing the script (240/240
structural assertions passed), but it is not committed while the UI is changing.
After #159–#164 merge:

```bash
pnpm --filter demo... build
python3 scripts/ci/capture-visual-matrix.py --skip-build --update-baseline \
  --baseline-label post-ui-migration
```

## Blocking policy

- **Blocking**: route readiness, root visibility, no visible error slots, theme,
  zoom, 100%-width overflow, DOCX page surfaces/parity, XLSX grid rows/canvas,
  PDF pages, PPTX slides, and evidence-specific assertions.
- **Non-blocking**: pixel diffs (`threshold=8`, `maxDiffRatio=0.002`). Nightly
  CI uploads diff images and the Markdown/JSON report. `--fail-on-pixel` is
  reserved for the #165 gate.

## Nightly CI

The `blackbox` job in `.github/workflows/ci.yml` (nightly/manual only) captures
the full matrix after the workspace build and uploads
`output/visual-matrix/` as the `visual-matrix-<sha>` artifact. Structural
failures fail the job; pixel differences only appear in the report.

## Report

The checked-in report for the current baseline is
[`product-visual-matrix-report-pre-ui-migration.md`](./product-visual-matrix-report-pre-ui-migration.md).
It contains the route × scenario → screenshot → status table, the evidence
scenario table and the confirmed/refuted/needs-another-fixture verdict table.

Regenerate it without a browser:

```bash
python3 scripts/ci/capture-visual-matrix.py --report-only \
  --output output/visual-matrix/report
```

## Related

- [`tests/baseline/visual/README.md`](../../tests/baseline/visual/README.md) —
  naming, determinism, commands and known gaps.
- [`tests/baseline/visual/verdicts.json`](../../tests/baseline/visual/verdicts.json)
  — machine-readable verdicts.
- `tests/blackbox/ux_parity.py` (#155) — upstream Extend reference parity; kept
  independent from this product-route matrix.
