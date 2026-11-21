# `tests/visual/extend` — committed Extend-parity baselines

These PNGs are the subject side of the Extend UI parity matrix (#165): they
capture **this repository's** components, so a PR that changes a toolbar's
spacing or colour fails without waiting for the upstream reference app to
change.

- Naming: `<surface>/<state>/<theme>-<viewport>.png` (e.g.
  `pdf-toolbar/default/dark-1440x900.png`).
- Matrix: all four parity viewports — `1440x900`, `1280x720`, `768x1024` and
  `390x844` — × `light` and `dark`, DPR 2, 12px capture padding around the
  measured target (#174). 25 gate-blocking surfaces × their states × 4
  viewports × 2 themes = **864 PNGs** (~14 MiB).
- CI coverage: the pull-request job captures and compares the fast pair
  (`coreSubset.viewports` = `1440x900` + `390x844`); the nightly full matrix
  captures and compares all four viewports, so a 1280x720 or 768x1024
  regression blocks there even when the upstream reference is unchanged.
- Only surfaces with `gateStatus: blocking` in
  `tools/parity-fixtures/manifest.json` have baselines.
- `manifest.json` records every file's SHA-256, size and dimensions, the
  `treeSha256` of the whole tree, and the approval that produced it.
- `APPROVALS.md` is generated from `manifest.json.approvals`.

## Updating a baseline

There is deliberately no `--update` flag. A refresh is an approval:

```bash
# 1. Inspect: file count, missing files, whether the tree matches the manifest
pnpm ui:baseline:status

# 2. Capture + approve (refuses without --approve, --reviewer, --issue, --reason)
pnpm ui:baseline capture --only file-upload \
  --approve --reviewer <github-handle> --issue "#163" \
  --reason "dropzone padding now uses --office-control-padding-inline"

# 3. Verify the contract (this is the CI step)
pnpm ui:baseline:verify
```

`capture` runs the real harness (`scripts/ci/pixel-parity.mjs --write-baseline`)
for the requested surfaces, rewrites `manifest.json`, and appends an approval
record. Review the PNG diff and `output/ui-parity/report.html` in the pull
request; the reason has to explain why the visual change is correct.

`verify --require-approval` fails when:

- a committed PNG no longer matches the SHA-256 in `manifest.json`;
- the directory contains a PNG the manifest does not list;
- a gate-blocking surface/state/theme/viewport combination has no baseline;
- the current `treeSha256` is not covered by the latest approval record.

## When a surface becomes gate-blocking

1. Flip `gateStatus` to `blocking` in `tools/parity-fixtures/manifest.json`
   (and add `core: true` if it should run on pull requests).
2. `pnpm ui:baseline capture --only <surface> --approve ...`.
3. `pnpm ui:baseline:verify` and
   `pnpm ui:parity -- --only <surface> --baseline --gate-mode ready` must pass.
4. Update `docs/testing/ui-parity-ci.md` §4 and
   `docs/components/ui-component-map.md`.

Background and thresholds: [docs/testing/ui-parity-ci.md](../../../docs/testing/ui-parity-ci.md).
