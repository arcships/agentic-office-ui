# L3 visual difference audit

This is the vision-first audit that sits beside L1 component parity and L2
route parity. It answers a different question:

> What is visually different, which element is extra/missing/misaligned, and
> is it a native control, border, divider, or composition problem?

It is not an aesthetic review, and it does not replace the pixel/geometry
gates.

## Rules

- **No checklist is given to the reviewer.** The reviewer finds differences
  first; categories are assigned only after the fact.
- **Use a vision model.** `deepseek-v4.1-flash` or another model with
  `vision: true` is required. A text-only model receives only image references
  and cannot perform visual confirmation.
- **One image is not enough.** A full screenshot may be downscaled. Use an
  overview plus overlapping native-resolution tiles, then candidate crops.
- **Visual discovery and DOM/component confirmation are separate.** DOM and
  source explain a candidate; they do not replace the visual observation.
- **Every confirmed candidate needs a visual crop.** Measurements alone are
  not a visual verdict.

## Atlas

Build an atlas from a reference/subject pair:

```bash
/tmp/bb-venv/bin/python scripts/ci/visual-audit-tiles.py \
  --reference audit/pdf-viewer/1440x900/light/reference.png \
  --subject audit/pdf-viewer/1440x900/light/subject.png \
  --out audit/pdf-viewer/1440x900/light \
  --route /pdf-viewer --viewport 1440x900 --theme light --state default \
  --locale en --zoom 100 \
  --fixture apps/demo/public/samples/sample.pdf \
  --fixture-sha256 <sha256> \
  --capture-sha <git-sha>
```

The command writes:

```text
overview/A.png, B.png, AB.png
tiles/<tile-id>/{A,B,AB}.png
manifest.json
coverage.json
```

Defaults are 512 px tiles with 64 px overlap. Use 640 px tiles only if the
reviewer's image input keeps them at native resolution; verify with the
calibration set.

## Review flow

```text
overview + all tiles
  -> blocker: every tile must be marked reviewedBy at least one reviewer
  -> reviewer A: find candidates in its own words
  -> candidate crops (1:1, 4x, 8x) + annotations + measurements
  -> DOM/component confirmation (element, native control, wrapper, style)
  -> reviewer B: independent blind pass, shuffled tile order
  -> adjudication: confirmed / rejected / needs-more-evidence
```

The A/B mapping is randomized per case and kept outside the reviewer
workspace. Reviewer B must not see reviewer A's candidate list.

## Candidate contract

```json
{
  "id": "pdf-toolbar-extra-divider-01",
  "route": "/pdf-viewer",
  "viewport": "1440x900",
  "theme": "light",
  "state": "default",
  "tile": "r2c3",
  "region": {"x": 120, "y": 16, "w": 1, "h": 32},
  "observation": "reference has no vertical line here; subject has a pale line",
  "measurement": "subject 1px #e5e5e5; reference 0 lines",
  "domFinding": "subject wrapper has border-inline-end",
  "componentFinding": "OfficeToolbarSeparator is rendered twice",
  "visualVerdict": "visual-confirmed",
  "severity": "major",
  "confidence": 0.95,
  "evidence": ["candidate/pdf-toolbar-extra-divider-01/ref-crop.png",
               "candidate/pdf-toolbar-extra-divider-01/subj-crop.png",
               "candidate/pdf-toolbar-extra-divider-01/overlay.png"]
}
```

`visualVerdict` is one of `candidate`, `measured`, `dom-confirmed`,
`visual-confirmed`, `rejected`, `needs-more-evidence`.

## Evidence upload and issue reporting

- **Full audit** stays in a CI artifact or a GitHub Release asset:
  `visual-audit-<sha>.zip` with all tiles, crops, reports and DOM snapshots.
- **Curated issue evidence** is committed only for confirmed candidates under
  `docs/assets/visual-audit/<issue>/`: reference crop, subject crop, overlay.
  Use WebP or compressed PNG; keep each image under 200 KiB.
- **Do not embed local paths in an issue.** GitHub does not upload them.
  The REST API has no issue-attachment upload endpoint.
- **Recommended issue image source**:
  `https://raw.githubusercontent.com/arcships/agentic-office-ui/main/docs/assets/visual-audit/<issue>/<id>-overlay.webp`
  (for a PR before merge, use the branch raw URL or a Release asset).
- The issue table should contain only a few compact crops; full evidence is
  linked.

## Post-fix visual verification

Discovery is not acceptance. A confirmed candidate is closed only after:

1. the same reference, viewport, theme, state, zoom, locale and fixture are
   captured again;
2. the same tile/crop coordinates are used for `reference`, `before-subject`
   and `after-subject`;
3. a blind reviewer confirms the target difference is gone;
4. the whole route is swept again for new candidates;
5. L1/L2 gates still pass.

Allowed close states:

- `visual-verified` — target crop is clean;
- `accepted-divergence` — deliberate product difference with reason and
  evidence;
- never close with only a code merge or a measurement.

## Route plan

The route matrix, viewport/state phases, calibration thresholds and child
issue breakdown live in the tracking issue:
<https://github.com/arcships/agentic-office-ui/issues/212>.
