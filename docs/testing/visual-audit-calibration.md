# L3 calibration: PDF viewer, pre-#199 pair

This calibration used the historical `docs/assets/182-route-parity` pair. It
validates the review instructions; it is **not** a current product finding.
Current audits must recapture the reference and subject from the current build
and store a DOM snapshot with the screenshots.

## Inputs

- `/pdf-viewer`, 1440x900 light/dark and 390x844 light.
- A/B randomly assigned; the answer key was kept outside the reviewer
  workspace.
- No checklist, no source, no answer key.

## What the reviewers found

A text-only model (`deepseek-v4-flash`, `vision:false`) produced 17 measured
candidates but explicitly reported that it received image references, not
pixels. It could measure colours and boxes with PIL but could not visually
confirm them.

A vision model (`deepseek-v4.1-flash`) read the PNGs and produced 14 confirmed
candidates, including:

- workspace field colours (light/dark);
- page shadow/softening and dark-mode page inversion;
- the header-left label/icon module;
- zoom pill height and chevron presence;
- mobile page inset;
- the fact that the 390 A/B mapping was swapped relative to the desktop pairs.

It refused to transcribe the header text and described only the visual glyph
run, which is the desired behaviour. It still did not report the logo chip or
the sidebar/status-bar defaults; those need a second reviewer or a targeted
crop, and DOM confirmation.

## Conclusions

1. **Vision capability is mandatory.** Text-only models can only measure.
2. **A single pass is not enough.** Use two reviewers and targeted crops for
   low-contrast extras such as the logo chip.
3. **Locale/fixture must be normalized.** The header difference was
   `page 1 of 4` vs `1 共 4`, not a different component.
4. **The old pair is calibration only.** It cannot prove the current code has
   or lacks a defect.
5. **Tile/coverage is the production path.** The raw two-image experiment was
   useful for instruction validation; production audits use the atlas and
   coverage contract in [visual-audit.md](visual-audit.md).
