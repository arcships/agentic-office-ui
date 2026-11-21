# Visual baseline approvals

Every capture of `tests/visual/extend/**` must be recorded here with a
reviewer, an issue and a reason. `pnpm ui:baseline verify --require-approval`
fails when the committed tree is not covered by the most recent record.

| Approved | Reviewer | Issue | Reason | Files | Tree |
|---|---|---|---|---|---|
| 2026-11-20 | dim-agent | #165 | Initial 0.8 baselines for the gate-blocking Extend-parity surfaces: button, select, dialog, tooltip, file-upload, file-thumbnail, pdf-toolbar, pdf-sidebar at 1440x900 and 390x844, light and dark. | 152 | `aaed405b945a` |
| 2026-11-20 | dim-agent | #165 | Promote #157 input, #159 DOCX chrome, #160 XLSX chrome and #164 blocks to the gate after exact full-matrix parity on 0.8; committed subject baselines at 1440x900 and 390x844, light and dark. | 348 | `b7c1cea9ec84` |
| 2026-11-20 | dim-agent | #161 | PPTX toolbar and sorter chrome landed with #161b; exact 64/64 full-matrix parity rerun on the merged 0.8 tree, so commit the subject baselines at 1440x900 and 390x844, light and dark. | 380 | `39b91f9c0919` |
| 2026-11-20 | dim-agent | #165 | Merge #158b document-sidebar and four office-state surfaces into the final gate; 104/104 full-matrix parity on the merged 0.8 tree, so commit the 52 subject baselines at 1440x900 and 390x844, light and dark. | 432 | `87c5b815a240` |
| 2026-11-20 | dim-agent | #165 | pin grayscale font rasterization so CI and local baselines match | 432 | `58d6dda6fc1a` |
| 2026-11-20 | dim-agent | #165 | pin grayscale LCD text only; keep native font hinting metrics | 432 | `64ac2a03e7c0` |

`pnpm ui:baseline capture --only <surface> --approve --reviewer <handle>`
`--issue <#nnn> --reason <why>` is the only supported way to add a row.
