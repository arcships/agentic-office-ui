# Visual baseline approvals

Every capture of `tests/visual/extend/**` must be recorded here with a
reviewer, an issue and a reason. `pnpm ui:baseline verify --require-approval`
fails when the committed tree is not covered by the most recent record.

| Approved | Reviewer | Issue | Reason | Files | Tree |
|---|---|---|---|---|---|
| 2026-11-20 | dim-agent | #165 | Initial 0.8 baselines for the gate-blocking Extend-parity surfaces: button, select, dialog, tooltip, file-upload, file-thumbnail, pdf-toolbar, pdf-sidebar at 1440x900 and 390x844, light and dark. | 152 | `aaed405b945a` |

`pnpm ui:baseline capture --only <surface> --approve --reviewer <handle>`
`--issue <#nnn> --reason <why>` is the only supported way to add a row.
