# Visual baseline approvals

Every capture of `tests/visual/extend/**` must be recorded here with a
reviewer, an issue and a reason. `pnpm ui:baseline verify --require-approval`
fails when the committed tree is not covered by the most recent record.

| Approved | Reviewer | Issue | Reason | Files | Tree |
|---|---|---|---|---|---|
| 2026-11-20 | dim-agent | #165 | Initial 0.8 baselines for the gate-blocking Extend-parity surfaces: button, select, dialog, tooltip, file-upload, file-thumbnail, pdf-toolbar, pdf-sidebar at 1440x900 and 390x844, light and dark. | 152 | `aaed405b945a` |
| 2026-11-20 | dim-agent | #158 | First baselines for the #158 shared chrome evidence: the generic DocumentViewerSidebar (inline open/closed/skeleton + overlay open/closed) and the upstream-aligned empty/error/loading/status-bar surfaces are pixel-equal to the pinned Extend reference; 48 files at 1440x900 and 390x844, light and dark. | 200 | `8b9b1ade3859` |
| 2026-11-20 | dim-agent | #158 | Add the wrapped-message error card state: the description now keeps the upstream card width (no 14rem cap), so a long message wraps like the pinned xlsx error card. | 204 | `7a410ef75ef4` |

`pnpm ui:baseline capture --only <surface> --approve --reviewer <handle>`
`--issue <#nnn> --reason <why>` is the only supported way to add a row.
