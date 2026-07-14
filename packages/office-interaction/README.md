# @arcships/office-interaction

Framework-neutral contracts and pure utilities for turning Office UI selections into precise, serializable object references.

The package does not parse Office files, render Vue components, own a reference collection, call an Agent, or modify documents. Format packages create locators; host applications decide what a confirmed reference means and what happens next.

```ts
import {
  applyOfficeSelectionKeyboard,
  confirmOfficeCandidate,
  createOfficeSelectionSessionState,
  parseOfficeObjectReference,
  type OfficeReferenceCandidatePreview,
} from "@arcships/office-interaction"

declare const candidates: readonly OfficeReferenceCandidatePreview[]
const session = createOfficeSelectionSessionState({ mode: "object", candidates })
const { event, state } = confirmOfficeCandidate(session, {
  referenceId: "selection-1",
  trigger: "pointer",
})
```

Candidate navigation and keyboard handling are pure state transitions. A confirmation emits one `OfficeReferenceConfirmEvent` and clears transient candidates; it never appends the reference to host state or assigns a business role.

Public behavior is specified in [the technical design](../../docs/product/object-reference-and-selection-technical-design.md).
