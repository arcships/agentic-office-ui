// Vue composable: useDocxLineSpacing
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31448-31465
//
// Thin wrapper around the editor controller's line-spacing state.

import type { DocxEditorController, DocxLineSpacingInfo } from "@arcships/docx-core"

export interface UseDocxLineSpacingResult {
  lineSpacing: DocxLineSpacingInfo
  setLineSpacing: (lineMultiple: number) => void
}

export function useDocxLineSpacing(
  editor: Pick<DocxEditorController, "selectedLineSpacing" | "setLineSpacing">
): UseDocxLineSpacingResult {
  const setLineSpacing = (lineMultiple: number): void => {
    editor.setLineSpacing(lineMultiple)
  }

  return {
    get lineSpacing() { return editor.selectedLineSpacing },
    setLineSpacing,
  }
}
