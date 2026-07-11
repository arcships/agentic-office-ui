// Vue composable: useDocxParagraphStyles
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31346-31373
//
// Thin wrapper around the editor controller's paragraph style state.

import type { DocxEditorController, ParagraphStyleDefinition } from "@arcships/docx-core"

export interface UseDocxParagraphStylesResult {
  paragraphStyles: ParagraphStyleDefinition[]
  selectedParagraphStyleId?: string
  setParagraphStyle: (styleId?: string) => void
}

export function useDocxParagraphStyles(
  editor: Pick<
    DocxEditorController,
    "availableParagraphStyles" | "selectedParagraphStyleId" | "setParagraphStyle"
  >
): UseDocxParagraphStylesResult {
  const setParagraphStyle = (styleId?: string): void => {
    editor.setParagraphStyle(styleId)
  }

  return {
    get paragraphStyles() { return editor.availableParagraphStyles },
    get selectedParagraphStyleId() { return editor.selectedParagraphStyleId },
    setParagraphStyle,
  }
}
