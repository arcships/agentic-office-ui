// Vue composable: useDocxBorders
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31467-31491
//
// Thin wrapper around the editor controller's border state.

import type {
  DocxEditorController,
  DocxBorderContext,
  DocxBorderPreset,
  DocxBorderPresetState,
} from "@extend-ai/docx-core"

export interface UseDocxBordersResult {
  borderContext: DocxBorderContext
  activeBorderPresets: DocxBorderPresetState
  applyBorderPreset: (preset: DocxBorderPreset) => void
}

export function useDocxBorders(
  editor: Pick<
    DocxEditorController,
    "selectedBorderContext" | "activeBorderPresets" | "applyBorderPreset"
  >
): UseDocxBordersResult {
  const applyBorderPreset = (preset: DocxBorderPreset): void => {
    editor.applyBorderPreset(preset)
  }

  return {
    get borderContext() { return editor.selectedBorderContext },
    get activeBorderPresets() { return editor.activeBorderPresets },
    applyBorderPreset,
  }
}
