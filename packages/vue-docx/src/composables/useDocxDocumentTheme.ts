// Vue composable: useDocxDocumentTheme
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31321-31344
//
// Thin wrapper around the editor controller's document-theme state.
// Exposes theme getter/setter + toggle + dark-mode check.

import type { DocxEditorController, DocxDocumentTheme } from "@extend-ai/docx-core"

export interface UseDocxDocumentThemeResult {
  documentTheme: DocxDocumentTheme
  isDarkDocument: boolean
  setDocumentTheme: (theme: DocxDocumentTheme) => void
  toggleDocumentTheme: () => void
}

export function useDocxDocumentTheme(
  editor: Pick<DocxEditorController, "documentTheme" | "setDocumentTheme">
): UseDocxDocumentThemeResult {
  const setDocumentTheme = (theme: DocxDocumentTheme): void => {
    editor.setDocumentTheme(theme)
  }

  const toggleDocumentTheme = (): void => {
    editor.setDocumentTheme(editor.documentTheme === "dark" ? "light" : "dark")
  }

  return {
    get documentTheme() { return editor.documentTheme },
    get isDarkDocument() { return editor.documentTheme === "dark" },
    setDocumentTheme,
    toggleDocumentTheme,
  }
}
