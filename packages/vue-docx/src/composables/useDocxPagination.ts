// Vue composable: useDocxPagination
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31693-31705
//
// Thin wrapper around the editor controller's pagination info.

import type { DocxEditorController, DocxPaginationInfo } from "@arcships/docx-core"

export interface UseDocxPaginationResult {
  pagination: DocxPaginationInfo
}

export function useDocxPagination(
  editor: Pick<DocxEditorController, "currentPage" | "totalPages">
): UseDocxPaginationResult {
  return {
    get pagination() {
      return {
        currentPage: editor.currentPage,
        totalPages: editor.totalPages,
      }
    },
  }
}
