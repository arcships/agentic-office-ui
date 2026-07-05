import type { DocModel } from "@extend-ai/docx-core"
import type { useDocxEditor } from "./composables"

export interface DocxViewerProps {
  file?: ArrayBuffer
  model?: DocModel | null
  className?: string
  class?: string
  style?: Record<string, string | number>
  layoutOptions?: Record<string, unknown>
  emptyState?: unknown
}

export interface DocxEditorViewerProps {
  editor: ReturnType<typeof useDocxEditor>
  className?: string
  class?: string
  style?: Record<string, string | number>
}
