// vue-docx composables — pending modular remigration from upstream.
// The real useDocxEditor controller will be rebuilt module-by-module here.

import { ref } from "vue"
import type { DocModel } from "@extend-ai/docx-core"

export interface UseDocxEditorOptions {
  starterModel?: DocModel
  initialFileName?: string
  initialDocumentTheme?: { mode: "light" | "dark" }
}

export interface DocxEditorController {
  model: ReturnType<typeof ref<DocModel | null>>
  canUndo: ReturnType<typeof ref<boolean>>
  canRedo: ReturnType<typeof ref<boolean>>
  isLoading: ReturnType<typeof ref<boolean>>
  error: ReturnType<typeof ref<string | null>>
}

/**
 * Stub controller. Returns reactive refs only — no editing logic yet.
 * Real implementation lands incrementally during the remigration.
 */
export function useDocxEditor(options?: UseDocxEditorOptions): DocxEditorController {
  return {
    model: ref(options?.starterModel ?? null),
    canUndo: ref(false),
    canRedo: ref(false),
    isLoading: ref(false),
    error: ref(null),
  }
}
