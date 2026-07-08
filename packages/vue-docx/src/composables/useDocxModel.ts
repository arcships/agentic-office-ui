// Vue composable: useDocxModel
// Migrated from upstream @extend-ai/react-docx, react-viewer/src/index.tsx
// React useDocxModel hook → Vue composable
//
// Parses a .docx ArrayBuffer into a DocModel via the docx-import pipeline.

import { ref, watchEffect, onScopeDispose } from "vue"
import { importDocxBuffer } from "@extend-ai/docx-core"
import type { DocModel } from "@extend-ai/docx-core"

export interface UseDocxModelState {
  model?: DocModel
  isLoading: boolean
  error?: Error
}

export function useDocxModel(file?: () => ArrayBuffer | undefined): UseDocxModelState {
  const model = ref<DocModel | undefined>(undefined)
  const isLoading = ref(false)
  const error = ref<Error | undefined>(undefined)

  let abortController: AbortController | undefined
  let isCurrent = true

  watchEffect(async () => {
    const buffer = file?.()
    if (!buffer) {
      isLoading.value = false
      error.value = undefined
      model.value = undefined
      return
    }

    abortController?.abort()
    abortController = new AbortController()
    isCurrent = false
    isCurrent = true
    const currentIteration = isCurrent

    isLoading.value = true
    error.value = undefined

    try {
      const result = await importDocxBuffer(buffer, {
        signal: abortController.signal,
        transferBuffer: false,
      })
      if (!currentIteration) return
      model.value = result.model
      isLoading.value = false
    } catch (err) {
      if (!currentIteration) return
      error.value = err instanceof Error ? err : new Error("Unknown DOCX parse error")
      isLoading.value = false
    }
  })

  onScopeDispose(() => {
    isCurrent = false
    abortController?.abort()
  })

  return {
    get model() { return model.value },
    get isLoading() { return isLoading.value },
    get error() { return error.value },
  }
}
