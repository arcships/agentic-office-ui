// Vue composable: useDocxModel
// Migrated from upstream @extend-ai/react-docx, react-viewer/src/index.tsx
// React useDocxModel hook → Vue composable
//
// Parses a .docx ArrayBuffer into a DocModel via the docx-import pipeline.

import { ref, watchEffect, onScopeDispose } from "vue"
import { createDocxRuntime } from "@arcships/docx-core"
import type { DocModel } from "@arcships/docx-core"

export interface UseDocxModelState {
  model?: DocModel
  isLoading: boolean
  error?: Error
}

export function useDocxModel(file?: () => ArrayBuffer | undefined): UseDocxModelState {
  const model = ref<DocModel | undefined>(undefined)
  const isLoading = ref(false)
  const error = ref<Error | undefined>(undefined)

  const runtime = createDocxRuntime()
  const loader = runtime.createLoader()

  watchEffect((onCleanup) => {
    const buffer = file?.()
    let active = true
    onCleanup(() => { active = false })

    if (!buffer) {
      loader.cancel()
      isLoading.value = false
      error.value = undefined
      model.value = undefined
      return
    }

    isLoading.value = true
    error.value = undefined

    void loader.load(buffer, { transferBuffer: false })
      .then((result) => {
        if (!active) return
        model.value = result.model
        error.value = undefined
      })
      .catch((err: unknown) => {
        if (!active) return
        error.value = err instanceof Error ? err : new Error("Unknown DOCX parse error")
      })
      .finally(() => {
        if (!active) return
        isLoading.value = false
      })
  })

  onScopeDispose(() => {
    loader.dispose()
    runtime.dispose()
  })

  return {
    get model() { return model.value },
    get isLoading() { return isLoading.value },
    get error() { return error.value },
  }
}
