import { shallowRef, onScopeDispose, getCurrentScope } from "vue"
import type { ShallowRef } from "vue"
import type { XlsxViewerController } from "@arcships/xlsx-core"

export interface UseClipboardCopyNoticeOptions {
  /** How long the failure notice stays visible. */
  durationMs?: number
  /** Message shown when the browser refuses the copy. */
  failureMessage?: string | (() => string)
}

export interface ClipboardCopyNotice {
  readonly notice: Readonly<ShallowRef<string>>
  /** Copies the current selection, reporting a refusal instead of throwing. */
  copySelection(): Promise<boolean>
  clear(): void
}

const DEFAULT_FAILURE_MESSAGE = "Copy failed: this browser does not allow clipboard writes."
const DEFAULT_DURATION_MS = 3200

/**
 * Copies the selection and turns a refused clipboard write into a transient
 * status message, so a denied copy is visible instead of becoming an
 * unhandled rejection.
 */
export function useClipboardCopyNotice(
  controller: () => XlsxViewerController,
  options: UseClipboardCopyNoticeOptions = {},
): ClipboardCopyNotice {
  const notice = shallowRef("")
  const durationMs = options.durationMs ?? DEFAULT_DURATION_MS
  const resolveFailureMessage = () => typeof options.failureMessage === "function"
    ? options.failureMessage()
    : options.failureMessage ?? DEFAULT_FAILURE_MESSAGE
  let timer: ReturnType<typeof setTimeout> | undefined

  function clear() {
    if (timer !== undefined) clearTimeout(timer)
    timer = undefined
    notice.value = ""
  }

  function reportFailure() {
    clear()
    notice.value = resolveFailureMessage()
    timer = setTimeout(() => {
      timer = undefined
      notice.value = ""
    }, durationMs)
  }

  async function copySelection(): Promise<boolean> {
    let copied = false
    try {
      copied = await controller().copySelectionToClipboard()
    } catch {
      // A rejected write is a refusal, not a crash: fall through and report it.
      copied = false
    }
    if (!copied) reportFailure()
    return copied
  }

  if (getCurrentScope()) onScopeDispose(clear)

  return { notice, copySelection, clear }
}
