import { getCurrentScope, onScopeDispose, ref, type Ref } from "vue"

/**
 * Shared drag-and-drop file import for the DOCX, XLSX, PDF and PPTX viewers.
 *
 * Every viewer root owns one drop target: the accept list, the drag-over
 * outline and the rejected-drop feedback live here so the four formats cannot
 * drift apart. A host routes the accepted `File` through the same loader its
 * file picker uses.
 *
 * Drag state is tracked from `dragenter`/`dragleave` nesting depth. `dragover`
 * fires continuously while the pointer stays inside the element, so it must
 * never increment the depth: doing that leaves the outline stuck after a drag
 * leaves (see the DOCX multi-page shell outline bug, #27).
 */

/** Outline drawn on a viewer root while a file is dragged over it. */
export const FILE_DROP_OUTLINE = "3px dashed #2563eb"

/** Accepted file shapes; a match on either list accepts the file. */
export interface FileDropAccept {
  /** Lower-case extensions including the leading dot, e.g. `".xlsx"`. */
  extensions?: readonly string[]
  /** MIME types accepted in addition to the extensions. */
  mimeTypes?: readonly string[]
}

/** The subset of `File` the accept check reads. */
export interface FileDropCandidate {
  name?: string
  type?: string
}

export interface UseFileDropOptions {
  accept: FileDropAccept
  /**
   * Handles the first accepted file of a drop. Hosts pass the same loader the
   * file picker calls.
   */
  onFile: (file: File, event: DragEvent) => void | Promise<void>
  /**
   * Reports a drop that carried files but none matched `accept`, so the host
   * can show an unsupported-type message.
   */
  onReject?: (files: readonly File[], event: DragEvent) => void
  /**
   * Drops are ignored while this is `false`. A getter keeps the check live
   * for viewers that only accept files while editable.
   */
  enabled?: boolean | (() => boolean)
}

export interface UseFileDropReturn {
  /** `true` while a file is dragged over the target; binds the outline. */
  isDragOver: Readonly<Ref<boolean>>
  /** Matches one file against the accept list (extension or MIME type). */
  isSupportedFile: (file: FileDropCandidate) => boolean
  onDragEnter: (event: DragEvent) => void
  onDragOver: (event: DragEvent) => void
  onDragLeave: (event: DragEvent) => void
  onDrop: (event: DragEvent) => void
  /** Clears the outline and the depth counter. */
  reset: () => void
}

/** Lower-case extension including the leading dot, or undefined. */
export function fileDropExtension(fileName: string): string | undefined {
  const match = /(\.[^.\\/]+)$/.exec(fileName.trim())
  return match?.[1]?.toLowerCase()
}

/** Matches a file name extension or MIME type against the accept list. */
export function isSupportedDropFile(
  file: FileDropCandidate,
  accept: FileDropAccept,
): boolean {
  const name = file.name?.trim()
  if (name) {
    const extension = fileDropExtension(name)
    if (extension !== undefined) {
      const extensions = accept.extensions ?? []
      if (extensions.some((entry) => entry.toLowerCase() === extension)) return true
    }
  }
  const type = file.type?.trim().toLowerCase()
  if (type) {
    const mimeTypes = accept.mimeTypes ?? []
    if (mimeTypes.some((entry) => entry.toLowerCase() === type)) return true
  }
  return false
}

function isFileDrag(dataTransfer: DataTransfer | null | undefined): boolean {
  if (!dataTransfer) return false
  const types = dataTransfer.types
  if (types) {
    for (let index = 0; index < types.length; index++) {
      // `types` is a frozen array in current browsers but stays array-like.
      if (types[index] === "Files") return true
    }
  }
  const items = dataTransfer.items
  if (items) {
    for (let index = 0; index < items.length; index++) {
      if (items[index]?.kind === "file") return true
    }
  }
  const files = dataTransfer.files
  return files !== null && files !== undefined && files.length > 0
}

function filesOf(dataTransfer: DataTransfer | null | undefined): File[] {
  const list = dataTransfer?.files
  if (!list) return []
  const files: File[] = []
  for (let index = 0; index < list.length; index++) files.push(list[index])
  return files
}

/** `true` when `relatedTarget` is still inside the drop target. */
function isInsideTarget(event: DragEvent): boolean {
  const related = event.relatedTarget as Node | null | undefined
  const current = event.currentTarget as Node | null | undefined
  if (!related || !current || typeof current.contains !== "function") return false
  return current.contains(related)
}

export function useFileDrop(options: UseFileDropOptions): UseFileDropReturn {
  const isDragOver = ref(false)
  let dragDepth = 0
  let watchingWindow = false

  const isEnabled = (): boolean => (
    typeof options.enabled === "function" ? options.enabled() : options.enabled !== false
  )

  function stopWindowWatch(): void {
    if (!watchingWindow) return
    watchingWindow = false
    if (typeof window === "undefined" || typeof window.removeEventListener !== "function") return
    window.removeEventListener("dragend", reset)
    window.removeEventListener("drop", reset)
  }

  /**
   * A drag can end outside the target — a cancelled drag, a drop on the page
   * background, a tab switch. Watching the window for `dragend`/`drop` keeps
   * the outline from sticking in exactly those cases.
   */
  function startWindowWatch(): void {
    if (watchingWindow) return
    if (typeof window === "undefined" || typeof window.addEventListener !== "function") return
    watchingWindow = true
    window.addEventListener("dragend", reset)
    window.addEventListener("drop", reset)
  }

  function reset(): void {
    dragDepth = 0
    isDragOver.value = false
    stopWindowWatch()
  }

  function markDragOver(event: DragEvent): void {
    event.preventDefault?.()
    if (event.dataTransfer) event.dataTransfer.dropEffect = "copy"
    isDragOver.value = true
    startWindowWatch()
  }

  function onDragEnter(event: DragEvent): void {
    if (!isEnabled() || !isFileDrag(event.dataTransfer)) return
    // Entering a nested element fires `dragenter` again; counting the levels
    // keeps the outline on until the pointer leaves the whole target.
    dragDepth += 1
    markDragOver(event)
  }

  function onDragOver(event: DragEvent): void {
    if (!isEnabled() || !isFileDrag(event.dataTransfer)) return
    // Never touch `dragDepth` here: `dragover` repeats for as long as the
    // pointer stays inside and would grow the counter without bound.
    markDragOver(event)
  }

  function onDragLeave(event: DragEvent): void {
    if (isInsideTarget(event)) return
    dragDepth = Math.max(0, dragDepth - 1)
    if (dragDepth === 0) reset()
  }

  function onDrop(event: DragEvent): void {
    const files = filesOf(event.dataTransfer)
    reset()
    if (!isEnabled() || files.length === 0) return
    event.preventDefault?.()
    const accepted = files.find((file) => isSupportedDropFile(file, options.accept))
    if (accepted) {
      // The loader owns its own error reporting (a viewer renders its load
      // error); a rejected promise here must not surface as an unhandled
      // rejection from a DOM event handler.
      try {
        void Promise.resolve(options.onFile(accepted, event)).catch(() => undefined)
      } catch {
        // A synchronous throw from a host loader is reported by the host too.
      }
      return
    }
    options.onReject?.(files, event)
  }

  if (getCurrentScope()) onScopeDispose(reset)

  return {
    isDragOver,
    isSupportedFile: (file) => isSupportedDropFile(file, options.accept),
    onDragEnter,
    onDragOver,
    onDragLeave,
    onDrop,
    reset,
  }
}
