/**
 * Notify a host once per finished web-font batch so it can re-measure text.
 *
 * Text rendered before a webfont is available keeps the fallback metrics
 * (wrong line breaks and autofit sizes) until the host renders again. The
 * watcher coalesces every trigger of the same batch into one callback and
 * keeps listening, so a later lazily loaded font set refreshes again.
 */
export function watchFontLoading(onLoaded: () => void): () => void {
  const fonts = typeof document === "undefined"
    ? undefined
    : (document as Document & { fonts?: FontFaceSet }).fonts
  if (!fonts) return () => undefined

  let disposed = false
  let scheduled = false
  let finishedBatchSeen = false

  function flush(): void {
    scheduled = false
    if (!disposed) onLoaded()
  }

  function schedule(): void {
    if (disposed || scheduled) return
    scheduled = true
    if (typeof requestAnimationFrame !== "function") {
      queueMicrotask(flush)
      return
    }
    // Two frames let the browser apply the loaded font metrics before re-measuring.
    requestAnimationFrame(() => {
      if (typeof requestAnimationFrame === "function") requestAnimationFrame(flush)
      else queueMicrotask(flush)
    })
  }

  const onFontBatch = () => {
    finishedBatchSeen = true
    schedule()
  }
  fonts.addEventListener("loadingdone", onFontBatch)
  if (fonts.status === "loading" && fonts.ready) {
    // `ready` settles the batch already reported by `loadingdone`; only the
    // initial load needs this fallback for browsers that skipped the event.
    fonts.ready
      .then(() => {
        if (!finishedBatchSeen) schedule()
      })
      .catch(() => undefined)
  }

  return () => {
    disposed = true
    fonts.removeEventListener("loadingdone", onFontBatch)
  }
}
