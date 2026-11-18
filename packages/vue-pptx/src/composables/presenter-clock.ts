import { onScopeDispose, ref, watch, type Ref } from "vue"

/** How often the presenter clock refreshes the elapsed-time readout. */
const PRESENTER_TICK_MS = 500

/** `mm:ss`, or `h:mm:ss` once a talk runs past the hour. */
export function formatPptxPresenterElapsed(elapsedMs: number): string {
  const safe = Number.isFinite(elapsedMs) ? Math.max(0, elapsedMs) : 0
  const totalSeconds = Math.floor(safe / 1_000)
  const seconds = totalSeconds % 60
  const minutes = Math.floor(totalSeconds / 60) % 60
  const hours = Math.floor(totalSeconds / 3_600)
  const pad = (value: number) => String(value).padStart(2, "0")
  return hours > 0 ? `${hours}:${pad(minutes)}:${pad(seconds)}` : `${pad(minutes)}:${pad(seconds)}`
}

/**
 * Elapsed presentation time for the presenter view.
 *
 * The clock starts when `active` turns true, resets to `00:00` when it turns
 * false, and is released with the owning effect scope, so leaving presentation
 * mode never leaves an interval running.
 */
export function usePptxPresenterClock(active: () => boolean): Ref<number> {
  const elapsed = ref(0)
  let startedAt: number | null = null
  let handle: ReturnType<typeof setInterval> | null = null

  const clearTimer = () => {
    if (handle === null) return
    clearInterval(handle)
    handle = null
  }
  const stop = () => {
    clearTimer()
    startedAt = null
    elapsed.value = 0
  }
  const start = () => {
    clearTimer()
    startedAt = Date.now()
    elapsed.value = 0
    if (typeof setInterval !== "function") return
    handle = setInterval(() => {
      if (startedAt !== null) elapsed.value = Date.now() - startedAt
    }, PRESENTER_TICK_MS)
  }

  watch(active, (value) => {
    if (value) start()
    else stop()
  }, { flush: "sync", immediate: true })
  onScopeDispose(stop)

  return elapsed
}
