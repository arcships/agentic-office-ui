import { pptxTransitionSoundMediaItem } from "../playback/transitions"
import type { PptxMediaItem, PptxPlaybackSlide } from "../playback/types"

export interface PptxTransitionSoundPlayerCallbacks {
  onBlocked(mediaId: string): void
  onError(mediaId: string, error: unknown): void
  onStateChange(): void
  resolveSource?(item: PptxMediaItem): Promise<string | undefined>
  releaseSource?(item: PptxMediaItem, url: string): void
}

export interface PptxTransitionSoundPlayer {
  readonly blockedMediaIds: readonly string[]
  prepare(slide: PptxPlaybackSlide): void
  /** Stops the prepared sound without preparing another slide. */
  stop(): void
  /** Attempts playback; autoplay rejection is reported, never thrown. */
  play(): Promise<void>
  pause(): void
  resume(): Promise<void>
  resumeBlocked(mediaId?: string): Promise<void>
  dispose(): void
}

interface ControlledSound {
  item: PptxMediaItem
  element: HTMLAudioElement
  sourceUrl?: string
  sourcePending: boolean
  pausedForController: boolean
  failed: boolean
  ready: Promise<void>
}

function createAudio(item: PptxMediaItem): HTMLAudioElement {
  const element = document.createElement("audio")
  element.preload = "auto"
  element.loop = item.loop
  element.volume = Math.max(0, Math.min(1, item.volume))
  return element
}

/**
 * Slide transition sounds. They share the media resolver and the same autoplay
 * gate as audio shapes: a browser that rejects `play()` surfaces the sound id
 * through `blockedMediaIds`, and `resumeBlockedMedia()` retries it after a user
 * gesture instead of silently dropping the sound.
 */
export function createPptxTransitionSoundPlayer(
  callbacks: PptxTransitionSoundPlayerCallbacks,
): PptxTransitionSoundPlayer {
  let current: ControlledSound | null = null
  const blocked = new Set<string>()
  let disposed = false

  const notify = () => callbacks.onStateChange()

  const stop = () => {
    if (!current) return
    current.element.pause()
    if (current.sourceUrl) {
      // Do not revoke the blob while the element is still loading it: Chromium
      // reports ERR_FILE_NOT_FOUND for the in-flight media request.
      const item = current.item
      const url = current.sourceUrl
      setTimeout(() => callbacks.releaseSource?.(item, url), 1000)
    }
    current = null
  }

  const fail = (item: PptxMediaItem, reason: unknown) => {
    callbacks.onError(item.id, reason)
    notify()
  }

  const attempt = async (entry: ControlledSound): Promise<void> => {
    if (disposed || entry !== current || entry.failed) return
    await entry.ready
    if (disposed || entry !== current || entry.failed || entry.sourcePending) return
    entry.pausedForController = false
    try {
      await entry.element.play()
      if (blocked.delete(entry.item.id)) notify()
    } catch (reason) {
      if (reason instanceof DOMException && reason.name === "NotAllowedError") {
        blocked.add(entry.item.id)
        callbacks.onBlocked(entry.item.id)
      } else fail(entry.item, reason)
      notify()
    }
  }

  return {
    get blockedMediaIds() {
      return Object.freeze([...blocked])
    },
    stop,
    prepare(slide) {
      stop()
      blocked.clear()
      if (disposed) return
      const sound = slide.transition?.sound
      if (!sound) return
      const item = pptxTransitionSoundMediaItem(slide.index, sound)
      let releaseReady: () => void = () => undefined
      const ready = new Promise<void>((resolve) => {
        releaseReady = resolve
      })
      const entry: ControlledSound = {
        item,
        element: createAudio(item),
        sourcePending: true,
        pausedForController: false,
        failed: false,
        ready,
      }
      current = entry
      const resolveSource = callbacks.resolveSource
      if (!resolveSource) {
        entry.sourcePending = false
        entry.failed = true
        fail(item, new Error("没有媒体解析器。"))
        releaseReady()
        return
      }
      void resolveSource(item).then((url) => {
        if (disposed || current !== entry) {
          if (url) callbacks.releaseSource?.(item, url)
          releaseReady()
          return
        }
        entry.sourcePending = false
        if (!url) {
          entry.failed = true
          fail(item, new Error("切换声音资源不存在。"))
          releaseReady()
          return
        }
        entry.sourceUrl = url
        entry.element.src = url
        entry.element.load()
        releaseReady()
      }).catch((reason) => {
        if (disposed || current !== entry) {
          releaseReady()
          return
        }
        entry.sourcePending = false
        entry.failed = true
        fail(item, reason)
        releaseReady()
      })
    },
    async play() {
      const entry = current
      if (!entry || disposed) return
      await entry.ready
      if (disposed || current !== entry) return
      await attempt(entry)
    },
    pause() {
      if (!current || current.element.paused) return
      current.pausedForController = true
      current.element.pause()
    },
    async resume() {
      if (!current?.pausedForController) return
      await attempt(current)
    },
    async resumeBlocked(mediaId) {
      const entry = current
      if (!entry || mediaId && mediaId !== entry.item.id) return
      if (!blocked.has(entry.item.id)) return
      await attempt(entry)
    },
    dispose() {
      if (disposed) return
      disposed = true
      stop()
      blocked.clear()
    },
  }
}
