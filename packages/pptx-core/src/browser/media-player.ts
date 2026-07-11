import type { PptxObjectStateMap } from "../playback/track-compiler"
import type { PptxMediaItem, PptxPlaybackSlide } from "../playback/types"

export interface PptxMediaPlayerCallbacks {
  onBlocked(mediaId: string): void
  onError(mediaId: string, error: unknown): void
  onBookmark(mediaId: string, bookmarkName: string): void
  onStateChange(): void
  resolveSource?(item: PptxMediaItem): Promise<string | undefined>
  releaseSource?(item: PptxMediaItem, url: string): void
}

export interface PptxMediaPlayer {
  readonly blockedMediaIds: readonly string[]
  prepare(slide: PptxPlaybackSlide): void
  apply(state: PptxObjectStateMap): void
  pause(): void
  resume(): Promise<void>
  resumeBlocked(mediaId?: string): Promise<void>
  dispose(): void
}

interface ControlledMedia {
  item: PptxMediaItem
  element: HTMLMediaElement
  previousCommand?: string
  firedBookmarks: Set<string>
  previousTimeMs: number
  sourcePending: boolean
  sourceUrl?: string
  removeListeners(): void
}

function playbackEndMs(entry: ControlledMedia): number | undefined {
  if (entry.item.trimEndMs === undefined || !Number.isFinite(entry.element.duration)) return undefined
  return Math.max(entry.item.trimStartMs ?? 0, entry.element.duration * 1000 - entry.item.trimEndMs)
}

function mediaElementFor(root: HTMLElement, item: PptxMediaItem): HTMLMediaElement | undefined {
  if (item.objectKey) {
    for (const object of root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
      if (object.dataset.pptxObjectKey !== item.objectKey) continue
      if (object.matches("audio,video")) return object as HTMLMediaElement
      const nested = object.querySelector<HTMLMediaElement>(item.kind)
        ?? object.querySelector<HTMLMediaElement>("audio,video")
      if (nested) return nested
    }
  }
  return [...root.querySelectorAll<HTMLMediaElement>(item.kind)][0]
}

function mediaHost(root: HTMLElement, item: PptxMediaItem): HTMLElement | undefined {
  if (!item.objectKey) return undefined
  return [...root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")]
    .find((element) => element.dataset.pptxObjectKey === item.objectKey)
}

function normalizedCommand(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined
  const command = value.trim().toLowerCase().replace(/\s+/gu, "")
  if (!command) return undefined
  if (command.includes("togglepause")) return "toggle-pause"
  if (command.includes("pause")) return "pause"
  if (command.includes("stop")) return "stop"
  if (command.includes("play")) return "play"
  return command
}

export function createPptxMediaPlayer(
  root: HTMLElement,
  callbacks: PptxMediaPlayerCallbacks,
): PptxMediaPlayer {
  const controlled = new Map<string, ControlledMedia>()
  const blocked = new Set<string>()
  const pausedForController = new Set<string>()
  let disposed = false

  const notify = () => callbacks.onStateChange()
  const play = async (entry: ControlledMedia): Promise<void> => {
    if (entry.sourcePending) return
    try {
      await entry.element.play()
      if (blocked.delete(entry.item.id)) notify()
    } catch (reason) {
      if (reason instanceof DOMException && reason.name === "NotAllowedError") {
        blocked.add(entry.item.id)
        callbacks.onBlocked(entry.item.id)
      } else callbacks.onError(entry.item.id, reason)
      notify()
    }
  }
  const stopAll = () => {
    for (const entry of controlled.values()) {
      entry.removeListeners()
      entry.element.pause()
      if (entry.sourceUrl) callbacks.releaseSource?.(entry.item, entry.sourceUrl)
    }
    controlled.clear()
    blocked.clear()
    pausedForController.clear()
  }

  return {
    get blockedMediaIds() {
      return Object.freeze([...blocked])
    },
    prepare(slide) {
      stopAll()
      if (disposed) return
      for (const item of slide.media) {
        let element = mediaElementFor(root, item)
        let sourcePending = false
        if (!element) {
          const host = mediaHost(root, item)
          if (!host || !callbacks.resolveSource) {
            callbacks.onError(item.id, new Error("没有找到媒体元素。"))
            continue
          }
          element = document.createElement(item.kind)
          element.controls = true
          element.preload = "metadata"
          element.style.width = "100%"
          element.style.height = "100%"
          element.style.objectFit = "contain"
          host.replaceChildren(element)
          sourcePending = true
        }
        element.loop = item.loop && item.trimEndMs === undefined
        element.volume = Math.max(0, Math.min(1, item.volume))
        if (item.trimStartMs !== undefined) {
          try { element.currentTime = item.trimStartMs / 1000 } catch { /* 元数据就绪后再由播放命令定位。 */ }
        }
        const entry: ControlledMedia = {
          item,
          element,
          firedBookmarks: new Set(),
          previousTimeMs: element.currentTime * 1000,
          sourcePending,
          removeListeners: () => undefined,
        }
        const onTimeUpdate = () => {
          const currentMs = element.currentTime * 1000
          if (currentMs < entry.previousTimeMs) {
            for (const bookmark of item.bookmarks) {
              if (currentMs < bookmark.timeMs) entry.firedBookmarks.delete(bookmark.name)
            }
          }
          for (const bookmark of item.bookmarks) {
            if (
              !entry.firedBookmarks.has(bookmark.name)
              && entry.previousTimeMs < bookmark.timeMs
              && currentMs >= bookmark.timeMs
            ) {
              entry.firedBookmarks.add(bookmark.name)
              callbacks.onBookmark(item.id, bookmark.name)
            }
          }
          entry.previousTimeMs = currentMs
          const endMs = playbackEndMs(entry)
          if (endMs !== undefined && currentMs >= endMs) {
            if (item.loop) {
              element.currentTime = (item.trimStartMs ?? 0) / 1000
              void play(entry)
            } else element.pause()
          }
        }
        const onLoadedMetadata = () => {
          if (item.trimStartMs !== undefined && element.currentTime * 1000 < item.trimStartMs) {
            element.currentTime = item.trimStartMs / 1000
            entry.previousTimeMs = item.trimStartMs
          }
        }
        element.addEventListener("timeupdate", onTimeUpdate)
        element.addEventListener("loadedmetadata", onLoadedMetadata)
        entry.removeListeners = () => {
          element.removeEventListener("timeupdate", onTimeUpdate)
          element.removeEventListener("loadedmetadata", onLoadedMetadata)
        }
        controlled.set(item.id, entry)
        const resolveSource = callbacks.resolveSource
        if (sourcePending && resolveSource) void resolveSource(item).then((url) => {
          if (disposed || controlled.get(item.id) !== entry) {
            if (url) callbacks.releaseSource?.(item, url)
            return
          }
          entry.sourcePending = false
          if (!url) {
            callbacks.onError(item.id, new Error("媒体资源不存在。"))
            return
          }
          element.src = url
          entry.sourceUrl = url
          element.load()
          if (entry.previousCommand === "play") void play(entry)
        }).catch((reason) => {
          entry.sourcePending = false
          callbacks.onError(item.id, reason)
        })
      }
    },
    apply(state) {
      if (disposed) return
      for (const entry of controlled.values()) {
        const objectKey = entry.item.objectKey
        if (!objectKey) continue
        const values = state[objectKey]
        const volume = values?.["media-volume"]
        if (typeof volume === "number") entry.element.volume = Math.max(0, Math.min(1, volume))
        const time = values?.["media-time"]
        if (typeof time === "number" && Number.isFinite(time)) entry.element.currentTime = Math.max(0, time) / 1000
        const command = normalizedCommand(values?.["media-playback"])
        if (!command || command === entry.previousCommand) continue
        entry.previousCommand = command
        if (command === "play") void play(entry)
        else if (command === "pause") entry.element.pause()
        else if (command === "stop") {
          entry.element.pause()
          entry.element.currentTime = (entry.item.trimStartMs ?? 0) / 1000
        } else if (command === "toggle-pause") {
          if (entry.element.paused) void play(entry)
          else entry.element.pause()
        }
      }
    },
    pause() {
      pausedForController.clear()
      for (const [id, entry] of controlled) {
        if (!entry.element.paused) {
          pausedForController.add(id)
          entry.element.pause()
        }
      }
    },
    async resume() {
      const entries = [...pausedForController].flatMap((id) => controlled.get(id) ?? [])
      pausedForController.clear()
      await Promise.all(entries.map(play))
    },
    async resumeBlocked(mediaId) {
      const entries = mediaId
        ? [controlled.get(mediaId)].filter((entry): entry is ControlledMedia => entry !== undefined)
        : [...blocked].flatMap((id) => controlled.get(id) ?? [])
      await Promise.all(entries.map(play))
    },
    dispose() {
      if (disposed) return
      disposed = true
      stopAll()
    },
  }
}
