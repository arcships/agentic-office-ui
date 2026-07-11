import {
  PptxViewer,
  buildPresentation,
  parseZip,
  parseZipLazyMedia,
  type PptxFiles,
  type SearchHighlightHandle,
  type TextSearchResult,
  type ZipParseLimits,
} from "@aiden0z/pptx-renderer"
import {
  OfficeLoadError,
} from "@arcships/office-runtime/errors"
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
} from "@arcships/office-runtime/load-task"
import {
  loadOfficeSource,
  type OfficeSource,
} from "@arcships/office-runtime/source"
import {
  resolveOfficeLimits,
  type OfficeLimits,
} from "@arcships/office-runtime/limits"

import { DEFAULT_PPTX_PREVIEW_LIMITS } from "./index"
import { applyPptxObjectMarkers } from "./browser/renderer-markers"
import { parsePptxPlaybackDocument } from "./browser/playback-parser"
import { createPptxPlaybackController } from "./browser/controller"
import type {
  PptxDisposableHandle,
  PptxDocumentSession,
  PptxDocumentSessionOptions,
  PptxPreviewOpenOptions,
  PptxPreviewSession,
  PptxPreviewSessionOptions,
  PptxPreviewSource,
} from "./browser-types"
import { PptxPlaybackError } from "./playback/errors"
import type { PptxMediaItem, PptxPlaybackDocument } from "./playback/types"
import {
  PptxPreviewError,
  type PptxPreviewDocument,
  type PptxPreviewErrorCode,
  type PptxPreviewLimits,
  type PptxSearchResult,
} from "./types"

const HARD_LIMITS = DEFAULT_PPTX_PREVIEW_LIMITS

function toOfficeLimits(limits: PptxPreviewLimits | undefined): OfficeLimits {
  return {
    maxInputBytes: limits?.maxInputBytes,
    maxArchiveEntries: limits?.maxArchiveEntries,
    maxSingleEntryBytes: limits?.maxSingleEntryBytes,
    maxUncompressedBytes: limits?.maxUncompressedBytes,
    maxTotalImageBytes: limits?.maxMediaBytes,
  }
}

function resolveLimits(requested: PptxPreviewLimits | undefined): Readonly<PptxPreviewLimits> {
  const resolved = resolveOfficeLimits(
    toOfficeLimits(DEFAULT_PPTX_PREVIEW_LIMITS),
    toOfficeLimits(requested),
    toOfficeLimits(HARD_LIMITS),
  )
  const concurrency = requested?.maxConcurrency
  return Object.freeze({
    maxInputBytes: resolved.maxInputBytes,
    maxArchiveEntries: resolved.maxArchiveEntries,
    maxSingleEntryBytes: resolved.maxSingleEntryBytes,
    maxUncompressedBytes: resolved.maxUncompressedBytes,
    maxMediaBytes: resolved.maxTotalImageBytes,
    maxConcurrency:
      typeof concurrency === "number" && Number.isFinite(concurrency) && concurrency > 0
        ? Math.min(concurrency, HARD_LIMITS.maxConcurrency)
        : HARD_LIMITS.maxConcurrency,
  })
}

function tightenLimits(
  base: Readonly<PptxPreviewLimits>,
  requested: PptxPreviewLimits | undefined,
): Readonly<PptxPreviewLimits> {
  if (!requested) return base
  const take = (name: keyof PptxPreviewLimits): number | undefined => {
    const fallback = base[name]
    const candidate = requested[name]
    if (typeof candidate !== "number" || !Number.isFinite(candidate) || candidate <= 0) return fallback
    return fallback === undefined ? candidate : Math.min(candidate, fallback)
  }
  return Object.freeze({
    maxInputBytes: take("maxInputBytes"),
    maxArchiveEntries: take("maxArchiveEntries"),
    maxSingleEntryBytes: take("maxSingleEntryBytes"),
    maxUncompressedBytes: take("maxUncompressedBytes"),
    maxMediaBytes: take("maxMediaBytes"),
    maxConcurrency: take("maxConcurrency"),
  })
}

function toZipLimits(limits: Readonly<PptxPreviewLimits>): ZipParseLimits {
  return {
    maxEntries: limits.maxArchiveEntries,
    maxEntryUncompressedBytes: limits.maxSingleEntryBytes,
    maxTotalUncompressedBytes: limits.maxUncompressedBytes,
    maxMediaBytes: limits.maxMediaBytes,
    maxConcurrency: limits.maxConcurrency,
  }
}

function toOfficeSource(source: PptxPreviewSource): OfficeSource {
  if (source instanceof ArrayBuffer) return { kind: "bytes", bytes: source }
  if (source instanceof Uint8Array) {
    const bytes = new Uint8Array(source.byteLength)
    bytes.set(source)
    return { kind: "bytes", bytes: bytes.buffer }
  }
  if (source && typeof source === "object" && typeof source.arrayBuffer === "function") {
    return { kind: "file", file: source, name: source.name }
  }
  throw new PptxPreviewError("INVALID_SOURCE", "PPTX 来源无效。")
}

function errorCode(error: unknown): PptxPreviewErrorCode {
  if (error instanceof OfficeLoadError) {
    if (error.code === "ABORTED") return "ABORTED"
    if (error.code === "STALE_RESULT") return "STALE_RESULT"
    if (error.code === "LIMIT_EXCEEDED") return "LIMIT_EXCEEDED"
    if (error.code === "INVALID_SOURCE") return "INVALID_SOURCE"
  }
  if (typeof DOMException !== "undefined" && error instanceof DOMException && error.name === "AbortError") {
    return "ABORTED"
  }
  const message = error instanceof Error ? error.message : String(error)
  if (/limit|exceed|too large|zip bomb|archive contains/i.test(message)) return "LIMIT_EXCEEDED"
  return "PARSE_FAILED"
}

function toPreviewError(error: unknown): PptxPreviewError {
  if (error instanceof PptxPreviewError) return error
  const code = errorCode(error)
  const fallback = code === "ABORTED" || code === "STALE_RESULT"
    ? "PPTX 加载已取消。"
    : code === "LIMIT_EXCEEDED"
      ? "PPTX 超过资源限制。"
      : "无法解析 PPTX 文件。"
  const message = error instanceof OfficeLoadError && error.message ? error.message : fallback
  return new PptxPreviewError(code, message, error)
}

function normalizeSearchResult(result: TextSearchResult): PptxSearchResult {
  return {
    slideIndex: result.slideIndex,
    nodeId: result.nodeId,
    text: result.text,
    snippet: result.snippet,
    matchStart: result.matchStart,
    matchEnd: result.matchEnd,
  }
}

export function createPptxDocumentSession(
  container: HTMLElement,
  options: PptxDocumentSessionOptions = {},
): PptxDocumentSession {
  const sequence = createOfficeTaskSequence(
    `pptx-${globalThis.crypto?.randomUUID?.() ?? Date.now()}`,
  )
  const coordinator = createLatestTaskCoordinator(sequence)
  const baseLimits = resolveLimits(options.limits)
  const viewer = new PptxViewer(container, {
    fitMode: options.fitMode ?? "contain",
    zoomPercent: options.zoomPercent ?? 100,
    zipLimits: toZipLimits(baseLimits),
    lazyMedia: options.lazyMedia ?? true,
    lazySlides: options.lazySlides ?? true,
    pdfjs: false,
    onSlideRendered(index, element) {
      const slide = viewer.presentationData?.slides[index]
      if (slide) applyPptxObjectMarkers(element, slide.slidePath)
    },
  })
  let currentDocument: PptxPreviewDocument | null = null
  let currentArchive: PptxFiles | null = null
  let currentPlaybackDocument: PptxPlaybackDocument | null = null
  let disposed = false
  const playbackControllers = new Set<ReturnType<typeof createPptxPlaybackController>>()
  const searchCache = new Map<string, TextSearchResult>()
  const mediaUrls = new Map<string, string>()

  const releaseMediaUrls = () => {
    for (const url of mediaUrls.values()) URL.revokeObjectURL(url)
    mediaUrls.clear()
  }

  const resolveMediaSource = async (item: PptxMediaItem): Promise<string | undefined> => {
    const archive = currentArchive
    const path = item.sourcePath
    if (!archive || !path) return undefined
    if (!item.embedded) {
      return options.externalMedia === "allowed" && /^https?:\/\//iu.test(path) ? path : undefined
    }
    const cached = mediaUrls.get(path)
    if (cached) return cached
    let data = archive.media.get(path)
    if (!data && archive.mediaResolver) {
      const resolved = await archive.mediaResolver.resolve(path)
        ?? await archive.mediaResolver.resolve(`../media/${path.split("/").pop() ?? ""}`)
      data = resolved?.data
    }
    if (!data || currentArchive !== archive) return undefined
    const bytes = new Uint8Array(data.byteLength)
    bytes.set(data)
    const url = URL.createObjectURL(new Blob([bytes.buffer], { type: item.contentType ?? "application/octet-stream" }))
    mediaUrls.set(path, url)
    return url
  }
  const releaseMediaSource = (item: PptxMediaItem, url: string) => {
    if (!item.embedded || !item.sourcePath || mediaUrls.get(item.sourcePath) !== url) return
    URL.revokeObjectURL(url)
    mediaUrls.delete(item.sourcePath)
  }

  const searchKey = (result: Pick<TextSearchResult, "slideIndex" | "nodeId" | "matchStart">) =>
    `${result.slideIndex}:${result.nodeId}:${result.matchStart}`

  const assertUsable = () => {
    if (disposed) throw new PptxPreviewError("ABORTED", "PPTX 预览会话已经销毁。")
  }

  const session: PptxDocumentSession = {
    get document() {
      return currentDocument
    },
    get currentSlideIndex() {
      return viewer.currentSlideIndex
    },
    get zoomPercent() {
      return viewer.zoomPercent
    },
    get playbackDocument() {
      return currentPlaybackDocument
    },
    get capabilityReport() {
      return currentPlaybackDocument?.capability ?? null
    },
    async open(source, openOptions = {}) {
      assertUsable()
      for (const controller of playbackControllers) controller.dispose()
      playbackControllers.clear()
      releaseMediaUrls()
      currentDocument = null
      currentArchive = null
      currentPlaybackDocument = null
      const officeSource = toOfficeSource(source)
      const limits = tightenLimits(baseLimits, openOptions.limits)
      const task = coordinator.start(officeSource, {
        signal: openOptions.signal,
        limits: toOfficeLimits(limits),
      })
      try {
        const resolved = await loadOfficeSource(officeSource, {
          signal: task.signal,
          limits: task.context.limits,
        })
        task.assertCurrent()
        const lazyMedia = openOptions.lazyMedia ?? options.lazyMedia ?? true
        const files = lazyMedia
          ? await parseZipLazyMedia(resolved.buffer, toZipLimits(limits))
          : await parseZip(resolved.buffer, toZipLimits(limits))
        task.assertCurrent()
        const presentation = buildPresentation(files, {
          lazySlides: openOptions.lazySlides ?? options.lazySlides ?? true,
        })
        task.assertCurrent()
        const playbackDocument = parsePptxPlaybackDocument(files, presentation, {
          approximation: options.approximation ?? "off",
          externalMedia: options.externalMedia ?? "disabled",
        })
        task.assertCurrent()
        const initialSlide = Math.min(
          Math.max(Math.trunc(openOptions.initialSlide ?? 0), 0),
          Math.max(presentation.slides.length - 1, 0),
        )
        viewer.load(presentation)
        await viewer.renderSlide(initialSlide)
        task.assertCurrent()

        currentArchive = files
        currentPlaybackDocument = playbackDocument

        currentDocument = Object.freeze({
          fileName: resolved.fileName,
          width: presentation.width,
          height: presentation.height,
          slides: Object.freeze(presentation.slides.map((slide, index) => Object.freeze({
            index,
            number: index + 1,
            hidden: playbackDocument.slides[index]?.hidden ?? slide.hidden === true,
          }))),
          warnings: Object.freeze([Object.freeze({
            code: "NOTES_UNSUPPORTED" as const,
            message: "当前静态底座尚未解析演讲者备注。",
          })]),
        })
        return currentDocument
      } catch (error) {
        throw toPreviewError(error)
      } finally {
        task.finish()
      }
    },
    async renderSlide(index) {
      assertUsable()
      if (!currentDocument) throw new PptxPreviewError("RENDER_FAILED", "请先加载 PPTX 文件。")
      if (!Number.isInteger(index) || index < 0 || index >= currentDocument.slides.length) {
        throw new PptxPreviewError("RENDER_FAILED", "页面编号超出范围。")
      }
      try {
        await viewer.goToSlide(index, { behavior: "instant", block: "center" })
      } catch (error) {
        throw new PptxPreviewError("RENDER_FAILED", "无法渲染指定页面。", error)
      }
    },
    renderThumbnail(index, target, width = 148) {
      assertUsable()
      return viewer.renderThumbnailToContainer(index, target, { width })
    },
    searchText(query) {
      assertUsable()
      const value = query.trim()
      searchCache.clear()
      if (!value) return []
      return viewer.searchText(value).map((result) => {
        searchCache.set(searchKey(result), result)
        return normalizeSearchResult(result)
      })
    },
    async highlightSearchResult(result) {
      assertUsable()
      const sourceResult = searchCache.get(searchKey(result))
      if (!sourceResult) return null
      const handle: SearchHighlightHandle | null = await viewer.highlightSearchResult(sourceResult)
      return handle
    },
    clearSearchHighlights() {
      viewer.clearSearchHighlights()
    },
    async setZoom(percent) {
      assertUsable()
      await viewer.setZoom(Math.min(400, Math.max(10, Math.round(percent))))
    },
    cancel() {
      coordinator.cancel()
    },
    createPlaybackController(controllerOptions = {}) {
      assertUsable()
      if (!currentArchive || !currentPlaybackDocument) {
        throw new PptxPlaybackError(
          "PLAYBACK_NOT_READY",
          "PPTX 播放模型尚未准备好。",
        )
      }
      let controller: ReturnType<typeof createPptxPlaybackController>
      controller = createPptxPlaybackController({
        root: container,
        document: currentPlaybackDocument,
        initialSlideIndex: viewer.currentSlideIndex,
        renderSlide: (index) => session.renderSlide(index),
        resolveMediaSource,
        releaseMediaSource,
        onDispose: () => playbackControllers.delete(controller),
      }, controllerOptions)
      playbackControllers.add(controller)
      return controller
    },
    dispose() {
      if (disposed) return
      disposed = true
      currentDocument = null
      currentArchive = null
      currentPlaybackDocument = null
      for (const controller of playbackControllers) controller.dispose()
      playbackControllers.clear()
      searchCache.clear()
      releaseMediaUrls()
      coordinator.dispose()
      sequence.dispose()
      viewer.destroy()
    },
  }
  return session
}

export function createPptxPreviewSession(
  container: HTMLElement,
  options: PptxPreviewSessionOptions = {},
): PptxPreviewSession {
  return createPptxDocumentSession(container, options)
}

export type {
  PptxDisposableHandle,
  PptxActionRequest,
  PptxApproximationPolicy,
  PptxDocumentSession,
  PptxDocumentSessionFactory,
  PptxDocumentSessionOptions,
  PptxPreviewOpenOptions,
  PptxPreviewSession,
  PptxPreviewSessionOptions,
  PptxPreviewSource,
  PptxPreviewSessionFactory,
  PptxPlaybackController,
  PptxPlaybackControllerOptions,
  PptxPlaybackEvent,
  PptxPlaybackSnapshot,
  PptxPlaybackStatus,
  PptxFileLike,
} from "./browser-types"
export type {
  PptxPreviewDocument,
  PptxPreviewLimits,
  PptxSearchResult,
} from "./types"
