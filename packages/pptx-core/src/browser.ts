import {
  PptxViewer,
  buildPresentation,
  type PptxFiles,
  type SearchHighlightHandle,
  type TextSearchResult,
} from "@aiden0z/pptx-renderer"
import {
  OfficeLoadError,
} from "@arcships/office-runtime/errors"
import {
  isMacroEnabledOfficeFormat,
} from "@arcships/office-runtime"
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
import { loadPptxArchive, toOfficeLimits, type PptxArchiveZipLimits } from "./archive"
import { resolvePptxExternalMediaUrl } from "./media-policy"
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

const PREVIEW_LIMIT_KEYS = [
  "maxInputBytes",
  "maxArchiveEntries",
  "maxSingleEntryBytes",
  "maxUncompressedBytes",
  "maxMediaBytes",
  "maxConcurrency",
  "maxCompressionRatio",
  "maxArchivePathLength",
  "maxXmlBytes",
  "maxSingleXmlBytes",
  "maxXmlDepth",
  "maxXmlAttributeBytes",
  "maxTextNodeBytes",
  "maxRelationships",
] as const

function resolveConcurrency(requested: number | undefined): number {
  return typeof requested === "number" && Number.isFinite(requested) && requested > 0
    ? Math.min(requested, HARD_LIMITS.maxConcurrency)
    : HARD_LIMITS.maxConcurrency
}

function resolveLimits(requested: PptxPreviewLimits | undefined): Readonly<PptxPreviewLimits> {
  // resolveOfficeLimits only walks the keys present in `defaults`, which is now
  // the full documented PPTX subset instead of five archive fields.
  const resolved = resolveOfficeLimits(
    toOfficeLimits(DEFAULT_PPTX_PREVIEW_LIMITS),
    toOfficeLimits(requested),
    toOfficeLimits(HARD_LIMITS),
  )
  return Object.freeze({
    maxInputBytes: resolved.maxInputBytes,
    maxArchiveEntries: resolved.maxArchiveEntries,
    maxSingleEntryBytes: resolved.maxSingleEntryBytes,
    maxUncompressedBytes: resolved.maxUncompressedBytes,
    maxMediaBytes: resolved.maxTotalImageBytes,
    maxCompressionRatio: resolved.maxCompressionRatio,
    maxArchivePathLength: resolved.maxArchivePathLength,
    maxXmlBytes: resolved.maxXmlBytes,
    maxSingleXmlBytes: resolved.maxSingleXmlBytes,
    maxXmlDepth: resolved.maxXmlDepth,
    maxXmlAttributeBytes: resolved.maxXmlAttributeBytes,
    maxTextNodeBytes: resolved.maxTextNodeBytes,
    maxRelationships: resolved.maxRelationships,
    maxConcurrency: resolveConcurrency(requested?.maxConcurrency),
  })
}

function tightenLimits(
  base: Readonly<PptxPreviewLimits>,
  requested: PptxPreviewLimits | undefined,
): Readonly<PptxPreviewLimits> {
  if (!requested) return base
  const tightened: Record<string, number | undefined> = {}
  for (const name of PREVIEW_LIMIT_KEYS) {
    const fallback = base[name]
    const candidate = requested[name]
    tightened[name] = typeof candidate !== "number" || !Number.isFinite(candidate) || candidate <= 0
      ? fallback
      : fallback === undefined ? candidate : Math.min(candidate, fallback)
  }
  return Object.freeze(tightened) as Readonly<PptxPreviewLimits>
}

function sameLimits(
  left: Readonly<PptxPreviewLimits>,
  right: Readonly<PptxPreviewLimits>,
): boolean {
  return PREVIEW_LIMIT_KEYS.every((name) => left[name] === right[name])
}

function toZipLimits(limits: Readonly<PptxPreviewLimits>): PptxArchiveZipLimits {
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
    // The runtime no longer copies byte sources, so hand over the caller's own
    // buffer whenever the view already spans it; a subarray needs one copy.
    const buffer = source.buffer
    if (buffer instanceof ArrayBuffer && source.byteOffset === 0 && source.byteLength === buffer.byteLength) {
      return { kind: "bytes", bytes: buffer }
    }
    const copy = new Uint8Array(source.byteLength)
    copy.set(source)
    return { kind: "bytes", bytes: copy.buffer }
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
  const externalMediaAllowed = options.externalMedia === "allowed"
  const externalMediaPolicy = options.externalMediaUrlPolicy
  const createViewer = (
    limits: Readonly<PptxPreviewLimits>,
    zoomPercent = options.zoomPercent ?? 100,
  ): PptxViewer => new PptxViewer(container, {
    fitMode: options.fitMode ?? "contain",
    zoomPercent,
    scrollContainer: options.scrollContainer,
    zipLimits: toZipLimits(limits),
    lazyMedia: options.lazyMedia ?? true,
    lazySlides: options.lazySlides ?? true,
    pdfjs: false,
    onSlideChange(index) {
      options.onSlideChange?.(index)
    },
    onSlideRendered(index, element) {
      const slide = viewer.presentationData?.slides[index]
      if (slide) applyPptxObjectMarkers(element, slide.slidePath)
    },
  })
  let viewer = createViewer(baseLimits)
  let viewerLimits = baseLimits
  let currentLimits = baseLimits
  let currentDocument: PptxPreviewDocument | null = null
  let currentArchive: PptxFiles | null = null
  let currentPlaybackDocument: PptxPlaybackDocument | null = null
  let disposed = false
  const renderMode = options.renderMode ?? "slide"
  const listOptions = {
    windowed: options.listOptions?.windowed ?? true,
    batchSize: options.listOptions?.batchSize,
    initialSlides: options.listOptions?.initialSlides,
    overscanViewport: options.listOptions?.overscanViewport,
    showSlideLabels: options.listOptions?.showSlideLabels,
  }
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
      if (!externalMediaAllowed) return undefined
      // External targets go through the shared URL policy: opting in is not
      // enough, the host must also allow the resolved URL.
      return resolvePptxExternalMediaUrl(path, externalMediaPolicy)
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
    // Blob copies the bytes it keeps, so hand it the decoded view directly.
    // The renderer hands out ArrayBuffer-backed views; the cast only narrows the
    // lib.dom BlobPart type, it does not copy anything.
    const part = data as unknown as BlobPart
    const url = URL.createObjectURL(new Blob([part], { type: item.contentType ?? "application/octet-stream" }))
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
    get limits() {
      return currentLimits
    },
    get playbackDocument() {
      return currentPlaybackDocument
    },
    get capabilityReport() {
      return currentPlaybackDocument?.capability ?? null
    },
    async open(source, openOptions = {}) {
      assertUsable()
      const officeSource = toOfficeSource(source)
      const limits = tightenLimits(baseLimits, openOptions.limits)
      const task = coordinator.start(officeSource, {
        signal: openOptions.signal,
        limits: toOfficeLimits(limits),
      })
      let installing = false
      try {
        const resolved = await loadOfficeSource(officeSource, {
          signal: task.signal,
          limits: task.context.limits,
        })
        task.assertCurrent()
        const lazyMedia = openOptions.lazyMedia ?? options.lazyMedia ?? true
        // One decompression pass: the renderer reads the container and the shared
        // validator checks the decoded parts, instead of inflating twice.
        const archive = await loadPptxArchive(resolved.buffer, {
          limits,
          zipLimits: toZipLimits(limits),
          lazyMedia,
          signal: task.signal,
        })
        task.assertCurrent()
        const files = archive.files as PptxFiles
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
        // Replace the current document only after the new deck has parsed: a
        // failed parse keeps the previous deck fully usable. A failed render is
        // handled below by clearing both the viewport and the session state.
        for (const controller of playbackControllers) controller.dispose()
        playbackControllers.clear()
        releaseMediaUrls()
        searchCache.clear()
        installing = true
        currentDocument = null
        currentArchive = null
        currentPlaybackDocument = null
        // The viewer owns lazy slide/media parsing. Rebuild it when the open
        // request tightens limits so those parses cannot fall back to base
        // limits. Rebuild only after the new deck parsed: a failed parse keeps
        // the previous deck and its limits fully usable.
        if (!sameLimits(limits, viewerLimits)) {
          const zoomPercent = viewer.zoomPercent
          viewer.destroy()
          viewer = createViewer(limits, zoomPercent)
          viewerLimits = limits
        }
        currentLimits = limits
        viewer.load(presentation)
        if (renderMode === "list") {
          await viewer.renderList(listOptions)
          if (initialSlide > 0) {
            await viewer.goToSlide(initialSlide, { behavior: "instant", block: "start" })
          }
        } else {
          await viewer.renderSlide(initialSlide)
        }
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
          warnings: Object.freeze([
            Object.freeze({
              code: "NOTES_UNSUPPORTED" as const,
              message: "当前静态底座尚未解析演讲者备注。",
            }),
            ...(isMacroEnabledOfficeFormat(archive.validation.sourceFormat)
              ? [Object.freeze({
                  code: "UNSUPPORTED_CONTENT" as const,
                  message: "文件中的 VBA 宏已忽略且不会执行。",
                })]
              : []),
          ]),
        })
        return currentDocument
      } catch (error) {
        if (installing && task.isCurrent()) {
          viewer.destroy()
          currentDocument = null
          currentArchive = null
          currentPlaybackDocument = null
          searchCache.clear()
          releaseMediaUrls()
          for (const controller of playbackControllers) controller.dispose()
          playbackControllers.clear()
        }
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
    async refresh() {
      assertUsable()
      if (!currentDocument) return
      const slideIndex = viewer.currentSlideIndex
      try {
        if (renderMode === "list") {
          await viewer.renderList(listOptions)
          await viewer.goToSlide(slideIndex, { behavior: "instant", block: "center" })
        } else {
          await viewer.renderSlide(slideIndex)
        }
      } catch (error) {
        throw new PptxPreviewError("RENDER_FAILED", "无法重新渲染 PPTX。", error)
      }
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
  return createPptxDocumentSession(container, {
    ...options,
    renderMode: options.renderMode ?? "list",
  })
}

export { createPptxPlaybackController } from "./browser/controller"
export type { PptxPlaybackControllerHost } from "./browser/controller"
export { loadPptxArchive, PPTX_ARCHIVE_READER } from "./archive"
export type {
  LoadedPptxArchive,
  LoadPptxArchiveOptions,
  PptxArchiveReader,
  PptxArchiveSourceFormat,
  PptxArchiveSummary,
  PptxArchiveValidation,
  PptxArchiveZipLimits,
} from "./archive"
export { resolvePptxExternalMediaUrl } from "./media-policy"

export type {
  PptxDisposableHandle,
  PptxActionRequest,
  PptxApproximationPolicy,
  PptxDocumentSession,
  PptxDocumentSessionFactory,
  PptxDocumentSessionOptions,
  PptxExternalMediaUrlPolicy,
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
