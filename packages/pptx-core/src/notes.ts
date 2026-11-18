import { unzipSync } from "fflate"
import { validateOfficeXmlText } from "@arcships/office-runtime"
import { OfficeLoadError } from "@arcships/office-runtime/errors"

import { toOfficeLimits } from "./limits"
import type { PptxPreviewLimits } from "./types"

/** One slide the reader should look up notes for. */
export interface PptxSpeakerNoteSlide {
  /** Zero-based slide index in presentation order. */
  slideIndex: number
  /** Slide part path, such as `ppt/slides/slide3.xml`. */
  slidePath: string
}

/**
 * Reads notes parts out of the container. Injectable so hosts can move the
 * second (notes-only) unzip into a worker and so tests can replace it.
 */
export interface PptxNotesArchiveReader {
  read(buffer: ArrayBuffer, paths: readonly string[]): Promise<ReadonlyMap<string, string>>
}

/**
 * Default notes reader: one central-directory pass that inflates only the
 * notes parts it is asked for, so slide XML and media are never decoded twice.
 */
export const PPTX_NOTES_ARCHIVE_READER: PptxNotesArchiveReader = Object.freeze({
  async read(buffer: ArrayBuffer, paths: readonly string[]): Promise<ReadonlyMap<string, string>> {
    const wanted = new Set(paths)
    const entries = unzipSync(new Uint8Array(buffer), {
      filter: (entry) => wanted.has(entry.name),
    })
    const decoder = new TextDecoder("utf-8")
    const parts = new Map<string, string>()
    for (const [name, bytes] of Object.entries(entries)) {
      parts.set(name, decoder.decode(bytes))
    }
    return parts
  },
})

export interface ReadPptxSpeakerNotesOptions {
  /** Slide relationship XML, keyed by the `.rels` path (the renderer's `PptxFiles.slideRels`). */
  slideRels?: ReadonlyMap<string, string>
  limits: PptxPreviewLimits
  signal?: AbortSignal
  reader?: PptxNotesArchiveReader
}

/** Notes placeholders that never carry speaker copy. */
const NON_CONTENT_PLACEHOLDERS = new Set(["sldimg", "sldnum", "hdr", "ftr", "dt"])

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return
  throw new OfficeLoadError({
    code: "ABORTED",
    message: "PPTX 加载已取消。",
    format: "pptx",
    phase: "notes",
  })
}

/** `ppt/slides/slide1.xml` → `ppt/slides/_rels/slide1.xml.rels`. */
function relationshipsPath(partPath: string): string {
  const slash = partPath.lastIndexOf("/")
  const directory = slash < 0 ? "" : partPath.slice(0, slash + 1)
  const file = slash < 0 ? partPath : partPath.slice(slash + 1)
  return `${directory}_rels/${file}.rels`
}

/**
 * Resolve a relationship target against the part that owns the relationship.
 * Targets outside the package are rejected instead of climbing out of it.
 */
function resolvePartPath(ownerPath: string, target: string): string | undefined {
  const value = target.trim()
  if (!value || /^[a-z][a-z0-9+.-]*:/iu.test(value)) return undefined
  // A target that starts with `/` is already package-absolute; the caller's
  // `ppt/` guard decides whether it may be read.
  if (value.startsWith("/")) return value.slice(1).replace(/^\.\//u, "")
  const segments = ownerPath.split("/").slice(0, -1)
  for (const segment of value.split("/")) {
    if (!segment || segment === ".") continue
    if (segment === "..") {
      if (!segments.length) return undefined
      segments.pop()
      continue
    }
    segments.push(segment)
  }
  return segments.length ? segments.join("/") : undefined
}

function parseXmlDocument(xml: string, init: DOMParser, what: string): Document {
  const document = init.parseFromString(xml, "application/xml")
  const root = document?.documentElement
  if (!root || root.nodeName.toLowerCase() === "parsererror") {
    throw new Error(`PPTX ${what} XML 无法解析。`)
  }
  return document
}

function elements(root: Document | Element, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS("*", localName))
}

function directChildren(root: Element, localName: string): Element[] {
  return Array.from(root.children).filter((child) => child.localName === localName)
}

/** Text of one `a:p`: runs and line breaks in document order. */
function paragraphText(paragraph: Element): string {
  let text = ""
  const visit = (node: Element): void => {
    for (const child of Array.from(node.children)) {
      if (child.localName === "t") text += child.textContent ?? ""
      else if (child.localName === "br") text += "\n"
      else if (child.localName === "r" || child.localName === "fld") visit(child)
    }
  }
  visit(paragraph)
  return text.trim()
}

/**
 * Speaker note text of one notes-slide part.
 *
 * Every text body except the slide-image/slide-number/header/footer/date
 * placeholders is read; one paragraph per line, empty paragraphs dropped.
 */
export function parsePptxNotesXml(xml: string, parser?: DOMParser): string {
  const domParser = parser ?? (typeof DOMParser === "undefined" ? undefined : new DOMParser())
  if (!domParser) throw new Error("当前环境不支持解析 PPTX 演讲者备注。")
  const document = parseXmlDocument(xml, domParser, "演讲者备注")
  const paragraphs: string[] = []
  for (const shape of elements(document, "sp")) {
    const placeholder = elements(shape, "ph")[0]
    const placeholderType = (placeholder?.getAttribute("type") ?? "body").toLowerCase()
    if (NON_CONTENT_PLACEHOLDERS.has(placeholderType)) continue
    for (const body of directChildren(shape, "txBody")) {
      for (const paragraph of directChildren(body, "p")) {
        const text = paragraphText(paragraph)
        if (text) paragraphs.push(text)
      }
    }
  }
  return paragraphs.join("\n")
}

/** Notes part path of one slide, taken from its relationship list. */
export function pptxNotesPartPath(
  slidePath: string,
  slideRels: ReadonlyMap<string, string> | undefined,
  parser?: DOMParser,
): string | undefined {
  const rels = slideRels?.get(relationshipsPath(slidePath))
  if (!rels) return undefined
  const domParser = parser ?? (typeof DOMParser === "undefined" ? undefined : new DOMParser())
  if (!domParser) return undefined
  let document: Document
  try {
    document = parseXmlDocument(rels, domParser, "幻灯片关系")
  } catch {
    return undefined
  }
  for (const relationship of elements(document, "Relationship")) {
    const type = (relationship.getAttribute("Type") ?? "").toLowerCase()
    if (!type.endsWith("/notesslide")) continue
    if ((relationship.getAttribute("TargetMode") ?? "").toLowerCase() === "external") continue
    const path = resolvePartPath(slidePath, relationship.getAttribute("Target") ?? "")
    // Notes always live under `ppt/`; a target that climbs out of it is treated
    // as a broken relationship instead of reading an unrelated part.
    if (path?.startsWith("ppt/")) return path
  }
  return undefined
}

function isLimitOrAbort(error: unknown): boolean {
  if (error instanceof OfficeLoadError) {
    return error.code === "LIMIT_EXCEEDED" || error.code === "ABORTED"
  }
  return typeof DOMException !== "undefined"
    && error instanceof DOMException
    && error.name === "AbortError"
}

/**
 * Speaker notes for every slide that has a notes part.
 *
 * Notes are optional content: a malformed or unparsable notes part leaves that
 * slide without notes instead of failing the whole deck. Budget violations and
 * cancellations still propagate, because they are archive-level decisions the
 * caller must see.
 */
export async function readPptxSpeakerNotes(
  buffer: ArrayBuffer,
  slides: readonly PptxSpeakerNoteSlide[],
  options: ReadPptxSpeakerNotesOptions,
): Promise<ReadonlyMap<number, string>> {
  const notesBySlide = new Map<number, string>()
  throwIfAborted(options.signal)
  const notesPathBySlide = new Map<number, string>()
  for (const slide of slides) {
    const path = pptxNotesPartPath(slide.slidePath, options.slideRels)
    if (path) notesPathBySlide.set(slide.slideIndex, path)
  }
  if (notesPathBySlide.size === 0) return notesBySlide
  const paths = [...new Set(notesPathBySlide.values())]
  const reader = options.reader ?? PPTX_NOTES_ARCHIVE_READER
  let parts: ReadonlyMap<string, string>
  try {
    parts = await reader.read(buffer, paths)
  } catch (error) {
    if (isLimitOrAbort(error)) throw error
    return notesBySlide
  }
  throwIfAborted(options.signal)
  const officeLimits = toOfficeLimits(options.limits)
  const texts = new Map<string, string>()
  for (const path of paths) {
    const xml = parts.get(path)
    if (!xml) continue
    try {
      validateOfficeXmlText(path, xml, officeLimits)
      texts.set(path, parsePptxNotesXml(xml))
    } catch (error) {
      if (isLimitOrAbort(error)) throw error
    }
  }
  for (const [slideIndex, path] of notesPathBySlide) {
    const text = texts.get(path)
    if (text) notesBySlide.set(slideIndex, text)
  }
  return notesBySlide
}
