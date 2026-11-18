// OpenDocument text (`.odt`, `.ott`, `.odm`, `.otm`, `.oth`, `.fodt`) input.
//
// Upstream react-docx opens OpenDocument text through the same `DocModel`
// builder as DOCX and marks the resulting model view-only. The local DOCX WASM
// parser only understands OOXML packages, so this reader performs the same
// model conversion in TypeScript: it accepts the ODF ZIP container or the flat
// XML form, validates the same archive/XML/image budgets as the DOCX path, and
// builds a model with no Word package. The model is display-only because there
// is no ODF writer.

import {
  assertOfficeInputBytes,
  createOfficeImageBudget,
  inspectOfficeArchive,
  validateOfficeXmlText,
} from "@arcships/office-runtime"
import { OfficeLoadError } from "@arcships/office-runtime"
import { strFromU8, unzipSync } from "fflate"
import type { DocModel } from "../engine/types"
import { normalizeDocModel } from "../engine/normalize"
import type { DocxRuntimeLimits } from "../resource-limits"
import type { OoxmlPackage } from "../engine/ooxml-core"
import {
  buildOpenDocumentModelFromDocuments,
  parseOpenDocumentXml,
  type OpenDocumentImageAsset,
  type OpenDocumentReadResult,
} from "./open-document-content"
import {
  attribute,
  base64ToBytes,
  descendantElements,
  firstChildElement,
  parseXmlDocument,
  textContent,
} from "./open-document-xml"

const OPEN_DOCUMENT_MIMETYPE = /application\/vnd\.oasis\.opendocument\./i

/** Shown wherever editing or export is refused for an OpenDocument file. */
export const OPEN_DOCUMENT_VIEW_ONLY_REASON =
  "OpenDocument files are view-only: they cannot be edited or saved as DOCX."

const IMAGE_EXTENSION_TYPES: Readonly<Record<string, string>> = Object.freeze({
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
})

export interface OpenDocumentImportResult {
  package: OoxmlPackage
  model: DocModel
  warnings: string[]
}

export class OpenDocumentImportError extends OfficeLoadError {
  constructor(message: string, cause?: unknown) {
    super({
      code: "INVALID_SOURCE",
      message,
      phase: "parse",
      cause,
    })
    this.name = "OpenDocumentImportError"
  }
}

function bytesOf(input: ArrayBuffer | Uint8Array): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input)
}

function startsWithZip(bytes: Uint8Array): boolean {
  return bytes.length >= 4 &&
    bytes[0] === 0x50 && bytes[1] === 0x4b &&
    bytes[2] === 0x03 && bytes[3] === 0x04
}

function decodeUtf8(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes)
  } catch {
    return strFromU8(bytes)
  }
}

/**
 * Reads the first ZIP local entry when it is small enough to inspect without
 * inflating the archive. ODF requires `mimetype` to be the first entry, stored
 * uncompressed; the byte scan keeps the fallback for the other producers that
 * still put it there with a data descriptor.
 */
function firstZipEntry(bytes: Uint8Array): { name: string; content: Uint8Array } | undefined {
  if (!startsWithZip(bytes) || bytes.length < 30) return undefined
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const flags = view.getUint16(6, true)
  const method = view.getUint16(8, true)
  const compressedBytes = view.getUint32(18, true)
  const nameBytes = view.getUint16(26, true)
  const extraBytes = view.getUint16(28, true)
  const nameStart = 30
  const dataStart = nameStart + nameBytes + extraBytes
  if (nameStart + nameBytes > bytes.length || dataStart > bytes.length) return undefined
  const name = decodeUtf8(bytes.subarray(nameStart, nameStart + nameBytes))
  if (method !== 0 || (flags & 0x8) !== 0 || compressedBytes === 0 || dataStart + compressedBytes > bytes.length) {
    return { name, content: new Uint8Array() }
  }
  return { name, content: bytes.subarray(dataStart, dataStart + compressedBytes) }
}

/** True for an OpenDocument text ZIP or flat XML document. */
export function isOpenDocumentBytes(input: ArrayBuffer | Uint8Array): boolean {
  const bytes = bytesOf(input)
  if (startsWithZip(bytes)) {
    const entry = firstZipEntry(bytes)
    if (entry?.name === "mimetype" && entry.content.length > 0) {
      return decodeUtf8(entry.content).trim().toLowerCase()
        .startsWith("application/vnd.oasis.opendocument.")
    }
    // Some writers place the mandatory mimetype first but use a data
    // descriptor; the first few kilobytes still name the format.
    const head = decodeUtf8(bytes.subarray(0, Math.min(bytes.length, 64 * 1024)))
    if (head.includes("mimetype") && OPEN_DOCUMENT_MIMETYPE.test(head)) return true
    // Nothing in the scanned directory names a mimetype entry: this is an
    // OOXML package or another ZIP, not ODF.
    if (!head.includes("mimetype")) return false
    // The mimetype entry may itself be deflated (non-conforming producers and
    // repacking tools). Inflate only that tiny entry, never document content.
    try {
      const mimetype = unzipSync(bytes, {
        filter: (file) => file.name === "mimetype",
      }).mimetype
      return mimetype !== undefined &&
        decodeUtf8(mimetype).trim().toLowerCase()
          .startsWith("application/vnd.oasis.opendocument.")
    } catch {
      return false
    }
  }

  // Flat `.fodt` files are a single XML document.
  const head = decodeUtf8(bytes.subarray(0, Math.min(bytes.length, 4096)))
  if (!/<office:document\b/i.test(head)) return false
  if (OPEN_DOCUMENT_MIMETYPE.test(head)) return true
  // The `office:mimetype` attribute is mandatory in flat ODF, but the element
  // prefix may differ in a malformed producer. Accept the root plus text body.
  return /<office:body\b/i.test(head) && /<office:text\b/i.test(head)
}

/** True for a model built from OpenDocument text. */
export function isOpenDocumentModel(model: Pick<DocModel, "metadata">): boolean {
  return model.metadata?.sourcePartNames?.document === "content.xml"
}

function assertMimetype(mimetype: string | undefined): void {
  if (!mimetype) {
    throw new OpenDocumentImportError("OpenDocument file does not declare a mimetype.")
  }
  const normalized = mimetype.trim().toLowerCase()
  if (!normalized.startsWith("application/vnd.oasis.opendocument.text")) {
    throw new OpenDocumentImportError(
      "OpenDocument file is not a text document; spreadsheets, presentations and drawings are not text documents and are not supported.",
    )
  }
}

function assertArchiveEntryBudgets(
  entries: Readonly<Record<string, Uint8Array>>,
  limits: Readonly<DocxRuntimeLimits>,
): void {
  let total = 0
  for (const [name, data] of Object.entries(entries)) {
    const single = limits.maxSingleEntryBytes
    if (single !== undefined && data.byteLength > single) {
      throw new OfficeLoadError({
        code: "LIMIT_EXCEEDED",
        message: `OpenDocument entry ${name} exceeds maxSingleEntryBytes.`,
        phase: "archive",
        limit: "maxSingleEntryBytes",
        actual: data.byteLength,
        allowed: single,
      })
    }
    total += data.byteLength
    const allowedTotal = limits.maxUncompressedBytes
    if (allowedTotal !== undefined && total > allowedTotal) {
      throw new OfficeLoadError({
        code: "LIMIT_EXCEEDED",
        message: "OpenDocument archive exceeds maxUncompressedBytes.",
        phase: "archive",
        limit: "maxUncompressedBytes",
        actual: total,
        allowed: allowedTotal,
      })
    }
  }
}

function imageContentType(name: string, explicit?: string): string {
  if (explicit) return explicit
  const extension = /\.([^./\\]+)$/.exec(name)?.[1]?.toLowerCase()
  return extension ? IMAGE_EXTENSION_TYPES[extension] ?? "application/octet-stream" : "application/octet-stream"
}

function validateImageAssets(
  images: Map<string, OpenDocumentImageAsset>,
  limits: Readonly<DocxRuntimeLimits>,
): void {
  const budget = createOfficeImageBudget(limits, "docx")
  try {
    for (const [name, asset] of images) {
      budget.inspectAndReserve(asset.bytes, name)
    }
  } finally {
    budget.dispose()
  }
}

function buildFromZip(
  bytes: Uint8Array,
  limits: Readonly<DocxRuntimeLimits>,
  signal: AbortSignal | undefined,
): OpenDocumentReadResult {
  inspectOfficeArchive(bytes, limits, signal)
  const entries = unzipSync(bytes)
  assertArchiveEntryBudgets(entries, limits)
  let mimetype: string | undefined
  let contentXml: string | undefined
  let stylesXml: string | undefined
  const images = new Map<string, OpenDocumentImageAsset>()

  for (const [name, data] of Object.entries(entries)) {
    if (name === "mimetype") {
      mimetype = decodeUtf8(data).trim()
      continue
    }
    if (name === "content.xml") {
      contentXml = decodeUtf8(data)
      validateOfficeXmlText(name, contentXml, limits, data.byteLength)
      continue
    }
    if (name === "styles.xml") {
      stylesXml = decodeUtf8(data)
      validateOfficeXmlText(name, stylesXml, limits, data.byteLength)
      continue
    }
    if (/\.xml$/i.test(name) || /\.rdf$/i.test(name)) {
      validateOfficeXmlText(name, decodeUtf8(data), limits, data.byteLength)
      continue
    }
    if (/\.(?:png|jpe?g|gif|bmp|svg|webp|tiff?|emf|wmf)$/i.test(name)) {
      images.set(name, {
        bytes: data,
        contentType: imageContentType(name),
      })
    }
  }

  assertMimetype(mimetype)
  if (!contentXml) {
    throw new OpenDocumentImportError("OpenDocument archive does not contain content.xml.")
  }
  validateImageAssets(images, limits)
  return parseOpenDocumentXml(contentXml, stylesXml, images)
}

function flatMimetype(document: Document): string | undefined {
  const root = document.documentElement
  return root.getAttribute("office:mimetype")
    ?? root.getAttributeNS("urn:oasis:names:tc:opendocument:xmlns:office:1.0", "mimetype")
    ?? undefined
}

function buildFromFlatXml(
  xml: string,
  byteLength: number,
  limits: Readonly<DocxRuntimeLimits>,
): OpenDocumentReadResult {
  validateOfficeXmlText("content.xml", xml, limits, byteLength)
  const document = parseXmlDocument(xml)
  assertMimetype(flatMimetype(document))
  // Flat ODF embeds pictures as base64 `office:binary-data`; reserve their
  // bytes/pixels before the model exposes them to a browser decoder.
  const images = new Map<string, OpenDocumentImageAsset>()
  descendantElements(document, "image").forEach((image, index) => {
    const binary = firstChildElement(image, "binary-data")
    if (!binary) return
    const bytes = base64ToBytes(textContent(binary))
    if (!bytes) return
    const href = attribute(image, "href") ?? `binary-data-${index}`
    images.set(href, {
      bytes,
      contentType: imageContentType(href, attribute(image, "mime-type")),
    })
  })
  validateImageAssets(images, limits)
  return buildOpenDocumentModelFromDocuments({ contentDocument: document, images })
}

/**
 * Parses OpenDocument text bytes into a display-only `DocModel` and the empty
 * package DOCX consumers expect for a model without a Word source package.
 *
 * `isOpenDocumentBytes()` is the binary format check; callers should use it
 * before invoking this reader so OOXML bytes keep going through the WASM path.
 */
export async function importOpenDocumentBuffer(
  input: ArrayBuffer | Uint8Array,
  limits: Readonly<DocxRuntimeLimits>,
  signal?: AbortSignal,
): Promise<OpenDocumentImportResult> {
  if (signal?.aborted) {
    throw new OfficeLoadError({
      code: "ABORTED",
      message: "OpenDocument import was aborted.",
      phase: "parse",
    })
  }
  const bytes = bytesOf(input)
  assertOfficeInputBytes(bytes.byteLength, limits)
  let result: OpenDocumentReadResult
  if (startsWithZip(bytes)) {
    result = buildFromZip(bytes, limits, signal)
  } else {
    result = buildFromFlatXml(decodeUtf8(bytes), bytes.byteLength, limits)
  }
  const model = normalizeDocModel(result.model)
  return {
    package: {
      parts: new Map(),
      binaryAssets: new Map(),
    },
    model,
    warnings: result.warnings,
  }
}

export type { OpenDocumentImageAsset } from "./open-document-content"
