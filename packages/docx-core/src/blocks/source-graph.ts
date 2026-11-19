// Document source graph — the inspectable structure of the package a DOCX
// import actually read.
//
// Upstream: @extend-ai/react-docx `doc-model/src/blocks.ts` and
// `rust-docx/src/source_graph.rs`. Upstream parses the package in Rust/WASM;
// this port keeps the same public contract in TypeScript so hosts can inspect
// any package even when the bundled `docx_wasm_bg.wasm` predates
// `inspect_document_source`.

import { strFromU8, unzipSync } from "fflate"
import type { OoxmlPackage } from "../engine/ooxml-core"

export interface SourceAttribute {
  namespace: string | null
  name: string
  value: string
}

export interface SourceElement {
  id: string
  kind: "document" | "element" | "text" | "comment" | "processing_instruction"
  namespace?: string
  name?: string
  text?: string
  byteStart: number
  byteEnd: number
  attributes: SourceAttribute[]
  children: SourceElement[]
}

export interface DocumentSourcePart {
  name: string
  byteLength: number
  root: SourceElement | null
  error: string | null
}

export interface DocumentSourceAsset {
  name: string
  byteLength: number
  root: null
  error: null
}

export interface PreservedFeature {
  partName: string
  feature: string
  occurrences: number
  support: "preserved_only" | "approximate_rendering"
}

export interface DocumentSourceGraph {
  schemaVersion: 1
  documentId: string
  parts: DocumentSourcePart[]
  assets: DocumentSourceAsset[]
  warnings: string[]
  preservedFeatures: PreservedFeature[]
}

export interface SourceGraphPackage {
  parts: ReadonlyMap<string, { name: string; content: string }>
  binaryAssets: ReadonlyMap<string, Uint8Array>
  warnings?: readonly string[]
}

export async function inspectDocumentSource(
  bytes: ArrayBuffer | Uint8Array,
): Promise<DocumentSourceGraph> {
  return buildDocumentSourceGraph(readOoxmlPackage(bytes))
}

/**
 * Build the graph from an already decoded package. The import Worker uses this
 * after parsing so a configured `onSourceGraph` callback costs no second ZIP
 * decode.
 */
export function buildDocumentSourceGraph(pkg: SourceGraphPackage): DocumentSourceGraph {
  const documentId = packageIdentity(pkg)
  const parsed = new Map<string, ParsedXml>()
  const parts = [...pkg.parts.entries()]
    .map(([key, part]) => {
      const name = part.name || key
      const content = part.content
      const result = parseXml(content)
      if (result.root) assignSourceIds(result.root, `${documentId}/${name}/0`)
      parsed.set(name, result)
      return {
        name,
        byteLength: utf8ByteLength(content),
        root: result.root ?? null,
        error: result.error ?? null,
      }
    })
    .sort((left, right) => compareNames(left.name, right.name))

  const assets = [...pkg.binaryAssets.entries()]
    .map(([key, data]) => ({
      name: key,
      byteLength: data.byteLength,
      root: null,
      error: null,
    }))
    .sort((left, right) => compareNames(left.name, right.name))

  return {
    schemaVersion: 1,
    documentId,
    parts,
    assets,
    warnings: [...(pkg.warnings ?? [])],
    preservedFeatures: preservedFeatures(parsed),
  }
}

function assignSourceIds(element: SourceElement, id: string): void {
  element.id = id
  element.children.forEach((child, index) => assignSourceIds(child, `${id}/${index}`))
}

// ── Package identity ──────────────────────────────────────────────

function packageIdentity(pkg: SourceGraphPackage): string {
  const encoder = new TextEncoder()
  const chunks: Uint8Array[] = []
  for (const [key, part] of [...pkg.parts.entries()].sort(([left], [right]) =>
    compareNames(left, right)
  )) {
    const name = encoder.encode(part.name || key)
    chunks.push(u64le(name.byteLength), name)
    const content = encoder.encode(part.content)
    chunks.push(u64le(content.byteLength), content)
  }
  for (const [name, data] of [...pkg.binaryAssets.entries()].sort(([left], [right]) =>
    compareNames(left, right)
  )) {
    const encodedName = encoder.encode(name)
    chunks.push(u64le(encodedName.byteLength), encodedName)
    chunks.push(u64le(data.byteLength), data)
  }
  return sha256Hex(chunks)
}

function u64le(value: number): Uint8Array {
  const bytes = new Uint8Array(8)
  const view = new DataView(bytes.buffer)
  view.setUint32(0, value >>> 0, true)
  view.setUint32(4, Math.floor(value / 0x1_0000_0000), true)
  return bytes
}

const SHA256_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
])

function sha256Hex(chunks: readonly Uint8Array[]): string {
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0)
  const paddedLength = Math.ceil((total + 9) / 64) * 64
  const data = new Uint8Array(paddedLength)
  let offset = 0
  for (const chunk of chunks) {
    data.set(chunk, offset)
    offset += chunk.byteLength
  }
  data[total] = 0x80
  const bitLength = total * 8
  const view = new DataView(data.buffer)
  view.setUint32(paddedLength - 8, Math.floor(bitLength / 0x1_0000_0000), false)
  view.setUint32(paddedLength - 4, bitLength >>> 0, false)

  const state = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ])
  const words = new Uint32Array(64)
  for (let block = 0; block < paddedLength; block += 64) {
    for (let index = 0; index < 16; index += 1) {
      words[index] = view.getUint32(block + index * 4, false)
    }
    for (let index = 16; index < 64; index += 1) {
      const previous15 = words[index - 15]!
      const previous2 = words[index - 2]!
      const sigma0 = rotateRight(previous15, 7) ^ rotateRight(previous15, 18) ^ (previous15 >>> 3)
      const sigma1 = rotateRight(previous2, 17) ^ rotateRight(previous2, 19) ^ (previous2 >>> 10)
      words[index] = (words[index - 16]! + sigma0 + words[index - 7]! + sigma1) >>> 0
    }

    let a = state[0]!
    let b = state[1]!
    let c = state[2]!
    let d = state[3]!
    let e = state[4]!
    let f = state[5]!
    let g = state[6]!
    let h = state[7]!
    for (let index = 0; index < 64; index += 1) {
      const sum1 = rotateRight(e, 6) ^ rotateRight(e, 11) ^ rotateRight(e, 25)
      const choice = (e & f) ^ (~e & g)
      const temp1 = (h + sum1 + choice + SHA256_K[index]! + words[index]!) >>> 0
      const sum0 = rotateRight(a, 2) ^ rotateRight(a, 13) ^ rotateRight(a, 22)
      const majority = (a & b) ^ (a & c) ^ (b & c)
      const temp2 = (sum0 + majority) >>> 0
      h = g
      g = f
      f = e
      e = (d + temp1) >>> 0
      d = c
      c = b
      b = a
      a = (temp1 + temp2) >>> 0
    }
    state[0] = (state[0]! + a) >>> 0
    state[1] = (state[1]! + b) >>> 0
    state[2] = (state[2]! + c) >>> 0
    state[3] = (state[3]! + d) >>> 0
    state[4] = (state[4]! + e) >>> 0
    state[5] = (state[5]! + f) >>> 0
    state[6] = (state[6]! + g) >>> 0
    state[7] = (state[7]! + h) >>> 0
  }
  return [...state].map((word) => word.toString(16).padStart(8, "0")).join("")
}

function rotateRight(value: number, amount: number): number {
  return (value >>> amount) | (value << (32 - amount))
}

// ── ZIP package decode ────────────────────────────────────────────

function readOoxmlPackage(input: ArrayBuffer | Uint8Array): OoxmlPackage {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  const head = new TextDecoder("utf-8", { fatal: false }).decode(
    bytes.subarray(0, Math.min(bytes.byteLength, 4096)),
  )
  if (!startsWithZip(bytes) && /<office:document\b/i.test(head) && /opendocument/i.test(head)) {
    throw new Error(
      "OpenDocument text has no Word package; read it with parse_document_model",
    )
  }

  let archive: Record<string, Uint8Array>
  try {
    archive = unzipSync(bytes)
  } catch (cause) {
    const message = cause instanceof Error ? cause.message : String(cause)
    throw new Error(`Invalid DOCX ZIP: ${message}`, { cause })
  }

  const mimetype = archive.mimetype
  if (mimetype) {
    const value = strFromU8(mimetype).trim().toLowerCase()
    if (value.startsWith("application/vnd.oasis.opendocument.")) {
      throw new Error(
        value.startsWith("application/vnd.oasis.opendocument.text")
          ? "OpenDocument text has no Word package; read it with parse_document_model"
          : "OpenDocument files are not text documents",
      )
    }
  }

  const parts = new Map<string, { name: string; content: string }>()
  const binaryAssets = new Map<string, Uint8Array>()
  for (const [name, data] of Object.entries(archive)) {
    if (name.endsWith("/")) continue
    if (isTextPart(name)) {
      parts.set(name, { name, content: decodeXmlPart(data) })
    } else {
      binaryAssets.set(name, data)
    }
  }

  const mainPart = mainDocumentPartName(parts)
  if (!parts.has(mainPart)) {
    throw new Error(`Invalid DOCX: missing ${mainPart}`)
  }
  return { parts, binaryAssets }
}

function startsWithZip(bytes: Uint8Array): boolean {
  return bytes.byteLength >= 4 &&
    bytes[0] === 0x50 && bytes[1] === 0x4b &&
    bytes[2] === 0x03 && bytes[3] === 0x04
}

function isTextPart(name: string): boolean {
  if (name === "[Content_Types].xml") return true
  const dot = name.lastIndexOf(".")
  if (dot < 0) return false
  const extension = name.slice(dot + 1).toLowerCase()
  return extension === "xml" || extension === "rels" || extension === "txt"
}

function decodeXmlPart(bytes: Uint8Array): string {
  const bigEndian = (bytes[0] === 0xfe && bytes[1] === 0xff) ||
    (bytes[0] === 0x00 && bytes[1] === 0x3c && bytes[2] === 0x00 && bytes[3] === 0x3f)
  const littleEndian = (bytes[0] === 0xff && bytes[1] === 0xfe) ||
    (bytes[0] === 0x3c && bytes[1] === 0x00 && bytes[2] === 0x3f && bytes[3] === 0x00)
  if (bigEndian || littleEndian) {
    const hasBom = (bytes[0] === 0xfe && bytes[1] === 0xff) ||
      (bytes[0] === 0xff && bytes[1] === 0xfe)
    const body = bytes.subarray(hasBom ? 2 : 0)
    if (body.byteLength % 2 !== 0) {
      throw new Error("Invalid UTF-16 XML byte length")
    }
    let xml: string
    try {
      xml = new TextDecoder(bigEndian ? "utf-16be" : "utf-16le", { fatal: true }).decode(body)
    } catch (cause) {
      throw new Error(
        `Invalid UTF-16 XML: ${cause instanceof Error ? cause.message : String(cause)}`,
        { cause },
      )
    }
    return xml.replace(
      /encoding[\x09-\x0D\x20]*=[\x09-\x0D\x20]*["']UTF-16(?:LE|BE)?["']/gi,
      'encoding="UTF-8"',
    )
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes)
  } catch (cause) {
    throw new Error(
      `Invalid or unsupported XML encoding: ${cause instanceof Error ? cause.message : String(cause)}`,
      { cause },
    )
  }
}

function mainDocumentPartName(
  parts: ReadonlyMap<string, { name: string; content: string }>,
): string {
  const rels = parts.get("_rels/.rels")?.content
  if (rels) {
    const parsed = parseXml(rels)
    if (parsed.root) {
      const target = findRelationshipTarget(parsed.root, "officeDocument")
      if (target) return resolvePartPath("", target)
    }
  }
  return "word/document.xml"
}

function findRelationshipTarget(root: SourceElement, kind: string): string | undefined {
  let found: string | undefined
  const visit = (element: SourceElement): void => {
    if (found || element.kind !== "element") return
    if (element.name === "Relationship") {
      const type = element.attributes.find((attribute) => attribute.name === "Type")?.value
      const target = element.attributes.find((attribute) => attribute.name === "Target")?.value
      const mode = element.attributes.find((attribute) => attribute.name === "TargetMode")?.value
      if (target && mode !== "External" && type?.split("/").pop() === kind) {
        found = target
        return
      }
    }
    for (const child of element.children) visit(child)
  }
  visit(root)
  return found
}

function resolvePartPath(basePartName: string, target: string): string {
  if (!target) return ""
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith("#")) return target
  if (target.startsWith("/")) return target.slice(1)
  const output = basePartName.split("/")
  output.pop()
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") continue
    if (segment === "..") {
      output.pop()
      continue
    }
    output.push(segment)
  }
  return output.join("/")
}

// ── XML index ─────────────────────────────────────────────────────

interface ParsedXml {
  root?: SourceElement
  error?: string
}

interface XmlFrame {
  qname: string
  node: SourceElement
  scope: Record<string, string>
}

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"

function parseXml(content: string): ParsedXml {
  try {
    return { root: parseXmlDocument(content) }
  } catch (cause) {
    return { error: cause instanceof Error ? cause.message : String(cause) }
  }
}

function parseXmlDocument(content: string): SourceElement {
  const offsets = utf8ByteOffsets(content)
  let index = content.charCodeAt(0) === 0xfeff ? 1 : 0
  const document: SourceElement = {
    id: "",
    kind: "document",
    byteStart: 0,
    byteEnd: offsets.byteLength,
    attributes: [],
    children: [],
  }
  const stack: XmlFrame[] = []
  const scopes: Array<Record<string, string>> = [Object.assign(Object.create(null), { xml: XML_NAMESPACE })]
  let rootCount = 0
  let sawRoot = false
  let declarationAllowed = true

  const parent = (): SourceElement => stack.length > 0 ? stack[stack.length - 1]!.node : document
  const addChild = (node: Omit<SourceElement, "id">): SourceElement => {
    const target = parent()
    const withId: SourceElement = { ...node, id: `${target.id}/${target.children.length}` }
    target.children.push(withId)
    return withId
  }

  while (index < content.length) {
    if (content[index] === "<") {
      if (content.startsWith("<!--", index)) {
        const end = content.indexOf("-->", index + 4)
        if (end < 0) fail("unterminated comment", index)
        if (stack.length === 0) declarationAllowed = false
        addChild({
          kind: "comment",
          text: content.slice(index + 4, end),
          byteStart: offsets.at(index),
          byteEnd: offsets.at(end + 3),
          attributes: [],
          children: [],
        })
        index = end + 3
        continue
      }

      if (content.startsWith("<![CDATA[", index)) {
        if (stack.length === 0) fail("CDATA section outside the document element", index)
        const end = content.indexOf("]]>", index + 9)
        if (end < 0) fail("unterminated CDATA section", index)
        const text = content.slice(index + 9, end)
        appendText(text, index, end + 3)
        index = end + 3
        continue
      }

      if (/^<!doctype/i.test(content.slice(index, index + 9))) {
        throw new Error("XML with DTD detected")
      }
      if (content.startsWith("<!", index)) fail("unsupported XML declaration", index)

      if (content.startsWith("<?", index)) {
        const end = content.indexOf("?>", index + 2)
        if (end < 0) fail("unterminated processing instruction", index)
        const body = content.slice(index + 2, end)
        const target = body.match(/^([^\s?]+)/)?.[1] ?? ""
        if (!target) fail("processing instruction has no target", index)
        if (target.toLowerCase() === "xml") {
          if (!declarationAllowed || stack.length > 0) {
            fail("XML declaration must be at the beginning of the document", index)
          }
        } else {
          const value = body.slice(target.length).replace(/^[\x09-\x0D\x20]+/, "")
          addChild({
            kind: "processing_instruction",
            name: target,
            ...(value ? { text: value } : {}),
            byteStart: offsets.at(index),
            byteEnd: offsets.at(end + 2),
            attributes: [],
            children: [],
          })
        }
        declarationAllowed = false
        index = end + 2
        continue
      }

      if (content.startsWith("</", index)) {
        const tag = parseEndTag(content, index)
        const frame = stack.pop()
        if (!frame) fail(`unexpected closing tag </${tag.name}>`, index)
        if (frame.qname !== tag.name) {
          fail(`mismatched closing tag </${tag.name}> for <${frame.qname}>`, index)
        }
        frame.node.byteEnd = offsets.at(tag.end)
        scopes.pop()
        declarationAllowed = false
        index = tag.end
        continue
      }

      const tag = parseStartTag(content, index, scopes)
      if (stack.length === 0) {
        if (rootCount > 0) fail("multiple root elements", index)
        rootCount += 1
        sawRoot = true
      }
      const node = addChild({
        kind: "element",
        ...(tag.namespace ? { namespace: tag.namespace } : {}),
        name: tag.localName,
        byteStart: offsets.at(index),
        byteEnd: offsets.at(tag.end),
        attributes: tag.attributes,
        children: [],
      })
      declarationAllowed = false
      if (!tag.selfClosing) {
        stack.push({ qname: tag.qname, node, scope: tag.scope })
        scopes.push(tag.scope)
      }
      index = tag.end
      continue
    }

    const next = content.indexOf("<", index)
    const end = next < 0 ? content.length : next
    const raw = content.slice(index, end)
    if (stack.length === 0) {
      if (raw.trim()) fail("text outside the document element", index)
    } else {
      appendText(decodeEntities(raw), index, end)
    }
    index = end
  }

  if (stack.length > 0) {
    fail("the root node was opened but never closed", content.length)
  }
  if (!sawRoot) fail("the document has no root element", 0)
  return document

  function appendText(text: string, start: number, end: number): void {
    if (!text) return
    const target = parent()
    const last = target.children[target.children.length - 1]
    if (last?.kind === "text") {
      last.text = `${last.text ?? ""}${text}`
      return
    }
    addChild({
      kind: "text",
      text,
      byteStart: offsets.at(start),
      byteEnd: offsets.at(end),
      attributes: [],
      children: [],
    })
  }
}

interface StartTag {
  qname: string
  localName: string
  namespace?: string
  attributes: SourceAttribute[]
  scope: Record<string, string>
  selfClosing: boolean
  end: number
}

function parseStartTag(
  content: string,
  start: number,
  scopes: readonly Record<string, string>[],
): StartTag {
  let index = start + 1
  const name = readName(content, index, "expected an element name")
  index = name.end
  const attributes: Array<{ qname: string; value: string }> = []
  const seen = new Set<string>()
  let selfClosing = false

  while (true) {
    index = skipWhitespace(content, index)
    if (index >= content.length) fail("unterminated start tag", start)
    if (content[index] === ">") {
      index += 1
      break
    }
    if (content[index] === "/") {
      if (content[index + 1] !== ">") fail("expected '/>'", index)
      selfClosing = true
      index += 2
      break
    }
    const attribute = readName(content, index, "expected an attribute name")
    if (seen.has(attribute.name)) fail(`duplicate attribute '${attribute.name}'`, index)
    seen.add(attribute.name)
    index = skipWhitespace(content, attribute.end)
    if (content[index] !== "=") fail(`expected '=' after attribute '${attribute.name}'`, index)
    index = skipWhitespace(content, index + 1)
    const quote = content[index]
    if (quote !== '"' && quote !== "'") fail("expected a quoted attribute value", index)
    const valueEnd = content.indexOf(quote, index + 1)
    if (valueEnd < 0) fail("unterminated attribute value", index)
    const raw = content.slice(index + 1, valueEnd)
    if (raw.includes("<")) fail("invalid '<' in attribute value", index)
    attributes.push({ qname: attribute.name, value: decodeEntities(raw) })
    index = valueEnd + 1
  }

  const scope = Object.assign(Object.create(null), ...scopes) as Record<string, string>
  const elementAttributes: Array<{ qname: string; value: string }> = []
  for (const attribute of attributes) {
    if (attribute.qname === "xmlns") {
      scope[""] = attribute.value
    } else if (attribute.qname.startsWith("xmlns:")) {
      const prefix = attribute.qname.slice(6)
      if (!prefix) fail("invalid namespace declaration", start)
      scope[prefix] = attribute.value
    } else {
      elementAttributes.push(attribute)
    }
  }

  const resolved = resolveQName(name.name, [scope], true)
  const resolvedAttributes = elementAttributes.map((attribute) => {
    const parts = attribute.qname.indexOf(":")
    if (parts < 0) {
      return { namespace: null, name: attribute.qname, value: attribute.value }
    }
    const resolvedAttribute = resolveQName(attribute.qname, [scope], false)
    return {
      namespace: resolvedAttribute.namespace ?? null,
      name: resolvedAttribute.localName,
      value: attribute.value,
    }
  })

  return {
    qname: name.name,
    localName: resolved.localName,
    ...(resolved.namespace ? { namespace: resolved.namespace } : {}),
    attributes: resolvedAttributes,
    scope,
    selfClosing,
    end: index,
  }
}

function parseEndTag(content: string, start: number): { name: string; end: number } {
  let index = skipWhitespace(content, start + 2)
  const name = readName(content, index, "expected a closing tag name")
  index = skipWhitespace(content, name.end)
  if (content[index] !== ">") fail("expected '>' in closing tag", index)
  return { name: name.name, end: index + 1 }
}

function readName(content: string, start: number, message: string): { name: string; end: number } {
  let index = start
  while (index < content.length) {
    const code = content.charCodeAt(index)
    if (code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d ||
      code === 0x3e || code === 0x2f || code === 0x3d) {
      break
    }
    index += 1
  }
  const name = content.slice(start, index)
  if (!name) fail(message, start)
  if (name.includes("<") || name.includes("&")) fail(`invalid XML name '${name}'`, start)
  return { name, end: index }
}

function resolveQName(
  qname: string,
  scopes: readonly Record<string, string>[],
  element: boolean,
): { localName: string; namespace?: string } {
  const separator = qname.indexOf(":")
  if (separator < 0) {
    if (!element) return { localName: qname }
    const namespace = lookupNamespace(scopes, "")
    return namespace ? { localName: qname, namespace } : { localName: qname }
  }
  if (separator === 0 || separator === qname.length - 1 || qname.indexOf(":", separator + 1) >= 0) {
    throw new Error(`invalid XML name '${qname}'`)
  }
  const prefix = qname.slice(0, separator)
  const localName = qname.slice(separator + 1)
  const namespace = lookupNamespace(scopes, prefix)
  if (namespace === undefined) {
    throw new Error(`unknown namespace prefix '${prefix}'`)
  }
  return namespace ? { localName, namespace } : { localName }
}

function lookupNamespace(
  scopes: readonly Record<string, string>[],
  prefix: string,
): string | undefined {
  for (let index = scopes.length - 1; index >= 0; index -= 1) {
    const scope = scopes[index]!
    if (Object.prototype.hasOwnProperty.call(scope, prefix)) return scope[prefix]
  }
  return undefined
}

function skipWhitespace(content: string, start: number): number {
  let index = start
  while (index < content.length) {
    const code = content.charCodeAt(index)
    if (code !== 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) break
    index += 1
  }
  return index
}

function decodeEntities(raw: string): string {
  if (!raw.includes("&")) return raw
  let output = ""
  let index = 0
  while (index < raw.length) {
    const ampersand = raw.indexOf("&", index)
    if (ampersand < 0) {
      output += raw.slice(index)
      break
    }
    output += raw.slice(index, ampersand)
    const semicolon = raw.indexOf(";", ampersand + 1)
    if (semicolon < 0) fail("unterminated entity reference", ampersand)
    const entity = raw.slice(ampersand + 1, semicolon)
    output += decodeEntity(entity, ampersand)
    index = semicolon + 1
  }
  return output
}

function decodeEntity(entity: string, position: number): string {
  switch (entity) {
    case "lt":
      return "<"
    case "gt":
      return ">"
    case "amp":
      return "&"
    case "quot":
      return '"'
    case "apos":
      return "'"
  }
  if (entity.startsWith("#")) {
    const hex = entity[1] === "x" || entity[1] === "X"
    const digits = entity.slice(hex ? 2 : 1)
    if (!digits || !(hex ? /^[0-9a-f]+$/i : /^[0-9]+$/).test(digits)) {
      fail("invalid character reference", position)
    }
    const codePoint = Number.parseInt(digits, hex ? 16 : 10)
    if (codePoint === 0 || codePoint > 0x10ffff ||
      (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
      fail("invalid character reference", position)
    }
    return String.fromCodePoint(codePoint)
  }
  fail(`unknown entity '&${entity};'`, position)
}

function fail(message: string, position: number): never {
  throw new Error(`${message} at byte ${Math.max(0, position)}`)
}

function utf8ByteLength(value: string): number {
  return utf8ByteOffsets(value).byteLength
}

function utf8ByteOffsets(value: string): { at(index: number): number; byteLength: number } {
  const offsets = new Int32Array(value.length + 1)
  let bytes = 0
  for (let index = 0; index < value.length; index += 1) {
    offsets[index] = bytes
    const code = value.charCodeAt(index)
    if (code < 0x80) {
      bytes += 1
    } else if (code < 0x800) {
      bytes += 2
    } else if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1)
      if (next >= 0xdc00 && next <= 0xdfff) {
        bytes += 4
        offsets[index + 1] = bytes - 2
        index += 1
        offsets[index + 1] = bytes
      } else {
        bytes += 3
      }
    } else {
      bytes += 3
    }
  }
  offsets[value.length] = bytes
  return { at: (index: number) => offsets[Math.max(0, Math.min(index, value.length))]!, byteLength: bytes }
}

// ── Preserved feature inventory ───────────────────────────────────

interface FeatureMatch {
  feature: string
  support: PreservedFeature["support"]
}

function preservedFeatureKind(namespace: string, localName: string): FeatureMatch | undefined {
  const word = namespace.endsWith("wordprocessingml/2006/main") ||
    namespace.endsWith("wordprocessingml/main")
  if (word) {
    if (localName === "altChunk" || localName === "subDoc" ||
      localName === "object" || localName === "control" ||
      localName === "movie" || localName === "sound") {
      return { feature: localName, support: "preserved_only" }
    }
  }
  if (namespace.includes("drawingml") && namespace.includes("diagram") &&
    localName === "dataModel") {
    return { feature: "diagram", support: "preserved_only" }
  }
  if (namespace.endsWith("/math") && localName === "oMath") {
    return { feature: "equation", support: "approximate_rendering" }
  }
  return undefined
}

function preservedFeatures(parsed: ReadonlyMap<string, ParsedXml>): PreservedFeature[] {
  const counts = new Map<string, PreservedFeature>()
  for (const [partName, result] of parsed) {
    if (!result.root) continue
    walk(result.root, partName)
  }
  return [...counts.values()].sort((left, right) =>
    compareNames(left.partName, right.partName) ||
    compareNames(left.feature, right.feature) ||
    compareNames(left.support, right.support)
  )

  function walk(element: SourceElement, partName: string): void {
    if (element.kind === "element" && element.name) {
      const match = preservedFeatureKind(element.namespace ?? "", element.name)
      if (match) {
        const key = `${partName}\u0000${match.feature}\u0000${match.support}`
        const existing = counts.get(key)
        if (existing) existing.occurrences += 1
        else {
          counts.set(key, {
            partName,
            feature: match.feature,
            occurrences: 1,
            support: match.support,
          })
        }
      }
    }
    for (const child of element.children) walk(child, partName)
  }
}

function compareNames(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}
