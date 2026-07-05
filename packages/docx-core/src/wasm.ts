/**
 * DOCX WASM Bridge & OOXML Package utilities
 *
 * Compatible with the public @extend-ai/react-docx WASM integration surface; see docs/upstream-extend-ui.md for attribution.
 * Framework-agnostic pure functions for:
 * - WASM initialization and lifecycle
 * - OOXML package parsing/serialization
 * - Document model building from WASM output
 */

import { strFromU8, unzipSync } from "fflate"
import type {
  WasmSource, WasmInitOutput,
  OoxmlPackage, OoxmlPart,
  DocModel,
  DocNode,
  ParagraphChildNode,
  ParagraphNode,
  TableCellNode,
  TableNode,
  TableRowNode,
} from "./types"

// ============================================================
// Default empty document XML
// ============================================================

const DEFAULT_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t></w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`

// ============================================================
// WASM State
// ============================================================

let wasmOutput: WasmInitOutput | null = null
let wasmSourceConfigured: WasmSource | undefined = undefined

// ============================================================
// WASM Initialization
// ============================================================

export function setWasmSource(source: WasmSource): void {
  wasmSourceConfigured = source
}

export function getWasmSource(): WasmSource | undefined {
  return wasmSourceConfigured
}

/**
 * Initialize the WASM engine.
 * In browser: loads docx_wasm_bg.wasm
 * In Node.js: uses the bundled WASM binary
 */
export async function initWasm(source?: WasmSource): Promise<WasmInitOutput> {
  if (wasmOutput) return wasmOutput

  const src = source ?? wasmSourceConfigured
  if (!src && typeof window === "undefined") {
    throw new Error("WASM source must be provided (call setWasmSource() or pass to initWasm())")
  }

  // Dynamic import — the actual WASM glue code is loaded lazily
  // In production, this would load the wasm-bindgen generated JS + .wasm file
  const module = await loadWasmModule(src)

  wasmOutput = module
  return module
}

/**
 * Load the WASM module. In a real deployment, this loads the compiled
 * Rust→WASM output via wasm-bindgen.
 *
 * For the Vue port, we reuse the existing WASM binary from @extend-ai/react-docx
 * or compile the Rust source independently.
 */
async function loadWasmModule(source?: WasmSource): Promise<WasmInitOutput> {
  // In the real implementation, this calls the wasm-bindgen generated init function:
  //   import init, { ... } from "./docx_wasm"
  //   await init(source)

  // For now, we provide a stub that throws with clear instructions
  throw new Error(
    "WASM module not loaded. To use docx-core in production:\n" +
    "1. Copy docx_wasm_bg.wasm from @extend-ai/react-docx/dist/ to your public directory\n" +
    "2. Import the wasm-bindgen glue: await init('/path/to/docx_wasm_bg.wasm')\n" +
    "3. Or use the pre-built WASM package from npm"
  )
}

// ============================================================
// OOXML Package Operations
// ============================================================

export function createMinimalDocxPackage(documentXml?: string): OoxmlPackage {
  const docXml = documentXml ?? DEFAULT_DOCUMENT_XML
  const parts = new Map<string, OoxmlPart>()
  parts.set("/word/document.xml", { name: "/word/document.xml", content: docXml })
  return {
    parts,
    binaryAssets: new Map(),
  }
}

export function getPart(pkg: OoxmlPackage, partName: string): OoxmlPart | undefined {
  return pkg.parts.get(partName)
}

export function withPart(pkg: OoxmlPackage, part: OoxmlPart): OoxmlPackage {
  const parts = new Map(pkg.parts)
  parts.set(part.name, part)
  return { ...pkg, parts }
}

// ============================================================
// DOCX Parsing (from ArrayBuffer)
// ============================================================

/**
 * Parse a .docx ArrayBuffer into an OoxmlPackage.
 * Internally uses JSZip-like decompression + XML parsing.
 */
export async function parseDocx(input: ArrayBuffer): Promise<OoxmlPackage> {
  // In real implementation, this uses fflate or JSZip to decompress the ZIP
  // and extract all XML parts + binary assets
  //
  // For WASM-accelerated parsing, this delegates to parse_docx_wasm()
  // after initWasm() has been called.

  if (wasmOutput) {
    return wasmParseDocx(input)
  }

  // Fallback: pure JS ZIP decompression
  return parseDocxFallback(input)
}

async function parseDocxFallback(input: ArrayBuffer): Promise<OoxmlPackage> {
  const files = unzipSync(new Uint8Array(input))
  const parts = new Map<string, OoxmlPart>()
  const binaryAssets = new Map<string, Uint8Array>()

  for (const [name, data] of Object.entries(files)) {
    const partName = `/${name}`
    if (name.endsWith(".xml") || name.endsWith(".rels")) {
      parts.set(partName, { name: partName, content: strFromU8(data) })
    } else {
      binaryAssets.set(partName, data)
    }
  }

  if (!parts.has("/word/document.xml")) {
    throw new Error("Invalid DOCX: missing /word/document.xml")
  }

  return { parts, binaryAssets }
}

async function wasmParseDocx(buffer: ArrayBuffer): Promise<OoxmlPackage> {
  if (!wasmOutput) throw new Error("WASM not initialized")
  // Call wasmOutput.parse_docx_wasm(...) and convert to OoxmlPackage
  throw new Error("WASM bridge not yet wired")
}

// ============================================================
// DOCX Serialization (to ArrayBuffer)
// ============================================================

export async function packageToArrayBuffer(pkg: OoxmlPackage): Promise<ArrayBuffer> {
  if (wasmOutput) {
    return wasmPackageToArrayBuffer(pkg)
  }
  throw new Error("WASM not initialized")
}

async function wasmPackageToArrayBuffer(_pkg: OoxmlPackage): Promise<ArrayBuffer> {
  throw new Error("WASM bridge not yet wired")
}

// ============================================================
// Document Model Building
// ============================================================

export async function buildDocModel(pkg: OoxmlPackage): Promise<DocModel> {
  if (wasmOutput) {
    return wasmBuildDocModel(pkg)
  }
  return buildDocModelFallback(pkg)
}

async function wasmBuildDocModel(_pkg: OoxmlPackage): Promise<DocModel> {
  throw new Error("WASM bridge not yet wired")
}

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml")
  const parserError = doc.querySelector("parsererror")
  if (parserError) throw new Error(parserError.textContent || "Invalid XML")
  return doc
}

function textFromElement(el: Element): string {
  return Array.from(el.getElementsByTagName("w:t")).map((t) => t.textContent ?? "").join("")
}

function runFromElement(run: Element): ParagraphChildNode {
  const text = textFromElement(run)
  const rPr = run.getElementsByTagName("w:rPr")[0]
  return {
    type: "text",
    text,
    style: {
      bold: !!rPr?.getElementsByTagName("w:b").length,
      italic: !!rPr?.getElementsByTagName("w:i").length,
      underline: !!rPr?.getElementsByTagName("w:u").length,
      strike: !!rPr?.getElementsByTagName("w:strike").length,
      color: rPr?.getElementsByTagName("w:color")[0]?.getAttribute("w:val") ? `#${rPr.getElementsByTagName("w:color")[0]!.getAttribute("w:val")}` : undefined,
    },
  }
}

function paragraphFromElement(p: Element): ParagraphNode {
  const pPr = Array.from(p.children).find((child) => child.localName === "pPr")
  const pStyle = pPr?.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") ?? ""
  const headingMatch = /Heading([1-6])|标题\s*([1-6])/i.exec(pStyle)
  const children = Array.from(p.children)
    .filter((child) => child.localName === "r")
    .map((run) => runFromElement(run))
    .filter((run) => run.type !== "text" || run.text)

  if (children.length === 0) children.push({ type: "text", text: textFromElement(p) })

  return {
    type: "paragraph",
    children,
    style: headingMatch ? { headingLevel: Number(headingMatch[1] || headingMatch[2]) as 1 | 2 | 3 | 4 | 5 | 6, styleId: pStyle } : { styleId: pStyle || undefined },
  }
}

function tableFromElement(tbl: Element): TableNode {
  const rows: TableRowNode[] = Array.from(tbl.children)
    .filter((child) => child.localName === "tr")
    .map((tr) => ({
      type: "table-row",
      cells: Array.from(tr.children)
        .filter((child) => child.localName === "tc")
        .map((tc): TableCellNode => ({
          type: "table-cell",
          nodes: Array.from(tc.children)
            .filter((child) => child.localName === "p" || child.localName === "tbl")
            .map((child) => child.localName === "tbl" ? tableFromElement(child) : paragraphFromElement(child)),
        })),
    }))
  return { type: "table", rows }
}

function buildDocModelFallback(pkg: OoxmlPackage): DocModel {
  const xml = pkg.parts.get("/word/document.xml")?.content
  if (!xml) throw new Error("Invalid DOCX: missing /word/document.xml")
  const doc = parseXml(xml)
  const body = doc.getElementsByTagName("w:body")[0]
  const nodes: DocNode[] = []

  for (const child of Array.from(body?.children ?? [])) {
    if (child.localName === "p") nodes.push(paragraphFromElement(child))
    if (child.localName === "tbl") nodes.push(tableFromElement(child))
  }

  return {
    nodes: nodes.length ? nodes : [{ type: "paragraph", children: [{ type: "text", text: "Empty document" }] }],
    metadata: { headerSections: [], footerSections: [] },
  }
}

export async function buildDocModelFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<{ package: OoxmlPackage; model: DocModel }> {
  const buffer = bytes instanceof ArrayBuffer ? bytes : bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
  const pkg = await parseDocx(buffer)
  const model = await buildDocModel(pkg)
  return { package: pkg, model }
}

// ============================================================
// Model → XML Serialization
// ============================================================

export async function modelToDocumentXml(
  model: DocModel,
  basePackage?: OoxmlPackage,
): Promise<string> {
  if (wasmOutput) {
    return wasmModelToDocumentXml(model, basePackage)
  }
  throw new Error("WASM not initialized")
}

async function wasmModelToDocumentXml(
  _model: DocModel,
  _basePackage?: OoxmlPackage,
): Promise<string> {
  throw new Error("WASM bridge not yet wired")
}

export async function serializeDocModel(
  model: DocModel,
  basePackage?: OoxmlPackage,
): Promise<OoxmlPackage> {
  const documentXml = await modelToDocumentXml(model, basePackage)
  const base = basePackage ? {
    parts: new Map(basePackage.parts),
    binaryAssets: new Map(basePackage.binaryAssets),
  } : createMinimalDocxPackage()
  return withPart(base, { name: "/word/document.xml", content: documentXml })
}

export async function serializeDocx(
  model: DocModel,
  basePackage?: OoxmlPackage,
): Promise<ArrayBuffer> {
  const pkg = await serializeDocModel(model, basePackage)
  return packageToArrayBuffer(pkg)
}

// ============================================================
// WASM Internal Bridge (stubs)
// ============================================================

// These would be the actual WASM function calls in production:

// function wasmParseDocx(input: ArrayBuffer): OoxmlPackage { ... }
// function wasmBuildDocModel(pkg: OoxmlPackage): DocModel { ... }
// function wasmModelToDocumentXml(model: DocModel, basePkg?: OoxmlPackage): string { ... }
// function wasmSerializeDocx(model: DocModel, basePkg?: OoxmlPackage): ArrayBuffer { ... }
// function wasmPackageToArrayBuffer(pkg: OoxmlPackage): ArrayBuffer { ... }

