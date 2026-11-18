// OpenDocument text body reader.
//
// Converts `office:text` content into the platform-neutral `DocModel` used by
// the rest of the package. OpenDocument files have no Word package, so every
// paragraph and table is built without `sourceXml`; the model carries the text,
// resolved styles, numbering, pictures and page setup needed to render it.

import type {
  DocModel,
  DocNode,
  DocumentNoteDefinition,
  ImageRunNode,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
  ParagraphChildNode,
  ParagraphNode,
  ParagraphStyle,
  ParagraphStyleDefinition,
  TableCellNode,
  TableNode,
  TableRowNode,
  TextBoxContent,
  TextBoxParagraph,
  TextStyle,
} from "../engine/types"
import {
  attribute,
  base64ToBytes,
  childElements,
  firstChildElement,
  firstDescendantElement,
  lengthToPx,
  localName,
  parseXmlDocument,
  textContent,
} from "./open-document-xml"
import {
  collectOdfStyles,
  resolveOdfListDefinition,
  resolveOdfPageLayout,
  resolveOdfParagraphStyle,
  resolveOdfTableColumnWidth,
  resolveOdfTableCellStyle,
  resolveOdfTableRowStyle,
  resolveOdfTableStyle,
  resolveOdfTextStyle,
  type OdfStyleSheet,
} from "./open-document-styles"

export interface OpenDocumentImageAsset {
  bytes: Uint8Array
  contentType: string
}

export interface OpenDocumentReadResult {
  model: DocModel
  warnings: string[]
}

interface ImageContext {
  images: Map<string, OpenDocumentImageAsset>
}

interface NumberingContext {
  definitions: NumberingDefinitionSet
  byStyle: Map<string, number>
  nextNumId: number
  nextAbstractId: number
}

interface ParseContext extends ImageContext {
  sheet: OdfStyleSheet
  numbering: NumberingContext
  paragraphStyles: Map<string, ParagraphStyleDefinition>
  footnotes: DocumentNoteDefinition[]
  endnotes: DocumentNoteDefinition[]
  nextNoteId: number
  warnings: string[]
}

interface ListContext {
  styleName?: string
  level: number
}

const MAX_REPEATED_ROWS = 1_000
const MAX_REPEATED_CELLS = 10_000

const IMAGE_CONTENT_TYPES: Readonly<Record<string, string>> = Object.freeze({
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

function imageContentType(href: string, explicit: string | undefined): string {
  if (explicit) return explicit
  const extension = /\.([^./\\]+)$/.exec(href)?.[1]?.toLowerCase()
  return extension ? IMAGE_CONTENT_TYPES[extension] ?? "application/octet-stream" : "application/octet-stream"
}

function bytesToBase64(bytes: Uint8Array): string | undefined {
  if (typeof btoa === "function") {
    let binary = ""
    const chunk = 0x8000
    for (let index = 0; index < bytes.length; index += chunk) {
      binary += String.fromCharCode(...bytes.subarray(index, index + chunk))
    }
    return btoa(binary)
  }
  const buffer = (globalThis as { Buffer?: { from(value: Uint8Array): { toString(encoding: string): string } } }).Buffer
  return buffer ? buffer.from(bytes).toString("base64") : undefined
}

function imageDataUri(asset: OpenDocumentImageAsset): string | undefined {
  const base64 = bytesToBase64(asset.bytes)
  return base64 ? `data:${asset.contentType};base64,${base64}` : undefined
}

function normalizeImageHref(href: string): string {
  const unescaped = href.replace(/^\.\//, "")
  try {
    return decodeURIComponent(unescaped)
  } catch {
    return unescaped
  }
}

function resolveImageAsset(context: ImageContext, href: string): OpenDocumentImageAsset | undefined {
  if (href.startsWith("data:")) {
    const match = /^data:([^;,]+)?(?:;base64)?,(.*)$/s.exec(href)
    if (!match) return undefined
    const bytes = base64ToBytes(match[2] ?? "")
    return bytes ? { bytes, contentType: match[1] ?? "application/octet-stream" } : undefined
  }
  const normalized = normalizeImageHref(href)
  return context.images.get(normalized)
    ?? context.images.get(normalized.replace(/^Pictures\//, ""))
    ?? context.images.get(`Pictures/${normalized}`)
}

function styleOrUndefined(style: TextStyle): TextStyle | undefined {
  return Object.keys(style).length > 0 ? style : undefined
}

function pushText(
  children: ParagraphChildNode[],
  text: string,
  style: TextStyle | undefined,
  link: string | undefined,
): void {
  if (!text) return
  const last = children[children.length - 1]
  if (
    last?.type === "text" &&
    last.link === link &&
    JSON.stringify(last.style ?? null) === JSON.stringify(style ?? null)
  ) {
    last.text += text
    return
  }
  children.push({
    type: "text",
    text,
    ...(style ? { style } : {}),
    ...(link ? { link } : {}),
  })
}

function whitespaceCollapsed(value: string | null | undefined): string {
  return (value ?? "").replace(/[\t\n\r ]+/g, " ")
}

function ensureNumbering(context: ParseContext, styleName: string | undefined, level: number): number {
  const key = styleName ?? "default"
  const existing = context.numbering.byStyle.get(key)
  if (existing !== undefined) return existing

  const definition = resolveOdfListDefinition(context.sheet, styleName)
  const numId = context.numbering.nextNumId++
  const abstractNumId = context.numbering.nextAbstractId++
  const levels = definition ? definition.levels.map((entry) => ({
    ilvl: entry.level,
    start: entry.start ?? 1,
    format: entry.format,
    text: entry.text,
    suffix: entry.suffix,
    ...(entry.indent ? { indent: entry.indent } : {}),
    ...(entry.runStyle ? { runStyle: entry.runStyle } : {}),
    ...(entry.bulletFontFamily ? { bulletFontFamily: entry.bulletFontFamily } : {}),
  })) : fallbackListLevels()
  context.numbering.definitions.abstracts.push({ abstractNumId, levels })
  context.numbering.definitions.instances.push({ numId, abstractNumId })
  context.numbering.byStyle.set(key, numId)
  // The level requested by the first paragraph is part of the contract only
  // through `ParagraphNumbering.ilvl`; the definition carries every level.
  void level
  return numId
}

function fallbackListLevels(): NumberingLevelDefinition[] {
  return Array.from({ length: 9 }, (_unused, index) => ({
    ilvl: index,
    start: 1,
    format: "bullet",
    text: index % 2 === 0 ? "•" : "◦",
    suffix: "space" as const,
    indent: { leftTwips: 720 * (index + 1), hangingTwips: 360 },
  }))
}

function applyListing(
  context: ParseContext,
  style: ParagraphNode["style"],
  list: ListContext | undefined,
  paragraphStyleList: string | undefined,
): void {
  const styleName = list?.styleName ?? paragraphStyleList
  const level = list?.level ?? 0
  if (!style || styleName === undefined) return
  const numId = ensureNumbering(context, styleName, level)
  style.numbering = { numId, ilvl: level }
  const definition = resolveOdfListDefinition(context.sheet, styleName)
  const levelDefinition = definition?.levels.find((entry) => entry.level === level)
  if (levelDefinition?.indent && !style.indent) style.indent = levelDefinition.indent
}

function noteFromElement(
  element: Element,
  context: ParseContext,
  kind: "footnote" | "endnote",
): string | undefined {
  const body = firstChildElement(element, "note-body")
  if (!body) return undefined
  const id = context.nextNoteId++
  const nodes = parseBlockChildren(body, context, undefined)
  const citation = firstChildElement(element, "note-citation")
  const label = citation ? textContent(citation).trim() : String(id + 1)
  const text = nodes
    .map((node) => node.type === "paragraph"
      ? node.children.map((child) => child.type === "text" ? child.text : "").join("")
      : "")
    .join("\n")
  const definition: DocumentNoteDefinition = {
    id,
    text,
    nodes,
  }
  if (kind === "footnote") context.footnotes.push(definition)
  else context.endnotes.push(definition)
  return label || String(id + 1)
}

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
}

function textBoxFromFrame(frame: Element, context: ParseContext): { content: TextBoxContent; sourceXml: string } | undefined {
  const textBox = firstDescendantElement(frame, "text-box")
  if (!textBox) return undefined
  const paragraphs: TextBoxParagraph[] = []
  const xmlParagraphs: string[] = []
  for (const element of childElements(textBox)) {
    const name = localName(element)
    if (name !== "p" && name !== "h") continue
    const paragraph = paragraphFromElement(element, context, undefined)
    const segments = paragraph.children
      .filter((child): child is Extract<ParagraphChildNode, { type: "text" }> => child.type === "text")
      .map((child) => ({ text: child.text, ...(child.style ? { style: child.style } : {}) }))
    if (segments.length === 0) continue
    const fontSizePt = segments.find((segment) => segment.style?.fontSizePt)?.style?.fontSizePt ?? 12
    paragraphs.push({
      align: paragraph.style?.align,
      lineHeightPx: Math.max(12, Math.round((fontSizePt * 96) / 72 * 1.2)),
      segments,
    })
    const runs = segments
      .map((segment) => `<w:r><w:t xml:space="preserve">${xmlEscape(segment.text)}</w:t></w:r>`)
      .join("")
    xmlParagraphs.push(`<w:p>${runs}</w:p>`)
  }
  if (paragraphs.length === 0) return undefined
  return {
    content: { paragraphs },
    sourceXml: `<w:txbxContent>${xmlParagraphs.join("")}</w:txbxContent>`,
  }
}

function imageFromFrame(frame: Element, context: ParseContext): ImageRunNode | undefined {
  const image = firstDescendantElement(frame, "image")
  if (!image) {
    const textBox = textBoxFromFrame(frame, context)
    if (!textBox) return undefined
    const widthPx = lengthToPx(attribute(frame, "width"))
    const heightPx = lengthToPx(attribute(frame, "height"))
    const anchorType = attribute(frame, "anchor-type")
    const node: ImageRunNode = {
      type: "image",
      syntheticTextBox: true,
      textBox: textBox.content,
      textBoxSourceXml: textBox.sourceXml,
      ...(widthPx !== undefined ? { widthPx } : {}),
      ...(heightPx !== undefined ? { heightPx } : {}),
    }
    if (anchorType && anchorType !== "as-char" && anchorType !== "char") {
      const xPx = lengthToPx(attribute(frame, "x"))
      const yPx = lengthToPx(attribute(frame, "y"))
      node.floating = {
        ...(xPx !== undefined ? { xPx } : {}),
        ...(yPx !== undefined ? { yPx } : {}),
        wrapType: attribute(frame, "wrap") === "none" ? "none" : "square",
      }
    }
    return node
  }
  const href = attribute(image, "href") ?? attribute(image, "href", "http://www.w3.org/1999/xlink")
  const binary = firstChildElement(image, "binary-data")
  let asset: OpenDocumentImageAsset | undefined
  if (href && !href.startsWith("#")) asset = resolveImageAsset(context, href)
  if (!asset && binary) {
    const bytes = base64ToBytes(textContent(binary))
    if (bytes) {
      asset = {
        bytes,
        contentType: imageContentType(href ?? "", attribute(image, "mime-type")),
      }
    }
  }
  const widthPx = lengthToPx(attribute(frame, "width"))
  const heightPx = lengthToPx(attribute(frame, "height"))
  const anchorType = attribute(frame, "anchor-type")
  const description = firstChildElement(frame, "desc") ?? firstChildElement(frame, "title")
  const alt = description ? textContent(description).trim() : undefined
  const src = asset ? imageDataUri(asset) : undefined
  const node: ImageRunNode = {
    type: "image",
    ...(src ? { src } : {}),
    ...(asset ? { data: asset.bytes, contentType: asset.contentType } : {}),
    ...(alt ? { alt } : {}),
    ...(widthPx !== undefined ? { widthPx } : {}),
    ...(heightPx !== undefined ? { heightPx } : {}),
  }
  if (anchorType && anchorType !== "as-char" && anchorType !== "char") {
    const xPx = lengthToPx(attribute(frame, "x"))
    const yPx = lengthToPx(attribute(frame, "y"))
    node.floating = {
      ...(xPx !== undefined ? { xPx } : {}),
      ...(yPx !== undefined ? { yPx } : {}),
      wrapType: attribute(frame, "wrap") === "none" ? "none" : "square",
    }
  }
  return node
}

function inlineChildren(
  parent: Node,
  context: ParseContext,
  baseStyle: TextStyle | undefined,
  link: string | undefined,
  output: ParagraphChildNode[],
): void {
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 3) {
      pushText(output, whitespaceCollapsed(child.nodeValue), baseStyle, link)
      continue
    }
    if (child.nodeType !== 1) continue
    const element = child as Element
    switch (localName(element)) {
      case "s": {
        const count = Math.max(1, Math.min(10_000, Number(attribute(element, "c") ?? "1") || 1))
        pushText(output, " ".repeat(count), baseStyle, link)
        break
      }
      case "tab":
        pushText(output, "\t", baseStyle, link)
        break
      case "page-number":
      case "page-count":
        // OpenDocument page fields have no saved result. A static 1 keeps the
        // footer readable; the local model has no Word field sequence.
        pushText(output, "1", baseStyle, link)
        break
      case "date": {
        const value = attribute(element, "date-value")
        if (value) pushText(output, value.slice(0, 10), baseStyle, link)
        break
      }
      case "time": {
        const value = attribute(element, "time-value")
        if (value) pushText(output, value.slice(0, 8), baseStyle, link)
        break
      }
      case "line-break":
        pushText(output, "\n", baseStyle, link)
        break
      case "span": {
        const spanStyle = resolveOdfTextStyle(context.sheet, attribute(element, "style-name"), baseStyle)
        inlineChildren(element, context, spanStyle ?? baseStyle, link, output)
        break
      }
      case "a": {
        const href = attribute(element, "href")
          ?? attribute(element, "href", "http://www.w3.org/1999/xlink")
        inlineChildren(element, context, baseStyle, href ?? link, output)
        break
      }
      case "note": {
        const kind = attribute(element, "note-class") === "endnote" ? "endnote" : "footnote"
        const label = noteFromElement(element, context, kind)
        if (label !== undefined) {
          const id = (kind === "endnote" ? context.endnotes : context.footnotes).at(-1)?.id ?? context.nextNoteId - 1
          pushText(output, label, baseStyle, link)
          const run = output[output.length - 1]
          if (run?.type === "text") run.noteReference = { kind, id }
        }
        break
      }
      case "frame": {
        const hasPicture = firstDescendantElement(element, "image") !== undefined
        const image = imageFromFrame(element, context)
        if (image) output.push(image)
        // A frame that contains both a picture and a text box (a caption or
        // labelled figure) keeps the caption visible after the picture. The
        // model has no nested frame content, so the text joins the paragraph.
        if (hasPicture) {
          const textBox = textBoxFromFrame(element, context)
          const textStyle = textBox?.content.paragraphs[0]?.segments[0]?.style
          for (const paragraph of textBox?.content.paragraphs ?? []) {
            pushText(output, paragraph.segments.map((segment) => segment.text).join(""), textStyle, link)
          }
        }
        break
      }
      case "bookmark":
      case "bookmark-start":
      case "bookmark-end":
      case "reference-mark":
      case "reference-mark-start":
      case "reference-mark-end":
      case "sequence-decls":
      case "soft-page-break":
      case "meta":
        break
      case "note-body":
      case "note-citation":
        break
      default:
        inlineChildren(element, context, baseStyle, link, output)
    }
  }
}

function paragraphFromElement(
  element: Element,
  context: ParseContext,
  list: ListContext | undefined,
): ParagraphNode {
  const styleName = attribute(element, "style-name") ?? attribute(element, "cond-style-name")
  const resolved = resolveOdfParagraphStyle(context.sheet, styleName)
  const style = resolved.style
  const explicitOutline = Number(attribute(element, "outline-level") ?? "")
  if (localName(element) === "h" && Number.isFinite(explicitOutline) && explicitOutline > 0) {
    style.headingLevel = Math.min(6, Math.max(1, explicitOutline)) as ParagraphStyle["headingLevel"]
  }
  if (styleName) {
    context.paragraphStyles.set(styleName, resolved.definition)
  }
  applyListing(context, style, list, resolved.listStyleName)
  const children: ParagraphChildNode[] = []
  inlineChildren(element, context, resolved.runStyle, undefined, children)
  return {
    type: "paragraph",
    children,
    ...(Object.keys(style).length > 0 ? { style } : {}),
  }
}

function paragraphFromText(text: string, style?: ParagraphNode["style"]): ParagraphNode {
  return {
    type: "paragraph",
    children: text ? [{ type: "text", text: whitespaceCollapsed(text) }] : [],
    ...(style && Object.keys(style).length > 0 ? { style } : {}),
  }
}

function listFromElement(
  element: Element,
  context: ParseContext,
  inherited: ListContext | undefined,
): DocNode[] {
  const styleName = attribute(element, "style-name") ?? inherited?.styleName
  const level = inherited?.level ?? 0
  const nodes: DocNode[] = []
  for (const item of childElements(element)) {
    const name = localName(item)
    if (name === "list-header") {
      for (const paragraph of childElements(item)) {
        if (localName(paragraph) === "p" || localName(paragraph) === "h") {
          nodes.push(paragraphFromElement(paragraph, context, undefined))
        }
      }
      continue
    }
    if (name !== "list-item") continue
    for (const child of childElements(item)) {
      const childName = localName(child)
      if (childName === "p" || childName === "h") {
        nodes.push(paragraphFromElement(child, context, { styleName, level }))
      } else if (childName === "list") {
        nodes.push(...listFromElement(child, context, { styleName: attribute(child, "style-name") ?? styleName, level: level + 1 }))
      }
    }
  }
  return nodes
}

interface RawTableCell {
  originRow: number
  originColumn: number
  colSpan: number
  rowSpan: number
  node: TableCellNode
}

interface RawTableRow {
  cells: RawTableCell[]
  styleName?: string
}

function repeatedCount(element: Element, name: string, max: number): number {
  const parsed = Number(attribute(element, name) ?? "1")
  if (!Number.isFinite(parsed) || parsed < 1) return 1
  return Math.min(max, Math.floor(parsed))
}

function parseTableCellContent(cell: Element, context: ParseContext): DocNode[] {
  const nodes = parseBlockChildren(cell, context, undefined)
  if (nodes.length > 0) return nodes
  const text = textContent(cell).trim()
  return [paragraphFromText(text)]
}

function emptyTableCell(): TableCellNode {
  return { type: "table-cell", nodes: [paragraphFromText("")] }
}

function tableFromElement(element: Element, context: ParseContext): TableNode {
  const tableStyleName = attribute(element, "style-name")
  const tableStyle = resolveOdfTableStyle(context.sheet, tableStyleName)
  const columnWidths: number[] = []
  for (const column of childElements(element)) {
    if (localName(column) !== "table-column") continue
    const width = resolveOdfTableColumnWidth(context.sheet, attribute(column, "style-name"))
    const repeat = repeatedCount(column, "number-columns-repeated", MAX_REPEATED_CELLS)
    for (let index = 0; index < repeat; index += 1) {
      if (width !== undefined) columnWidths.push(width)
    }
  }

  const rawRows: RawTableRow[] = []
  for (const rowElement of childElements(element)) {
    if (localName(rowElement) !== "table-row") continue
    const repeat = repeatedCount(rowElement, "number-rows-repeated", MAX_REPEATED_ROWS)
    for (let repeatIndex = 0; repeatIndex < repeat; repeatIndex += 1) {
      const rowIndex = rawRows.length
      const row: RawTableRow = {
        cells: [],
        styleName: attribute(rowElement, "style-name"),
      }
      let column = 0
      let pendingSpanCovers = 0
      for (const cellElement of childElements(rowElement)) {
        const name = localName(cellElement)
        if (name !== "table-cell" && name !== "covered-table-cell") continue
        const repeatCount = repeatedCount(cellElement, "number-columns-repeated", MAX_REPEATED_CELLS)
        for (let cellRepeat = 0; cellRepeat < repeatCount; cellRepeat += 1) {
          if (name === "covered-table-cell") {
            // A covered cell after a horizontal span names a column the span
            // already consumed. Only a vertical cover (no pending span) owns
            // the next grid column.
            if (pendingSpanCovers > 0) pendingSpanCovers -= 1
            else column += 1
            continue
          }
          const colSpan = Math.max(1, Number(attribute(cellElement, "number-columns-spanned") ?? "1") || 1)
          const rowSpan = Math.max(1, Number(attribute(cellElement, "number-rows-spanned") ?? "1") || 1)
          const contentNodes = parseTableCellContent(cellElement, context)
          const cellStyle = resolveOdfTableCellStyle(context.sheet, attribute(cellElement, "style-name")) ?? {}
          if (colSpan > 1) cellStyle.gridSpan = colSpan
          if (rowSpan > 1) cellStyle.rowSpan = rowSpan
          row.cells.push({
            originRow: rowIndex,
            originColumn: column,
            colSpan,
            rowSpan,
            node: {
              type: "table-cell",
              nodes: contentNodes,
              ...(Object.keys(cellStyle).length > 0 ? { style: cellStyle } : {}),
            },
          })
          column += colSpan
          if (colSpan > 1) pendingSpanCovers += colSpan - 1
        }
      }
      rawRows.push(row)
    }
  }

  const origins = new Map<string, RawTableCell>()
  const continuations = new Map<string, RawTableCell>()
  const covered = new Set<string>()
  let columnCount = columnWidths.length
  for (const row of rawRows) {
    for (const cell of row.cells) {
      origins.set(`${cell.originRow}:${cell.originColumn}`, cell)
      columnCount = Math.max(columnCount, cell.originColumn + cell.colSpan)
      for (let rowOffset = 0; rowOffset < cell.rowSpan; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < cell.colSpan; columnOffset += 1) {
          if (rowOffset === 0 && columnOffset === 0) continue
          const key = `${cell.originRow + rowOffset}:${cell.originColumn + columnOffset}`
          if (rowOffset > 0 && columnOffset === 0) continuations.set(key, cell)
          else covered.add(key)
        }
      }
    }
  }
  if (columnCount < 1) columnCount = 1
  if (!tableStyle.columnWidthsTwips) {
    tableStyle.columnWidthsTwips = Array.from({ length: columnCount }, (_unused, index) =>
      columnWidths[index] ?? Math.max(1, Math.round((tableStyle.widthTwips ?? 8640) / columnCount))
    )
  }

  const rows: TableRowNode[] = rawRows.map((row, rowIndex) => {
    const cells: TableCellNode[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const key = `${rowIndex}:${column}`
      const origin = origins.get(key)
      if (origin) {
        cells.push(origin.node)
        continue
      }
      const continuation = continuations.get(key)
      if (continuation) {
        cells.push({
          type: "table-cell",
          nodes: [paragraphFromText("")],
          style: {
            ...(continuation.colSpan > 1 ? { gridSpan: continuation.colSpan } : {}),
            vMergeContinuation: true,
          },
        })
        continue
      }
      if (covered.has(key)) continue
      cells.push(emptyTableCell())
    }
    const style = resolveOdfTableRowStyle(context.sheet, row.styleName)
    return {
      type: "table-row",
      cells,
      ...(style ? { style } : {}),
    }
  })

  return {
    type: "table",
    rows,
    style: tableStyle,
  }
}

function parseBlockChildren(
  parent: Node,
  context: ParseContext,
  list: ListContext | undefined,
): DocNode[] {
  const nodes: DocNode[] = []
  for (const child of childElements(parent)) {
    const name = localName(child)
    if (name === "p" || name === "h") {
      nodes.push(paragraphFromElement(child, context, list))
    } else if (name === "list") {
      nodes.push(...listFromElement(child, context, list))
    } else if (name === "table") {
      nodes.push(tableFromElement(child, context))
    } else if (name === "section") {
      nodes.push(...parseBlockChildren(child, context, list))
    } else if (name === "soft-page-break") {
      // Soft breaks describe pagination in the source file; they are not
      // content. The layout engine repaginates the model itself.
      continue
    }
  }
  if (nodes.length === 0) {
    const text = textContent(parent).trim()
    if (text) nodes.push(paragraphFromText(text))
  }
  return nodes
}

function findTextBody(document: Document): Element | undefined {
  const body = firstDescendantElement(document, "body")
  if (!body) return undefined
  return firstChildElement(body, "text")
}

export interface BuildOpenDocumentContext {
  stylesDocument?: Document
  contentDocument: Document
  images: Map<string, OpenDocumentImageAsset>
}

export function buildOpenDocumentModelFromDocuments(
  input: BuildOpenDocumentContext,
): OpenDocumentReadResult {
  const documents = input.stylesDocument
    ? [input.contentDocument, input.stylesDocument]
    : [input.contentDocument]
  const sheet = collectOdfStyles(documents)
  const numbering: NumberingContext = {
    definitions: { abstracts: [], instances: [] },
    byStyle: new Map(),
    nextNumId: 1,
    nextAbstractId: 0,
  }
  const context: ParseContext = {
    sheet,
    images: input.images,
    numbering,
    paragraphStyles: new Map(),
    footnotes: [],
    endnotes: [],
    nextNoteId: 0,
    warnings: [],
  }

  const body = findTextBody(input.contentDocument)
  if (!body) {
    throw new Error("OpenDocument file does not contain a text body.")
  }
  const nodes = parseBlockChildren(body, context, undefined)
  const pageLayout = resolveOdfPageLayout(sheet, undefined)
  const masterPage = sheet.masterPages.values().next().value as Element | undefined
  const header = masterPage ? firstChildElement(masterPage, "header") : undefined
  const footer = masterPage ? firstChildElement(masterPage, "footer") : undefined
  const headerNodes = header ? parseBlockChildren(header, context, undefined) : []
  const footerNodes = footer ? parseBlockChildren(footer, context, undefined) : []

  const metadata: DocModel["metadata"] = {
    sourceParts: 0,
    warnings: context.warnings,
    sourcePartNames: { document: "content.xml" },
    documentPageCount: undefined,
    sectionPropertiesXml: pageLayout.sectionPropertiesXml,
    headerSections: headerNodes.length > 0
      ? [{ partName: "styles.xml/header1", referenceType: "default", nodes: headerNodes }]
      : [],
    footerSections: footerNodes.length > 0
      ? [{ partName: "styles.xml/footer1", referenceType: "default", nodes: footerNodes }]
      : [],
    paragraphStyles: [...context.paragraphStyles.values()],
    numberingDefinitions: numbering.definitions.abstracts.length > 0
      ? numbering.definitions
      : undefined,
    footnotes: context.footnotes.length > 0 ? context.footnotes : undefined,
    endnotes: context.endnotes.length > 0 ? context.endnotes : undefined,
  }

  return {
    model: {
      nodes,
      metadata,
    },
    warnings: context.warnings,
  }
}

export function parseOpenDocumentXml(
  contentXml: string,
  stylesXml?: string,
  images: Map<string, OpenDocumentImageAsset> = new Map(),
): OpenDocumentReadResult {
  const contentDocument = parseXmlDocument(contentXml)
  const stylesDocument = stylesXml ? parseXmlDocument(stylesXml) : undefined
  return buildOpenDocumentModelFromDocuments({ contentDocument, stylesDocument, images })
}
