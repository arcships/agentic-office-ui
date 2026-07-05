/**
 * DOCX Layout Engine — Pure functions for document pagination and layout
 *
 * Inspired by the public @extend-ai/react-docx API; see docs/upstream-extend-ui.md for attribution.
 * Framework-agnostic, zero dependencies (pure TypeScript).
 */

import type {
  DocModel, DocNode, ParagraphNode, TableNode,
  LayoutOptions, LayoutPage, LayoutBlock, LayoutRun,
  LayoutTableRow, LayoutTableCell, TextStyle, ImageFloatingOptions,
} from "./types"

// ============================================================
// Constants
// ============================================================

export const DEFAULT_LAYOUT_OPTIONS: Required<LayoutOptions> = {
  pageWidth: 816,    // US Letter @ 96dpi
  pageHeight: 1056,
  margin: 72,         // 1 inch
  minLineHeight: 22,
  paragraphSpacing: 8,
  tableCellPadding: 8,
}

export const DEFAULT_PAGE_OVERFLOW_TOLERANCE_PX = 2
export const DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX = 14

// ============================================================
// Heading Scale
// ============================================================

export function headingScale(level?: number): number {
  if (!level) return 1
  switch (level) {
    case 1: return 2.15
    case 2: return 1.75
    case 3: return 1.45
    case 4: return 1.28
    case 5: return 1.15
    case 6: return 1.05
    default: return 1
  }
}

// ============================================================
// Run Height Calculation
// ============================================================

function runHeightPx(run: LayoutRun): number {
  if (run.kind === "image") {
    if (run.floating) return 0
    return run.heightPx ?? 96
  }
  const fontSizePt = run.style?.fontSizePt ?? 12
  return Math.round(fontSizePt * 1.6)
}

function lineHeightFromRuns(runs: LayoutRun[], minLineHeight: number, headingLevel?: number): number {
  const base = runs.reduce((largest, run) => Math.max(largest, runHeightPx(run)), 12)
  return Math.max(minLineHeight, Math.round(base * headingScale(headingLevel)))
}

function spacingForBlock(baseSpacing: number, headingLevel?: number): number {
  if (!headingLevel) return baseSpacing
  return baseSpacing + Math.max(4, (7 - headingLevel) * 2)
}

// ============================================================
// Paragraph → Layout Block
// ============================================================

function formFieldDisplayText(field: import("./types").FormFieldRunNode): string {
  switch (field.fieldType) {
    case "checkbox":
      return field.checked ? field.checkedSymbol ?? "☒" : field.uncheckedSymbol ?? "☐"
    case "dropdown":
    case "date":
    case "text":
    default:
      return field.value ?? ""
  }
}

export function paragraphToLayout(
  paragraph: ParagraphNode,
  idPrefix: string,
  x: number,
  y: number,
  width: number,
  minLineHeight: number,
): LayoutBlock {
  const runs: LayoutRun[] = paragraph.children.map((child, runIndex) => {
    if (child.type === "text") {
      return {
        kind: "text",
        id: `${idPrefix}-run-${runIndex}`,
        text: child.text,
        style: child.style,
        link: child.link ? { url: child.link.url, text: child.link.text } : undefined,
      }
    }
    if (child.type === "form-field") {
      return {
        kind: "form-field",
        id: `${idPrefix}-run-${runIndex}`,
        text: formFieldDisplayText(child),
        fieldType: child.fieldType,
        checked: child.checked,
        value: child.value,
      }
    }
    // image
    return {
      kind: "image",
      id: `${idPrefix}-run-${runIndex}`,
      src: child.src,
      widthPx: child.widthPx,
      heightPx: child.heightPx,
      floating: child.floating,
      style: undefined,
    }
  })

  const height = lineHeightFromRuns(runs, minLineHeight, paragraph.style?.headingLevel)

  return {
    kind: "paragraph",
    id: idPrefix,
    x,
    y,
    width,
    height,
    children: runs,
    headingLevel: paragraph.style?.headingLevel,
  }
}

// ============================================================
// Table → Layout Block
// ============================================================

export function tableToLayout(
  table: TableNode,
  idPrefix: string,
  x: number,
  y: number,
  contentWidth: number,
  options: Required<LayoutOptions>,
): LayoutBlock {
  const { tableCellPadding, minLineHeight } = options

  const colWidths = computeColumnWidths(table, contentWidth, tableCellPadding)

  let cursorY = y
  const rows: LayoutTableRow[] = []

  for (const [rowIndex, row] of table.rows.entries()) {
    let maxRowHeight = minLineHeight * 2
    const cells: LayoutTableCell[] = []

    let cursorX = x
    for (const [cellIndex, cell] of row.cells.entries()) {
      const cellWidth = colWidths[cellIndex] ?? contentWidth / row.cells.length
      const cellContentWidth = cellWidth - tableCellPadding * 2

      const blocks: LayoutBlock[] = []
      let blockY = cursorY + tableCellPadding

      for (const [nodeIndex, node] of cell.nodes.entries()) {
        const block = node.type === "paragraph"
          ? paragraphToLayout(node, `${idPrefix}-r${rowIndex}-c${cellIndex}-b${nodeIndex}`, cursorX + tableCellPadding, blockY, cellContentWidth, minLineHeight)
          : tableToLayout(node, `${idPrefix}-r${rowIndex}-c${cellIndex}-t${nodeIndex}`, cursorX + tableCellPadding, blockY, cellContentWidth, options)
        blocks.push(block)
        blockY += block.height + (block.kind === "paragraph" ? spacingForBlock(options.paragraphSpacing, block.headingLevel) : options.paragraphSpacing + 10)
      }

      const cellHeight = blockY - cursorY + tableCellPadding
      maxRowHeight = Math.max(maxRowHeight, cellHeight)

      cells.push({
        x: cursorX,
        y: cursorY,
        width: cellWidth,
        height: cellHeight,
        nodes: blocks,
        colSpan: cell.style?.colSpan,
        rowSpan: cell.style?.rowSpan,
      })

      cursorX += cellWidth
    }

    rows.push({ height: maxRowHeight, cells })
    cursorY += maxRowHeight
  }

  const tableHeight = cursorY - y

  return {
    kind: "table",
    id: idPrefix,
    x,
    y,
    width: contentWidth,
    height: Math.max(tableHeight, options.minLineHeight * 2),
    rows,
  }
}

function computeColumnWidths(table: TableNode, contentWidth: number, padding: number): number[] {
  const colCount = table.rows.reduce((max, row) => Math.max(max, row.cells.length), 0)
  let remaining = contentWidth

  // Use cell style widths where available
  const widths: (number | null)[] = Array(colCount).fill(null)
  for (const row of table.rows) {
    for (const [i, cell] of row.cells.entries()) {
      if (widths[i] == null && cell.style?.width) {
        widths[i] = cell.style.width
        remaining -= cell.style.width
      }
    }
  }

  return widths.map(w => w ?? remaining / colCount)
}

// ============================================================
// Document → Pages (Layout Engine)
// ============================================================

export function estimateBlockHeight(block: LayoutBlock): number {
  return block.height
}

export function layoutDocument(model: DocModel, options: Partial<LayoutOptions> = {}): LayoutPage[] {
  const resolved = { ...DEFAULT_LAYOUT_OPTIONS, ...options }
  const pages: LayoutPage[] = [{ number: 1, blocks: [] }]
  const contentWidth = resolved.pageWidth - resolved.margin * 2
  const pageBottom = resolved.pageHeight - resolved.margin
  let cursorY = resolved.margin

  for (const [index, node] of model.nodes.entries()) {
    const block = node.type === "paragraph"
      ? paragraphToLayout(node, `paragraph-${index}`, resolved.margin, cursorY, contentWidth, resolved.minLineHeight)
      : tableToLayout(node, `table-${index}`, resolved.margin, cursorY, contentWidth, resolved)

    const blockHeight = estimateBlockHeight(block)
    const currentPage = pages[pages.length - 1]

    // Page break: if block doesn't fit and current page already has content
    if (cursorY + blockHeight > pageBottom && currentPage.blocks.length > 0) {
      pages.push({ number: pages.length + 1, blocks: [] })
      cursorY = resolved.margin
    }

    // Update block y position
    if (block.kind === "paragraph") block.y = cursorY
    else block.y = cursorY

    pages[pages.length - 1].blocks.push(block)

    cursorY += blockHeight + (
      block.kind === "paragraph"
        ? spacingForBlock(resolved.paragraphSpacing, block.headingLevel)
        : resolved.paragraphSpacing + 10
    )
  }

  return pages
}

// ============================================================
// Paragraph Spacing Analysis
// ============================================================

export function paragraphBeforeSpacingPx(paragraph: ParagraphNode): number {
  const spacing = paragraph.style?.spacing?.before
  if (!spacing) return 0
  const twips = spacing.twips ?? 0
  return Math.round(twips / 20) // twips → px @ 96dpi
}

export function paragraphAfterSpacingPx(paragraph: ParagraphNode): number {
  const spacing = paragraph.style?.spacing?.after
  if (!spacing) return 0
  const twips = spacing.twips ?? 0
  return Math.round(twips / 20)
}

export function resolveParagraphBeforeSpacingPx(
  _model: DocModel,
  _nodeIndex: number,
  paragraph: ParagraphNode,
  pageConsumedHeightPx: number,
  suppressSpacingBeforeAfterPageBreak: boolean,
): number {
  // At top of page, suppress spacing based on flag
  if (pageConsumedHeightPx === 0 && suppressSpacingBeforeAfterPageBreak) {
    return 0
  }
  return paragraphBeforeSpacingPx(paragraph)
}

// ============================================================
// Page Break Detection
// ============================================================

export function paragraphHasExplicitPageBreak(paragraph: ParagraphNode): boolean {
  // Check for page break character in children
  return paragraph.children.some(
    child => child.type === "text" && child.text.includes("\f")
  )
}

export function paragraphHasPageBreakBefore(paragraph: ParagraphNode): boolean {
  return paragraph.style?.pageBreakBefore === true
}

export function sectionBreakAfterParagraphStartsNewPage(paragraph: ParagraphNode): boolean {
  const sectionBreak = paragraph.style?.sectionBreak
  if (!sectionBreak) return false
  return sectionBreak.type === "nextPage" || sectionBreak.type === "evenPage" || sectionBreak.type === "oddPage"
}

export function sectionBreakPropertiesStartNewPage(sectionPropertiesXml?: string): boolean {
  if (!sectionPropertiesXml) return false
  return sectionPropertiesXml.includes('w:type="nextPage"') ||
    sectionPropertiesXml.includes('w:type="evenPage"') ||
    sectionPropertiesXml.includes('w:type="oddPage"')
}

export function sectionTitlePageEnabled(sectionPropertiesXml?: string): boolean {
  if (!sectionPropertiesXml) return false
  return sectionPropertiesXml.includes("<w:titlePg")
}

// ============================================================
// Hard Page Break Collection
// ============================================================

export function collectDocxHardPageBreakStartNodeIndexes(model: DocModel): Set<number> {
  const breakIndexes = new Set<number>()

  for (const [index, node] of model.nodes.entries()) {
    if (node.type === "paragraph") {
      if (paragraphHasExplicitPageBreak(node) || paragraphHasPageBreakBefore(node)) {
        breakIndexes.add(index)
      }
    }
  }

  return breakIndexes
}

export function collectTopLevelExplicitPageBreakStartNodeIndexes(nodes: DocNode[]): Set<number> {
  const breakIndexes = new Set<number>()

  for (const [index, node] of nodes.entries()) {
    if (node.type === "paragraph") {
      if (paragraphHasExplicitPageBreak(node)) {
        breakIndexes.add(index)
      }
    }
  }

  return breakIndexes
}

// ============================================================
// Document Section Resolution
// ============================================================

export function resolveSectionIndexForNodeIndex(
  sections: { startNodeIndex: number }[],
  nodeIndex: number,
  previousSectionIndex: number = 0,
): number {
  // Binary search for the section that contains this node
  let lo = previousSectionIndex
  let hi = sections.length - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (sections[mid].startNodeIndex <= nodeIndex) {
      lo = mid + 1
    } else {
      hi = mid - 1
    }
  }
  return Math.max(0, hi)
}

// ============================================================
// Section Properties XML Resolution
// ============================================================

export function resolveSectionPropertiesXmlForNodeIndex(
  sections: { startNodeIndex: number; sectionPropertiesXml?: string }[],
  nodeIndex: number,
  fallback?: string,
): string | undefined {
  const idx = resolveSectionIndexForNodeIndex(sections, nodeIndex)
  return sections[idx]?.sectionPropertiesXml ?? fallback
}

// ============================================================
// Document Layout Metrics Resolution
// ============================================================

export interface DocumentLayoutMetrics {
  pageWidth: number
  pageHeight: number
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
  columns?: { count: number; space: number }[]
}

export function parseSectionLayout(sectionPropertiesXml?: string): DocumentLayoutMetrics {
  const metrics: DocumentLayoutMetrics = {
    pageWidth: DEFAULT_LAYOUT_OPTIONS.pageWidth,
    pageHeight: DEFAULT_LAYOUT_OPTIONS.pageHeight,
    marginTop: DEFAULT_LAYOUT_OPTIONS.margin,
    marginBottom: DEFAULT_LAYOUT_OPTIONS.margin,
    marginLeft: DEFAULT_LAYOUT_OPTIONS.margin,
    marginRight: DEFAULT_LAYOUT_OPTIONS.margin,
  }

  if (!sectionPropertiesXml) return metrics

  // Parse <w:pgSz w:w="12240" w:h="15840"/>
  const pgSzMatch = sectionPropertiesXml.match(/<w:pgSz[^>]*w:w="(\d+)"[^>]*w:h="(\d+)"/)
  if (pgSzMatch) {
    metrics.pageWidth = Math.round(parseInt(pgSzMatch[1]) / 20)
    metrics.pageHeight = Math.round(parseInt(pgSzMatch[2]) / 20)
  }

  // Parse <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>
  const pgMarMatch = sectionPropertiesXml.match(
    /<w:pgMar[^>]*w:top="(\d+)"[^>]*w:bottom="(\d+)"[^>]*w:left="(\d+)"[^>]*w:right="(\d+)"/
  )
  if (pgMarMatch) {
    metrics.marginTop = Math.round(parseInt(pgMarMatch[1]) / 20)
    metrics.marginBottom = Math.round(parseInt(pgMarMatch[2]) / 20)
    metrics.marginLeft = Math.round(parseInt(pgMarMatch[3]) / 20)
    metrics.marginRight = Math.round(parseInt(pgMarMatch[4]) / 20)
  }

  // Parse <w:cols w:num="2" w:space="720"/>
  const colsMatch = sectionPropertiesXml.match(/<w:cols[^>]*w:num="(\d+)"(?:[^>]*w:space="(\d+)")?/)
  if (colsMatch) {
    metrics.columns = [{
      count: parseInt(colsMatch[1]),
      space: colsMatch[2] ? Math.round(parseInt(colsMatch[2]) / 20) : 720,
    }]
  }

  return metrics
}

export function resolveDocumentLayout(model: DocModel): DocumentLayoutMetrics {
  const fallback = model.metadata.sections?.[0]?.sectionPropertiesXml
  // Default section (before any explicit sections)
  if (model.metadata.headerSections[0]?.sectionPropertiesXml) {
    return parseSectionLayout(model.metadata.headerSections[0].sectionPropertiesXml)
  }
  return parseSectionLayout(fallback)
}

// ============================================================
// Thumbnail Resolution
// ============================================================

export function resolveDocxPageThumbnailResolution(
  options: { pageWidth: number; pageHeight: number; containerWidth: number; containerHeight: number; gap?: number; imagePadding?: number; labelHeight?: number }
): { width: number; height: number; scale: number; pixelRatio: number } {
  const { pageWidth, pageHeight, containerWidth, containerHeight } = options
  const imagePadding = options.imagePadding ?? 4
  const labelHeight = options.labelHeight ?? 16

  const maxThumbWidth = containerWidth - imagePadding * 2
  const maxThumbHeight = containerHeight - imagePadding * 2 - labelHeight

  const scaleX = maxThumbWidth / pageWidth
  const scaleY = maxThumbHeight / pageHeight
  const scale = Math.min(scaleX, scaleY)

  return {
    width: Math.round(pageWidth * scale),
    height: Math.round(pageHeight * scale),
    scale,
    pixelRatio: typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1,
  }
}

