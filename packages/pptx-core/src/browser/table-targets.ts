export interface PptxRenderedTableCell {
  element: HTMLElement
  /** Zero-based table grid row. */
  row: number
  /** Zero-based table grid column, honouring merged cells. */
  column: number
  /** Paragraphs that belong to this cell, in document order. */
  paragraphs: readonly HTMLElement[]
}

const PARAGRAPH_SELECTOR = "[data-pptx-paragraph-index]"

function tableOf(element: HTMLElement): HTMLTableElement | undefined {
  if (element.localName === "table") return element as HTMLTableElement
  if (typeof element.querySelector !== "function") return undefined
  return element.querySelector<HTMLTableElement>("table") ?? undefined
}

function numericAttribute(element: HTMLElement, names: readonly string[]): number | undefined {
  for (const name of names) {
    const raw = element.getAttribute(name)
    if (raw === null || raw === "") continue
    const value = Number(raw)
    if (Number.isFinite(value) && value > 0) return Math.trunc(value)
  }
  return undefined
}

function spanValue(element: HTMLElement, property: "colSpan" | "rowSpan", attributes: readonly string[]): number {
  const propertyValue = (element as unknown as Record<string, unknown>)[property]
  const numericProperty = typeof propertyValue === "number" && Number.isFinite(propertyValue) && propertyValue > 0
    ? Math.trunc(propertyValue)
    : undefined
  const attributeValue = numericAttribute(element, attributes)
  if (attributeValue !== undefined && (numericProperty === undefined || numericProperty === 1)) {
    return attributeValue
  }
  return numericProperty ?? 1
}

function ownsTableRow(row: HTMLElement, table: HTMLElement): boolean {
  return typeof row.closest !== "function" || row.closest("table") === table
}

function ownsTableCell(cell: HTMLElement, row: HTMLElement, table: HTMLElement): boolean {
  const inside = typeof cell.closest !== "function" || cell.closest("table") === table
  return cell.parentElement === row && inside
}

function mergeFlag(element: HTMLElement, localName: string): boolean {
  // DrawingML table continuation cells are not rendered, but a host-built DOM
  // may keep them; treat both the attributes and the renderer's data markers.
  return element.localName === localName
    || element.dataset.pptxMerge === localName
}

function paragraphsInCell(cell: HTMLElement, table: HTMLElement): readonly HTMLElement[] {
  return [...cell.querySelectorAll<HTMLElement>(PARAGRAPH_SELECTOR)]
    .filter((paragraph) => typeof paragraph.closest !== "function" || paragraph.closest("table") === table)
}

/**
 * Enumerates rendered table cells with their real grid position and the
 * paragraphs each cell owns. Row-span and col-span cells are expanded so a
 * row/column target addresses the visual grid PowerPoint uses, and continuation
 * cells from a host-built DOM are skipped instead of shifting later columns.
 */
export function pptxRenderedTableCells(object: HTMLElement): readonly PptxRenderedTableCell[] {
  const table = tableOf(object)
  if (!table) return Object.freeze([])
  // A table nested in a group belongs to the table object itself, not the group.
  const owner = typeof table.closest === "function" ? table.closest("[data-pptx-object-key]") : null
  if (owner && owner !== object) return Object.freeze([])
  const rows = [...table.querySelectorAll<HTMLElement>("tr")]
    .filter((row) => ownsTableRow(row, table))
  const result: PptxRenderedTableCell[] = []
  const occupied = new Set<string>()
  rows.forEach((row, rowIndex) => {
    let column = 0
    for (const cell of [...row.children] as HTMLElement[]) {
      if (cell.localName !== "td" && cell.localName !== "th") continue
      if (!ownsTableCell(cell, row, table)) continue
      if (mergeFlag(cell, "hmerge")) continue
      const columnSpan = spanValue(cell, "colSpan", ["gridSpan", "gridspan", "colSpan", "colspan"])
      if (mergeFlag(cell, "vmerge")) {
        column += columnSpan
        continue
      }
      while (occupied.has(`${rowIndex}:${column}`)) column += 1
      const rowSpan = spanValue(cell, "rowSpan", ["rowSpan", "rowspan"])
      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < columnSpan; columnOffset += 1) {
          occupied.add(`${rowIndex + rowOffset}:${column + columnOffset}`)
        }
      }
      result.push(Object.freeze({
        element: cell,
        row: rowIndex,
        column,
        paragraphs: Object.freeze(paragraphsInCell(cell, table)),
      }))
      column += columnSpan
    }
  })
  return Object.freeze(result)
}

/** Paragraphs of a table in row-major cell order, which is `pRg` order. */
export function pptxTableParagraphs(object: HTMLElement): readonly HTMLElement[] {
  const cells = pptxRenderedTableCells(object)
  if (!cells.length) return Object.freeze([])
  return Object.freeze(cells.flatMap((cell) => cell.paragraphs))
}
