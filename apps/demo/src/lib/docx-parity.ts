// Shared shape/compare logic for the DOCX viewer/editor parity page.
//
// The page used to report PASS whenever both sides normalized to an empty
// structure: `pageCount: 0` on both sides compared equal. It also capped its
// virtual-scroll scan at 64 steps, so a long document could report "equal"
// after only part of it had been mounted and compared. Both are acceptance
// bugs, so the checks live here and the page only supplies the DOM scan.

export type ImageShape = { alt: string; sourceKind: string }
export type CellShape = { text: string; images: ImageShape[] }
export type NodeShape =
  | { type: "paragraph"; text: string; images: ImageShape[] }
  | { type: "table"; text: string; rows: CellShape[][]; images: ImageShape[] }
export type SectionShape = { text: string; images: ImageShape[] }
export type PageShape = {
  index: number
  nodes: NodeShape[]
  text: string
  tables: number
  cells: number
  images: number
  header: SectionShape
  footer: SectionShape
}
export type InteractionShape = {
  contentEditableTrue: number
  inputs: number
  selects: number
  resizeHandles: number
}
export type RenderStructure = {
  pageCount: number
  /** What the surface declared it was rendering; 0 when nothing was declared. */
  declaredPageCount: number
  /** Virtual-scroll steps the page needed to mount every page. */
  scannedSteps: number
  pages: PageShape[]
  totals: { nodes: number; paragraphs: number; tables: number; cells: number; images: number }
  interaction: InteractionShape
}
export type ParityDiff = { equal: boolean; differences: string[] }

export const EMPTY_INTERACTION = (): InteractionShape => ({
  contentEditableTrue: 0,
  inputs: 0,
  selects: 0,
  resizeHandles: 0,
})

export const EMPTY_STRUCTURE = (): RenderStructure => ({
  pageCount: 0,
  declaredPageCount: 0,
  scannedSteps: 0,
  pages: [],
  totals: { nodes: 0, paragraphs: 0, tables: 0, cells: 0, images: 0 },
  interaction: EMPTY_INTERACTION(),
})

export function buildRenderStructure(
  pages: PageShape[],
  interaction: InteractionShape,
  meta: { declaredPageCount: number; scannedSteps: number },
): RenderStructure {
  const orderedPages = [...pages].sort((left, right) => left.index - right.index)
  const nodes = orderedPages.flatMap((page) => page.nodes)
  return {
    pageCount: orderedPages.length,
    declaredPageCount: Number.isFinite(meta.declaredPageCount)
      ? Math.max(0, Math.floor(meta.declaredPageCount))
      : 0,
    scannedSteps: Math.max(0, Math.floor(meta.scannedSteps)),
    pages: orderedPages,
    totals: {
      nodes: nodes.length,
      paragraphs: nodes.filter((node) => node.type === "paragraph").length,
      tables: nodes.filter((node) => node.type === "table").length,
      cells: nodes.reduce(
        (total, node) => total + (node.type === "table" ? node.rows.flat().length : 0),
        0,
      ),
      images: orderedPages.reduce((total, page) => total + page.images, 0),
    },
    interaction,
  }
}

function deepDifferences(left: unknown, right: unknown): string[] {
  const differences: string[] = []
  const visit = (leftValue: unknown, rightValue: unknown, path: string): void => {
    if (differences.length >= 200 || Object.is(leftValue, rightValue)) return
    if (Array.isArray(leftValue) || Array.isArray(rightValue)) {
      if (!Array.isArray(leftValue) || !Array.isArray(rightValue)) {
        differences.push(`${path}: 类型不同`)
        return
      }
      if (leftValue.length !== rightValue.length) {
        differences.push(`${path}.length: ${leftValue.length} != ${rightValue.length}`)
      }
      const length = Math.max(leftValue.length, rightValue.length)
      for (let index = 0; index < length; index += 1) {
        visit(leftValue[index], rightValue[index], `${path}[${index}]`)
      }
      return
    }
    if (
      typeof leftValue === "object" && leftValue !== null &&
      typeof rightValue === "object" && rightValue !== null
    ) {
      const keys = new Set([
        ...Object.keys(leftValue as Record<string, unknown>),
        ...Object.keys(rightValue as Record<string, unknown>),
      ])
      for (const key of keys) {
        visit(
          (leftValue as Record<string, unknown>)[key],
          (rightValue as Record<string, unknown>)[key],
          `${path}.${key}`,
        )
      }
      return
    }
    differences.push(`${path}: ${JSON.stringify(leftValue)} != ${JSON.stringify(rightValue)}`)
  }
  visit(left, right, "structure")
  return differences
}

/** Fields that describe the rendered document, not how it was collected. */
function comparable(structure: RenderStructure) {
  return {
    pageCount: structure.pageCount,
    pages: structure.pages,
    totals: structure.totals,
    interaction: structure.interaction,
  }
}

export function compareRenderStructures(
  viewer: RenderStructure,
  editor: RenderStructure,
): ParityDiff {
  const differences: string[] = []
  if (viewer.pageCount === 0) {
    differences.push("Viewer 未渲染任何页面（0 页），空结构不能判为一致")
  }
  if (editor.pageCount === 0) {
    differences.push("Editor 未渲染任何页面（0 页），空结构不能判为一致")
  }
  if (viewer.declaredPageCount > 0 && viewer.pageCount < viewer.declaredPageCount) {
    differences.push(
      `Viewer 只挂载了 ${viewer.pageCount}/${viewer.declaredPageCount} 页，未完成的结构不能判为一致`,
    )
  }
  if (editor.declaredPageCount > 0 && editor.pageCount < editor.declaredPageCount) {
    differences.push(
      `Editor 只挂载了 ${editor.pageCount}/${editor.declaredPageCount} 页，未完成的结构不能判为一致`,
    )
  }
  if (viewer.pageCount !== editor.pageCount) {
    differences.push(`pageCount: ${viewer.pageCount} != ${editor.pageCount}`)
  }
  differences.push(...deepDifferences(comparable(viewer), comparable(editor)))
  return { equal: differences.length === 0, differences }
}
