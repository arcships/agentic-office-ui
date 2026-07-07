// docx-core — upstream DOCX engine port (modular remigration in progress).
//
// Migration source: @extend-ai/react-docx @ commit 6f70b92
// Upstream path: /Users/eric8810/Code/extend-ui-upstream/react-docx/packages
//
// Unlike the previous mechanical copy, this port rebuilds the engine
// module-by-module with sane file boundaries. Each module is ported from
// upstream and verified independently before the next one lands.

// Stage 0: minimal DocModel type so downstream packages (vue-docx, demo)
// can typecheck while the engine is being rebuilt.
export type DocModel = {
  nodes: DocNode[]
  metadata: DocModelMetadata
}

export type DocNode = ParagraphNode | TableNode

export interface ParagraphNode {
  type: "paragraph"
  children: ParagraphChildNode[]
  style?: ParagraphStyle
  sourceXml?: string
}

export interface TableNode {
  type: "table"
  rows: TableRowNode[]
  style?: TableStyle
  sourceXml?: string
}

export type ParagraphChildNode = TextRunNode | ImageRunNode | FormFieldRunNode

export interface TextRunNode {
  type: "text"
  text: string
  style?: TextStyle
  sourceXml?: string
}

export interface ImageRunNode {
  type: "image"
  src: string
  alt?: string
  widthPx?: number
  heightPx?: number
  data: Uint8Array
  mimeType: string
}

export interface FormFieldRunNode {
  type: "form-field"
  fieldType: "checkbox" | "text" | "date" | "dropdown"
  // detailed shape filled in during editor-ops port
  [key: string]: unknown
}

export interface TextStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  superscript?: boolean
  subscript?: boolean
  color?: string
  highlight?: string
  fontFamily?: string
  fontSizePt?: number
  linkHref?: string
}

export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6

export type ParagraphAlignment = "left" | "center" | "right" | "justify"

export interface ParagraphStyle {
  headingLevel?: HeadingLevel
  styleId?: string
  alignment?: ParagraphAlignment
  lineHeight?: number
  spaceBeforeTwips?: number
  spaceAfterTwips?: number
  indentTwips?: number
}

export interface ParagraphStyleDefinition {
  styleId: string
  name: string
  basedOn?: string
  headingLevel?: HeadingLevel
  alignment?: ParagraphAlignment
  [key: string]: unknown
}

export interface TableStyle {
  [key: string]: unknown
}

export interface TableRowNode {
  type: "table-row"
  cells: TableCellNode[]
}

export type TableCellContentNode = ParagraphNode | TableNode

export interface TableCellNode {
  type: "table-cell"
  nodes: TableCellContentNode[]
}

export interface DocModelMetadata {
  sections?: unknown[]
  headerSections?: unknown[]
  footerSections?: unknown[]
  paragraphStyles?: ParagraphStyleDefinition[]
  numberingDefinitions?: unknown[]
  footnotes?: unknown[]
  endnotes?: unknown[]
  comments?: unknown[]
  compatibility?: Record<string, unknown>
}
