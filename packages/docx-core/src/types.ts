/**
 * @extend-ai/docx-core — Framework-agnostic DOCX engine
 *
 * Inspired by the public @extend-ai/react-docx API; see docs/upstream-extend-ui.md for attribution.
 * All types and functions are pure JS/TS, zero framework dependencies.
 */

// ============================================================
// 1. Document Model Types
// ============================================================

export interface TextStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  color?: string
  highlight?: string
  fontSizePt?: number
  fontFamily?: string
  subscript?: boolean
  superscript?: boolean
}

export interface TextRunNode {
  type: "text"
  text: string
  style?: TextStyle
  link?: {
    url: string
    text?: string
  }
  noteReference?: string
}

export interface ImageRunNode {
  type: "image"
  src?: string
  alt?: string
  widthPx?: number
  heightPx?: number
  data?: Uint8Array
  floating?: ImageFloatingOptions
  crop?: ImageCrop
  cssFilter?: string
  mimeType?: string
}

export interface ImageFloatingOptions {
  horizontalPosition?: string | number
  verticalPosition?: string | number
  wrap?: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behindText" | "inFrontOfText"
  horizontalOffsetPx?: number
  verticalOffsetPx?: number
}

export interface ImageCrop {
  left?: number
  right?: number
  top?: number
  bottom?: number
}

export interface FormFieldRunNode {
  type: "form-field"
  fieldType: "checkbox" | "text" | "date" | "dropdown"
  id?: string
  checked?: boolean
  checkedSymbol?: string
  uncheckedSymbol?: string
  value?: string
  options?: string[]
  widget?: string
}

export type ParagraphChildNode = TextRunNode | ImageRunNode | FormFieldRunNode

export interface ParagraphStyle {
  align?: "left" | "center" | "right" | "justify"
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6
  styleId?: string
  styleName?: string
  numbering?: ParagraphNumbering
  spacing?: ParagraphSpacing
  indent?: ParagraphIndent
  borders?: ParagraphBorders
  dropCap?: DropCap
  sectionBreak?: {
    type: "nextPage" | "continuous" | "evenPage" | "oddPage"
  }
  sectionPropertiesXml?: string
  pageBreakBefore?: boolean
}

export interface ParagraphNumbering {
  level?: number
  numId?: number
}

export interface ParagraphSpacing {
  before?: {
    twips: number
    rule?: "auto" | "exact" | "atLeast"
  }
  after?: {
    twips: number
    rule?: "auto" | "exact" | "atLeast"
  }
  line?: {
    twips?: number
    multiple?: number
    rule?: "auto" | "exact" | "atLeast"
  }
}

export interface ParagraphIndent {
  left?: number
  right?: number
  firstLine?: number
  hanging?: number
}

export interface ParagraphBorders {
  top?: BorderProperties
  bottom?: BorderProperties
  left?: BorderProperties
  right?: BorderProperties
}

export interface BorderProperties {
  style?: string
  color?: string
  size?: number
}

export interface DropCap {
  lines?: number
  fontName?: string
}

export interface ParagraphNode {
  type: "paragraph"
  children: ParagraphChildNode[]
  style?: ParagraphStyle
  revisionId?: string
  revisionType?: "insert" | "delete" | "moveFrom" | "moveTo"
}

export interface TableCellStyle {
  width?: number
  vAlign?: "top" | "center" | "bottom"
  borders?: ParagraphBorders
  shading?: {
    color?: string
    fill?: string
  }
  colSpan?: number
  rowSpan?: number
}

export interface TableCellNode {
  type: "table-cell"
  nodes: (ParagraphNode | TableNode)[]
  style?: TableCellStyle
}

export interface TableRowNode {
  type: "table-row"
  cells: TableCellNode[]
}

export interface TableStyle {
  width?: number
  borders?: ParagraphBorders
  alignment?: "left" | "center" | "right"
  indent?: number
  cellSpacing?: number
}

export interface TableNode {
  type: "table"
  rows: TableRowNode[]
  style?: TableStyle
}

export type DocNode = ParagraphNode | TableNode

export interface HeaderSection {
  nodes: DocNode[]
  sectionPropertiesXml?: string
}

export interface FooterSection {
  nodes: DocNode[]
  sectionPropertiesXml?: string
}

export interface DocumentSectionInfo {
  startNodeIndex: number
  sectionPropertiesXml?: string
  headerSections: HeaderSection[]
  footerSections: FooterSection[]
}

export interface NumberingLevelDefinition {
  level: number
  format?: string
  text?: string
  alignment?: "left" | "center" | "right"
  start?: number
}

export interface NumberingDefinition {
  numId: number
  abstractNumId: number
  levels: NumberingLevelDefinition[]
}

export interface DocumentStyleDefinition {
  styleId: string
  name: string
  type: "paragraph" | "character" | "table" | "numbering"
  basedOn?: string
  next?: string
  paragraphStyle?: ParagraphStyle
  textStyle?: TextStyle
}

export interface DocumentCommentDefinition {
  id: string
  commentId: string
  author: string
  text: string
  parentId?: string
  resolved?: boolean
}

export interface DocumentNoteDefinition {
  id: string
  type: "footnote" | "endnote"
  nodes?: DocNode[]
}

export interface DocumentCompatibilitySettings {
  [key: string]: unknown
}

export interface DocModelMetadata {
  headerSections: HeaderSection[]
  footerSections: FooterSection[]
  sections?: DocumentSectionInfo[]
  styles?: DocumentStyleDefinition[]
  numberings?: NumberingDefinition[]
  comments?: DocumentCommentDefinition[]
  footnotes?: DocumentNoteDefinition[]
  endnotes?: DocumentNoteDefinition[]
  compatibilitySettings?: DocumentCompatibilitySettings
  properties?: Record<string, string>
}

export interface DocModel {
  nodes: DocNode[]
  metadata: DocModelMetadata
}

// ============================================================
// 2. Layout Engine Types
// ============================================================

export interface LayoutOptions {
  pageWidth?: number
  pageHeight?: number
  margin?: number
  minLineHeight?: number
  paragraphSpacing?: number
  tableCellPadding?: number
}

export interface LayoutBlock {
  kind: "paragraph" | "table"
  id: string
  x: number
  y: number
  width: number
  height: number
  children?: LayoutRun[]
  headingLevel?: number
  rows?: LayoutTableRow[]
}

export interface LayoutRun {
  kind: "text" | "image" | "form-field"
  id: string
  text?: string
  style?: TextStyle
  link?: { url: string; text?: string }
  src?: string
  widthPx?: number
  heightPx?: number
  floating?: ImageFloatingOptions
  fieldType?: string
  checked?: boolean
  value?: string
}

export interface LayoutTableCell {
  x: number
  y: number
  width: number
  height: number
  nodes: LayoutBlock[]
  colSpan?: number
  rowSpan?: number
}

export interface LayoutTableRow {
  height: number
  cells: LayoutTableCell[]
}

export interface LayoutPage {
  number: number
  blocks: LayoutBlock[]
}

// ============================================================
// 3. OOXML Package Types
// ============================================================

export interface OoxmlPart {
  name: string
  content: string
}

export interface OoxmlPackage {
  parts: Map<string, OoxmlPart>
  binaryAssets: Map<string, Uint8Array>
}

// ============================================================
// 4. Editor Controller Types (framework-agnostic core)
// ============================================================

export interface DocxEditorController {
  // Model
  readonly model: DocModel
  readonly fileName: string
  readonly selection: DocxSelection

  // Status
  readonly status: DocxEditorStatus
  readonly isDirty: boolean

  // Undo/Redo
  undo(): void
  redo(): void
  readonly canUndo: boolean
  readonly canRedo: boolean

  // File operations
  importDocxFile(file: File | Blob): Promise<void>
  importDocxBuffer(buffer: ArrayBuffer, fileName?: string): Promise<void>
  exportDocx(): Promise<Blob>
  updateParagraphText?(nodeIndex: number, text: string): void
  selectParagraph?(nodeIndex: number, runIndex?: number): void

  // Text formatting
  toggleBold(): void
  toggleItalic(): void
  toggleUnderline(): void
  toggleStrike(): void
  setFontSize(size: number): void
  setFontFamily(family: string): void
  setTextColor(color: string | null): void
  setHighlight(color: string | null): void
  setSuperscript(): void
  setSubscript(): void

  // Paragraph operations
  insertParagraph(): void
  removeParagraph(): void
  duplicateParagraph(): void
  setParagraphHeading(level: number | null): void
  setParagraphAlignment(align: "left" | "center" | "right" | "justify" | null): void
  setLineSpacing(info: DocxLineSpacingInfo): void
  setParagraphStyle(styleId: string): void

  // Table operations
  insertTable(rows: number, cols: number): void
  removeTable(): void
  insertTableRow(): void
  insertTableColumn(): void

  // Image operations
  insertImage(file: File): Promise<void>

  // Borders
  applyBorderPreset(preset: DocxBorderPreset): void
  readonly selectedBorderContext: DocxBorderContext | null
  readonly activeBorderPresets: DocxBorderPreset[]

  // Form fields
  setFormFieldValue(fieldId: string, value: string): void
  toggleFormCheckbox(fieldId: string): void
  selectFormField(fieldId: string | null): void
  readonly selectedFormField: FormFieldRunNode | null

  // Track changes
  readonly trackedChanges: DocxTrackedChange[]
  readonly showTrackedChanges: boolean
  readonly setShowTrackedChanges: (show: boolean) => void
  readonly toggleShowTrackedChanges: () => void

  // Comments
  readonly comments: DocxComment[]
  readonly showComments: boolean
  readonly setShowComments: (show: boolean) => void
  readonly toggleShowComments: () => void

  // Theme
  readonly documentTheme: DocxDocumentTheme
  readonly setDocumentTheme: (theme: DocxDocumentTheme) => void
  readonly toggleDocumentTheme: () => void

  // Pagination
  readonly currentPage: number
  readonly totalPages: number

  // Paragraph styles
  readonly availableParagraphStyles: ParagraphStyleDefinition[]
  readonly selectedParagraphStyleId: string | null
  readonly selectedLineSpacing: DocxLineSpacingInfo | null
}

export interface DocxSelection {
  readonly text: string
  readonly nodeIndex: number | null
  readonly runIndex: number | null
  readonly textRange: DocxTextRange | null
  readonly isCollapsed: boolean
}

export type DocxEditorStatus =
  | "idle"
  | "loading"
  | "ready"
  | "error"

export interface DocxLineSpacingInfo {
  lineRule: "auto" | "exact" | "atLeast"
  lineTwips?: number
  multiple?: number
}

export interface DocxBorderContext {
  type: "paragraph" | "table" | "table-cell"
  nodeIndex: number
}

export interface DocxBorderPreset {
  id: string
  label: string
  appliesTo: ("paragraph" | "table" | "table-cell")[]
}

export interface DocxDocumentTheme {
  mode: "light" | "dark"
  colors?: Record<string, string>
}

export interface DocxTrackedChange {
  id: string
  kind: "insertion" | "deletion" | "moveFrom" | "moveTo"
  author: string
  date: string
  text: string
  nodeIndex: number
  runIndex?: number
}

export interface DocxComment {
  id: string
  commentId: string
  author: string
  text: string
  parentId?: string
  resolved?: boolean
  anchorText?: string
  nodeIndex: number
  runIndex?: number
}

export interface ParagraphStyleDefinition {
  id: string
  name: string
  headingLevel?: number
  type: string
}

export interface DocxTextRangeBoundary {
  nodeIndex: number
  runIndex: number
  offset: number
}

export interface DocxTextRange {
  start: DocxTextRangeBoundary
  end: DocxTextRangeBoundary
}

// ============================================================
// 5. Thumbnail Types
// ============================================================

export interface DocxPageThumbnailResolutionOptions {
  pageWidth: number
  pageHeight: number
  containerWidth: number
  containerHeight: number
  gap?: number
  imagePadding?: number
  labelHeight?: number
}

export interface DocxPageThumbnailResolution {
  width: number
  height: number
  scale: number
  pixelRatio: number
}

export interface DocxPageThumbnailItem {
  pageNumber: number
  width: number
  height: number
  paint(canvas: HTMLCanvasElement): void
  canvasRef: HTMLCanvasElement | null
  renderToCanvas(canvas: HTMLCanvasElement): Promise<void>
}

export interface UseDocxPageThumbnailsOptions {
  width?: number
  gap?: number
  imagePadding?: number
  labelHeight?: number
  paddingY?: number
  buffer?: number
  autoScroll?: boolean
  scrollBehavior?: "auto" | "smooth"
}

// ============================================================
// 6. WASM Types
// ============================================================

export type WasmSource = string | URL | Request | Response | ArrayBuffer | BufferSource | WebAssembly.Module

export interface WasmInitOutput {
  readonly memory: WebAssembly.Memory
  readonly build_doc_model_from_bytes: (a: number, b: number, c: number) => void
  readonly build_doc_model_from_package: (a: number, b: number) => void
  readonly model_to_document_xml_from_json_wasm: (a: number, b: number, c: number, d: number) => void
  readonly model_to_document_xml_wasm: (a: number, b: number, c: number) => void
  readonly package_to_array_buffer_wasm: (a: number, b: number) => void
  readonly parse_docx_wasm: (a: number, b: number, c: number) => void
  readonly serialize_docx_from_json_wasm: (a: number, b: number, c: number, d: number) => void
  readonly serialize_docx_wasm: (a: number, b: number, c: number) => void
}

