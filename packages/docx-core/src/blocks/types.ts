// Document block and word model — the addressable layer between a DOCX model
// and the geometry the viewer renders.
//
// Upstream: @extend-ai/react-docx `doc-model/src/blocks.ts` and
// `react-viewer/src/block-geometry.tsx`. The upstream parser lives in Rust/WASM;
// this port derives the same contract from the local `DocModel` so the browser
// bundle needs no extra parser build.

/**
 * Document block type names. These are the Extend parser's public type names;
 * the local extractor emits the subset the DOCX model can prove and keeps the
 * full enum so consumers can switch without version checks.
 */
export const DOCUMENT_BLOCK_TYPES = [
  "text",
  "heading",
  "section_heading",
  "figure",
  "table",
  "table_head",
  "table_cell",
  "key_value",
  "page_number",
  "barcode",
  "formula",
  "header",
  "footer",
  "watermark",
  "legend",
] as const;

export type DocumentBlockType = (typeof DOCUMENT_BLOCK_TYPES)[number];

/** Story a source reference belongs to. */
export type DocxBlockStory =
  | "body"
  | "header"
  | "footer"
  | "footnote"
  | "endnote";

/**
 * A model path segment. Structurally identical to the unified reference
 * contract's `DocxPathSegment`, kept local so the block model stays usable
 * without importing the interaction package.
 */
export interface DocxBlockPathSegment {
  kind: "node" | "table" | "row" | "cell" | "paragraph" | "run";
  index: number;
  nativeId?: string;
}

/** Part locator for a story, mirroring the unified reference contract. */
export type DocxBlockPart =
  | { kind: "body" }
  | { kind: "header" | "footer"; partName: string }
  | { kind: "footnote" | "endnote"; noteId: string };

/**
 * Source identity of one occurrence of a block (or word range) in the original
 * document. `nodeId` is the same value the viewer writes to
 * `data-docx-source-id`, so geometry resolution can find the rendered element.
 */
export interface DocxSourceRef {
  documentId: string;
  nodeId: string;
  partName: string;
  story: DocxBlockStory;
  /** Inclusive UTF-16 offset into the source paragraph (word ranges only). */
  start?: number;
  /** Exclusive UTF-16 offset into the source paragraph (word ranges only). */
  end?: number;
  offsetEncoding: "utf-16";
}

export interface ExtractDocumentBlocksOptions {
  /**
   * Also segment every paragraph into words. Opt-in to match the upstream
   * parser contract; results without it keep no `words` array.
   */
  includeWords?: boolean;
  /** Reserved for 3D model content; the local model has none. */
  include3dModels?: boolean;
  /** Identity persisted in every source reference. Defaults to `"docx-document"`. */
  documentId?: string;
  /** Include header/footer story containers. Defaults to true. */
  includeHeadersAndFooters?: boolean;
  /** Include footnote/endnote story paragraphs. Defaults to true. */
  includeNotes?: boolean;
}

export interface DocumentWord {
  object: "word";
  id: string;
  blockId: string;
  content: string;
  sourceRefs: DocxSourceRef[];
}

export type DocumentBlockDetails = Record<string, unknown> & {
  /**
   * Structural discriminator. `"paragraph"` marks text blocks whose geometry is
   * measured from visible text ranges; `"node"` marks a bare source id.
   */
  kind?:
    | "paragraph"
    | "table"
    | "table-row"
    | "table-cell"
    | "figure"
    | "header"
    | "footer"
    | "node"
    | "word";
  /** Top-level model node index. */
  nodeIndex?: number;
  /** Table this block belongs to (cells, header rows). */
  tableIndex?: number;
  /** Row index inside the owning table. */
  rowIndex?: number;
  /** Cell index inside the owning row. */
  cellIndex?: number;
  /** Image child index inside the owning paragraph. */
  childIndex?: number;
  /** Paragraph index inside the owning story/cell. */
  paragraphIndex?: number;
  /** Story part and model path used by reference drafts. */
  part?: DocxBlockPart;
  modelPath?: readonly DocxBlockPathSegment[];
  headingLevel?: number;
  style?: {
    styleId?: string;
    styleName?: string;
    headingLevel?: number;
  };
  list?: {
    numId: number;
    level: number;
    format?: string;
  };
  formField?: {
    type: string;
    value?: string;
  };
  dropCap?: boolean;
  rowCount?: number;
  columnCount?: number;
  isHeaderRow?: boolean;
  alt?: string;
};

export interface DocumentBlock {
  object: "block";
  id: string;
  type: DocumentBlockType;
  content: string;
  details: DocumentBlockDetails;
  sourceRefs: DocxSourceRef[];
  parentBlockId?: string;
  children: DocumentBlock[];
  words?: DocumentWord[];
}

export interface StructuredDocument {
  schemaVersion: 1;
  parserVersion: string;
  documentId: string;
  blocks: DocumentBlock[];
  warnings: string[];
  geometryStatus: "requires_layout";
}

export interface BlockBoundingBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** One page-space rectangle for a source block or word range. */
export interface DocxBlockFragment {
  nodeId: string;
  pageIndex: number;
  pageWidth: number;
  pageHeight: number;
  boundingBox: BlockBoundingBox;
  polygon: Array<{ x: number; y: number }>;
  coordinateSpace: "page-css-px";
}

export interface DocxBlockGeometry {
  status:
    | "rendered"
    | "partially_rendered"
    | "not_mounted"
    | "not_visible"
    | "unresolved";
  scope: "mounted_pages";
  issues: Array<{
    nodeId: string;
    reason:
      | "unsupported_range"
      | "document_mismatch"
      | "not_on_page"
      | "range_text_mismatch";
  }>;
  fragments: DocxBlockFragment[];
  unresolvedSourceIds: string[];
}

/** A word selector for `DocxWordHighlights`, independent of extraction. */
export interface DocxWordRange {
  /** Paragraph/block source id the offsets belong to. */
  blockId: string;
  /** Inclusive UTF-16 offset into the paragraph text. */
  start: number;
  /** Exclusive UTF-16 offset into the paragraph text. */
  end: number;
  /** Expected rendered text; a mismatch is reported instead of a box. */
  content?: string;
  /** Optional stable id; defaults to `<blockId>/word/<start>-<end>`. */
  id?: string;
}
