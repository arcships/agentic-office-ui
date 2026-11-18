// Public prop types for the block/word highlight overlays.

import type {
  DocxBlockGeometry,
  DocxWordRange,
  DocumentBlock,
  DocumentWord,
} from "@arcships/docx-core"

export interface DocxBlockHighlightsProps {
  /** Viewer root that contains `[data-docx-page-surface]` elements. */
  root?: HTMLElement | null
  /** Emphasized block (stronger border/background). */
  block?: DocumentBlock
  /** Blocks to outline. */
  blocks?: readonly DocumentBlock[]
  /** Selects blocks by id; ids double as source ids without `blocks`. */
  blockIds?: readonly string[]
  /** Only render fragments on this zero-based page. */
  pageIndex?: number
  /** Border and tint color, as a CSS color with a `#rrggbb` readable alpha suffix. */
  color?: string
  colorForBlock?: (block: DocumentBlock) => string
  onGeometry?: (geometries: Map<string, DocxBlockGeometry>) => void
}

export interface DocxWordHighlightsProps {
  root?: HTMLElement | null
  /** Emphasized word. */
  word?: DocumentWord
  /** Words to outline. */
  words?: readonly DocumentWord[]
  /** Offsets to highlight without running word extraction. */
  wordRanges?: readonly DocxWordRange[]
  /** Only render fragments on this zero-based page. */
  pageIndex?: number
  color?: string
  onGeometry?: (geometries: Map<string, DocxBlockGeometry>) => void
}
