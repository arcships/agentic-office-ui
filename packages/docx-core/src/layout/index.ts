import type {
  DocModel,
  HeadingLevel,
  ParagraphAlignment,
  TextStyle
} from "../engine/types";
import {
  layoutDocument,
  paragraphToLayout,
  tableToLayout,
  type LayoutOptions,
  type LayoutPage,
  type LayoutParagraphBlock,
  type LayoutRun,
  type LayoutTableBlock,
  type LayoutBlock,
} from "./layout-engine";
import {
  paginateDocumentLayout,
  paginateDocumentLayoutWithNoteReserve,
  resolvePaginationSectionGeometry,
  type PaginatedPageBlock,
  type PaginationMeasurementFeedback,
  type PaginationSectionGeometry
} from "./pagination-layout";
import type { ParagraphLineRange, TableRowRange } from "./page-segmentation-core";
import {
  resolvePaginationPageCountTarget,
  type PaginationPageCountTarget
} from "./pagination-target";
export type {
  LayoutOptions,
  LayoutPage,
  LayoutParagraphBlock,
  LayoutRun,
  LayoutTableBlock,
  LayoutBlock,
};
export { layoutDocument };
export type { PaginationMeasurementFeedback } from "./pagination-layout";
export {
  paginateDocumentLayout,
  resolvePaginationSectionGeometry
} from "./pagination-layout";
export type {
  PaginatedDocumentLayout,
  PaginatedPage,
  PaginatedPageBlock,
  PaginationSectionGeometry
} from "./pagination-layout";
export * from "./pagination";
export * from "./pagination-iteration";
export * from "./pagination-stability";
export * from "./pagination-target";
export * from "./page-segmentation-core";
export * from "./page-segmentation-table";
export * from "./section-page-numbers";
// `pagination-layout` is intentionally not re-exported: it is the adapter that
// wires `buildDocumentPageNodeSegments` into `buildLayoutSnapshot`, not new
// public API surface.

const DEFAULT_LAYOUT_OPTIONS: Required<LayoutOptions> = {
  pageWidth: 816,
  pageHeight: 1056,
  margin: 72,
  minLineHeight: 22,
  paragraphSpacing: 8,
  tableCellPadding: 8,
  respectSectionLayout: true
};

export interface LayoutSnapshotOptions extends LayoutOptions {
  /**
   * Reserve the height of the notes printed at the bottom of a page so body
   * text does not run into them (default: true).
   */
  reserveNoteHeight?: boolean;
  /** Measured-height feedback from the previous render pass (#99). */
  measurementFeedback?: PaginationMeasurementFeedback;
  /**
   * Advisory page-count target. Defaults to `model.metadata.documentPageCount`.
   * It is reported on the snapshot but never forces page count (#99).
   */
  pageCountTarget?: number;
}

export interface LayoutSnapshotRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface LayoutSnapshotPageSize {
  width: number;
  height: number;
}

export interface LayoutSnapshotMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface ResolvedLayoutMetrics {
  pageSizePx: LayoutSnapshotPageSize;
  marginsPx: LayoutSnapshotMargins;
  contentBoxPx: LayoutSnapshotRect;
}

export interface ResolvedDocumentSection {
  index: number;
  startNodeIndex: number;
  synthetic: boolean;
  layout: ResolvedLayoutMetrics;
}

export interface ResolvedDocument {
  source: "layout-engine-adapter";
  nodeCount: number;
  pageCountHint?: number;
  layout: ResolvedLayoutMetrics;
  sections: ResolvedDocumentSection[];
  metadata: {
    sourceParts: number;
    warningCount: number;
    warnings: string[];
    headerSectionCount: number;
    footerSectionCount: number;
  };
}

export interface LayoutFragmentSource {
  kind: "paragraph" | "table" | "table-cell" | "table-cell-paragraph";
  nodeIndex: number;
  rowIndex?: number;
  cellIndex?: number;
  paragraphIndex?: number;
  /** Body column this fragment is rendered in (0 for single-column pages). */
  columnIndex?: number;
}

export interface LayoutSnapshotTextRun {
  kind: "text";
  id: string;
  text: string;
  style?: TextStyle;
  link?: string;
}

export interface LayoutSnapshotImageRun {
  kind: "image";
  id: string;
  src?: string;
  alt?: string;
  widthPx?: number;
  heightPx?: number;
  floating?: boolean;
}

export type LayoutSnapshotRun = LayoutSnapshotTextRun | LayoutSnapshotImageRun;

export interface LayoutSnapshotParagraphFragment {
  kind: "paragraph";
  id: string;
  framePx: LayoutSnapshotRect;
  align: ParagraphAlignment;
  headingLevel?: HeadingLevel;
  runs: LayoutSnapshotRun[];
  /** Set when the paragraph is split across pages by the pagination engine. */
  paragraphLineRange?: ParagraphLineRange;
  source?: LayoutFragmentSource;
}

export interface LayoutSnapshotTableCellFragment {
  id: string;
  colSpan?: number;
  backgroundColor?: string;
  paragraphs: LayoutSnapshotParagraphFragment[];
  source?: LayoutFragmentSource;
}

export interface LayoutSnapshotTableRowFragment {
  id: string;
  backgroundColor?: string;
  cells: LayoutSnapshotTableCellFragment[];
}

export interface LayoutSnapshotTableFragment {
  kind: "table";
  id: string;
  framePx: LayoutSnapshotRect;
  rows: LayoutSnapshotTableRowFragment[];
  /** Set when the table rows are split across pages by the pagination engine. */
  tableRowRange?: TableRowRange;
  source?: LayoutFragmentSource;
}

export type LayoutSnapshotBlockFragment =
  | LayoutSnapshotParagraphFragment
  | LayoutSnapshotTableFragment;

export interface LayoutSnapshotPage {
  index: number;
  number: number;
  pageSizePx: LayoutSnapshotPageSize;
  marginsPx: LayoutSnapshotMargins;
  contentBoxPx: LayoutSnapshotRect;
  /** Real body column layout for this page's section. */
  columnCount: number;
  columnGapPx: number;
  columnWidthsPx: number[];
  blocks: LayoutSnapshotBlockFragment[];
}

export interface LayoutSnapshot {
  version: 1;
  source: "layout-engine-adapter";
  resolvedDocument: ResolvedDocument;
  /**
   * Advisory page-count target from `documentPageCount` (#99). It is never
   * enforced: neither page heights nor page boundaries are scaled/merged to
   * reach it. A mismatch means the cached count is stale or the measurement
   * differs, and the rendered count wins.
   */
  pageCountTarget?: PaginationPageCountTarget;
  pages: LayoutSnapshotPage[];
}

function normalizeLayoutOptions(options: LayoutSnapshotOptions = {}): Required<LayoutOptions> & {
  measurementFeedback?: PaginationMeasurementFeedback;
  pageCountTarget?: number;
} {
  return {
    ...DEFAULT_LAYOUT_OPTIONS,
    ...options
  };
}

function resolvedLayoutMetricsFromGeometry(
  geometry: PaginationSectionGeometry
): ResolvedLayoutMetrics {
  return {
    pageSizePx: {
      width: geometry.pageWidthPx,
      height: geometry.pageHeightPx
    },
    marginsPx: { ...geometry.marginsPx },
    contentBoxPx: { ...geometry.contentBoxPx }
  };
}

function resolveSectionsFromGeometry(
  model: DocModel,
  sectionGeometries: PaginationSectionGeometry[],
  fallbackLayout: ResolvedLayoutMetrics,
  respectSectionLayout: boolean
): ResolvedDocumentSection[] {
  const hasExplicitSections = (model.metadata.sections ?? []).length > 0;
  if (sectionGeometries.length === 0) {
    return [
      {
        index: 0,
        startNodeIndex: 0,
        synthetic: true,
        layout: fallbackLayout
      }
    ];
  }

  return sectionGeometries.map((geometry, index) => ({
    index,
    startNodeIndex: geometry.startNodeIndex,
    synthetic: !hasExplicitSections,
    layout: respectSectionLayout
      ? resolvedLayoutMetricsFromGeometry(geometry)
      : fallbackLayout
  }));
}

export function resolveDocumentForLayout(
  model: DocModel,
  options: LayoutSnapshotOptions = {}
): ResolvedDocument {
  const normalizedOptions = normalizeLayoutOptions(options);
  const sectionGeometries = resolvePaginationSectionGeometry(model, {
    pageWidthPx: normalizedOptions.pageWidth,
    pageHeightPx: normalizedOptions.pageHeight,
    marginPx: normalizedOptions.margin
  });
  const layout = resolvedLayoutMetricsFromGeometry(
    sectionGeometries[0] ?? {
      startNodeIndex: 0,
      pageWidthPx: normalizedOptions.pageWidth,
      pageHeightPx: normalizedOptions.pageHeight,
      marginsPx: {
        top: normalizedOptions.margin,
        right: normalizedOptions.margin,
        bottom: normalizedOptions.margin,
        left: normalizedOptions.margin
      },
      contentBoxPx: {
        x: normalizedOptions.margin,
        y: normalizedOptions.margin,
        width: Math.max(0, normalizedOptions.pageWidth - normalizedOptions.margin * 2),
        height: Math.max(0, normalizedOptions.pageHeight - normalizedOptions.margin * 2)
      }
    }
  );

  const sections = resolveSectionsFromGeometry(
    model,
    sectionGeometries,
    layout,
    normalizedOptions.respectSectionLayout
  );

  return {
    source: "layout-engine-adapter",
    nodeCount: model.nodes.length,
    pageCountHint: model.metadata.documentPageCount,
    layout,
    sections,
    metadata: {
      sourceParts: model.metadata.sourceParts,
      warningCount: model.metadata.warnings.length,
      warnings: [...model.metadata.warnings],
      headerSectionCount: model.metadata.headerSections.length,
      footerSectionCount: model.metadata.footerSections.length
    }
  };
}

function legacyBlockSource(id: string): LayoutFragmentSource | undefined {
  const paragraphMatch = /^paragraph-(\d+)$/.exec(id);
  if (paragraphMatch) {
    return {
      kind: "paragraph",
      nodeIndex: Number.parseInt(paragraphMatch[1], 10)
    };
  }

  const tableMatch = /^table-(\d+)$/.exec(id);
  if (tableMatch) {
    return {
      kind: "table",
      nodeIndex: Number.parseInt(tableMatch[1], 10)
    };
  }

  return undefined;
}

function snapshotRectFromLegacyBlock(
  block: Pick<LayoutParagraphBlock | LayoutTableBlock, "x" | "y" | "width" | "height">
): LayoutSnapshotRect {
  return {
    x: block.x,
    y: block.y,
    width: block.width,
    height: block.height
  };
}

function snapshotRunFromLegacyRun(run: LayoutRun): LayoutSnapshotRun {
  if (run.kind === "image") {
    return {
      kind: "image",
      id: run.id,
      src: run.src,
      alt: run.alt,
      widthPx: run.widthPx,
      heightPx: run.heightPx,
      floating: run.floating
    };
  }

  return {
    kind: "text",
    id: run.id,
    text: run.text,
    style: run.style,
    link: run.link
  };
}

function snapshotParagraphFragment(params: {
  block: LayoutParagraphBlock;
  framePx: LayoutSnapshotRect;
  source?: LayoutFragmentSource;
  paragraphLineRange?: ParagraphLineRange;
}): LayoutSnapshotParagraphFragment {
  return {
    kind: "paragraph",
    id: params.block.id,
    framePx: params.framePx,
    align: params.block.align,
    headingLevel: params.block.headingLevel,
    runs: params.block.runs.map(snapshotRunFromLegacyRun),
    paragraphLineRange: params.paragraphLineRange,
    source: params.source
  };
}

function snapshotParagraphFromLegacyBlock(
  block: LayoutParagraphBlock,
  source?: LayoutFragmentSource
): LayoutSnapshotParagraphFragment {
  return snapshotParagraphFragment({
    block,
    framePx: snapshotRectFromLegacyBlock(block),
    source
  });
}

function snapshotTableFragment(params: {
  block: LayoutTableBlock;
  framePx: LayoutSnapshotRect;
  source?: LayoutFragmentSource;
  tableRowRange?: TableRowRange;
}): LayoutSnapshotTableFragment {
  const tableSource = params.source;
  const rowStartIndex = Math.max(0, Math.round(params.tableRowRange?.startRowIndex ?? 0));
  const rowEndIndex = Math.max(
    rowStartIndex,
    Math.round(params.tableRowRange?.endRowIndex ?? params.block.rows.length)
  );

  return {
    kind: "table",
    id: params.block.id,
    framePx: params.framePx,
    rows: params.block.rows.slice(rowStartIndex, rowEndIndex).map((row, visibleRowIndex) => {
      const rowIndex = rowStartIndex + visibleRowIndex;
      return {
        id: row.id,
        backgroundColor: row.backgroundColor,
        cells: row.cells.map((cell, cellIndex) => ({
          id: cell.id,
          colSpan: cell.colSpan,
          backgroundColor: cell.backgroundColor,
          paragraphs: cell.paragraphs.map((paragraph, paragraphIndex) =>
            snapshotParagraphFromLegacyBlock(paragraph, {
              kind: "table-cell-paragraph",
              nodeIndex: tableSource?.nodeIndex ?? -1,
              rowIndex,
              cellIndex,
              paragraphIndex
            })
          ),
          source:
            tableSource !== undefined
              ? {
                  kind: "table-cell",
                  nodeIndex: tableSource.nodeIndex,
                  rowIndex,
                  cellIndex
                }
              : undefined
        }))
      };
    }),
    tableRowRange: params.tableRowRange,
    source: tableSource
  };
}

function snapshotPageBlock(params: {
  block: PaginatedPageBlock;
  model: DocModel;
  options: Required<LayoutOptions>;
}): LayoutSnapshotBlockFragment | undefined {
  const node = params.model.nodes[params.block.nodeIndex];
  if (!node) {
    return undefined;
  }

  const source: LayoutFragmentSource = {
    kind: params.block.kind,
    nodeIndex: params.block.nodeIndex,
    columnIndex: params.block.columnIndex
  };

  if (node.type === "table") {
    const block = tableToLayout(
      node,
      `table-${params.block.nodeIndex}`,
      params.block.framePx.x,
      params.block.framePx.y,
      params.block.framePx.width,
      params.options
    );
    return snapshotTableFragment({
      block,
      framePx: params.block.framePx,
      source,
      tableRowRange: params.block.tableRowRange
    });
  }

  const block = paragraphToLayout(
    node,
    `paragraph-${params.block.nodeIndex}`,
    params.block.framePx.x,
    params.block.framePx.y,
    params.block.framePx.width,
    params.options.minLineHeight
  );
  return snapshotParagraphFragment({
    block,
    framePx: params.block.framePx,
    source,
    paragraphLineRange: params.block.paragraphLineRange
  });
}

/**
 * Builds the viewer layout snapshot from the real pagination engine
 * (`buildDocumentPageNodeSegments`) rather than the legacy one-line-per-paragraph
 * estimator in `layoutDocument()`.
 *
 * `layoutDocument()` remains exported for the public `core` entry point and for
 * static block stacking, but it is no longer the page-boundary source of truth.
 */
export function buildLayoutSnapshot(
  model: DocModel,
  options: LayoutSnapshotOptions = {}
): LayoutSnapshot {
  const resolvedDocument = resolveDocumentForLayout(model, options);
  const normalizedOptions = normalizeLayoutOptions(options);
  const paginationOptions = {
    pageWidthPx: normalizedOptions.pageWidth,
    pageHeightPx: normalizedOptions.pageHeight,
    marginPx: normalizedOptions.margin,
    suppressSpacingBeforeAfterPageBreak:
      model.metadata.compatibility?.suppressSpacingBeforeAfterPageBreak === true,
    measurementFeedback: normalizedOptions.measurementFeedback
  };
  // Notes render at the bottom of their page; reserve that height during
  // pagination so body content (and the notes themselves) never overflow.
  const paginated =
    options.reserveNoteHeight === false
      ? paginateDocumentLayout(model, paginationOptions)
      : paginateDocumentLayoutWithNoteReserve(model, paginationOptions);

  return {
    version: 1,
    source: "layout-engine-adapter",
    resolvedDocument,
    pageCountTarget: resolvePaginationPageCountTarget(
      options.pageCountTarget !== undefined
        ? options.pageCountTarget
        : model.metadata.documentPageCount
    ),
    pages: paginated.pages.map((page) => ({
      index: page.index,
      number: page.index + 1,
      pageSizePx: {
        width: page.geometry.pageWidthPx,
        height: page.geometry.pageHeightPx
      },
      marginsPx: page.geometry.marginsPx,
      contentBoxPx: page.geometry.contentBoxPx,
      columnCount: page.geometry.columnCount,
      columnGapPx: page.geometry.columnGapPx,
      columnWidthsPx: page.geometry.columnWidthsPx,
      blocks: page.blocks
        .map((block) =>
          snapshotPageBlock({ block, model, options: normalizedOptions })
        )
        .filter((block): block is LayoutSnapshotBlockFragment => block !== undefined)
    }))
  };
}
