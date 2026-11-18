import type {
  DocModel,
  DocNode,
  FormFieldRunNode,
  HeadingLevel,
  ParagraphAlignment,
  ParagraphNode,
  TableCellContentNode,
  TableNode,
  TextStyle
} from "../engine/types";
import { shareImageData } from "../engine/clone";
import type { DocumentLayoutMetrics } from "../viewer/section-layout";
import { parseSectionLayout } from "../viewer/section-layout";
import {
  resolveDocumentSectionsFromMetadata,
  resolveSectionIndexForNodeIndex,
  sectionBreakPropertiesStartNewPage,
  sectionBreakRequiredPageParity,
  type SectionStartPageParity
} from "./sections";

export interface LayoutOptions {
  pageWidth?: number;
  pageHeight?: number;
  margin?: number;
  minLineHeight?: number;
  paragraphSpacing?: number;
  tableCellPadding?: number;
  /**
   * Use each section's own `w:pgSz`/`w:pgMar` (parsed from its `w:sectPr`)
   * instead of applying the single document-level geometry to every section.
   * Defaults to `true`.
   */
  respectSectionLayout?: boolean;
}

export interface LayoutTextRun {
  kind: "text";
  id: string;
  text: string;
  style?: TextStyle;
  link?: string;
}

export interface LayoutImageRun {
  kind: "image";
  id: string;
  src?: string;
  alt?: string;
  widthPx?: number;
  heightPx?: number;
  contentType?: string;
  data?: Uint8Array;
  floating?: boolean;
}

export type LayoutRun = LayoutTextRun | LayoutImageRun;

export interface LayoutParagraphBlock {
  kind: "paragraph";
  id: string;
  runs: LayoutRun[];
  align: ParagraphAlignment;
  headingLevel?: HeadingLevel;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface LayoutTableCell {
  id: string;
  colSpan?: number;
  backgroundColor?: string;
  paragraphs: LayoutParagraphBlock[];
}

export interface LayoutTableRow {
  id: string;
  backgroundColor?: string;
  cells: LayoutTableCell[];
}

export interface LayoutTableBlock {
  kind: "table";
  id: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rows: LayoutTableRow[];
}

export type LayoutBlock = LayoutParagraphBlock | LayoutTableBlock;

export interface LayoutPageLayout {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
}

export interface LayoutPage {
  number: number;
  blocks: LayoutBlock[];
  /** Zero-based model section active when the page was started. */
  sectionIndex?: number;
  /** Per-section page geometry this page is rendered with. */
  layout?: LayoutPageLayout;
}

const DEFAULT_OPTIONS: Required<LayoutOptions> = {
  pageWidth: 816,
  pageHeight: 1056,
  margin: 72,
  minLineHeight: 22,
  paragraphSpacing: 8,
  tableCellPadding: 8,
  respectSectionLayout: true
};

function headingScale(level?: HeadingLevel): number {
  if (!level) {
    return 1;
  }

  switch (level) {
    case 1:
      return 2.15;
    case 2:
      return 1.75;
    case 3:
      return 1.45;
    case 4:
      return 1.28;
    case 5:
      return 1.15;
    case 6:
      return 1.05;
    default:
      return 1;
  }
}

function runHeightPx(run: LayoutRun): number {
  if (run.kind === "image") {
    if (run.floating) {
      return 0;
    }
    return run.heightPx ?? 96;
  }

  const fontSizePt = run.style?.fontSizePt ?? 12;
  return Math.round(fontSizePt * 1.6);
}

function lineHeightFromRuns(
  runs: LayoutRun[],
  minLineHeight: number,
  headingLevel?: HeadingLevel
): number {
  const base =
    runs.reduce((largest, run) => {
      return Math.max(largest, runHeightPx(run));
    }, 12) * headingScale(headingLevel);

  return Math.max(minLineHeight, Math.round(base));
}

function spacingForBlock(baseSpacing: number, headingLevel?: HeadingLevel): number {
  if (!headingLevel) {
    return baseSpacing;
  }
  return baseSpacing + Math.max(4, (7 - headingLevel) * 2);
}

function formFieldDisplayText(field: FormFieldRunNode): string {
  switch (field.fieldType) {
    case "checkbox":
      return field.checked ? field.checkedSymbol ?? "☒" : field.uncheckedSymbol ?? "☐";
    case "dropdown":
    case "date":
    case "text":
    default:
      return field.value ?? "";
  }
}

export function paragraphToLayout(
  paragraph: ParagraphNode,
  idPrefix: string,
  x: number,
  y: number,
  width: number,
  minLineHeight: number
): LayoutParagraphBlock {
  const runs: LayoutRun[] = paragraph.children.map((child, runIndex) => {
    if (child.type === "text") {
      return {
        kind: "text",
        id: `${idPrefix}-run-${runIndex}`,
        text: child.text,
        style: child.style,
        link: child.link
      };
    }

    if (child.type === "form-field") {
      return {
        kind: "text",
        id: `${idPrefix}-run-${runIndex}`,
        text: formFieldDisplayText(child),
        style: child.style,
        link: child.link
      };
    }

    return {
      kind: "image",
      id: `${idPrefix}-run-${runIndex}`,
      src: child.src,
      alt: child.alt,
      widthPx: child.widthPx,
      heightPx: child.heightPx,
      contentType: child.contentType,
      data: shareImageData(child.data),
      floating: Boolean(child.floating)
    };
  });

  return {
    kind: "paragraph",
    id: idPrefix,
    runs,
    align: paragraph.style?.align ?? "left",
    headingLevel: paragraph.style?.headingLevel,
    x,
    y,
    width,
    height: lineHeightFromRuns(runs, minLineHeight, paragraph.style?.headingLevel)
  };
}

function paragraphsForTableCellNodes(nodes: TableCellContentNode[]): ParagraphNode[] {
  const paragraphs: ParagraphNode[] = [];

  const walk = (items: typeof nodes): void => {
    for (const item of items) {
      if (item.type === "paragraph") {
        paragraphs.push(item);
        continue;
      }

      for (const row of item.rows) {
        for (const cell of row.cells) {
          walk(cell.nodes);
        }
      }
    }
  };

  walk(nodes);
  return paragraphs;
}

export function tableToLayout(
  table: TableNode,
  idPrefix: string,
  x: number,
  y: number,
  width: number,
  options: Required<LayoutOptions>
): LayoutTableBlock {
  const columnCount = Math.max(
    1,
    ...table.rows.map((row) =>
      row.cells.reduce(
        (sum, cell) => sum + (cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1),
        0
      )
    )
  );
  const baseCellWidth = width / columnCount;
  let tableHeight = 0;

  const rows: LayoutTableRow[] = table.rows.map((row, rowIndex) => {
    let rowHeight = options.minLineHeight + options.tableCellPadding * 2;
    let columnCursor = 0;

    const cells: LayoutTableCell[] = row.cells.map((cell, cellIndex) => {
      const colSpan = cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1;
      const cellWidth = baseCellWidth * colSpan;
      const cellParagraphs = paragraphsForTableCellNodes(cell.nodes);

      const paragraphBlocks = cellParagraphs.map((paragraph, paragraphIndex) =>
        paragraphToLayout(
          paragraph,
          `${idPrefix}-r${rowIndex}-c${cellIndex}-p${paragraphIndex}`,
          x + columnCursor * baseCellWidth,
          y,
          cellWidth,
          options.minLineHeight
        )
      );

      const paragraphsHeight = paragraphBlocks.reduce((sum, block) => sum + block.height, 0);
      rowHeight = Math.max(rowHeight, paragraphsHeight + options.tableCellPadding * 2);
      columnCursor += colSpan;

      return {
        id: `${idPrefix}-r${rowIndex}-c${cellIndex}`,
        colSpan,
        backgroundColor: cell.style?.backgroundColor ?? row.style?.backgroundColor,
        paragraphs: paragraphBlocks
      };
    });

    tableHeight += rowHeight;

    return {
      id: `${idPrefix}-row-${rowIndex}`,
      backgroundColor: row.style?.backgroundColor,
      cells
    };
  });

  return {
    kind: "table",
    id: idPrefix,
    x,
    y,
    width,
    height: Math.max(tableHeight, options.minLineHeight * 2),
    rows
  };
}

function estimateBlockHeight(block: LayoutBlock): number {
  if (block.kind === "paragraph") {
    return block.height;
  }
  return block.height;
}

// Header/footer content occupies the page margin band. When it grows past the
// nominal margin, Word pushes the body down instead of growing the paper. The
// legacy paginator must reserve that overflow so body lines do not slide under
// the footer.
const PAGE_FURNITURE_INTER_PARAGRAPH_GAP_PX = 8;

function pageFurnitureNodeHasVisibleContent(node: DocNode): boolean {
  if (node.type === "table") {
    return true;
  }

  return node.children.some((child) => {
    if (child.type === "text") {
      return child.text.trim().length > 0;
    }
    return child.type === "image" || child.type === "form-field";
  });
}

function estimatePageFurnitureSectionHeightPx(
  nodes: DocNode[] | undefined,
  availableWidthPx: number,
  options: Required<LayoutOptions>
): number {
  const visibleNodes = (nodes ?? []).filter(pageFurnitureNodeHasVisibleContent);
  if (visibleNodes.length === 0) {
    return 0;
  }

  const heightPx = visibleNodes.reduce((sum, node) => {
    const block =
      node.type === "paragraph"
        ? paragraphToLayout(
            node,
            "page-furniture",
            0,
            0,
            availableWidthPx,
            options.minLineHeight
          )
        : tableToLayout(node, "page-furniture", 0, 0, availableWidthPx, options);
    return sum + block.height;
  }, 0);

  return (
    heightPx +
    Math.max(0, visibleNodes.length - 1) * PAGE_FURNITURE_INTER_PARAGRAPH_GAP_PX
  );
}

export function resolvePageFurnitureReservePx(
  model: DocModel,
  options: Required<LayoutOptions>
): { headerReservePx: number; footerReservePx: number } {
  const metadata = model.metadata;
  if (!metadata) {
    return { headerReservePx: 0, footerReservePx: 0 };
  }
  const section = metadata.sections?.[0];
  const sectionPropertiesXml =
    section?.sectionPropertiesXml ?? metadata.sectionPropertiesXml;
  const sectionLayout = parseSectionLayout(sectionPropertiesXml);
  const headerSections =
    section?.headerSections && section.headerSections.length > 0
      ? section.headerSections
      : metadata.headerSections;
  const footerSections =
    section?.footerSections && section.footerSections.length > 0
      ? section.footerSections
      : metadata.footerSections;
  const availableWidthPx = Math.max(
    24,
    options.pageWidth - options.margin * 2
  );
  const estimateGroupsHeightPx = (
    groups: Array<{ nodes?: DocNode[] }> | undefined
  ): number =>
    (groups ?? []).reduce(
      (largest, group) =>
        Math.max(
          largest,
          estimatePageFurnitureSectionHeightPx(
            group.nodes,
            availableWidthPx,
            options
          )
        ),
      0
    );
  const headerHeightPx = estimateGroupsHeightPx(headerSections);
  const footerHeightPx = estimateGroupsHeightPx(footerSections);

  return {
    headerReservePx:
      headerHeightPx > 0
        ? Math.max(
            0,
            Math.round(
              headerHeightPx + sectionLayout.headerDistancePx - options.margin
            )
          )
        : 0,
    footerReservePx:
      footerHeightPx > 0
        ? Math.max(
            0,
            Math.round(
              footerHeightPx + sectionLayout.footerDistancePx - options.margin
            )
          )
        : 0
  };
}

interface ResolvedLayoutSection {
  startNodeIndex: number;
  layout: DocumentLayoutMetrics;
  startsNewPage: boolean;
  parity?: SectionStartPageParity;
}

function baseLayoutMetrics(options: Required<LayoutOptions>): DocumentLayoutMetrics {
  const marginPx = Math.max(0, Math.round(options.margin));
  return {
    pageWidthPx: Math.max(1, Math.round(options.pageWidth)),
    pageHeightPx: Math.max(1, Math.round(options.pageHeight)),
    marginsPx: {
      top: marginPx,
      right: marginPx,
      bottom: marginPx,
      left: marginPx
    },
    headerDistancePx: marginPx,
    footerDistancePx: marginPx,
    docGridLinePitchPx: undefined
  };
}

function resolveLayoutSections(
  model: DocModel,
  baseLayout: DocumentLayoutMetrics,
  respectSectionLayout: boolean,
  useSectionZeroLayout: boolean
): ResolvedLayoutSection[] {
  return resolveDocumentSectionsFromMetadata(model.metadata).map(
    (section, sectionIndex) => {
      // Explicit layout options win for section 0; otherwise its own `w:sectPr`
      // is authoritative, like every later section.
      const layout =
        respectSectionLayout && (sectionIndex > 0 || useSectionZeroLayout)
          ? parseSectionLayout(section.sectionPropertiesXml, baseLayout)
          : baseLayout;
      return {
        startNodeIndex: Math.max(0, Math.round(section.startNodeIndex)),
        layout,
        startsNewPage:
          sectionIndex > 0 &&
          sectionBreakPropertiesStartNewPage(section.sectionPropertiesXml ?? ""),
        parity:
          sectionIndex > 0
            ? sectionBreakRequiredPageParity(section.sectionPropertiesXml)
            : undefined
      };
    }
  );
}

function layoutPageGeometry(layout: DocumentLayoutMetrics): LayoutPageLayout {
  return {
    pageWidth: layout.pageWidthPx,
    pageHeight: layout.pageHeightPx,
    marginTop: layout.marginsPx.top,
    marginRight: layout.marginsPx.right,
    marginBottom: layout.marginsPx.bottom,
    marginLeft: layout.marginsPx.left
  };
}

function layoutSectionGeometryChanged(
  left: DocumentLayoutMetrics,
  right: DocumentLayoutMetrics
): boolean {
  return (
    left.pageWidthPx !== right.pageWidthPx ||
    left.pageHeightPx !== right.pageHeightPx ||
    left.marginsPx.top !== right.marginsPx.top ||
    left.marginsPx.right !== right.marginsPx.right ||
    left.marginsPx.bottom !== right.marginsPx.bottom ||
    left.marginsPx.left !== right.marginsPx.left
  );
}

export function layoutDocument(model: DocModel, options: LayoutOptions = {}): LayoutPage[] {
  const resolved = { ...DEFAULT_OPTIONS, ...options };
  const baseLayout = baseLayoutMetrics(resolved);
  const useSectionZeroLayout =
    options.pageWidth === undefined &&
    options.pageHeight === undefined &&
    options.margin === undefined;
  const sections = resolveLayoutSections(
    model,
    baseLayout,
    resolved.respectSectionLayout,
    useSectionZeroLayout
  );
  if (sections.length === 0) {
    sections.push({
      startNodeIndex: 0,
      layout: baseLayout,
      startsNewPage: false
    });
  }

  const pages: LayoutPage[] = [];
  let currentSectionIndex = 0;
  let currentLayout = sections[0].layout;
  const { headerReservePx, footerReservePx } = resolvePageFurnitureReservePx(
    model,
    resolved
  );
  let contentWidthPx = 1;
  let pageBottomPx = 1;
  let cursorY = 0;

  const applyLayoutBounds = (): void => {
    contentWidthPx = Math.max(
      1,
      currentLayout.pageWidthPx -
        currentLayout.marginsPx.left -
        currentLayout.marginsPx.right
    );
    pageBottomPx =
      currentLayout.pageHeightPx -
      currentLayout.marginsPx.bottom -
      footerReservePx;
    cursorY = currentLayout.marginsPx.top + headerReservePx;
  };

  const startPage = (): void => {
    pages.push({
      number: pages.length + 1,
      blocks: [],
      sectionIndex: currentSectionIndex,
      layout: layoutPageGeometry(currentLayout)
    });
    applyLayoutBounds();
  };

  applyLayoutBounds();

  const startSectionPage = (
    parity: SectionStartPageParity | undefined,
    previousSectionIndex: number,
    previousLayout: DocumentLayoutMetrics
  ): void => {
    if (parity) {
      // The number the new section's first page would get without a blank page.
      const nextPageNumber = pages.length + 1;
      const startsOnEvenPage = nextPageNumber % 2 === 0;
      if (startsOnEvenPage !== (parity === "even")) {
        // Word inserts a blank page at the end of the section that ends here, so
        // the page keeps that section's header/footer and continues its
        // numbering; the new section's first content page then lands on the
        // required parity.
        pages.push({
          number: nextPageNumber,
          blocks: [],
          sectionIndex: previousSectionIndex,
          layout: layoutPageGeometry(previousLayout)
        });
      }
    }
    startPage();
  };

  startPage();

  for (const [index, node] of model.nodes.entries()) {
    const sectionIndex = resolveSectionIndexForNodeIndex(
      sections,
      index,
      currentSectionIndex
    );
    if (sectionIndex !== currentSectionIndex) {
      const previousLayout = currentLayout;
      const previousSectionIndex = currentSectionIndex;
      const nextSection = sections[sectionIndex];
      currentSectionIndex = sectionIndex;
      currentLayout = nextSection.layout;
      const startsNewPage =
        nextSection.startsNewPage ||
        layoutSectionGeometryChanged(previousLayout, currentLayout);
      if (startsNewPage) {
        if (pages[pages.length - 1].blocks.length > 0) {
          startSectionPage(
            nextSection.parity,
            previousSectionIndex,
            previousLayout
          );
        } else {
          pages[pages.length - 1].sectionIndex = currentSectionIndex;
          pages[pages.length - 1].layout = layoutPageGeometry(currentLayout);
          applyLayoutBounds();
        }
      }
    }

    const block =
      node.type === "paragraph"
        ? paragraphToLayout(
            node,
            `paragraph-${index}`,
            currentLayout.marginsPx.left,
            cursorY,
            contentWidthPx,
            resolved.minLineHeight
          )
        : tableToLayout(
            node,
            `table-${index}`,
            currentLayout.marginsPx.left,
            cursorY,
            contentWidthPx,
            resolved
          );

    const blockHeight = estimateBlockHeight(block);
    const currentPage = pages[pages.length - 1];

    if (cursorY + blockHeight > pageBottomPx && currentPage.blocks.length > 0) {
      startPage();
      block.y = cursorY;
    }

    pages[pages.length - 1].blocks.push(block);
    cursorY +=
      blockHeight +
      (block.kind === "paragraph"
        ? spacingForBlock(resolved.paragraphSpacing, block.headingLevel)
        : resolved.paragraphSpacing + 10);
  }

  return pages;
}
