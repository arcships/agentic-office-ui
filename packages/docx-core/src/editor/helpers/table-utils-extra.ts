// Table border presets and cell-border rendering.
// Upstream editor.tsx: lines 21353-21913.

import type {
  ParagraphBorderSet,
  ParagraphBorderStyle,
  TableBorderSet,
  TableBorderStyle,
  TableNode
} from "../../engine/types";
import type {
  DocxBorderPreset,
  DocxBorderPresetState
} from "./editor-types";
import {
  cloneParagraphBorderSet,
  cloneTableBorderSet
} from "./numbering";
import {
  tableBorderToCss,
  tableBorderStrokeWidthPx,
  borderTypeVisible,
  paragraphBorderVisible,
  tableBorderVisible,
  tableBorderSetHasVisibleEdges,
  tableUsesSeparateBorderModel,
  resolveTableSeparateBorderSpacingPx,
  tableElementBorderStyle,
  resolvePreferredParagraphBorder,
  resolvePreferredTableBorder,
  toolbarParagraphBorderStyle,
  toolbarTableBorderStyle,
  nilParagraphBorderStyle,
  nilTableBorderStyle,
  tableColumnCount,
  normalizeBorderType,
  type TableBorderSide
} from "./table-utils";

export function paragraphBorderPresetState(
  borders: ParagraphBorderSet | undefined
): DocxBorderPresetState {
  const top = paragraphBorderVisible(borders?.top);
  const right = paragraphBorderVisible(borders?.right);
  const bottom = paragraphBorderVisible(borders?.bottom);
  const left = paragraphBorderVisible(borders?.left);
  const between = paragraphBorderVisible(borders?.between);
  const bar = paragraphBorderVisible(borders?.bar);
  const hasAny = top || right || bottom || left || between || bar;

  return {
    bottom,
    top,
    left,
    right,
    none: !hasAny,
    all: top && right && bottom && left,
    outside: top && right && bottom && left,
    inside: between || bar,
    "inside-horizontal": between,
    "inside-vertical": bar,
    "diagonal-down": false,
    "diagonal-up": false,
    "horizontal-line": bottom,
  };
}

export function paragraphRangePresetActive(
  preset: DocxBorderPreset,
  bordersByParagraph: Array<ParagraphBorderSet | undefined>
): boolean {
  if (bordersByParagraph.length === 0) {
    return false;
  }

  const firstBorders = bordersByParagraph[0];
  const lastBorders = bordersByParagraph[bordersByParagraph.length - 1];
  const allLeft = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.left)
  );
  const allRight = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.right)
  );
  const allTop = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.top)
  );
  const allBottom = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.bottom)
  );
  const allBetween = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.between)
  );
  const allBar = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.bar)
  );

  switch (preset) {
    case "top":
      return paragraphBorderVisible(firstBorders?.top);
    case "bottom":
      return paragraphBorderVisible(lastBorders?.bottom);
    case "left":
      return allLeft;
    case "right":
      return allRight;
    case "all":
    case "outside":
      return (
        paragraphBorderVisible(firstBorders?.top) &&
        paragraphBorderVisible(lastBorders?.bottom) &&
        allLeft &&
        allRight
      );
    case "none":
      return bordersByParagraph.every(
        (borders) => paragraphBorderPresetState(borders).none
      );
    case "inside":
      return allBetween && allBar;
    case "inside-horizontal":
      return allBetween;
    case "inside-vertical":
      return allBar;
    case "horizontal-line":
      return paragraphBorderVisible(lastBorders?.bottom);
    case "diagonal-down":
    case "diagonal-up":
      return false;
    default:
      return allTop && allBottom;
  }
}

export function applyParagraphBorderPresetForRangeEntry(
  borders: ParagraphBorderSet | undefined,
  preset: DocxBorderPreset,
  remove: boolean,
  index: number,
  total: number
): ParagraphBorderSet | undefined {
  if (total <= 1) {
    return applyParagraphBorderPreset(borders, preset, remove);
  }

  const nextBorders = cloneParagraphBorderSet(borders) ?? {};
  const visibleBorder = toolbarParagraphBorderStyle(
    resolvePreferredParagraphBorder(nextBorders)
  );
  const nilBorder = nilParagraphBorderStyle();
  const borderToApply = remove ? nilBorder : visibleBorder;
  const isFirst = index === 0;
  const isLast = index === total - 1;

  switch (preset) {
    case "top":
      nextBorders.top = { ...(isFirst ? borderToApply : nilBorder) };
      return nextBorders;
    case "bottom":
      nextBorders.bottom = { ...(isLast ? borderToApply : nilBorder) };
      return nextBorders;
    case "left":
      nextBorders.left = { ...borderToApply };
      return nextBorders;
    case "right":
      nextBorders.right = { ...borderToApply };
      return nextBorders;
    case "all":
    case "outside":
      if (remove) {
        nextBorders.top = { ...nilBorder };
        nextBorders.right = { ...nilBorder };
        nextBorders.bottom = { ...nilBorder };
        nextBorders.left = { ...nilBorder };
      } else {
        nextBorders.top = { ...(isFirst ? borderToApply : nilBorder) };
        nextBorders.right = { ...borderToApply };
        nextBorders.bottom = { ...(isLast ? borderToApply : nilBorder) };
        nextBorders.left = { ...borderToApply };
      }
      return nextBorders;
    case "horizontal-line":
      nextBorders.bottom = { ...(isLast ? borderToApply : nilBorder) };
      return nextBorders;
    case "none":
      nextBorders.top = { ...nilBorder };
      nextBorders.right = { ...nilBorder };
      nextBorders.bottom = { ...nilBorder };
      nextBorders.left = { ...nilBorder };
      nextBorders.between = { ...nilBorder };
      nextBorders.bar = { ...nilBorder };
      return nextBorders;
    default:
      return applyParagraphBorderPreset(borders, preset, remove);
  }
}

export function tableBorderPresetState(
  borders: TableBorderSet | undefined,
  selectedCellBorders?: TableBorderSet
): DocxBorderPresetState {
  const top = tableBorderVisible(borders?.top);
  const right = tableBorderVisible(borders?.right);
  const bottom = tableBorderVisible(borders?.bottom);
  const left = tableBorderVisible(borders?.left);
  const insideH = tableBorderVisible(borders?.insideH);
  const insideV = tableBorderVisible(borders?.insideV);
  const diagonalDown = tableBorderVisible(
    selectedCellBorders?.tl2br ?? borders?.tl2br
  );
  const diagonalUp = tableBorderVisible(
    selectedCellBorders?.tr2bl ?? borders?.tr2bl
  );
  const hasAny =
    top ||
    right ||
    bottom ||
    left ||
    insideH ||
    insideV ||
    diagonalDown ||
    diagonalUp;

  return {
    bottom,
    top,
    left,
    right,
    none: !hasAny,
    all: top && right && bottom && left && insideH && insideV,
    outside: top && right && bottom && left,
    inside: insideH && insideV,
    "inside-horizontal": insideH,
    "inside-vertical": insideV,
    "diagonal-down": diagonalDown,
    "diagonal-up": diagonalUp,
    "horizontal-line": insideH,
  };
}

export function applyParagraphBorderPreset(
  borders: ParagraphBorderSet | undefined,
  preset: DocxBorderPreset,
  remove = false
): ParagraphBorderSet | undefined {
  const nextBorders = cloneParagraphBorderSet(borders) ?? {};
  const visibleBorder = toolbarParagraphBorderStyle(
    resolvePreferredParagraphBorder(nextBorders)
  );
  const nilBorder = nilParagraphBorderStyle();
  const borderToApply = remove ? nilBorder : visibleBorder;

  switch (preset) {
    case "top":
      nextBorders.top = { ...borderToApply };
      break;
    case "right":
      nextBorders.right = { ...borderToApply };
      break;
    case "bottom":
      nextBorders.bottom = { ...borderToApply };
      break;
    case "left":
      nextBorders.left = { ...borderToApply };
      break;
    case "all":
    case "outside":
      nextBorders.top = { ...borderToApply };
      nextBorders.right = { ...borderToApply };
      nextBorders.bottom = { ...borderToApply };
      nextBorders.left = { ...borderToApply };
      break;
    case "inside":
      nextBorders.between = { ...borderToApply };
      nextBorders.bar = { ...borderToApply };
      break;
    case "inside-horizontal":
      nextBorders.between = { ...borderToApply };
      break;
    case "inside-vertical":
      nextBorders.bar = { ...borderToApply };
      break;
    case "horizontal-line":
      nextBorders.bottom = { ...borderToApply };
      break;
    case "none":
      nextBorders.top = { ...nilBorder };
      nextBorders.right = { ...nilBorder };
      nextBorders.bottom = { ...nilBorder };
      nextBorders.left = { ...nilBorder };
      nextBorders.between = { ...nilBorder };
      nextBorders.bar = { ...nilBorder };
      break;
    case "diagonal-down":
    case "diagonal-up":
      return undefined;
    default:
      return nextBorders;
  }

  return nextBorders;
}

export function applyTableBorderPreset(
  borders: TableBorderSet | undefined,
  preset: DocxBorderPreset,
  remove = false
): TableBorderSet | undefined {
  const nextBorders = cloneTableBorderSet(borders) ?? {};
  const visibleBorder = toolbarTableBorderStyle(
    resolvePreferredTableBorder(nextBorders)
  );
  const nilBorder = nilTableBorderStyle();
  const borderToApply = remove ? nilBorder : visibleBorder;

  switch (preset) {
    case "top":
      nextBorders.top = { ...borderToApply };
      break;
    case "right":
      nextBorders.right = { ...borderToApply };
      break;
    case "bottom":
      nextBorders.bottom = { ...borderToApply };
      break;
    case "left":
      nextBorders.left = { ...borderToApply };
      break;
    case "all":
      nextBorders.top = { ...borderToApply };
      nextBorders.right = { ...borderToApply };
      nextBorders.bottom = { ...borderToApply };
      nextBorders.left = { ...borderToApply };
      nextBorders.insideH = { ...borderToApply };
      nextBorders.insideV = { ...borderToApply };
      break;
    case "outside":
      nextBorders.top = { ...borderToApply };
      nextBorders.right = { ...borderToApply };
      nextBorders.bottom = { ...borderToApply };
      nextBorders.left = { ...borderToApply };
      if (!remove) {
        nextBorders.insideH = { ...nilBorder };
        nextBorders.insideV = { ...nilBorder };
      }
      break;
    case "inside":
      nextBorders.insideH = { ...borderToApply };
      nextBorders.insideV = { ...borderToApply };
      break;
    case "inside-horizontal":
      nextBorders.insideH = { ...borderToApply };
      break;
    case "inside-vertical":
      nextBorders.insideV = { ...borderToApply };
      break;
    case "horizontal-line":
      nextBorders.insideH = { ...borderToApply };
      break;
    case "none":
      nextBorders.top = { ...nilBorder };
      nextBorders.right = { ...nilBorder };
      nextBorders.bottom = { ...nilBorder };
      nextBorders.left = { ...nilBorder };
      nextBorders.insideH = { ...nilBorder };
      nextBorders.insideV = { ...nilBorder };
      nextBorders.tl2br = { ...nilBorder };
      nextBorders.tr2bl = { ...nilBorder };
      break;
    case "diagonal-down":
      nextBorders.tl2br = { ...borderToApply };
      break;
    case "diagonal-up":
      nextBorders.tr2bl = { ...borderToApply };
      break;
    default:
      return nextBorders;
  }

  return nextBorders;
}

export function resolveTableBorder(
  tableBorders: TableBorderSet | undefined,
  cellBorders: TableBorderSet | undefined,
  side: TableBorderSide,
  rowIndex: number,
  rowCount: number,
  startColumnIndex: number,
  endColumnIndex: number,
  columnCount: number
): TableBorderStyle | undefined {
  const directCellBorder = cellBorders?.[side];
  if (directCellBorder) {
    return directCellBorder;
  }

  const isTopRow = rowIndex === 0;
  const isBottomRow = rowIndex >= rowCount - 1;
  const isFirstColumn = startColumnIndex === 0;
  const isLastColumn = endColumnIndex >= columnCount - 1;

  if (side === "top") {
    return isTopRow ? tableBorders?.top : tableBorders?.insideH;
  }
  if (side === "bottom") {
    return isBottomRow ? tableBorders?.bottom : tableBorders?.insideH;
  }
  if (side === "left") {
    return isFirstColumn ? tableBorders?.left : tableBorders?.insideV;
  }

  return isLastColumn ? tableBorders?.right : tableBorders?.insideV;
}

export function resolveTableCellBorderCss(
  tableBorders: TableBorderSet | undefined,
  cellBorders: TableBorderSet | undefined,
  rowIndex: number,
  rowCount: number,
  startColumnIndex: number,
  endColumnIndex: number,
  columnCount: number
): Record<string, string | number | undefined> {
  const top = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "top",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );
  const right = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "right",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );
  const bottom = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "bottom",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );
  const left = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "left",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );

  return {
    ...(top !== undefined ? { borderTop: top } : undefined),
    ...(right !== undefined ? { borderRight: right } : undefined),
    ...(bottom !== undefined ? { borderBottom: bottom } : undefined),
    ...(left !== undefined ? { borderLeft: left } : undefined),
  };
}

export function resolveCollapsedTableHorizontalOuterBleedPx(
  table: TableNode,
  columnCount = tableColumnCount(table)
): number {
  if (columnCount <= 0 || tableUsesSeparateBorderModel(table)) {
    return 0;
  }

  const rowCount = table.rows.length;
  const tableBorders = table.style?.borders;
  let maxLeftBorderWidthPx = tableBorderStrokeWidthPx(tableBorders?.left);
  let maxRightBorderWidthPx = tableBorderStrokeWidthPx(tableBorders?.right);

  table.rows.forEach((row, rowIndex) => {
    let columnCursor = 0;

    row.cells.forEach((cell) => {
      const columnSpan =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const startColumnIndex = columnCursor;
      const endColumnIndex = Math.min(
        columnCount - 1,
        startColumnIndex + columnSpan - 1
      );
      columnCursor += columnSpan;

      if (cell.style?.vMergeContinuation) {
        return;
      }

      if (startColumnIndex === 0) {
        maxLeftBorderWidthPx = Math.max(
          maxLeftBorderWidthPx,
          tableBorderStrokeWidthPx(
            resolveTableBorder(
              tableBorders,
              cell.style?.borders,
              "left",
              rowIndex,
              rowCount,
              startColumnIndex,
              endColumnIndex,
              columnCount
            )
          )
        );
      }

      if (endColumnIndex >= columnCount - 1) {
        maxRightBorderWidthPx = Math.max(
          maxRightBorderWidthPx,
          tableBorderStrokeWidthPx(
            resolveTableBorder(
              tableBorders,
              cell.style?.borders,
              "right",
              rowIndex,
              rowCount,
              startColumnIndex,
              endColumnIndex,
              columnCount
            )
          )
        );
      }
    });
  });

  return Math.max(
    0,
    Math.ceil((maxLeftBorderWidthPx + maxRightBorderWidthPx) / 2)
  );
}

export function resolveTableCellDiagonalOverlayCss(
  tableBorders: TableBorderSet | undefined,
  cellBorders: TableBorderSet | undefined
): Record<string, string | number | undefined> {
  const diagonalDownBorder = cellBorders?.tl2br ?? tableBorders?.tl2br;
  const diagonalUpBorder = cellBorders?.tr2bl ?? tableBorders?.tr2bl;
  const layers: string[] = [];

  const addLayer = (
    border: TableBorderStyle | undefined,
    direction: "to bottom right" | "to bottom left"
  ): void => {
    if (!tableBorderVisible(border)) {
      return;
    }

    const color = border?.color ?? "#000000";
    const widthPx =
      Number.isFinite(border?.sizeEighthPt) &&
      (border?.sizeEighthPt as number) > 0
        ? Math.max(
            0.75,
            Number(((border?.sizeEighthPt as number) / 6).toFixed(2))
          )
        : 1;
    const halfWidthPx = Number((widthPx / 2).toFixed(2));
    layers.push(
      `linear-gradient(${direction}, transparent calc(50% - ${halfWidthPx}px), ${color} calc(50% - ${halfWidthPx}px), ${color} calc(50% + ${halfWidthPx}px), transparent calc(50% + ${halfWidthPx}px))`
    );
  };

  addLayer(diagonalDownBorder, "to bottom right");
  addLayer(diagonalUpBorder, "to bottom left");

  if (layers.length === 0) {
    return {};
  }

  return {
    backgroundImage: layers.join(", "),
    backgroundRepeat: "no-repeat",
    backgroundSize: "100% 100%",
  };
}
