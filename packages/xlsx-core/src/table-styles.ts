import type {
  XlsxCellAddress,
  XlsxResolvedCellStyle,
  XlsxTable,
  XlsxTableStyleDefinition,
  XlsxTableStyleInfo,
} from "./types";
import {
  BUILTIN_TABLE_STYLES,
  resolveBuiltinTableStyle,
} from "./builtin-table-styles";

export { BUILTIN_TABLE_STYLES, resolveBuiltinTableStyle };

/**
 * Resolved table-style elements that apply to one cell of a worksheet table.
 *
 * Excel layers the elements from lowest to highest precedence: whole table,
 * column stripes, row stripes, last/first column, header row, total row, and
 * finally the corner header/total cells. Excel also treats an explicitly
 * formatted cell as taking precedence over the table style, so
 * {@link withTableCellStyle} merges only the cell's explicit fill, font and
 * borders over the table style.
 */
type TableStyleRegion = {
  endCol: number;
  endRow: number;
  startCol: number;
  startRow: number;
};

type TableStyleOptions = {
  /** Workbook-defined table styles; a name here wins over the shipped preset. */
  tableStyleByName?: Record<string, XlsxTableStyleDefinition> | null;
  /** Named cell styles for `table/@headerRowCellStyle`. */
  namedCellStyleByName?: Record<string, XlsxResolvedCellStyle> | null;
  /**
   * Subtrees the cell's own `xf` record declares. When present, only these
   * properties override the table style; when absent, the merge falls back to
   * comparing the resolved font against the workbook defaults.
   */
  declaredTableStyleOverrides?: XlsxResolvedCellStyle | null;
};

const FONT_DEFAULTS: Record<string, unknown> = {
  bold: false,
  italic: false,
  name: "Calibri",
  size: 11,
  strikethrough: false,
  underline: "none",
  verticalAlign: "baseline",
};

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" ? value as Record<string, unknown> : null;
}

/**
 * Merges a style overlay over a base style the way the viewer's sheet model
 * does: nested font/border objects merge key by key, fill and other top-level
 * keys are replaced.
 */
export function mergeTableStyle(
  base: Record<string, unknown> | null | undefined,
  overlay: Record<string, unknown> | null | undefined,
): Record<string, unknown> | null {
  if (!base && !overlay) {
    return null;
  }
  const merged: Record<string, unknown> = {
    ...(base ?? {}),
    ...(overlay ?? {}),
  };
  for (const key of ["alignment", "font"] as const) {
    const basePart = asRecord(base?.[key]);
    const overlayPart = asRecord(overlay?.[key]);
    if (basePart || overlayPart) {
      merged[key] = { ...(basePart ?? {}), ...(overlayPart ?? {}) };
    }
  }
  const baseBorder = asRecord(base?.border);
  const overlayBorder = asRecord(overlay?.border);
  if (baseBorder || overlayBorder) {
    merged.border = { ...(baseBorder ?? {}), ...(overlayBorder ?? {}) };
  }
  const baseFill = asRecord(base?.fill);
  const overlayFill = asRecord(overlay?.fill);
  if (baseFill || overlayFill) {
    merged.fill = overlayFill ?? baseFill;
  }
  return merged;
}

/**
 * Clips one table-style element's border to the region that element covers.
 *
 * A table style element's border describes the region as a whole: `top` is the
 * region's top edge, `horizontal` its inner row separators. A renderer drawing
 * a single cell needs the per-cell subset, otherwise every cell would paint
 * the region's outer border.
 */
function normalizeTableStyleEdges(
  style: Record<string, unknown> | null | undefined,
  region: TableStyleRegion,
  row: number,
  col: number,
): Record<string, unknown> | null {
  if (!style) {
    return null;
  }
  const border = asRecord(style.border);
  if (!border) {
    return style;
  }

  const nextBorder: Record<string, unknown> = {};
  if (border.top && row === region.startRow) {
    nextBorder.top = border.top;
  }
  if (border.bottom && row === region.endRow) {
    nextBorder.bottom = border.bottom;
  }
  if (border.left && col === region.startCol) {
    nextBorder.left = border.left;
  }
  if (border.right && col === region.endCol) {
    nextBorder.right = border.right;
  }
  if (border.horizontal && row < region.endRow) {
    nextBorder.bottom = nextBorder.bottom ?? border.horizontal;
  }
  if (border.vertical && col < region.endCol) {
    nextBorder.right = nextBorder.right ?? border.vertical;
  }

  const nextStyle: Record<string, unknown> = { ...style };
  if (Object.keys(nextBorder).length === 0) {
    delete nextStyle.border;
  } else {
    nextStyle.border = nextBorder;
  }
  return nextStyle;
}

function isFirstColumnVisible(styleInfo: XlsxTableStyleInfo | undefined) {
  return styleInfo?.showFirstColumn === true;
}

function isLastColumnVisible(styleInfo: XlsxTableStyleInfo | undefined) {
  return styleInfo?.showLastColumn === true;
}

/** Resolves the table-style overlay of one cell, or `null` when none applies. */
export function resolveTableCellStyle(
  table: XlsxTable | null | undefined,
  cell: XlsxCellAddress,
  options: TableStyleOptions = {},
): XlsxResolvedCellStyle | null {
  if (!table) {
    return null;
  }
  const styleName = table.styleInfo?.name;
  if (!styleName) {
    return null;
  }
  // A workbook may define its own style under any name, including a built-in
  // one, so the workbook's definition wins over the shipped preset.
  const tableStyle = options.tableStyleByName?.[styleName] ?? resolveBuiltinTableStyle(styleName);
  if (!tableStyle) {
    return null;
  }

  const headerRowCount = Math.max(table.headerRowCount, 1);
  const totalsRowCount = Math.max(table.totalsRowCount, 0);
  const headerEndRow = table.start.row + headerRowCount - 1;
  const totalsStartRow = table.end.row - totalsRowCount + 1;
  const bodyStartRow = headerEndRow + 1;
  const bodyEndRow = totalsStartRow - 1;
  const row = cell.row;
  const col = cell.col;
  const isHeaderRow = row >= table.start.row && row <= headerEndRow;
  const isTotalsRow = totalsRowCount > 0 && row >= totalsStartRow && row <= table.end.row;
  const isBodyRow = row >= bodyStartRow && row <= bodyEndRow;
  const isFirstColumn = col === table.start.col;
  const isLastColumn = col === table.end.col;

  const tableRegion: TableStyleRegion = {
    endCol: table.end.col,
    endRow: table.end.row,
    startCol: table.start.col,
    startRow: table.start.row,
  };
  const headerRegion: TableStyleRegion = { ...tableRegion, endRow: headerEndRow };
  const totalsRegion: TableStyleRegion = { ...tableRegion, startRow: totalsStartRow };

  let resolved: Record<string, unknown> | null = null;
  const applyElement = (elementType: string, region: TableStyleRegion, enabled = true) => {
    if (!enabled) {
      return;
    }
    const nextStyle = normalizeTableStyleEdges(tableStyle[elementType] ?? null, region, row, col);
    if (nextStyle) {
      resolved = mergeTableStyle(resolved, nextStyle);
    }
  };

  applyElement("wholeTable", tableRegion);

  if (isBodyRow && table.styleInfo?.showColumnStripes) {
    const isFirstStripe = (col - table.start.col) % 2 === 0;
    applyElement(isFirstStripe ? "firstColumnStripe" : "secondColumnStripe", {
      endCol: col,
      endRow: bodyEndRow,
      startCol: col,
      startRow: bodyStartRow,
    });
  }

  if (isBodyRow && table.styleInfo?.showRowStripes) {
    // Excel bands from the first body row, so an even offset is the first
    // stripe even when the table has one or more header rows.
    const isFirstStripe = (row - bodyStartRow) % 2 === 0;
    applyElement(isFirstStripe ? "firstRowStripe" : "secondRowStripe", {
      ...tableRegion,
      endRow: row,
      startRow: row,
    });
  }

  applyElement(
    "lastColumn",
    { ...tableRegion, startCol: table.end.col },
    isLastColumnVisible(table.styleInfo) && isLastColumn,
  );
  applyElement(
    "firstColumn",
    { ...tableRegion, endCol: table.start.col },
    isFirstColumnVisible(table.styleInfo) && isFirstColumn,
  );

  applyElement("headerRow", headerRegion, isHeaderRow);
  if (isHeaderRow && table.headerRowCellStyle) {
    const namedHeaderStyle = options.namedCellStyleByName?.[table.headerRowCellStyle];
    if (namedHeaderStyle) {
      resolved = mergeTableStyle(resolved, namedHeaderStyle as Record<string, unknown>);
    }
  }

  applyElement("totalRow", totalsRegion, isTotalsRow);

  applyElement("firstHeaderCell", { ...headerRegion, endCol: table.start.col }, isHeaderRow && isFirstColumn);
  applyElement("lastHeaderCell", { ...headerRegion, startCol: table.end.col }, isHeaderRow && isLastColumn);
  applyElement("firstTotalCell", { ...totalsRegion, endCol: table.start.col }, isTotalsRow && isFirstColumn);
  applyElement("lastTotalCell", { ...totalsRegion, startCol: table.end.col }, isTotalsRow && isLastColumn);

  return resolved as XlsxResolvedCellStyle | null;
}

/**
 * The cell's explicit font properties.
 *
 * The engine resolves a default font for every styled cell and the viewer's
 * style resolver falls back to the row/column style, so a table style's bold
 * white header would be overwritten by the default black Calibri 11 font if the
 * whole font object were merged over it. Comparing against the workbook's
 * defaults keeps the header emphasis while a truly customised cell still wins,
 * matching Excel's precedence.
 */
function explicitFont(
  value: unknown,
  declared: Record<string, unknown> | null | undefined,
): Record<string, unknown> {
  const font = asRecord(value);
  if (!font) {
    return {};
  }
  if (declared) {
    // The `xf` says which font subtree it owns; every key of that subtree is
    // explicit, including a deliberate `bold: false`.
    return { ...font };
  }
  const explicit: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(font)) {
    if (key === "color") {
      const color = asRecord(entry);
      const isDefaultBlack = !color
        || color.colorType === "auto"
        || (color.theme === 1 && color.tint === undefined && color.rgb === undefined && color.argb === undefined);
      if (!isDefaultBlack) {
        explicit.color = entry;
      }
      continue;
    }
    if (key === "family" || key === "scheme" || key === "charset") {
      // Font metadata only affects font selection, not table emphasis.
      continue;
    }
    if (entry !== FONT_DEFAULTS[key]) {
      explicit[key] = entry;
    }
  }
  return explicit;
}

function explicitBorder(
  value: unknown,
  declared: Record<string, unknown> | null | undefined,
): Record<string, unknown> {
  const border = asRecord(value);
  if (!border) {
    return {};
  }
  const edges = declared ? Object.keys(declared) : null;
  return Object.fromEntries(
    Object.entries(border).filter(([edge]) => {
      if (edges && !edges.includes(edge)) {
        return false;
      }
      const style = asRecord(border[edge])?.style;
      return typeof style === "string" && style !== "none";
    }),
  );
}

/** Layers explicit cell formatting over a resolved table style. */
function mergeCellOverTableStyle(
  base: Record<string, unknown>,
  tableStyle: XlsxResolvedCellStyle,
  declaredOverrides?: XlsxResolvedCellStyle | null,
): XlsxResolvedCellStyle {
  const tableRecord = tableStyle as Record<string, unknown>;
  const result: Record<string, unknown> = { ...tableRecord, ...base };
  const declaredFill = declaredOverrides?.fill;
  const baseFill = asRecord(base.fill);
  const cellFillIsExplicit = declaredOverrides
    ? declaredFill !== undefined
    : Boolean(baseFill && baseFill.fillType && baseFill.fillType !== "none");
  if (cellFillIsExplicit && baseFill) {
    result.fill = base.fill;
  } else if (tableRecord.fill) {
    result.fill = tableRecord.fill;
  } else {
    delete result.fill;
  }

  const tableFont = asRecord(tableRecord.font);
  const declaredFont = declaredOverrides?.font as Record<string, unknown> | null | undefined;
  const font = { ...(tableFont ?? {}), ...explicitFont(base.font, declaredFont) };
  if (Object.keys(font).length > 0) {
    result.font = font;
  } else {
    delete result.font;
  }

  const tableBorder = asRecord(tableRecord.border);
  const declaredBorder = declaredOverrides?.border as Record<string, unknown> | null | undefined;
  const border = { ...(tableBorder ?? {}), ...explicitBorder(base.border, declaredBorder) };
  if (Object.keys(border).length > 0) {
    result.border = border;
  } else {
    delete result.border;
  }

  if (base.alignment !== undefined) {
    result.alignment = base.alignment;
  } else if (tableRecord.alignment !== undefined) {
    result.alignment = tableRecord.alignment;
  }
  return result as XlsxResolvedCellStyle;
}

/** First worksheet table that contains `cell`, or `null`. */
export function findTableAtCell(
  tables: readonly XlsxTable[] | null | undefined,
  cell: XlsxCellAddress,
): XlsxTable | null {
  for (const table of tables ?? []) {
    if (
      cell.row >= table.start.row
      && cell.row <= table.end.row
      && cell.col >= table.start.col
      && cell.col <= table.end.col
    ) {
      return table;
    }
  }
  return null;
}

/**
 * Applies the table style of the table covering `cell` to a resolved cell style.
 *
 * Returns `base` unchanged when the sheet has no table at the cell, so sheets
 * without tables pay no allocation cost on the paint path.
 */
export function withTableCellStyle(
  base: XlsxResolvedCellStyle | undefined,
  tables: readonly XlsxTable[] | null | undefined,
  cell: XlsxCellAddress,
  options: TableStyleOptions = {},
): XlsxResolvedCellStyle | undefined {
  if (!tables || tables.length === 0) {
    return base;
  }
  const table = findTableAtCell(tables, cell);
  if (!table) {
    return base;
  }
  const tableStyle = resolveTableCellStyle(table, cell, options);
  if (!tableStyle) {
    return base;
  }
  return mergeCellOverTableStyle(
    (base ?? {}) as Record<string, unknown>,
    tableStyle,
    options.declaredTableStyleOverrides,
  );
}
