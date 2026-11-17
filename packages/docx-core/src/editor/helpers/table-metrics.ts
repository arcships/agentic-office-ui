// Shared table render metrics.
//
// The DOM table host and the pagination estimators must resolve the same table
// width and column widths. When they disagree (tblW vs tcW vs the grid, a
// 100%-wide render, or a table that was scaled down to fit the page) body
// wrapping and row heights drift away from what is painted.
//
// Upstream editor.tsx resolved these numbers inline in several places; this
// module is the single source of truth.

import type { TableNode } from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
import {
  clampTableWidthPx,
  columnWidthsFromTableDefinition,
  defaultColumnWidthsPx,
  fitColumnWidthsToWidth,
  normalizeColumnWidthsPx,
  tableColumnCount,
} from "./table-utils";
import { resolveCollapsedTableHorizontalOuterBleedPx } from "./table-utils-extra";

const MIN_TABLE_RENDER_WIDTH_PX = 120;

export interface TableRenderMetricsPx {
  columnCount: number;
  /** Table width from tblW or, without one, the sum of the column widths. */
  rawTableWidthPx: number;
  /** Table width after clamping to the available content width. */
  tableWidthPx: number;
  /** Column widths derived from tblGrid/tcW before fitting to the table width. */
  rawColumnWidthsPx: number[];
  /** Column widths the table is actually laid out with. */
  columnWidthsPx: number[];
  /** True when the model carries no tblW and no column widths at all. */
  autoFitWidth: boolean;
}

export function resolveTableRenderMetricsPx(
  table: TableNode,
  maxAvailableWidthPx?: number
): TableRenderMetricsPx {
  const columnCount = tableColumnCount(table);
  const declaredTableWidthPx = twipsToPixels(table.style?.widthTwips);
  const definedWidthsTwips = columnWidthsFromTableDefinition(table, columnCount);
  const hasDefinedColumnWidths = Boolean(
    definedWidthsTwips && definedWidthsTwips.length > 0
  );
  const autoFitWidth =
    !hasDefinedColumnWidths &&
    !(Number.isFinite(declaredTableWidthPx) &&
      (declaredTableWidthPx as number) > 0);

  const collapsedHorizontalBorderBleedPx =
    resolveCollapsedTableHorizontalOuterBleedPx(table, columnCount);
  const availableWidthPx =
    Number.isFinite(maxAvailableWidthPx) && (maxAvailableWidthPx as number) > 0
      ? Math.max(
          MIN_TABLE_RENDER_WIDTH_PX,
          (maxAvailableWidthPx as number) - collapsedHorizontalBorderBleedPx
        )
      : undefined;

  // A table without tblW/tblGrid is auto-fitted by Word: it fills the content
  // width instead of collapsing to the per-column default.
  const rawTableWidthPx = autoFitWidth
    ? availableWidthPx ??
      defaultColumnWidthsPx(columnCount).reduce((sum, width) => sum + width, 0)
    : declaredTableWidthPx ??
      (definedWidthsTwips ?? [])
        .map((widthTwips) => twipsToPixels(widthTwips) ?? 0)
        .reduce((sum, width) => sum + width, 0);

  const rawColumnWidthsPx = hasDefinedColumnWidths
    ? normalizeColumnWidthsPx(
        (definedWidthsTwips ?? []).map(
          (widthTwips) => twipsToPixels(widthTwips) ?? 0
        ),
        columnCount,
        rawTableWidthPx > 0 ? rawTableWidthPx : declaredTableWidthPx,
        1
      )
    : defaultColumnWidthsPx(
        columnCount,
        rawTableWidthPx > 0 ? rawTableWidthPx : undefined
      );

  const tableWidthPx = clampTableWidthPx(rawTableWidthPx, availableWidthPx);

  return {
    columnCount,
    rawTableWidthPx,
    tableWidthPx,
    rawColumnWidthsPx,
    columnWidthsPx: fitColumnWidthsToWidth(rawColumnWidthsPx, tableWidthPx),
    autoFitWidth,
  };
}
