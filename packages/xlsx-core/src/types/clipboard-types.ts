import type { XlsxChart } from "./chart-types";
import type { XlsxImage, XlsxImageAnchor, XlsxShape } from "./image-types";

/** Drawing families the clipboard can move between sheets. */
export type XlsxDrawingClipboardKind = "chart" | "image" | "shape";

/**
 * One drawing captured by copy/cut.
 *
 * `model` is the JSON model rendered at copy time, while `sourceId` and
 * `sourceSheetId` let paste clone the DrawingML anchor inside the same
 * workbook. `relativeAnchor` is the drawing's top-left offset (EMU) from the
 * copied group's top-left, so pasting several drawings keeps their spacing.
 */
export interface XlsxDrawingClipboardItem {
  /** Absolute source anchor, retained so paste can preserve the anchor kind. */
  anchor: XlsxImageAnchor;
  kind: XlsxDrawingClipboardKind;
  model: XlsxChart | XlsxImage | XlsxShape;
  /** Top-left offset from the copied drawing group, in EMU. */
  relativeAnchor: {
    x: number;
    y: number;
  };
  /** Rendered size in EMU. */
  size: {
    cx: number;
    cy: number;
  };
  /** Stable identity of the copied drawing in the source workbook. */
  sourceId: string;
  /** Workbook sheet index the drawing was copied from. */
  sourceSheetId: number;
}
