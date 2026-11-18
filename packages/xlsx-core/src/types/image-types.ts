// ── Image anchor / marker ────────────────────────────────────────────────

export interface XlsxImageMarker {
  col: number;
  colOffsetEmu: number;
  row: number;
  rowOffsetEmu: number;
}

export type XlsxImageAnchor =
  | {
      from: XlsxImageMarker;
      kind: "one-cell";
      sizeEmu: {
        cx: number;
        cy: number;
      };
    }
  | {
      kind: "absolute";
      positionEmu: {
        x: number;
        y: number;
      };
      sizeEmu: {
        cx: number;
        cy: number;
      };
    }
  | {
      from: XlsxImageMarker;
      kind: "two-cell";
      to: XlsxImageMarker;
    };

// ── Image ─────────────────────────────────────────────────────────────────

/**
 * Source crop of a picture, expressed as fractions of the intrinsic image
 * (0 = no crop, 0.25 = hide a quarter of that edge). Mirrors DrawingML
 * `a:srcRect`.
 */
export interface XlsxImageSourceRect {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export interface XlsxImage {
  anchor: XlsxImageAnchor;
  description?: string;
  editable?: boolean;
  flipH?: boolean;
  flipV?: boolean;
  hyperlink?: string;
  id: string;
  mediaPath?: string;
  mimeType: string;
  name?: string;
  /** Clockwise rotation in degrees, as stored by DrawingML `rot`. */
  rotationDeg?: number;
  sheetIndex: number;
  src: string;
  srcRect?: XlsxImageSourceRect;
  workbookSheetIndex: number;
  zIndex: number;
}

// ── Shape ─────────────────────────────────────────────────────────────────

export interface XlsxShapeFill {
  color?: string;
  none?: boolean;
  opacity?: number;
}

export interface XlsxShapeStroke {
  color?: string;
  dash?: string;
  headEndType?: string;
  none?: boolean;
  opacity?: number;
  tailEndType?: string;
  widthPx?: number;
}

export interface XlsxShapeTextRun {
  bold?: boolean;
  color?: string;
  fontFamily?: string;
  fontSizePt?: number;
  italic?: boolean;
  text: string;
  underline?: boolean;
}

export interface XlsxShapeParagraph {
  align?: "center" | "justify" | "left" | "right";
  runs: XlsxShapeTextRun[];
}

export interface XlsxShapeTextBox {
  horizontalAlign?: "center" | "left";
  insetPx?: {
    bottom: number;
    left: number;
    right: number;
    top: number;
  };
  verticalAlign?: "bottom" | "middle" | "top";
}

export interface XlsxShape {
  anchor: XlsxImageAnchor;
  description?: string;
  fill?: XlsxShapeFill;
  flipH?: boolean;
  flipV?: boolean;
  /** `a:prstGeom@prst` (or `custom` for `a:custGeom`); matched case-insensitively. */
  geometry: string;
  /** Raw `a:avLst/a:gd@fmla` values in DrawingML 1/100000 units. */
  geometryAdjustments?: Record<string, number>;
  hidden?: boolean;
  hyperlink?: string;
  id: string;
  name?: string;
  paragraphs: XlsxShapeParagraph[];
  rotationDeg?: number;
  scaleX?: number;
  scaleY?: number;
  sheetIndex: number;
  /** SVG path for custom geometry or the resolved preset (see `buildPresetShapePath`). */
  svgPath?: string;
  /** Coordinate space of `svgPath`; the preset path uses the shape's pixel extents. */
  svgViewBox?: {
    height: number;
    width: number;
  };
  stroke?: XlsxShapeStroke;
  textBox?: XlsxShapeTextBox;
  workbookSheetIndex: number;
  zIndex: number;
}

// ── Form controls ─────────────────────────────────────────────────────────

export type XlsxFormControlKind =
  | "button"
  | "checkbox"
  | "dropdown"
  | "editbox"
  | "group-box"
  | "label"
  | "listbox"
  | "radio"
  | "scrollbar"
  | "spinner"
  | "unknown";

export interface XlsxFormControl {
  anchor: XlsxImageAnchor;
  checked?: boolean;
  editable?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  hidden?: boolean;
  id: string;
  kind: XlsxFormControlKind;
  label?: string;
  linkedCell?: string;
  name?: string;
  sheetIndex: number;
  /** Worksheet `control@shapeId`, a stable identity in the sheet XML. */
  shapeId?: number;
  textAlign?: "center" | "left" | "right";
  textColor?: string;
  workbookSheetIndex: number;
  zIndex: number;
}

/** Patch accepted by programmable form-control CRUD. */
export interface XlsxFormControlInput {
  anchor: XlsxImageAnchor;
  checked?: boolean;
  kind?: XlsxFormControlKind;
  label?: string;
  linkedCell?: string;
  name?: string;
  zIndex?: number;
}

// ── Image rect / resize ──────────────────────────────────────────────────

export interface XlsxImageRect {
  height: number;
  left: number;
  top: number;
  width: number;
}

export type XlsxImageResizeHandlePosition = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";

// ── Render props ─────────────────────────────────────────────────────────

export interface XlsxImageRenderProps {
  /** The built-in image element that react-xlsx would render without customization. */
  defaultNode: unknown;
  /** Workbook image metadata, including source, anchor, name, alt text, and editability. */
  image: XlsxImage;
  /** The image rectangle in viewer pixels, including the current zoom level. */
  rect: XlsxImageRect;
  /** Absolute positioning styles to apply when rendering a replacement image node. */
  style: Record<string, string | number | undefined>;
}

export interface XlsxImageSelectionRenderProps {
  /** The built-in selected-image outline and resize handles. */
  defaultNode: unknown;
  /** Returns pointer handlers and styles for a custom resize handle. */
  getHandleProps: (
    position: XlsxImageResizeHandlePosition
  ) => {
    onPointerDown: (event: Record<string, unknown>) => void;
    style: Record<string, string | number | undefined>;
  };
  /** The currently selected image. */
  image: XlsxImage;
  /** The selected image rectangle in viewer pixels, including the current zoom level. */
  rect: XlsxImageRect;
}
