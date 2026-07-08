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

export interface XlsxImage {
  anchor: XlsxImageAnchor;
  description?: string;
  editable?: boolean;
  hyperlink?: string;
  id: string;
  mediaPath?: string;
  mimeType: string;
  name?: string;
  sheetIndex: number;
  src: string;
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
  geometry: string;
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
  svgPath?: string;
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
  fontFamily?: string;
  fontSizePt?: number;
  hidden?: boolean;
  id: string;
  kind: XlsxFormControlKind;
  label?: string;
  linkedCell?: string;
  name?: string;
  sheetIndex: number;
  textAlign?: "center" | "left" | "right";
  textColor?: string;
  workbookSheetIndex: number;
  zIndex: number;
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
