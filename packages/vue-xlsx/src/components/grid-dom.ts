import type { CSSProperties } from "vue";
import type { XlsxCellTextRun } from "@arcships/xlsx-core";
import { RICH_TEXT_VERTICAL_ALIGN_SCALE } from "./rich-text-layout";

/**
 * Opt-in DOM renderer for the grid body.
 *
 * The canvas path paints the visible range into four canvases. That is fast and
 * pixel-consistent, but it hides the cell text from browser find, native text
 * selection and the accessibility tree, and it leaves no fallback when a canvas
 * cannot be created. `renderMode="dom"` renders the same visible window as real
 * elements instead: one absolutely positioned `role="gridcell"` per visible
 * cell, positioned with the same column/row offsets the canvas uses, and styled
 * from the very same `CanvasCellStyle` the paint loop resolves.
 *
 * Everything in this module is pure: the grid builds the model and this module
 * maps a resolved cell style (plus the geometry options) to CSS so the DOM and
 * canvas renderers cannot drift apart in typography, fills, borders or
 * alignment.
 */
export type XlsxGridRenderMode = "canvas" | "dom";

export const DEFAULT_GRID_RENDER_MODE: XlsxGridRenderMode = "canvas";

/** Normalizes the public `renderMode` prop; anything but an explicit `dom` stays canvas. */
export function resolveGridRenderMode(value: unknown): XlsxGridRenderMode {
  return value === "dom" ? "dom" : DEFAULT_GRID_RENDER_MODE;
}

/**
 * Style one cell has after `resolveCanvasCellStyleInto` (the canvas paint path).
 * Kept here so both renderers share one type instead of copying the fields.
 */
export interface CanvasCellStyle {
  alignment: "left" | "center" | "right";
  backgroundColor: string;
  border?: Record<string, Record<string, unknown>>;
  color: string;
  fontFamily: string;
  fontSizePx: number;
  fontStyle: string;
  fontWeight: string;
  /** Canvas angle of rotated text, or `null` for horizontal text. */
  rotation: number | null;
  /** Whether Excel may shrink the font down to 0.4× to fit the cell. */
  shrinkToFit: boolean;
  strike: boolean;
  underline: boolean;
  vertical: "top" | "middle" | "bottom";
  wrap: boolean;
}

/** Horizontal/vertical cell padding, matching the canvas text chrome. */
export const GRID_DOM_CELL_PADDING_X_PX = 4;
export const GRID_DOM_CELL_PADDING_Y_PX = 1;

/** Excel's link colour, shared with `drawCellText`'s hyperlink argument. */
export const GRID_DOM_LINK_COLOR = "#2563eb";

/** Resolved CSS border of a workbook style edge, or `undefined` for no border. */
export function canvasBorderEdgeCss(
  edge: Record<string, unknown> | undefined,
  color: string | null | undefined,
): string | undefined {
  const style = typeof edge?.style === "string" ? edge.style : undefined;
  if (!style || style === "none") return undefined;
  const lowered = style.toLowerCase();
  const cssStyle = lowered.includes("dash")
    ? "dashed"
    : lowered === "dotted" || lowered === "hair"
      ? "dotted"
      : lowered === "double"
        ? "double"
        : "solid";
  const width = lowered.startsWith("medium") ? 2 : lowered === "thick" || lowered === "double" ? 3 : 1;
  return `${width}px ${cssStyle} ${color ?? "#52525b"}`;
}

/** CSS border of one cell's four edges, resolved from the workbook style. */
export interface GridDomCellBorders {
  top?: string;
  right?: string;
  bottom?: string;
  left?: string;
}

export interface CanvasCellCssOptions {
  /** Workbook borders for this cell (already resolved to CSS). */
  borders?: GridDomCellBorders;
  /** Conditional color-scale fill; wins over the cell fill like the canvas. */
  backgroundColor?: string;
  /** Excel shrink-to-fit size, resolved by the caller from the cell width. */
  fontSizePx?: number;
  /** Gridlines are drawn for this cell (sheet has them enabled). */
  gridLines?: boolean;
  gridColor?: string;
  /**
   * A merged cell paints its own outline in the canvas painter instead of
   * relying on the neighbouring gridlines, so it needs all four edges.
   */
  merged?: boolean;
  hyperlink?: boolean;
  hyperlinkColor?: string;
}

/**
 * Cell box CSS: background, font, alignment, padding and borders.
 *
 * The canvas fills the cell and then draws its border centred on the edge; the
 * DOM equivalent uses `border-box` borders and inset box-shadows for the
 * gridlines (the same trick the upstream DOM renderer uses), so the text keeps
 * exactly the resolved workbook style and the border never shifts the layout.
 */
export function canvasCellStyleToCss(
  style: CanvasCellStyle,
  options: CanvasCellCssOptions = {},
): CSSProperties {
  const fontSizePx = Math.max(1, options.fontSizePx ?? style.fontSizePx);
  const decoration = [
    style.underline || options.hyperlink ? "underline" : "",
    style.strike ? "line-through" : "",
  ].filter(Boolean).join(" ");
  const shadows: string[] = [];
  if (options.gridLines && options.gridColor) {
    if (options.merged) {
      shadows.push(`inset 0 0 0 1px ${options.gridColor}`);
    } else {
      shadows.push(
        `inset -1px 0 0 0 ${options.gridColor}`,
        `inset 0 -1px 0 0 ${options.gridColor}`,
      );
    }
  }
  return {
    alignItems: style.vertical === "top"
      ? "flex-start"
      : style.vertical === "middle" ? "center" : "flex-end",
    backgroundColor: options.backgroundColor ?? style.backgroundColor,
    borderBottom: options.borders?.bottom,
    borderLeft: options.borders?.left,
    borderRight: options.borders?.right,
    borderTop: options.borders?.top,
    boxShadow: shadows.length ? shadows.join(", ") : undefined,
    boxSizing: "border-box",
    color: options.hyperlink
      ? options.hyperlinkColor ?? GRID_DOM_LINK_COLOR
      : style.color,
    display: "flex",
    fontFamily: `${JSON.stringify(style.fontFamily)}, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif`,
    fontSize: `${fontSizePx}px`,
    fontStyle: style.fontStyle,
    fontWeight: style.fontWeight,
    justifyContent: style.alignment === "right"
      ? "flex-end"
      : style.alignment === "center" ? "center" : "flex-start",
    lineHeight: `${Math.max(12, fontSizePx * 1.25)}px`,
    overflow: "hidden",
    padding: `${GRID_DOM_CELL_PADDING_Y_PX}px ${GRID_DOM_CELL_PADDING_X_PX}px`,
    textAlign: style.alignment,
    textDecoration: decoration || undefined,
  };
}

export interface CanvasCellTextCssOptions {
  /** Excel shrink-to-fit size, resolved by the caller from the cell width. */
  fontSizePx?: number;
  /** Available width of the text column when an icon set reserves the right edge. */
  maxWidthPx?: number;
}

/**
 * Text box CSS inside the cell.
 *
 * Horizontal text is a flex child that clips or ellipsizes exactly like the
 * canvas line; rotated text is centred and transformed in place, so the rotated
 * bounding box is clipped by the cell instead of escaping it. Either way the
 * text stays a real, selectable text node.
 */
export function canvasCellTextStyle(
  style: CanvasCellStyle,
  options: CanvasCellTextCssOptions = {},
): CSSProperties {
  const fontSizePx = Math.max(1, options.fontSizePx ?? style.fontSizePx);
  const base: CSSProperties = {
    display: "block",
    fontSize: `${fontSizePx}px`,
    lineHeight: `${Math.max(12, fontSizePx * 1.25)}px`,
    // The canvas ends right-aligned text 5 px inside the cell (its text
    // padding); keep the same gutter before an icon-set column.
    marginRight: options.maxWidthPx === undefined ? undefined : GRID_DOM_CELL_PADDING_X_PX + 1,
    maxWidth: options.maxWidthPx === undefined ? "100%" : `${Math.max(0, options.maxWidthPx)}px`,
    minWidth: "0",
    overflow: "hidden",
    overflowWrap: style.wrap ? "break-word" : undefined,
    position: "relative",
    textOverflow: style.wrap ? "clip" : "ellipsis",
    userSelect: "text",
    whiteSpace: style.wrap ? "pre-wrap" : "nowrap",
    wordBreak: style.wrap ? "break-word" : undefined,
    zIndex: 1,
  };
  if (style.rotation === null) return base;
  return {
    ...base,
    display: "inline-block",
    left: "50%",
    maxWidth: "none",
    position: "absolute",
    textOverflow: undefined,
    top: "50%",
    transform: `translate(-50%, -50%) rotate(${style.rotation}deg)`,
    transformOrigin: "center center",
    userSelect: "text",
    whiteSpace: "nowrap",
    width: "max-content",
  };
}

/** CSS of one inline rich-text run, on top of the cell style. */
export function richTextRunCss(
  style: CanvasCellStyle,
  font: XlsxCellTextRun["font"],
  options: { color?: string | null; fontSizePx?: number } = {},
): CSSProperties {
  const cellFontSizePx = Math.max(1, options.fontSizePx ?? style.fontSizePx);
  const verticalAlign = (font?.verticalAlign ?? "").toLocaleLowerCase();
  const shrink = verticalAlign === "superscript" || verticalAlign === "subscript";
  const fontSizePx = Math.max(1, cellFontSizePx * (shrink ? RICH_TEXT_VERTICAL_ALIGN_SCALE : 1));
  const decoration = [
    font?.underline ? "underline" : "",
    font?.strikethrough ? "line-through" : "",
  ].filter(Boolean).join(" ");
  return {
    color: options.color ?? undefined,
    fontFamily: font?.name?.trim() ? font.name : undefined,
    fontSize: `${Math.round(fontSizePx)}px`,
    fontStyle: font?.italic ? "italic" : undefined,
    fontWeight: font?.bold ? "700" : undefined,
    lineHeight: shrink ? String(RICH_TEXT_VERTICAL_ALIGN_SCALE) : undefined,
    textDecoration: decoration || undefined,
    verticalAlign: verticalAlign === "superscript"
      ? "super"
      : verticalAlign === "subscript" ? "sub" : undefined,
  };
}

/** Selection tint laid over a cell, matching the canvas' inset fill. */
export function gridDomSelectionStyle(fill: string): CSSProperties {
  return {
    background: fill,
    inset: "1px",
    pointerEvents: "none",
    position: "absolute",
    zIndex: 0,
  };
}

/** Display slots one freeze-pane axis renders: pinned slots plus the scrolling window. */
export interface GridDomAxisWindow {
  frozen: number[];
  scrolling: number[];
}

/**
 * Splits the visible display range into the frozen strip and the scrolling
 * window. Frozen slots are pinned and always painted, even when the scroll
 * position left them behind; the scrolling window starts after them so a cell
 * is never rendered by two panes.
 */
export function resolveGridDomAxisWindow(
  visible: { start: number; end: number },
  frozenCount: number,
  total: number,
): GridDomAxisWindow {
  const count = Number.isFinite(total) ? Math.max(0, Math.floor(total)) : 0;
  const frozen: number[] = [];
  const pinned = Math.max(0, Math.min(Math.floor(frozenCount) || 0, count));
  for (let index = 0; index < pinned; index += 1) frozen.push(index);
  const scrolling: number[] = [];
  const start = Math.max(Math.floor(visible.start) || 0, pinned);
  const end = Math.min(Math.floor(visible.end) || 0, count - 1);
  for (let index = start; index <= end; index += 1) scrolling.push(index);
  return { frozen, scrolling };
}

/** Pane-local rectangle of a range or cell. */
export interface GridDomRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Pane-local box a fragment has to be clipped to. */
export interface GridDomClip {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface GridDomOutlineOptions {
  color: string;
  /** Inset of the stroke from the range edge, matching `drawRangeBorder`. */
  inset?: number;
  lineWidth?: number;
  dashed?: boolean;
  zIndex?: number;
}

/**
 * Overflow allowed before an outline box is clamped. The pane clips the box,
 * so keeping the clamp far outside the viewport preserves the real border on
 * the visible edges while a whole-row selection never allocates a million-pixel
 * element.
 */
const OUTLINE_CLAMP_PX = 4_000;

/**
 * Absolutely positioned border box for a range in one pane, or `null` when the
 * range is completely outside the pane. The caller places the result inside the
 * pane's translated content layer, so scrolling and the freeze line follow the
 * pane exactly as the canvas outline does.
 */
export function resolveGridDomOutlineStyle(
  rect: GridDomRect,
  clip: GridDomClip,
  options: GridDomOutlineOptions,
): CSSProperties | null {
  const inset = options.inset ?? 1;
  const left = rect.left + inset;
  const top = rect.top + inset;
  const right = rect.right - inset;
  const bottom = rect.bottom - inset;
  if (!(right > left) || !(bottom > top)) return null;
  if (right < clip.left - OUTLINE_CLAMP_PX || left > clip.right + OUTLINE_CLAMP_PX) return null;
  if (bottom < clip.top - OUTLINE_CLAMP_PX || top > clip.bottom + OUTLINE_CLAMP_PX) return null;
  // Clamping only bites far outside the pane, where the border is clipped away
  // anyway, so no visible edge is moved onto the freeze line.
  const clampedLeft = Math.max(left, clip.left - OUTLINE_CLAMP_PX);
  const clampedTop = Math.max(top, clip.top - OUTLINE_CLAMP_PX);
  const clampedRight = Math.min(right, clip.right + OUTLINE_CLAMP_PX);
  const clampedBottom = Math.min(bottom, clip.bottom + OUTLINE_CLAMP_PX);
  const lineWidth = Math.max(1, options.lineWidth ?? 2);
  return {
    border: `${lineWidth}px ${options.dashed ? "dashed" : "solid"} ${options.color}`,
    boxSizing: "border-box",
    height: `${clampedBottom - clampedTop}px`,
    left: `${clampedLeft}px`,
    pointerEvents: "none",
    position: "absolute",
    top: `${clampedTop}px`,
    width: `${clampedRight - clampedLeft}px`,
    zIndex: options.zIndex ?? 2,
  };
}

/** One rendered row inside a pane. */
export interface GridDomRowModel {
  key: string;
  displayRow: number;
  ariaRowIndex: number;
  style: CSSProperties;
  cells: GridDomCellModel[];
}

/** One rendered cell; every field is ready for the template. */
export interface GridDomCellModel {
  key: string;
  address: string;
  id?: string;
  active: boolean;
  ariaColIndex: number;
  ariaRowIndex: number;
  ariaLabel?: string;
  selected: boolean;
  style: CSSProperties;
  text?: string;
  textStyle?: CSSProperties;
  runs?: GridDomTextRunModel[];
  selectionStyle?: CSSProperties;
  dataBarStyle?: CSSProperties;
  dataBarAxisStyle?: CSSProperties;
  icon?: string;
  iconStyle?: CSSProperties;
  sparkline?: GridDomSparklineModel;
  comment?: boolean;
}

export interface GridDomTextRunModel {
  key: string;
  text: string;
  style: CSSProperties;
}

/** Tiny inline SVG description of a sparkline, painted instead of the value. */
export interface GridDomSparklineModel {
  viewBox: string;
  color: string;
  bars?: Array<{ key: string; x: number; y: number; width: number; height: number }>;
  line?: string;
  markers?: Array<{ key: string; cx: number; cy: number }>;
}

export interface GridDomOutlineModel {
  key: string;
  style: CSSProperties;
}

/** One freeze pane: a clipped box plus the cells that belong to it. */
export interface GridDomPaneModel {
  key: string;
  style: CSSProperties;
  scrollX: boolean;
  scrollY: boolean;
  rows: GridDomRowModel[];
  outlines: GridDomOutlineModel[];
}

export interface GridDomHeaderItemModel {
  key: string;
  label: string;
  style: CSSProperties;
}

export interface GridDomHeaderLayerModel {
  /** Columns/rows that scroll with the body. */
  scrolling: GridDomHeaderItemModel[];
  /** Columns/rows pinned in the frozen strip. */
  frozen: GridDomHeaderItemModel[];
  /** Freeze-line accent between the pinned and scrolling strips. */
  divider?: GridDomDividerModel;
}

export interface GridDomDividerModel {
  key: string;
  style: CSSProperties;
}

export interface GridDomLayerModel {
  panes: GridDomPaneModel[];
  dividers: GridDomDividerModel[];
  columnHeaders: GridDomHeaderLayerModel;
  rowHeaders: GridDomHeaderLayerModel;
  cornerStyle: CSSProperties;
  /**
   * Id the active cell carries when it is inside the rendered window. The grid
   * keeps its off-screen a11y cell only when this is empty, so assistive tech
   * never sees the active cell twice.
   */
  activeCellId?: string;
}
