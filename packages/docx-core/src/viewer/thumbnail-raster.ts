/**
 * DOCX page thumbnail rasterization helpers.
 *
 * The raster pipeline clones a live page surface, serializes it into an SVG
 * `foreignObject`, decodes it as an image, and draws it onto a canvas. The
 * helpers here keep that pipeline cheap:
 *
 * - editor-only chrome (selection rects, caret, handles) is stripped from the
 *   clone before serialization;
 * - scoped component CSS is carried into the clone as inline styles because a
 *   `foreignObject` cannot see the host document's stylesheets;
 * - large embedded data-URI images are swapped for cached, thumbnail-scale
 *   versions so serialize/encode/decode stop round-tripping megabytes;
 * - image decodes are bounded, so a stalled decoder cannot pin a thumbnail
 *   request in the rendering state forever;
 * - rasters land on reusable surface canvases held in an LRU cache;
 * - work runs through a serial idle-time queue that coalesces repeat requests
 *   per target canvas instead of fanning out `Promise.all` storms.
 */

/** Marks editor chrome that must never appear in rasterized thumbnails. */
export const DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE = "data-docx-thumbnail-exclude";

const THUMBNAIL_EXCLUDED_CLONE_SELECTOR = [
  `[${DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE}="true"]`,
  "textarea",
  '[data-image-resize-handle="true"]',
  '[data-docx-table-move-handle="true"]',
].join(",");

/**
 * Editor state classes that paint extra chrome on otherwise-document content.
 * They cannot be removed with an exclusion attribute because they sit on the
 * paragraph/table cell itself, so the raster pass drops the class and skips
 * the background it contributes.
 */
const THUMBNAIL_EDITOR_CHROME_CLASS_NAMES = new Set([
  "docx-paragraph-host--active",
  "docx-table-host--active",
  "docx-table-cell--active",
  "docx-floating-image--selected",
]);

/**
 * Computed properties copied into the clone. A data-URI SVG has no access to
 * the host document's stylesheets (scoped component CSS or `@font-face`), so
 * the values that actually shape the page have to travel as inline styles.
 * Layout-only state that the browser never paints (outline, box-shadow,
 * cursor…) is deliberately omitted.
 */
const THUMBNAIL_INLINE_STYLE_PROPERTIES = [
  "display",
  "position",
  "top",
  "right",
  "bottom",
  "left",
  "z-index",
  "overflow",
  "box-sizing",
  "min-width",
  "min-height",
  "flex-direction",
  "flex-wrap",
  "flex-grow",
  "flex-shrink",
  "flex-basis",
  "align-items",
  "align-self",
  "justify-content",
  "gap",
  "grid-template-columns",
  "grid-template-rows",
  "table-layout",
  "border-collapse",
  "border-spacing",
  "background-color",
  "border-top-width",
  "border-top-style",
  "border-top-color",
  "border-right-width",
  "border-right-style",
  "border-right-color",
  "border-bottom-width",
  "border-bottom-style",
  "border-bottom-color",
  "border-left-width",
  "border-left-style",
  "border-left-color",
  "padding-top",
  "padding-right",
  "padding-bottom",
  "padding-left",
  "margin-top",
  "margin-right",
  "margin-bottom",
  "margin-left",
  "color",
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "text-align",
  "text-decoration-line",
  "text-transform",
  "text-indent",
  "white-space",
  "word-break",
  "overflow-wrap",
  "vertical-align",
  "list-style-type",
  "list-style-position",
  "opacity",
] as const;

/** Inherited properties are skipped when the parent already carries the value. */
const THUMBNAIL_INHERITED_STYLE_PROPERTIES = new Set<string>([
  "color",
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "text-align",
  "text-transform",
  "text-indent",
  "white-space",
  "word-break",
  "overflow-wrap",
  "list-style-type",
  "list-style-position",
]);

/** Default values that do not need to be written into the clone. */
const THUMBNAIL_INLINE_STYLE_DEFAULTS: Readonly<Record<string, string>> = {
  display: "inline",
  position: "static",
  top: "auto",
  right: "auto",
  bottom: "auto",
  left: "auto",
  "z-index": "auto",
  overflow: "visible",
  "box-sizing": "content-box",
  "flex-direction": "row",
  "flex-wrap": "nowrap",
  "flex-grow": "0",
  "flex-shrink": "1",
  "flex-basis": "auto",
  "align-items": "normal",
  "align-self": "auto",
  "justify-content": "normal",
  gap: "normal",
  "grid-template-columns": "none",
  "grid-template-rows": "none",
  "table-layout": "auto",
  "border-collapse": "separate",
  "border-spacing": "0px",
  "background-color": "rgba(0, 0, 0, 0)",
  "border-top-width": "0px",
  "border-top-style": "none",
  "border-right-width": "0px",
  "border-right-style": "none",
  "border-bottom-width": "0px",
  "border-bottom-style": "none",
  "border-left-width": "0px",
  "border-left-style": "none",
  "padding-top": "0px",
  "padding-right": "0px",
  "padding-bottom": "0px",
  "padding-left": "0px",
  "margin-top": "0px",
  "margin-right": "0px",
  "margin-bottom": "0px",
  "margin-left": "0px",
  "letter-spacing": "normal",
  "word-spacing": "0px",
  "text-align": "start",
  "text-decoration-line": "none",
  "text-transform": "none",
  "text-indent": "0px",
  "white-space": "normal",
  "word-break": "normal",
  "overflow-wrap": "normal",
  "vertical-align": "baseline",
  "list-style-type": "disc",
  "list-style-position": "outside",
  opacity: "1",
};

const THUMBNAIL_IMAGE_DECODE_TIMEOUT_MS = 2_000;
const THUMBNAIL_SURFACE_DECODE_TIMEOUT_MS = 10_000;
const THUMBNAIL_TRANSPARENT_COLORS = new Set(["rgba(0, 0, 0, 0)", "transparent"]);

const THUMBNAIL_IMAGE_DOWNSCALE_MIN_DATA_URI_LENGTH = 32_768;
const THUMBNAIL_IMAGE_DOWNSCALE_MAX_DIMENSION_PX = 512;
const THUMBNAIL_IMAGE_JPEG_QUALITY = 0.78;
const THUMBNAIL_DIRECT_DEFAULT_FONT_FAMILY = "Calibri, Arial, sans-serif";
const THUMBNAIL_DIRECT_DEFAULT_TEXT_COLOR = "#111827";
const THUMBNAIL_DIRECT_TABLE_BORDER_COLOR = "#d1d5db";
const THUMBNAIL_DIRECT_IMAGE_BACKGROUND = "#f3f4f6";
const THUMBNAIL_DIRECT_MAX_ELEMENTS = 320;
const THUMBNAIL_DIRECT_MAX_TEXT_CHARS = 640;
const THUMBNAIL_DIRECT_MAX_LINES = 14;
const THUMBNAIL_DIRECT_MAX_LAYOUT_LINES = 80;

export type DocxPageThumbnailSnapshotTextAlign =
  | "left"
  | "center"
  | "right"
  | "justify";

export interface DocxPageThumbnailTextRunSnapshot {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  backgroundColor?: string;
  fontSizePx?: number;
  fontFamily?: string;
}

export interface DocxPageThumbnailParagraphSnapshot {
  kind: "paragraph";
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  align?: DocxPageThumbnailSnapshotTextAlign;
  backgroundColor?: string;
  lineHeightPx?: number;
  startLineIndex?: number;
  runs: readonly DocxPageThumbnailTextRunSnapshot[];
}

export interface DocxPageThumbnailImagePlaceholderSnapshot {
  kind: "image-placeholder";
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  backgroundColor?: string;
  borderColor?: string;
}

export interface DocxPageThumbnailImageSnapshot {
  kind: "image";
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  /** Renderable image source (data URI or URL) drawn into the thumbnail. */
  src: string;
  /** Painted if the image cannot be decoded/drawn (falls back to a box). */
  backgroundColor?: string;
  borderColor?: string;
}

export interface DocxPageThumbnailTableCellSnapshot {
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  backgroundColor?: string;
  runs?: readonly DocxPageThumbnailTextRunSnapshot[];
}

export interface DocxPageThumbnailTableSnapshot {
  kind: "table";
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  borderColor?: string;
  cells: readonly DocxPageThumbnailTableCellSnapshot[];
}

export type DocxPageThumbnailSnapshotElement =
  | DocxPageThumbnailParagraphSnapshot
  | DocxPageThumbnailImagePlaceholderSnapshot
  | DocxPageThumbnailImageSnapshot
  | DocxPageThumbnailTableSnapshot;

export interface DocxPageThumbnailRenderSnapshot {
  key: string;
  sourceWidthPx: number;
  sourceHeightPx: number;
  pageBackgroundColor?: string;
  elements: readonly DocxPageThumbnailSnapshotElement[];
}

function thumbnailSvgDataUri(svg: string): string {
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

function thumbnailComputedStyle(
  element: Element
): CSSStyleDeclaration | undefined {
  try {
    const view = element.ownerDocument?.defaultView;
    if (view && typeof view.getComputedStyle === "function") {
      return view.getComputedStyle(element);
    }
    if (typeof getComputedStyle === "function") {
      return getComputedStyle(element);
    }
  } catch {
    // Detached elements can refuse computed styles; the clone keeps whatever
    // inline styles it already had.
  }
  return undefined;
}

function thumbnailChromeClassNames(element: Element): string[] {
  const names: string[] = [];
  const classList = (element as HTMLElement).classList;
  if (!classList || typeof classList.contains !== "function") return names;
  for (const className of THUMBNAIL_EDITOR_CHROME_CLASS_NAMES) {
    if (classList.contains(className)) names.push(className);
  }
  return names;
}

function isDefaultThumbnailStyleValue(
  property: string,
  value: string
): boolean {
  if (property === "background-color") {
    return THUMBNAIL_TRANSPARENT_COLORS.has(value);
  }
  if (property === "border-spacing") {
    return value === "0px" || value === "0px 0px";
  }
  if (property === "min-width" || property === "min-height") {
    return value === "auto" || value === "0px";
  }
  return THUMBNAIL_INLINE_STYLE_DEFAULTS[property] === value;
}

function writeThumbnailComputedStyle(
  liveElement: Element,
  cloneElement: HTMLElement,
  computedStyle: CSSStyleDeclaration | undefined,
  parentStyle: CSSStyleDeclaration | undefined
): void {
  const chromeClassNames = thumbnailChromeClassNames(liveElement);
  if (chromeClassNames.length > 0 && cloneElement.classList) {
    cloneElement.classList.remove(...chromeClassNames);
  }
  if (!computedStyle) return;
  if (!cloneElement.style || typeof cloneElement.style.setProperty !== "function") {
    return;
  }

  for (const property of THUMBNAIL_INLINE_STYLE_PROPERTIES) {
    // The live computed background includes the active-paragraph / active-cell
    // tint. Dropping the class is not enough, so skip the background entirely
    // and let any document-supplied inline background survive.
    if (property === "background-color" && chromeClassNames.length > 0) {
      continue;
    }
    const value = computedStyle.getPropertyValue(property);
    if (!value) continue;
    if (
      THUMBNAIL_INHERITED_STYLE_PROPERTIES.has(property) &&
      parentStyle?.getPropertyValue(property) === value
    ) {
      continue;
    }
    if (isDefaultThumbnailStyleValue(property, value)) continue;
    cloneElement.style.setProperty(property, value);
  }
}

/**
 * Walks the live page and its clone in parallel, copying computed styles onto
 * the clone and dropping editor chrome. The live DOM is never mutated.
 */
function inlineThumbnailComputedStyles(
  liveElement: Element,
  cloneElement: HTMLElement,
  parentStyle?: CSSStyleDeclaration
): void {
  if (
    typeof liveElement.matches === "function" &&
    liveElement.matches(THUMBNAIL_EXCLUDED_CLONE_SELECTOR)
  ) {
    if (typeof cloneElement.remove === "function") cloneElement.remove();
    return;
  }

  const computedStyle = thumbnailComputedStyle(liveElement);
  writeThumbnailComputedStyle(
    liveElement,
    cloneElement,
    computedStyle,
    parentStyle
  );

  // Snapshot the child lists up front: removing an excluded node mutates the
  // clone's live `childNodes` collection while indexes still refer to the
  // original structure. Minimal non-DOM test doubles may not model children.
  if (!liveElement.childNodes || !cloneElement.childNodes) return;
  const liveChildren = Array.from(liveElement.childNodes);
  const cloneChildren = Array.from(cloneElement.childNodes);
  const childCount = Math.min(liveChildren.length, cloneChildren.length);
  for (let index = 0; index < childCount; index += 1) {
    const liveChild = liveChildren[index];
    const cloneChild = cloneChildren[index];
    if (!liveChild || !cloneChild) continue;
    if (liveChild.nodeType !== 1 || cloneChild.nodeType !== 1) continue;
    inlineThumbnailComputedStyles(
      liveChild as Element,
      cloneChild as HTMLElement,
      computedStyle
    );
  }
}

/**
 * True when an `<img>` source is a raster data URI large enough that swapping
 * it for a downscaled copy meaningfully shrinks the serialized page markup.
 */
export function thumbnailImageSourceQualifiesForDownscale(
  src: string
): boolean {
  return (
    src.length >= THUMBNAIL_IMAGE_DOWNSCALE_MIN_DATA_URI_LENGTH &&
    src.startsWith("data:image/") &&
    !src.startsWith("data:image/svg")
  );
}

/**
 * Decodes an image source, rejecting when the decoder does not settle inside
 * `timeoutMs`. A stalled decoder must never leave a thumbnail request pending
 * forever; callers fall back to a placeholder box instead.
 */
async function loadThumbnailImage(
  src: string,
  timeoutMs: number = THUMBNAIL_IMAGE_DECODE_TIMEOUT_MS
): Promise<HTMLImageElement> {
  const image = new Image();
  image.decoding = "async";

  const loaded = new Promise<HTMLImageElement>((resolve, reject) => {
    let settled = false;
    let timer: ReturnType<typeof setTimeout> | undefined;

    const cleanup = (): void => {
      image.onload = null;
      image.onerror = null;
      if (timer !== undefined) {
        clearTimeout(timer);
        timer = undefined;
      }
    };
    const succeed = (): void => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(image);
    };
    const fail = (error: Error): void => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(error);
    };

    image.onload = () => succeed();
    image.onerror = () => {
      fail(new Error("Failed to decode DOCX thumbnail image."));
    };
    if (Number.isFinite(timeoutMs) && timeoutMs > 0) {
      timer = setTimeout(() => {
        fail(new Error("Timed out decoding DOCX thumbnail image."));
      }, timeoutMs);
    }

    image.src = src;

    if (typeof image.decode === "function") {
      image.decode().then(succeed, () => {
        // Some engines reject decode() for sources they still paint correctly;
        // the load/error events above decide the final outcome.
      });
    }
  });

  return loaded;
}

async function downscaleThumbnailImageDataUri(
  src: string,
  decodeTimeoutMs?: number
): Promise<string | undefined> {
  if (typeof document === "undefined") {
    return undefined;
  }

  const image = await loadThumbnailImage(src, decodeTimeoutMs);
  const naturalWidth = image.naturalWidth || image.width;
  const naturalHeight = image.naturalHeight || image.height;
  if (!naturalWidth || !naturalHeight) {
    return undefined;
  }

  const scale =
    THUMBNAIL_IMAGE_DOWNSCALE_MAX_DIMENSION_PX /
    Math.max(naturalWidth, naturalHeight);
  if (scale >= 1) {
    return undefined;
  }

  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(naturalHeight * scale));
  const context = canvas.getContext("2d");
  if (!context) {
    return undefined;
  }

  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  context.drawImage(image, 0, 0, canvas.width, canvas.height);

  const sourceIsJpeg =
    src.startsWith("data:image/jpeg") || src.startsWith("data:image/jpg");
  const downscaled = sourceIsJpeg
    ? canvas.toDataURL("image/jpeg", THUMBNAIL_IMAGE_JPEG_QUALITY)
    : canvas.toDataURL("image/png");
  return downscaled.length < src.length ? downscaled : undefined;
}

/**
 * Returns a thumbnail-scale replacement for a large data-URI image source,
 * or `undefined` when the original should be kept. The default helper does
 * not retain module-level state; the owning thumbnail session caches the
 * resulting page surface with its own byte and entry limits.
 */
export function getDownscaledThumbnailImageDataUri(
  src: string,
  options: { decodeTimeoutMs?: number } = {}
): Promise<string | undefined> {
  return downscaleThumbnailImageDataUri(src, options.decodeTimeoutMs).catch(
    () => undefined
  );
}

/**
 * Decodes an image source for the direct snapshot painter. Decoded bitmaps are
 * intentionally scoped to one render call; the owning session retains only
 * the bounded final thumbnail surface.
 */
function getDecodedThumbnailImage(
  src: string
): Promise<HTMLImageElement | undefined> {
  return loadThumbnailImage(src).catch(() => undefined);
}

function directThumbnailPositivePx(value: number | undefined, fallback = 1): number {
  return Number.isFinite(value) && (value as number) > 0
    ? Math.max(1, Number(value))
    : fallback;
}

function setCanvasFillStyle(
  context: CanvasRenderingContext2D,
  color: string | undefined,
  fallback: string
): void {
  try {
    context.fillStyle = color || fallback;
  } catch {
    context.fillStyle = fallback;
  }
}

function setCanvasStrokeStyle(
  context: CanvasRenderingContext2D,
  color: string | undefined,
  fallback: string
): void {
  try {
    context.strokeStyle = color || fallback;
  } catch {
    context.strokeStyle = fallback;
  }
}

function directThumbnailFont(
  run: DocxPageThumbnailTextRunSnapshot | undefined,
  fallbackFontSizePx: number
): string {
  const fontSizePx = Math.max(
    6,
    Math.min(
      36,
      Math.round(
        directThumbnailPositivePx(run?.fontSizePx, fallbackFontSizePx)
      )
    )
  );
  const fontStyle = run?.italic ? "italic " : "";
  const fontWeight = run?.bold ? "700 " : "";
  return `${fontStyle}${fontWeight}${fontSizePx}px ${
    run?.fontFamily || THUMBNAIL_DIRECT_DEFAULT_FONT_FAMILY
  }`;
}

/**
 * Tokenizer for thumbnail text layout. A token is a hard break, a run of
 * whitespace, or a run of non-whitespace — the classes never mix, so a token's
 * kind is decided by its first character. Hoisted to module scope so the hot
 * layout loop does not allocate a fresh RegExp per run.
 */
const THUMBNAIL_DIRECT_TOKEN_REGEX =
  /(\r\n|\n|\t|[^\S\r\n\t]+|[^\s\r\n\t]+)/g;
const THUMBNAIL_DIRECT_LEADING_WHITESPACE_REGEX = /^\s/;
/**
 * Measures a token's advance width. The caller has already assigned
 * `context.font`; keeping this stateless prevents one document session from
 * retaining another document's text indefinitely.
 */
function measureDirectThumbnailToken(
  context: CanvasRenderingContext2D,
  font: string,
  text: string
): number {
  if (!text) {
    return 0;
  }
  void font;
  return context.measureText(text).width;
}

interface DirectThumbnailTextSegment {
  run: DocxPageThumbnailTextRunSnapshot;
  text: string;
  widthPx: number;
  /** The run's resolved canvas font shorthand, reused by the draw pass. */
  font: string;
}

interface DirectThumbnailTextLine {
  segments: DirectThumbnailTextSegment[];
  widthPx: number;
}

function appendDirectThumbnailTextLine(
  lines: DirectThumbnailTextLine[],
  currentSegments: DirectThumbnailTextSegment[],
  currentWidthPx: number
): DirectThumbnailTextLine {
  const line = {
    segments: currentSegments,
    widthPx: currentWidthPx,
  };
  lines.push(line);
  return line;
}

function layoutDirectThumbnailTextRuns(params: {
  context: CanvasRenderingContext2D;
  runs: readonly DocxPageThumbnailTextRunSnapshot[];
  widthPx: number;
  fallbackFontSizePx: number;
  maxLineCount: number;
}): DirectThumbnailTextLine[] {
  const { context, runs, fallbackFontSizePx } = params;
  const widthPx = Math.max(1, params.widthPx);
  const maxLineCount = Math.max(
    1,
    Math.min(THUMBNAIL_DIRECT_MAX_LAYOUT_LINES, params.maxLineCount)
  );
  const lines: DirectThumbnailTextLine[] = [];
  let currentSegments: DirectThumbnailTextSegment[] = [];
  let currentWidthPx = 0;
  let remainingChars = THUMBNAIL_DIRECT_MAX_TEXT_CHARS;

  const flushLine = (): void => {
    appendDirectThumbnailTextLine(lines, currentSegments, currentWidthPx);
    currentSegments = [];
    currentWidthPx = 0;
  };

  for (const run of runs) {
    if (lines.length >= maxLineCount || remainingChars <= 0) {
      break;
    }

    const text = run.text.slice(0, remainingChars);
    remainingChars -= text.length;
    // Set the run's font once for the whole run; every token below measures
    // against it (the cache and the wrap test assume context.font === runFont).
    const runFont = directThumbnailFont(run, fallbackFontSizePx);
    context.font = runFont;
    const tokens = text.match(THUMBNAIL_DIRECT_TOKEN_REGEX) ?? [];
    for (const token of tokens) {
      if (lines.length >= maxLineCount) {
        break;
      }
      if (token === "\n" || token === "\r\n") {
        flushLine();
        continue;
      }

      const drawableToken = token === "\t" ? "    " : token;
      const tokenWidthPx = measureDirectThumbnailToken(
        context,
        runFont,
        drawableToken
      );
      const tokenIsWhitespace =
        THUMBNAIL_DIRECT_LEADING_WHITESPACE_REGEX.test(drawableToken);
      if (
        currentSegments.length > 0 &&
        currentWidthPx + tokenWidthPx > widthPx &&
        !tokenIsWhitespace
      ) {
        flushLine();
      }
      if (currentSegments.length === 0 && tokenIsWhitespace) {
        continue;
      }
      currentSegments.push({
        run,
        text: drawableToken,
        widthPx: tokenWidthPx,
        font: runFont,
      });
      currentWidthPx += tokenWidthPx;
    }
  }

  if (currentSegments.length > 0 || lines.length === 0) {
    flushLine();
  }

  return lines;
}

function directThumbnailAlignedX(params: {
  xPx: number;
  widthPx: number;
  lineWidthPx: number;
  align?: DocxPageThumbnailSnapshotTextAlign;
}): number {
  const { xPx, widthPx, lineWidthPx, align } = params;
  if (align === "center") {
    return xPx + Math.max(0, (widthPx - lineWidthPx) / 2);
  }
  if (align === "right") {
    return xPx + Math.max(0, widthPx - lineWidthPx);
  }
  return xPx;
}

function drawDirectThumbnailTextRuns(params: {
  context: CanvasRenderingContext2D;
  runs: readonly DocxPageThumbnailTextRunSnapshot[];
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  align?: DocxPageThumbnailSnapshotTextAlign;
  lineHeightPx?: number;
  startLineIndex?: number;
}): void {
  const {
    context,
    runs,
    xPx,
    yPx,
    widthPx,
    heightPx,
    align,
    startLineIndex,
  } = params;
  const safeWidthPx = Math.max(1, widthPx);
  const safeHeightPx = Math.max(1, heightPx);
  const fallbackFontSizePx = Math.max(
    7,
    Math.min(
      20,
      Math.round(
        runs.find((run) => Number.isFinite(run.fontSizePx))?.fontSizePx ?? 12
      )
    )
  );
  const lineHeightPx = Math.max(
    fallbackFontSizePx + 1,
    Math.round(params.lineHeightPx ?? fallbackFontSizePx * 1.25)
  );
  const skippedLineCount = Math.max(
    0,
    Math.min(THUMBNAIL_DIRECT_MAX_LAYOUT_LINES - 1, Math.trunc(startLineIndex ?? 0))
  );
  const visibleLineCount = Math.max(
    1,
    Math.min(
      THUMBNAIL_DIRECT_MAX_LINES,
      Math.ceil(safeHeightPx / Math.max(1, lineHeightPx)) + 1
    )
  );
  const lines = layoutDirectThumbnailTextRuns({
    context,
    runs,
    widthPx: safeWidthPx,
    fallbackFontSizePx,
    maxLineCount: skippedLineCount + visibleLineCount,
  }).slice(skippedLineCount, skippedLineCount + visibleLineCount);

  context.save();
  context.beginPath();
  context.rect(xPx, yPx, safeWidthPx, safeHeightPx);
  context.clip();
  context.textBaseline = "alphabetic";

  // The layout pass already resolved each segment's font and width; only write
  // context.font when it actually changes (one assignment per run, not token).
  let lastAppliedFont: string | undefined;
  lines.forEach((line, lineIndex) => {
    const lineTopPx = yPx + lineIndex * lineHeightPx;
    if (lineTopPx > yPx + safeHeightPx) {
      return;
    }
    let cursorXPx = directThumbnailAlignedX({
      xPx,
      widthPx: safeWidthPx,
      lineWidthPx: line.widthPx,
      align,
    });
    const baselineYPx =
      lineTopPx + Math.max(1, Math.round(lineHeightPx * 0.78));
    line.segments.forEach((segment) => {
      const segmentWidthPx = segment.widthPx;
      if (segment.run.backgroundColor) {
        setCanvasFillStyle(context, segment.run.backgroundColor, "transparent");
        context.fillRect(cursorXPx, lineTopPx + 1, segmentWidthPx, lineHeightPx);
      }
      if (segment.font !== lastAppliedFont) {
        context.font = segment.font;
        lastAppliedFont = segment.font;
      }
      setCanvasFillStyle(
        context,
        segment.run.color,
        THUMBNAIL_DIRECT_DEFAULT_TEXT_COLOR
      );
      context.fillText(segment.text, cursorXPx, baselineYPx);
      cursorXPx += segmentWidthPx;
    });
  });

  context.restore();
}

function drawDirectThumbnailParagraph(
  context: CanvasRenderingContext2D,
  paragraph: DocxPageThumbnailParagraphSnapshot
): void {
  const xPx = Math.round(paragraph.xPx);
  const yPx = Math.round(paragraph.yPx);
  const widthPx = Math.max(1, Math.round(paragraph.widthPx));
  const heightPx = Math.max(1, Math.round(paragraph.heightPx));
  if (paragraph.backgroundColor) {
    setCanvasFillStyle(context, paragraph.backgroundColor, "transparent");
    context.fillRect(xPx, yPx, widthPx, heightPx);
  }
  drawDirectThumbnailTextRuns({
    context,
    runs: paragraph.runs,
    xPx: xPx + 1,
    yPx,
    widthPx: Math.max(1, widthPx - 2),
    heightPx,
    align: paragraph.align,
    lineHeightPx: paragraph.lineHeightPx,
    startLineIndex: paragraph.startLineIndex,
  });
}

function drawDirectThumbnailImagePlaceholder(
  context: CanvasRenderingContext2D,
  image: DocxPageThumbnailImagePlaceholderSnapshot,
  hairlineSourcePx: number
): void {
  const xPx = Math.round(image.xPx);
  const yPx = Math.round(image.yPx);
  const widthPx = Math.max(1, Math.round(image.widthPx));
  const heightPx = Math.max(1, Math.round(image.heightPx));
  setCanvasFillStyle(
    context,
    image.backgroundColor,
    THUMBNAIL_DIRECT_IMAGE_BACKGROUND
  );
  context.fillRect(xPx, yPx, widthPx, heightPx);
  setCanvasStrokeStyle(
    context,
    image.borderColor,
    THUMBNAIL_DIRECT_TABLE_BORDER_COLOR
  );
  context.lineWidth = hairlineSourcePx;
  context.strokeRect(xPx, yPx, widthPx, heightPx);
}

function drawDirectThumbnailImage(
  context: CanvasRenderingContext2D,
  image: DocxPageThumbnailImageSnapshot,
  decoded: HTMLImageElement | undefined,
  hairlineSourcePx: number
): void {
  const xPx = Math.round(image.xPx);
  const yPx = Math.round(image.yPx);
  const widthPx = Math.max(1, Math.round(image.widthPx));
  const heightPx = Math.max(1, Math.round(image.heightPx));
  if (decoded && decoded.naturalWidth > 0 && decoded.naturalHeight > 0) {
    try {
      context.drawImage(decoded, xPx, yPx, widthPx, heightPx);
      return;
    } catch {
      // Drawing can throw for sources the decoder accepted but the canvas
      // refuses (e.g. cross-origin taint); fall back to the placeholder box.
    }
  }
  setCanvasFillStyle(
    context,
    image.backgroundColor,
    THUMBNAIL_DIRECT_IMAGE_BACKGROUND
  );
  context.fillRect(xPx, yPx, widthPx, heightPx);
  setCanvasStrokeStyle(
    context,
    image.borderColor,
    THUMBNAIL_DIRECT_TABLE_BORDER_COLOR
  );
  context.lineWidth = hairlineSourcePx;
  context.strokeRect(xPx, yPx, widthPx, heightPx);
}

function drawDirectThumbnailTable(
  context: CanvasRenderingContext2D,
  table: DocxPageThumbnailTableSnapshot,
  hairlineSourcePx: number
): void {
  const tableXPx = Math.round(table.xPx);
  const tableYPx = Math.round(table.yPx);
  const tableWidthPx = Math.max(1, Math.round(table.widthPx));
  const tableHeightPx = Math.max(1, Math.round(table.heightPx));

  context.save();
  context.beginPath();
  context.rect(tableXPx, tableYPx, tableWidthPx, tableHeightPx);
  context.clip();
  setCanvasStrokeStyle(
    context,
    table.borderColor,
    THUMBNAIL_DIRECT_TABLE_BORDER_COLOR
  );
  context.lineWidth = hairlineSourcePx;

  table.cells.forEach((cell) => {
    const xPx = tableXPx + Math.round(cell.xPx);
    const yPx = tableYPx + Math.round(cell.yPx);
    const widthPx = Math.max(1, Math.round(cell.widthPx));
    const heightPx = Math.max(1, Math.round(cell.heightPx));
    if (cell.backgroundColor) {
      setCanvasFillStyle(context, cell.backgroundColor, "transparent");
      context.fillRect(xPx, yPx, widthPx, heightPx);
    }
    context.strokeRect(xPx, yPx, widthPx, heightPx);
    if (cell.runs?.length) {
      drawDirectThumbnailTextRuns({
        context,
        runs: cell.runs,
        xPx: xPx + 3,
        yPx: yPx + 2,
        widthPx: Math.max(1, widthPx - 6),
        heightPx: Math.max(1, heightPx - 4),
        lineHeightPx: 13,
      });
    }
  });

  context.restore();
}

/**
 * Paints a thumbnail directly from a layout/model snapshot. This skips the
 * expensive DOM clone -> SVG foreignObject -> image decode path and is intended
 * as the default fast path for virtualized thumbnail rails.
 *
 * Embedded images are decoded (and cached) before painting so image-dominated
 * pages — e.g. scanned `.doc` documents where every page is a single full-page
 * bitmap — render their real content instead of an empty placeholder box.
 * Pages with no images resolve without awaiting, keeping the text fast path
 * effectively synchronous.
 */
export async function renderDocxThumbnailSnapshotSurface(params: {
  snapshot: DocxPageThumbnailRenderSnapshot;
  widthPx: number;
  heightPx: number;
  pixelWidthPx: number;
  pixelHeightPx: number;
}): Promise<HTMLCanvasElement> {
  if (typeof document === "undefined") {
    throw new Error("DOCX thumbnails require a browser environment.");
  }

  const sourceWidthPx = directThumbnailPositivePx(params.snapshot.sourceWidthPx);
  const sourceHeightPx = directThumbnailPositivePx(params.snapshot.sourceHeightPx);
  const pixelWidthPx = Math.max(1, Math.round(params.pixelWidthPx));
  const pixelHeightPx = Math.max(1, Math.round(params.pixelHeightPx));
  const surface = document.createElement("canvas");
  surface.width = pixelWidthPx;
  surface.height = pixelHeightPx;
  const context = surface.getContext("2d");
  if (!context) {
    throw new Error("2D canvas context is unavailable for DOCX thumbnails.");
  }

  const elements = params.snapshot.elements.slice(0, THUMBNAIL_DIRECT_MAX_ELEMENTS);
  const decodedImagesBySrc = new Map<string, HTMLImageElement | undefined>();
  const imageSrcs = new Set<string>();
  for (const element of elements) {
    if (element.kind === "image" && element.src) {
      imageSrcs.add(element.src);
    }
  }
  if (imageSrcs.size > 0) {
    await Promise.all(
      Array.from(imageSrcs, async (src) => {
        decodedImagesBySrc.set(src, await getDecodedThumbnailImage(src));
      })
    );
  }

  const scaleX = pixelWidthPx / sourceWidthPx;
  const scaleY = pixelHeightPx / sourceHeightPx;
  const hairlineSourcePx = Math.max(0.75, 1 / Math.max(scaleX, scaleY));
  context.setTransform(scaleX, 0, 0, scaleY, 0, 0);
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  setCanvasFillStyle(
    context,
    params.snapshot.pageBackgroundColor,
    "#ffffff"
  );
  context.fillRect(0, 0, sourceWidthPx, sourceHeightPx);

  elements.forEach((element) => {
    switch (element.kind) {
      case "paragraph":
        drawDirectThumbnailParagraph(context, element);
        break;
      case "image-placeholder":
        drawDirectThumbnailImagePlaceholder(
          context,
          element,
          hairlineSourcePx
        );
        break;
      case "image":
        drawDirectThumbnailImage(
          context,
          element,
          decodedImagesBySrc.get(element.src),
          hairlineSourcePx
        );
        break;
      case "table":
        drawDirectThumbnailTable(context, element, hairlineSourcePx);
        break;
    }
  });

  return surface;
}

async function buildDocxThumbnailSvgMarkup(params: {
  pageElement: HTMLElement;
  sourceWidthPx: number;
  sourceHeightPx: number;
  widthPx: number;
  heightPx: number;
}): Promise<string> {
  const { pageElement, sourceWidthPx, sourceHeightPx, widthPx, heightPx } =
    params;
  const clone = pageElement.cloneNode(true) as HTMLElement;
  // Scoped component CSS never reaches the SVG foreignObject, so carry the
  // computed values that shape the page into the clone as inline styles while
  // removing editor-only chrome at the same time.
  inlineThumbnailComputedStyles(pageElement, clone);

  const cloneImages = Array.from(clone.querySelectorAll("img"));
  await Promise.all(
    cloneImages.map(async (cloneImage) => {
      const src = cloneImage.getAttribute("src");
      if (!src || !thumbnailImageSourceQualifiesForDownscale(src)) {
        return;
      }

      const downscaled = await getDownscaledThumbnailImageDataUri(src);
      if (downscaled) {
        cloneImage.setAttribute("src", downscaled);
      }
    })
  );

  const scaleX = widthPx / sourceWidthPx;
  const scaleY = heightPx / sourceHeightPx;
  const serializedPage = new XMLSerializer().serializeToString(clone);
  return `
    <svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">
      <foreignObject x="0" y="0" width="100%" height="100%">
        <div xmlns="http://www.w3.org/1999/xhtml" style="width:${widthPx}px;height:${heightPx}px;overflow:hidden;">
          <div style="width:${sourceWidthPx}px;height:${sourceHeightPx}px;transform-origin:top left;transform:scale(${scaleX}, ${scaleY});">
            ${serializedPage}
          </div>
        </div>
      </foreignObject>
    </svg>
  `;
}

/**
 * Rasterizes a live page surface element to a fresh offscreen surface canvas
 * at the requested pixel resolution. The surface is cacheable and can be
 * blitted to any number of target canvases with {@link blitDocxThumbnailSurface}.
 */
export async function rasterizeDocxThumbnailSurface(params: {
  pageElement: HTMLElement;
  sourceWidthPx: number;
  sourceHeightPx: number;
  widthPx: number;
  heightPx: number;
  pixelWidthPx: number;
  pixelHeightPx: number;
}): Promise<HTMLCanvasElement> {
  if (typeof window === "undefined" || typeof XMLSerializer === "undefined") {
    throw new Error("DOCX thumbnails require a browser environment.");
  }

  const safeSourceWidthPx = Math.max(1, Math.round(params.sourceWidthPx));
  const safeSourceHeightPx = Math.max(1, Math.round(params.sourceHeightPx));
  const svgMarkup = await buildDocxThumbnailSvgMarkup({
    pageElement: params.pageElement,
    sourceWidthPx: safeSourceWidthPx,
    sourceHeightPx: safeSourceHeightPx,
    widthPx: params.widthPx,
    heightPx: params.heightPx,
  });
  // Load the foreignObject SVG via a data: URL rather than a blob: URL.
  // WebKit/Safari still taints a canvas drawn from a blob:-backed SVG image
  // (bug 156176), but a data:-URI SVG is explicitly exempted (bug 180301), and
  // Chrome/Firefox never taint either way. Keeping the canvas clean lets
  // callers run toDataURL()/toBlob() for client-side thumbnail export.
  const image = await loadThumbnailImage(
    thumbnailSvgDataUri(svgMarkup),
    THUMBNAIL_SURFACE_DECODE_TIMEOUT_MS
  );

  const surface = document.createElement("canvas");
  surface.width = Math.max(1, Math.round(params.pixelWidthPx));
  surface.height = Math.max(1, Math.round(params.pixelHeightPx));
  const context = surface.getContext("2d");
  if (!context) {
    throw new Error("2D canvas context is unavailable for DOCX thumbnails.");
  }

  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  context.drawImage(image, 0, 0, surface.width, surface.height);
  return surface;
}

/** Copies a cached thumbnail surface onto a target canvas. */
export function blitDocxThumbnailSurface(
  surface: HTMLCanvasElement,
  canvas: HTMLCanvasElement,
  resolution: {
    widthPx: number;
    heightPx: number;
    pixelWidthPx: number;
    pixelHeightPx: number;
  }
): void {
  const pixelWidth = Math.max(1, Math.round(resolution.pixelWidthPx));
  const pixelHeight = Math.max(1, Math.round(resolution.pixelHeightPx));
  const cssWidth = `${Math.max(1, Math.round(resolution.widthPx))}px`;
  const cssHeight = `${Math.max(1, Math.round(resolution.heightPx))}px`;

  // Writing canvas.width/height resets (and clears) the backing store even when
  // the value is unchanged, so only do it on an actual size change.
  let bufferResized = false;
  if (canvas.width !== pixelWidth) {
    canvas.width = pixelWidth;
    bufferResized = true;
  }
  if (canvas.height !== pixelHeight) {
    canvas.height = pixelHeight;
    bufferResized = true;
  }
  if (canvas.style.width !== cssWidth) {
    canvas.style.width = cssWidth;
  }
  if (canvas.style.height !== cssHeight) {
    canvas.style.height = cssHeight;
  }

  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("2D canvas context is unavailable for DOCX thumbnails.");
  }

  context.setTransform(1, 0, 0, 1, 0, 0);
  // A resize already cleared the buffer; otherwise clear stale pixels first.
  if (!bufferResized) {
    context.clearRect(0, 0, canvas.width, canvas.height);
  }
  context.drawImage(surface, 0, 0, canvas.width, canvas.height);
}
