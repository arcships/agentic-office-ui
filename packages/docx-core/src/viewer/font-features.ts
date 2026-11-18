/**
 * OOXML ligature / contextual-alternate mapping and the canvas used for text
 * measurement.
 *
 * Word only applies ligatures the run asks for (`w14:ligatures`), and the
 * default is `none`. Keeping the same feature string on the measurement canvas
 * and on the rendered span is what makes the estimate and the DOM agree: a
 * `liga` pair that is 1 px narrower on one side must be 1 px narrower on the
 * other.
 */

import type { TextStyle } from "../engine/types";

export interface DocxTypographyOptions {
  /**
   * Host escape hatch: `false` disables every OpenType feature even when the
   * document asks for them (used by diagnostics and deterministic snapshots).
   */
  enableOpenTypeFeatures?: boolean;
  adjustLineHeightInTable?: boolean | null;
}

export const DISABLED_DOCX_FONT_FEATURE_SETTINGS =
  '"liga" 0, "clig" 0, "dlig" 0, "hlig" 0, "calt" 0';

/**
 * `font-feature-settings` for one run. `ligatures` may list several families
 * (`standardContextual`, `all`, …), so every token is matched by name.
 */
export function resolveDocxFontFeatureSettings(
  style?: TextStyle,
  options?: DocxTypographyOptions
): string {
  if (options?.enableOpenTypeFeatures === false) {
    return DISABLED_DOCX_FONT_FEATURE_SETTINGS;
  }
  const ligatures = style?.ligatures?.toLowerCase() ?? "none";
  const enabled = (kind: string): number =>
    ligatures === "all" || ligatures.includes(kind) ? 1 : 0;
  return `"liga" ${enabled("standard")}, "clig" ${enabled(
    "contextual"
  )}, "dlig" ${enabled("discretional")}, "hlig" ${enabled(
    "historical"
  )}, "calt" ${style?.contextualAlternates === true ? 1 : 0}`;
}

type CanvasFeatureContext = {
  canvas?: HTMLCanvasElement | OffscreenCanvas;
  textRendering?: string;
};

const contextsByDocument = new WeakMap<Document, CanvasRenderingContext2D>();
let fallbackContext: OffscreenCanvasRenderingContext2D | undefined;
const settingsByCanvas = new WeakMap<
  HTMLCanvasElement,
  { requested: string; serialized: string }
>();
const connectedCanvases = new WeakSet<HTMLCanvasElement>();

function connectMeasurementCanvas(canvas: HTMLCanvasElement): void {
  if (connectedCanvases.has(canvas) || canvas.isConnected) {
    connectedCanvases.add(canvas);
    return;
  }
  const root = canvas.ownerDocument?.documentElement;
  if (!root) {
    return;
  }
  canvas.setAttribute("aria-hidden", "true");
  canvas.style.all = "initial";
  canvas.style.position = "fixed";
  canvas.style.left = "0";
  canvas.style.top = "0";
  canvas.style.width = "0";
  canvas.style.height = "0";
  canvas.style.visibility = "hidden";
  canvas.style.pointerEvents = "none";
  root.appendChild(canvas);
  connectedCanvases.add(canvas);
}

/**
 * The one canvas used for DOCX text measurement. DOM canvas first (font
 * feature settings can be applied through its CSSOM), then an offscreen
 * context for Worker hosts that have no document.
 */
export function createDocxTextMeasurementContext():
  | CanvasRenderingContext2D
  | OffscreenCanvasRenderingContext2D
  | undefined {
  if (typeof document !== "undefined") {
    const cached = contextsByDocument.get(document);
    if (cached) return cached;
    try {
      const canvas = document.createElement("canvas");
      canvas.width = 1;
      canvas.height = 1;
      const context = canvas.getContext("2d");
      if (context) {
        connectMeasurementCanvas(canvas);
        contextsByDocument.set(document, context);
        return context;
      }
    } catch {
      // Nonvisual hosts can still provide an offscreen measurement context.
    }
    // A DOM host with a broken canvas keeps the DOM span fallback instead of
    // silently measuring on an unrelated offscreen context.
    return undefined;
  }
  if (fallbackContext) return fallbackContext;
  if (typeof OffscreenCanvas !== "undefined") {
    fallbackContext = new OffscreenCanvas(1, 1).getContext("2d") ?? undefined;
  }
  return fallbackContext;
}

/**
 * Apply `fontFeatureSettings` to a measurement context. A DOM canvas takes the
 * value through `style.fontFeatureSettings`, which the canvas text shaper
 * honours; offscreen contexts fall back to `textRendering` heuristics.
 */
export function setCanvasFontFeatureSettings(
  context: CanvasFeatureContext,
  settings?: string
): void {
  const canvas = context.canvas;
  if (
    canvas &&
    typeof HTMLCanvasElement !== "undefined" &&
    canvas instanceof HTMLCanvasElement
  ) {
    connectMeasurementCanvas(canvas);
    const value = settings ?? "normal";
    const previous = settingsByCanvas.get(canvas);
    if (
      previous?.requested !== value ||
      previous.serialized !== canvas.style.fontFeatureSettings
    ) {
      canvas.style.fontFeatureSettings = value;
      settingsByCanvas.set(canvas, {
        requested: value,
        serialized: canvas.style.fontFeatureSettings,
      });
    }
    if (typeof context.textRendering === "string") {
      context.textRendering = "auto";
    }
    return;
  }
  if (typeof context.textRendering === "string") {
    context.textRendering =
      settings === DISABLED_DOCX_FONT_FEATURE_SETTINGS
        ? "optimizeSpeed"
        : "auto";
  }
}
