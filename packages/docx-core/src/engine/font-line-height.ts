/**
 * Platform-neutral font vertical metrics.
 *
 * The real measurement path (`viewer/line-metrics`) asks Canvas for
 * `fontBoundingBoxAscent/Descent`. Headless hosts and the coarse legacy
 * `layout-engine` have no canvas, so they use the same em ratios Word's
 * single-spacing fallback uses: 0.9 em above the baseline and 0.25 em below it.
 * Keeping the constants in a pure module lets the core graph use the font
 * metrics without importing a browser capability.
 */

export const FALLBACK_FONT_ASCENT_EM = 0.9;
export const FALLBACK_FONT_DESCENT_EM = 0.25;

/** Line-box height of a font size under the platform-neutral metrics. */
export function estimateFontLineHeightPx(fontSizePx: number): number {
  const safeFontSizePx =
    Number.isFinite(fontSizePx) && fontSizePx > 0 ? fontSizePx : 12;
  return Math.max(
    1,
    Math.round(safeFontSizePx * (FALLBACK_FONT_ASCENT_EM + FALLBACK_FONT_DESCENT_EM))
  );
}
