// Document layout defaults that shift every paragraph: the compatibility mode
// (which decides whether Word's automatic paragraph spacing applies) and the
// default tab stop (the fallback tab grid). Both live in word/settings.xml and
// are resolved once per model into `DocumentLayoutDefaults`.
//
// Upstream: `tab-layout-context.ts:25-49` — `resolveDefaultTabStopPx(
// model.metadata.defaultTabStopTwips)` and the wrapped-left-tab refit.

import type {
  DocumentCompatibilitySettings,
  ParagraphTabStop,
} from "./types";

/** CSS pixels per twip (1440 twips per inch, 96 pixels per inch). */
const TWIPS_PER_PIXEL = 15;

/**
 * Whole, non-negative CSS pixels for a twips measurement — the same conversion
 * the model/renderer helpers use for every twips field.
 */
export function twipsToWholePixels(twips?: number): number {
  if (!Number.isFinite(twips)) {
    return 0;
  }
  return Math.max(0, Math.round((twips as number) / TWIPS_PER_PIXEL));
}

/** Fractional CSS pixels for a twips measurement (no rounding). */
export function twipsToPixelsExact(twips: number): number {
  return twips / TWIPS_PER_PIXEL;
}

/** Word's own `w:defaultTabStop` when the document names none: 720 twips. */
export const WORD_DEFAULT_TAB_STOP_TWIPS = 720;

/** 720 twips (0.5") in CSS pixels — Word's default tab grid. */
export const DEFAULT_TAB_STOP_PX = twipsToPixelsExact(
  WORD_DEFAULT_TAB_STOP_TWIPS
);

/**
 * Compatibility modes whose Word version honours HTML automatic paragraph
 * spacing (`w:beforeAutospacing` / `w:afterAutospacing`): Word 2007 (12) and
 * Word 2013+ (15). Modes 13/14 ignore the attributes entirely.
 */
const AUTO_SPACING_COMPATIBILITY_MODES = new Set([12, 15]);

export interface DocumentLayoutDefaults {
  /** `w:compatSetting[@w:name='compatibilityMode']`. */
  compatibilityMode?: number;
  /** `w:defaultTabStop/@w:val`, in twips. */
  defaultTabStopTwips?: number;
}

export interface DocumentLayoutDefaultsSource {
  compatibility?: DocumentCompatibilitySettings;
  defaultTabStopTwips?: number;
}

const EMPTY_LAYOUT_DEFAULTS: DocumentLayoutDefaults = Object.freeze({});

/**
 * Reads the layout defaults a document carries on its metadata. Kept stateless:
 * an edit may replace `metadata.compatibility` in place, and layout must pick
 * the change up on the next pass.
 */
export function resolveDocumentLayoutDefaults(
  metadata?: DocumentLayoutDefaultsSource
): DocumentLayoutDefaults {
  const compatibilityMode = metadata?.compatibility?.compatibilityMode;
  const defaultTabStopTwips = metadata?.defaultTabStopTwips;
  if (compatibilityMode === undefined && defaultTabStopTwips === undefined) {
    return EMPTY_LAYOUT_DEFAULTS;
  }
  return { compatibilityMode, defaultTabStopTwips };
}

/** True when the document's Word version applies HTML auto spacing. */
export function documentHonoursAutoSpacing(
  compatibilityMode?: number
): boolean {
  return (
    compatibilityMode !== undefined &&
    AUTO_SPACING_COMPATIBILITY_MODES.has(compatibilityMode)
  );
}

/**
 * Effective default tab stop in CSS pixels. Documents that never set
 * `w:defaultTabStop` keep Word's 720-twip grid.
 */
export function resolveDefaultTabStopPx(
  defaultTabStopTwips?: number
): number {
  if (
    !Number.isFinite(defaultTabStopTwips) ||
    (defaultTabStopTwips as number) <= 0
  ) {
    return DEFAULT_TAB_STOP_PX;
  }
  return Math.max(1, twipsToWholePixels(defaultTabStopTwips));
}

/**
 * Defaults that still apply where paragraph margins are not rendered: table
 * cells (`.docx-table-cell-paragraph`) and note bodies both pin their
 * paragraphs to zero margin, so Word's automatic paragraph spacing is
 * deliberately not measured there — the measured box would otherwise be taller
 * than the rendered one. The tab grid still applies.
 */
export function tabGridLayoutDefaults(
  defaults?: DocumentLayoutDefaults
): DocumentLayoutDefaults | undefined {
  if (defaults?.defaultTabStopTwips === undefined) {
    return undefined;
  }
  return { defaultTabStopTwips: defaults.defaultTabStopTwips };
}

/**
 * Cache discriminator for everything a height/line-count estimate depends on
 * beyond the paragraph itself. Returns an empty string for documents whose
 * defaults cannot change an estimate, so those keep the plain numeric keys.
 */
export function documentLayoutCacheKey(
  defaults?: DocumentLayoutDefaults
): string {
  if (!defaults) {
    return "";
  }
  const keyParts: string[] = [];
  if (documentHonoursAutoSpacing(defaults.compatibilityMode)) {
    // Only the Word versions that read the flags change any estimate.
    keyParts.push(`a${defaults.compatibilityMode}`);
  }
  const defaultTabStopPx = resolveDefaultTabStopPx(defaults.defaultTabStopTwips);
  if (defaultTabStopPx !== DEFAULT_TAB_STOP_PX) {
    keyParts.push(`t${defaultTabStopPx}`);
  }
  return keyParts.join(":");
}

/**
 * A left tab that does not fit the current line wraps, and on the new line it
 * re-aligns to the first default tab stop instead of collapsing (Word's
 * wrapped-tab behaviour). Centre/right/decimal stops anchor text to an absolute
 * offset measured from the line origin, so their layout is kept as-is.
 */
export function allowsWrappedLeftTabRefit(
  tabStops?: ParagraphTabStop[]
): boolean {
  return (tabStops ?? []).every(
    (tabStop) =>
      tabStop.alignment === undefined ||
      tabStop.alignment === "left" ||
      tabStop.alignment === "bar"
  );
}
