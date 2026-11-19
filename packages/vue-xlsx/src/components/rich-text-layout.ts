/**
 * Layout helpers for rich-text cells.
 *
 * A rich cell stores several runs with their own fonts, so its text width is
 * the sum of the per-run widths, not the width of the concatenated string in
 * the cell's base font. Wrapping and truncation run over those per-run widths:
 * a smaller superscript marker or a bold word measures differently from the
 * same characters at the cell font, and measuring the plain string would clip
 * or spill the line.
 */
import type { XlsxCellTextRun, XlsxCellTextRunFont } from "@arcships/xlsx-core";

/** Scale applied to a superscript/subscript run, shared by canvas and DOM. */
export const RICH_TEXT_VERTICAL_ALIGN_SCALE = 0.7;
/** Baseline shift of a superscript run, in ems of the cell font size. */
export const RICH_TEXT_SUPERSCRIPT_SHIFT_EM = 0.42;
/** Baseline shift of a subscript run, in ems of the cell font size. */
export const RICH_TEXT_SUBSCRIPT_SHIFT_EM = 0.16;

/** One drawable run after wrapping: a slice of an input run. */
export interface RichTextSegment {
  /** Run font of the source run, copied so callers can resolve per segment. */
  font?: XlsxCellTextRunFont;
  text: string;
}

/** Measures one run's text with that run's own font applied. */
export type RichTextMeasure = (text: string, font: XlsxCellTextRunFont | undefined) => number;

/** `true` for the run's `vertAlign` value. */
export function isRichTextVerticalAlign(
  font: XlsxCellTextRunFont | undefined,
  value: "superscript" | "subscript",
): boolean {
  return (font?.verticalAlign ?? "").toLocaleLowerCase() === value;
}

/** Font-size multiplier of a run: superscript/subscript runs shrink. */
export function richTextFontScale(font: XlsxCellTextRunFont | undefined): number {
  return isRichTextVerticalAlign(font, "superscript") || isRichTextVerticalAlign(font, "subscript")
    ? RICH_TEXT_VERTICAL_ALIGN_SCALE
    : 1;
}

/**
 * Baseline offset of a run in pixels, positive downwards.
 *
 * @param fontSizePx Cell font size the run is drawn against.
 */
export function richTextBaselineOffset(font: XlsxCellTextRunFont | undefined, fontSizePx: number): number {
  if (isRichTextVerticalAlign(font, "superscript")) {
    return -fontSizePx * RICH_TEXT_SUPERSCRIPT_SHIFT_EM;
  }
  if (isRichTextVerticalAlign(font, "subscript")) {
    return fontSizePx * RICH_TEXT_SUBSCRIPT_SHIFT_EM;
  }
  return 0;
}

/**
 * Breaks runs into display lines the way `wrapCellText` breaks a plain string.
 *
 * Hard newlines inside a run split the line unconditionally; without wrap text
 * (`maxLines <= 1`) only the first hard line shows and it is truncated with an
 * ellipsis when it does not fit. Every character is measured with the font of
 * the run that owns it, so a narrowed superscript run wraps where it actually
 * ends on screen.
 */
export function wrapRichTextRuns(
  runs: readonly XlsxCellTextRun[],
  measure: RichTextMeasure,
  maxWidth: number,
  maxLines: number,
  ellipsis = "\u2026",
): RichTextSegment[][] {
  const hardLines = splitRunsOnHardBreaks(runs);
  if (maxLines <= 1) {
    return [truncateRichTextRuns(hardLines[0] ?? [], measure, maxWidth, ellipsis)];
  }

  const lines: RichTextSegment[][] = [];
  for (const hardLine of hardLines) {
    if (lines.length >= maxLines) break;
    let line: RichTextSegment[] = [];
    let lineWidth = 0;
    let stopped = false;
    for (const run of hardLine) {
      for (const character of Array.from(run.text)) {
        const characterWidth = measure(character, run.font);
        const segment = { text: character, ...(run.font ? { font: run.font } : {}) };
        if (line.length > 0 && lineWidth + characterWidth > maxWidth) {
          lines.push(line);
          line = [segment];
          lineWidth = characterWidth;
          if (lines.length >= maxLines - 1) {
            stopped = true;
            break;
          }
        } else {
          appendSegment(line, segment);
          lineWidth += characterWidth;
        }
      }
      if (stopped) break;
    }
    if (lines.length >= maxLines) break;
    if (line.length > 0 || hardLines.length > 1) lines.push(line);
  }

  return lines;
}

/**
 * Keeps the runs that fit `maxWidth` and appends an ellipsis to the last one.
 *
 * The whole run list is returned unchanged when it already fits, so callers can
 * draw it without a second measurement pass.
 */
export function truncateRichTextRuns(
  runs: readonly RichTextSegment[],
  measure: RichTextMeasure,
  maxWidth: number,
  ellipsis = "\u2026",
): RichTextSegment[] {
  if (maxWidth <= 0) {
    return [];
  }

  if (measureRichTextSegments(runs, measure) <= maxWidth) {
    return runs.map(cloneSegment);
  }

  const output: RichTextSegment[] = [];
  const ellipsisWidth = measure(ellipsis, undefined);
  let width = 0;
  for (const run of runs) {
    let kept = "";
    for (const character of Array.from(run.text)) {
      const characterWidth = measure(character, run.font);
      if (width + characterWidth + ellipsisWidth > maxWidth) {
        if (kept) output.push({ text: kept, ...(run.font ? { font: run.font } : {}) });
        if (output.length === 0) {
          // Nothing fits: still show the ellipsis so the cell is not silent.
          return [{ text: ellipsis }];
        }
        const last = output[output.length - 1]!;
        last.text += ellipsis;
        return output;
      }
      kept += character;
      width += characterWidth;
    }
    if (kept) output.push({ text: kept, ...(run.font ? { font: run.font } : {}) });
  }

  return output;
}

/**
 * Text width of a laid-out line, measured run by run.
 *
 * This is the width a rich cell actually needs: the sum of the run widths in
 * their own fonts, which is what wrapping, truncation and any caller that fits
 * a column to its content must use. The concatenated string in the cell font
 * measures differently as soon as one run is bold, larger or a superscript.
 */
export function measureRichTextSegments(
  segments: readonly RichTextSegment[],
  measure: RichTextMeasure,
): number {
  let width = 0;
  for (const segment of segments) {
    width += measure(segment.text, segment.font);
  }
  return width;
}

/** Concatenated text of a laid-out line. */
export function segmentsToText(segments: readonly RichTextSegment[]): string {
  let text = "";
  for (const segment of segments) {
    text += segment.text;
  }
  return text;
}

function splitRunsOnHardBreaks(runs: readonly XlsxCellTextRun[]): RichTextSegment[][] {
  const lines: RichTextSegment[][] = [[]];
  for (const run of runs) {
    const parts = run.text.replace(/\r\n?/g, "\n").split("\n");
    parts.forEach((part, index) => {
      if (index > 0) lines.push([]);
      if (part) {
        lines[lines.length - 1]!.push({ text: part, ...(run.font ? { font: run.font } : {}) });
      }
    });
  }
  return lines;
}

function appendSegment(line: RichTextSegment[], segment: RichTextSegment): RichTextSegment[] {
  const last = line[line.length - 1];
  if (last && last.font === undefined && segment.font === undefined) {
    last.text += segment.text;
    return line;
  }
  line.push(segment);
  return line;
}

function cloneSegment(segment: RichTextSegment): RichTextSegment {
  return { text: segment.text, ...(segment.font ? { font: segment.font } : {}) };
}
