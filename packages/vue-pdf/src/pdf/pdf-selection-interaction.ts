import type { PdfPageGeometry, Position } from "@embedpdf/models"

export interface PdfCharacterRange {
  from: number
  to: number
}

interface WordSegment {
  index: number
  segment: string
  isWordLike?: boolean
}

interface WordSegmenter {
  segment(input: string): Iterable<WordSegment>
}

function containsPoint(
  rect: { x: number; y: number; width: number; height: number },
  point: Position,
): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.width
    && point.y >= rect.y && point.y <= rect.y + rect.height
}

function distanceSquaredToRect(
  rect: { x: number; y: number; width: number; height: number },
  point: Position,
): number {
  const dx = point.x < rect.x
    ? rect.x - point.x
    : point.x > rect.x + rect.width
      ? point.x - (rect.x + rect.width)
      : 0
  const dy = point.y < rect.y
    ? rect.y - point.y
    : point.y > rect.y + rect.height
      ? point.y - (rect.y + rect.height)
      : 0
  return dx * dx + dy * dy
}

function runContainsCharacter(
  run: PdfPageGeometry["runs"][number],
  charIndex: number,
): boolean {
  return charIndex >= run.charStart && charIndex < run.charStart + run.glyphs.length
}

/**
 * Resolve a point that lands in the visual gap between glyphs, while keeping
 * the correction scoped to the text run that actually contains the point.
 */
export function nearestGlyphWithinContainingRun(
  geometry: PdfPageGeometry,
  point: Position,
  previousFocus = -1,
): number {
  const containingRuns = geometry.runs.filter((run) => containsPoint(run.rect, point))
  if (containingRuns.length === 0) return -1

  const previousRun = containingRuns.find((run) => runContainsCharacter(run, previousFocus))
  const runs = previousRun ? [previousRun] : containingRuns
  let nearest = -1
  let nearestDistance = Number.POSITIVE_INFINITY

  for (const run of runs) {
    for (let index = 0; index < run.glyphs.length; index++) {
      const glyph = run.glyphs[index]
      if (glyph.flags === 2) continue
      const distance = distanceSquaredToRect({
        x: glyph.tightX ?? glyph.x,
        y: glyph.tightY ?? glyph.y,
        width: glyph.tightWidth ?? glyph.width,
        height: glyph.tightHeight ?? glyph.height,
      }, point)
      if (distance < nearestDistance) {
        nearestDistance = distance
        nearest = run.charStart + index
      }
    }
  }

  return nearest
}

/**
 * Refine a coarse PDF character range with a Unicode word segmenter. PDF
 * character indices are mapped outward so a multi-code-unit character is
 * never split by a UTF-16 segment boundary.
 */
export function wordRangeFromCharacterTexts(
  candidateFrom: number,
  clickedCharIndex: number,
  characterTexts: readonly string[],
  segmenter: WordSegmenter,
): PdfCharacterRange | null {
  const clickedOffset = clickedCharIndex - candidateFrom
  if (clickedOffset < 0 || clickedOffset >= characterTexts.length) return null

  const utf16Offsets = [0]
  for (const text of characterTexts) {
    utf16Offsets.push(utf16Offsets[utf16Offsets.length - 1] + text.length)
  }
  const joined = characterTexts.join("")
  const clickedText = characterTexts[clickedOffset]
  if (!joined || !clickedText) return null

  const clickedUtf16Offset = utf16Offsets[clickedOffset]
  const segment = Array.from(segmenter.segment(joined)).find((entry) =>
    entry.index <= clickedUtf16Offset
      && clickedUtf16Offset < entry.index + entry.segment.length
      && entry.isWordLike !== false,
  )
  if (!segment) return null

  const segmentEnd = segment.index + segment.segment.length
  let first = -1
  let last = -1
  for (let index = 0; index < characterTexts.length; index++) {
    const charStart = utf16Offsets[index]
    const charEnd = utf16Offsets[index + 1]
    if (charEnd <= segment.index || charStart >= segmentEnd) continue
    if (first === -1) first = index
    last = index
  }
  if (first === -1 || last === -1) return null

  return {
    from: candidateFrom + first,
    to: candidateFrom + last,
  }
}
