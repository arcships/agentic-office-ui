import type {
  ManualRegionLocator,
  NormalizedPoint,
  NormalizedRect,
  SheetRegionPoint,
} from "./types"

function clampUnit(value: number): number {
  if (!Number.isFinite(value)) throw new TypeError("normalized coordinates must be finite")
  return Math.min(1, Math.max(0, value))
}

export function normalizedRectFromPoints(a: NormalizedPoint, b: NormalizedPoint): NormalizedRect {
  const ax = clampUnit(a.x)
  const ay = clampUnit(a.y)
  const bx = clampUnit(b.x)
  const by = clampUnit(b.y)
  const x = Math.min(ax, bx)
  const y = Math.min(ay, by)
  const width = Math.abs(ax - bx)
  const height = Math.abs(ay - by)
  if (width === 0 || height === 0) throw new RangeError("normalized rectangle must have positive width and height")
  return { x, y, width, height }
}

export function normalizedRectArea(rect: NormalizedRect): number {
  return rect.width * rect.height
}

export function normalizedRectContains(rect: NormalizedRect, point: NormalizedPoint): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.width &&
    point.y >= rect.y && point.y <= rect.y + rect.height
}

export function intersectNormalizedRects(a: NormalizedRect, b: NormalizedRect): NormalizedRect | null {
  const x = Math.max(a.x, b.x)
  const y = Math.max(a.y, b.y)
  const right = Math.min(a.x + a.width, b.x + b.width)
  const bottom = Math.min(a.y + a.height, b.y + b.height)
  if (right <= x || bottom <= y) return null
  return { x, y, width: right - x, height: bottom - y }
}

function axisStart(
  firstIndex: number,
  firstOffset: number,
  secondIndex: number,
  secondOffset: number,
): readonly [number, number] {
  if (firstIndex < secondIndex) return [firstIndex, firstOffset]
  if (firstIndex > secondIndex) return [secondIndex, secondOffset]
  return [firstIndex, Math.min(firstOffset, secondOffset)]
}

function axisEnd(
  firstIndex: number,
  firstOffset: number,
  secondIndex: number,
  secondOffset: number,
): readonly [number, number] {
  if (firstIndex > secondIndex) return [firstIndex, firstOffset]
  if (firstIndex < secondIndex) return [secondIndex, secondOffset]
  return [firstIndex, Math.max(firstOffset, secondOffset)]
}

export function normalizeSheetRegion(
  sheetId: string,
  first: SheetRegionPoint,
  second: SheetRegionPoint,
): Extract<ManualRegionLocator, { space: "sheet" }> {
  if (!sheetId) throw new TypeError("sheetId must not be empty")
  for (const [label, point] of [["first", first], ["second", second]] as const) {
    if (!Number.isSafeInteger(point.row) || point.row < 0 || !Number.isSafeInteger(point.col) || point.col < 0) {
      throw new TypeError(`${label} sheet coordinates must use non-negative safe integer rows and columns`)
    }
  }
  const [startRow, startYOffset] = axisStart(first.row, clampUnit(first.yOffset), second.row, clampUnit(second.yOffset))
  const [endRow, endYOffset] = axisEnd(first.row, clampUnit(first.yOffset), second.row, clampUnit(second.yOffset))
  const [startCol, startXOffset] = axisStart(first.col, clampUnit(first.xOffset), second.col, clampUnit(second.xOffset))
  const [endCol, endXOffset] = axisEnd(first.col, clampUnit(first.xOffset), second.col, clampUnit(second.xOffset))
  const rowSpan = endRow - startRow + endYOffset - startYOffset
  const colSpan = endCol - startCol + endXOffset - startXOffset
  if (rowSpan <= 0 || colSpan <= 0) throw new RangeError("sheet region must have a positive area")
  return {
    space: "sheet",
    sheetId,
    start: { row: startRow, col: startCol, xOffset: startXOffset, yOffset: startYOffset },
    end: { row: endRow, col: endCol, xOffset: endXOffset, yOffset: endYOffset },
  }
}
