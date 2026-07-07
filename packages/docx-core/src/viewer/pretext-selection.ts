import type {
  PretextLineFragment,
  PretextSelectionRect,
  PretextVariableWidthLayout,
} from "./pretext-layout";
import { cachedFragmentOffsetAdvances } from "./pretext-layout";

function nearestLineIndexForY(
  layout: PretextVariableWidthLayout,
  y: number
): number {
  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  if (layout.lines.length === 0) {
    return 0;
  }

  let nearestIndex = 0;
  let nearestDistance = Number.POSITIVE_INFINITY;
  layout.lines.forEach((line, index) => {
    const centerY = line.y + lineHeightPx / 2;
    const distance = Math.abs(y - centerY);
    if (distance < nearestDistance) {
      nearestDistance = distance;
      nearestIndex = index;
    }
  });

  return nearestIndex;
}

function fragmentOffsetAtX(
  font: string,
  fragment: PretextLineFragment,
  xWithinFragment: number
): number {
  if (xWithinFragment <= 0) {
    return fragment.startOffset;
  }

  if (xWithinFragment >= fragment.width) {
    return fragment.endOffset;
  }

  let bestOffset = fragment.startOffset;
  let bestDistance = Number.POSITIVE_INFINITY;
  const advances = cachedFragmentOffsetAdvances(font, fragment);
  for (let localOffset = 0; localOffset < advances.length; localOffset += 1) {
    const advancePx = advances[localOffset] ?? 0;
    const distance = Math.abs(xWithinFragment - advancePx);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestOffset = fragment.startOffset + localOffset;
    }
  }

  return bestOffset;
}

export function resolveOffsetAtPoint(
  layout: PretextVariableWidthLayout,
  x: number,
  y: number
): number {
  const textLength = layout.text?.length ?? 0;
  if (layout.lines.length === 0) {
    return 0;
  }

  const lineIndex = nearestLineIndexForY(layout, y);
  const line = layout.lines[lineIndex];
  if (!line || line.fragments.length === 0) {
    return Math.max(0, Math.min(textLength, 0));
  }

  const firstFragment = line.fragments[0]!;
  const lastFragment = line.fragments[line.fragments.length - 1]!;
  if (x <= firstFragment.x) {
    return firstFragment.startOffset;
  }

  if (x >= lastFragment.x + lastFragment.width) {
    return lastFragment.endOffset;
  }

  for (
    let fragmentIndex = 0;
    fragmentIndex < line.fragments.length;
    fragmentIndex += 1
  ) {
    const fragment = line.fragments[fragmentIndex]!;
    const fragmentLeft = fragment.x;
    const fragmentRight = fragment.x + fragment.width;
    if (x >= fragmentLeft && x <= fragmentRight) {
      return fragmentOffsetAtX(layout.font ?? "", fragment, x - fragmentLeft);
    }

    const nextFragment = line.fragments[fragmentIndex + 1];
    if (nextFragment && x > fragmentRight && x < nextFragment.x) {
      const gapMidpoint = fragmentRight + (nextFragment.x - fragmentRight) / 2;
      return x < gapMidpoint ? fragment.endOffset : nextFragment.startOffset;
    }
  }

  return Math.max(0, Math.min(textLength, lastFragment.endOffset));
}

export function resolveCaretRectAtOffset(
  layout: PretextVariableWidthLayout,
  offset: number
): PretextSelectionRect | undefined {
  if (layout.lines.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(
    0,
    Math.min(Math.round(offset), layout.text?.length ?? 0)
  );
  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));

  for (const line of layout.lines) {
    for (const fragment of line.fragments) {
      if (
        safeOffset < fragment.startOffset ||
        safeOffset > fragment.endOffset
      ) {
        continue;
      }

      const localOffset = safeOffset - fragment.startOffset;
      const advances = cachedFragmentOffsetAdvances(
        layout.font ?? "",
        fragment
      );
      const left = fragment.x + (advances[localOffset] ?? 0);
      return {
        left,
        top: line.y,
        width: 1,
        height: lineHeightPx,
      };
    }
  }

  const lastLine = layout.lines[layout.lines.length - 1];
  const lastFragment = lastLine?.fragments[lastLine.fragments.length - 1];
  if (!lastLine || !lastFragment) {
    return undefined;
  }

  return {
    left: lastFragment.x + lastFragment.width,
    top: lastLine.y,
    width: 1,
    height: lineHeightPx,
  };
}

export function resolveSelectionRects(
  layout: PretextVariableWidthLayout,
  startOffset: number,
  endOffset: number
): PretextSelectionRect[] {
  const safeStart = Math.max(
    0,
    Math.min(Math.round(startOffset), layout.text?.length ?? 0)
  );
  const safeEnd = Math.max(
    safeStart,
    Math.min(Math.round(endOffset), layout.text?.length ?? 0)
  );
  if (safeStart === safeEnd) {
    return [];
  }

  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  const rects: PretextSelectionRect[] = [];

  layout.lines.forEach((line) => {
    line.fragments.forEach((fragment) => {
      const overlapStart = Math.max(safeStart, fragment.startOffset);
      const overlapEnd = Math.min(safeEnd, fragment.endOffset);
      if (overlapStart >= overlapEnd) {
        return;
      }

      const advances = cachedFragmentOffsetAdvances(
        layout.font ?? "",
        fragment
      );
      const leadingWidthPx = advances[overlapStart - fragment.startOffset] ?? 0;
      const selectedWidthPx =
        (advances[overlapEnd - fragment.startOffset] ?? 0) - leadingWidthPx;
      rects.push({
        left: fragment.x + leadingWidthPx,
        top: line.y,
        width: Math.max(1, selectedWidthPx),
        height: lineHeightPx,
      });
    });
  });

  return rects;
}

export function sliceLayoutToLineRange(
  layout: PretextVariableWidthLayout,
  startLineIndex: number,
  endLineIndex: number
): PretextVariableWidthLayout {
  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), layout.lines.length)
  );
  const safeEnd = Math.max(
    safeStart,
    Math.min(Math.round(endLineIndex), layout.lines.length)
  );
  const slicedLines = layout.lines.slice(safeStart, safeEnd);
  const yOffset = slicedLines[0]?.y ?? 0;
  const normalizedLines = slicedLines.map((line) => ({
    ...line,
    y: line.y - yOffset,
    fragments: line.fragments.map((fragment) => ({ ...fragment })),
  }));
  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  const height =
    normalizedLines.length > 0
      ? (normalizedLines[normalizedLines.length - 1]?.y ?? 0) + lineHeightPx
      : 0;

  return {
    ...layout,
    lineCount: normalizedLines.length,
    height,
    lines: normalizedLines,
  };
}