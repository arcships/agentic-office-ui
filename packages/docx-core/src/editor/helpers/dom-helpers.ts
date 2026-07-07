// DOM manipulation helpers: scheduled writes and selection placement.
// Upstream editor.tsx: lines 981-993 (scheduleDomWrite), 1155-1282
// (placeCaretInsideElementDom / selectionOffsetsWithinElementDom /
//  setSelectionWithinElementByTextOffsetsDom).

/** Schedule a DOM-mutating callback on the next animation frame, falling back
 *  to setTimeout(0) when requestAnimationFrame is unavailable. */
export function scheduleDomWrite(callback: () => void): void {
  if (
    typeof window !== "undefined" &&
    typeof window.requestAnimationFrame === "function"
  ) {
    window.requestAnimationFrame(() => {
      callback();
    });
    return;
  }

  setTimeout(callback, 0);
}

/** Collapse the caret to the end of `element`'s contents. */
export function placeCaretInsideElementDom(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  range.selectNodeContents(element);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

/** Resolve the current selection as text offsets within `element`, ignoring
 *  numbering-label spans. Returns undefined when the selection is outside the
 *  element. */
export function selectionOffsetsWithinElementDom(
  element: HTMLElement
): { start: number; end: number } | undefined {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return undefined;
  }

  const selectedRange = selection.getRangeAt(0);
  if (
    !element.contains(selectedRange.startContainer) ||
    !element.contains(selectedRange.endContainer)
  ) {
    return undefined;
  }

  try {
    const textLengthWithoutNumberingLabels = (range: Range): number => {
      const fragment = range.cloneContents();
      fragment
        .querySelectorAll("[data-docx-numbering-label='true']")
        .forEach((label) => {
          label.remove();
        });
      return fragment.textContent?.length ?? 0;
    };

    const startRange = document.createRange();
    startRange.setStart(element, 0);
    startRange.setEnd(selectedRange.startContainer, selectedRange.startOffset);

    const endRange = document.createRange();
    endRange.setStart(element, 0);
    endRange.setEnd(selectedRange.endContainer, selectedRange.endOffset);

    const startOffset = textLengthWithoutNumberingLabels(startRange);
    const endOffset = textLengthWithoutNumberingLabels(endRange);

    return {
      start: Math.min(startOffset, endOffset),
      end: Math.max(startOffset, endOffset),
    };
  } catch {
    return undefined;
  }
}

/** Set the DOM selection inside `element` to the given text offsets, skipping
 *  numbering-label spans. Falls back to placing the caret at the end. */
export function setSelectionWithinElementByTextOffsetsDom(
  element: HTMLElement,
  startOffset: number,
  endOffset: number
): void {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const resolveDomPosition = (
    container: HTMLElement,
    targetOffset: number
  ): {
    node: Node;
    offset: number;
  } => {
    const safeOffset = Math.max(0, Math.round(targetOffset));
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    let traversed = 0;
    let currentTextNode: Node | null = walker.nextNode();

    while (currentTextNode) {
      if (
        currentTextNode instanceof Text &&
        currentTextNode.parentElement?.closest(
          "[data-docx-numbering-label='true']"
        )
      ) {
        currentTextNode = walker.nextNode();
        continue;
      }

      const textLength = currentTextNode.textContent?.length ?? 0;
      if (traversed + textLength >= safeOffset) {
        return {
          node: currentTextNode,
          offset: Math.max(0, safeOffset - traversed),
        };
      }

      traversed += textLength;
      currentTextNode = walker.nextNode();
    }

    return {
      node: container,
      offset: container.childNodes.length,
    };
  };

  try {
    const range = document.createRange();
    const normalizedStart = Math.max(0, Math.min(startOffset, endOffset));
    const normalizedEnd = Math.max(
      normalizedStart,
      Math.max(startOffset, endOffset)
    );
    const start = resolveDomPosition(element, normalizedStart);
    const end = resolveDomPosition(element, normalizedEnd);
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    selection.removeAllRanges();
    selection.addRange(range);
  } catch {
    placeCaretInsideElementDom(element);
  }
}
