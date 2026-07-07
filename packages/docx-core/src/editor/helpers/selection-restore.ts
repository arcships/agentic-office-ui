// DOM selection restore heuristics: deciding whether to reissue a DOM selection
// restore after model/selection changes, whether a key-up should sync the active
// range, and detecting collapsed selections that sit at the start of an element
// (used to suppress spurious caret placements).
// Upstream editor.tsx: lines 24148-24259.

import type {
  DocxSelectionSessionKind,
  DocxTextRange
} from "./editor-types";
import { ACTIVE_RANGE_SYNC_KEYS, MODIFIER_ONLY_KEYS } from "./constants";

// Minimal framework-agnostic view of a keyboard event, matching the subset of
// `React.KeyboardEvent<HTMLElement>` / native `KeyboardEvent` this helper reads
// (`nativeEvent.isComposing`, `key`, `ctrlKey`, `metaKey`). Keeps the helper
// free of React type imports.
export interface EditorKeyboardEventLike {
  nativeEvent: KeyboardEvent;
  key: string;
  ctrlKey: boolean;
  metaKey: boolean;
}

export function shouldReissueDomSelectionRestore(options: {
  modelChanged: boolean;
  selectionChanged: boolean;
  rangeChanged: boolean;
  activeTextRange?: DocxTextRange;
  suppressNext: boolean;
  selectionSessionKind: DocxSelectionSessionKind;
}): boolean {
  if (options.suppressNext) {
    return false;
  }

  if (
    options.selectionSessionKind === "pointer" ||
    options.selectionSessionKind === "keyboard" ||
    options.selectionSessionKind === "composition"
  ) {
    return false;
  }

  if (!options.activeTextRange) {
    return false;
  }

  if (options.selectionChanged || options.rangeChanged) {
    return false;
  }

  return options.modelChanged;
}

export function shouldSyncActiveRangeOnKeyUp(
  event: EditorKeyboardEventLike
): boolean {
  const nativeKeyboardEvent = event.nativeEvent as KeyboardEvent | undefined;
  if (nativeKeyboardEvent?.isComposing) {
    return false;
  }

  if (MODIFIER_ONLY_KEYS.has(event.key)) {
    return false;
  }

  if (ACTIVE_RANGE_SYNC_KEYS.has(event.key)) {
    return true;
  }

  if (event.ctrlKey || event.metaKey) {
    const normalizedKey = event.key.toLowerCase();
    return (
      normalizedKey === "a" ||
      normalizedKey === "x" ||
      normalizedKey === "z" ||
      normalizedKey === "y"
    );
  }

  return false;
}

export function isCollapsedSelectionAtElementStart(
  element: HTMLElement,
  range: Range,
  pointX?: number
): boolean {
  if (!range.collapsed) {
    return false;
  }

  if (range.startContainer === element && range.startOffset === 0) {
    return true;
  }

  if (!element.contains(range.startContainer)) {
    return false;
  }

  try {
    const offsetProbe = document.createRange();
    offsetProbe.selectNodeContents(element);
    offsetProbe.setEnd(range.startContainer, range.startOffset);
    const textOffset = offsetProbe.toString().length;
    if (textOffset !== 0) {
      return false;
    }

    if (Number.isFinite(pointX)) {
      const elementRect = element.getBoundingClientRect();
      if ((pointX as number) <= elementRect.left + 12) {
        return false;
      }
    }

    return true;
  } catch {
    return false;
  }
}

export function isSuspiciousCollapsedSelectionAtElementStart(
  element: HTMLElement,
  range: Range,
  pointX?: number
): boolean {
  if (!isCollapsedSelectionAtElementStart(element, range, pointX)) {
    return false;
  }

  return (
    range.startContainer === element || !(range.startContainer instanceof Text)
  );
}
