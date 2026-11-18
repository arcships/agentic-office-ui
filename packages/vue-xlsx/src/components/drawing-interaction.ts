import { onUnmounted, ref, type CSSProperties } from "vue";
import {
  resizeImageRect,
  type XlsxImageRect,
  type XlsxImageResizeHandlePosition,
} from "@arcships/xlsx-core";

/** The eight Excel resize handles, clockwise from the top-left corner. */
export const DRAWING_RESIZE_HANDLES: XlsxImageResizeHandlePosition[] = [
  "nw",
  "n",
  "ne",
  "e",
  "se",
  "s",
  "sw",
  "w",
];

/** Position and cursor for a resize handle inside an object frame. */
export function drawingHandleStyle(handle: XlsxImageResizeHandlePosition): CSSProperties {
  const horizontal = handle.includes("w") ? "left" : handle.includes("e") ? "right" : "center";
  const vertical = handle.includes("n") ? "top" : handle.includes("s") ? "bottom" : "center";
  const cursor = ({
    n: "ns-resize",
    s: "ns-resize",
    e: "ew-resize",
    w: "ew-resize",
    ne: "nesw-resize",
    sw: "nesw-resize",
    nw: "nwse-resize",
    se: "nwse-resize",
  } as Record<XlsxImageResizeHandlePosition, string>)[handle];
  const style: CSSProperties = { cursor };
  if (horizontal === "left") style.left = "-4px";
  else if (horizontal === "right") style.right = "-4px";
  else style.left = "calc(50% - 4px)";
  if (vertical === "top") style.top = "-4px";
  else if (vertical === "bottom") style.bottom = "-4px";
  else style.top = "calc(50% - 4px)";
  return style;
}

type GestureMode = "move" | "resize";

interface DrawingGesture {
  handle: XlsxImageResizeHandlePosition | null;
  id: string;
  mode: GestureMode;
  moved: boolean;
  pointerId: number;
  startRect: XlsxImageRect;
  startX: number;
  startY: number;
  x: number;
  y: number;
}

export interface DrawingGestureOptions {
  onMove: (id: string, deltaX: number, deltaY: number) => void;
  onResize: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
  onSelect: (id: string) => void;
}

/**
 * Local drag / resize preview for anchored drawing objects.
 *
 * The workbook is mutated once, on pointer-up, so a gesture produces exactly
 * one undo entry instead of one per pointermove frame. The preview rect is in
 * grid-body content pixels (zoom applied), matching `resolveAnchorRect`.
 */
export function useDrawingGestures(options: DrawingGestureOptions) {
  const gesture = ref<DrawingGesture | null>(null);
  const preview = ref<(XlsxImageRect & { id: string }) | null>(null);

  function stopListening() {
    if (typeof window === "undefined") {
      return;
    }
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
    window.removeEventListener("pointercancel", onPointerCancel);
    window.removeEventListener("keydown", onKeydown);
  }

  function reset() {
    stopListening();
    gesture.value = null;
    preview.value = null;
  }

  function updatePreview(deltaX: number, deltaY: number) {
    const current = gesture.value;
    if (!current) {
      return;
    }
    const rect = current.mode === "move"
      ? {
          ...current.startRect,
          left: current.startRect.left + deltaX,
          top: current.startRect.top + deltaY,
        }
      : resizeImageRect(current.startRect, current.handle!, deltaX, deltaY);
    preview.value = { ...rect, id: current.id };
  }

  function onPointerMove(event: PointerEvent) {
    const current = gesture.value;
    if (!current || event.pointerId !== current.pointerId) {
      return;
    }
    const deltaX = event.clientX - current.startX;
    const deltaY = event.clientY - current.startY;
    if (deltaX !== 0 || deltaY !== 0) {
      current.moved = true;
    }
    current.x = event.clientX;
    current.y = event.clientY;
    updatePreview(deltaX, deltaY);
  }

  function onPointerUp(event: PointerEvent) {
    const current = gesture.value;
    if (!current || event.pointerId !== current.pointerId) {
      return;
    }
    const deltaX = event.clientX - current.startX;
    const deltaY = event.clientY - current.startY;
    const shouldCommit = current.moved || deltaX !== 0 || deltaY !== 0;
    reset();
    if (!shouldCommit) {
      return;
    }
    if (current.mode === "move") {
      options.onMove(current.id, deltaX, deltaY);
    } else if (current.handle) {
      options.onResize(current.id, current.handle, deltaX, deltaY);
    }
  }

  function onPointerCancel(event: PointerEvent) {
    if (gesture.value?.pointerId === event.pointerId) {
      reset();
    }
  }

  function onKeydown(event: KeyboardEvent) {
    if (event.key === "Escape") {
      reset();
    }
  }

  function begin(
    id: string,
    rect: XlsxImageRect,
    mode: GestureMode,
    handle: XlsxImageResizeHandlePosition | null,
    event: PointerEvent,
  ) {
    if (event.button !== 0) {
      return;
    }
    event.preventDefault?.();
    options.onSelect(id);
    const target = event.currentTarget as HTMLElement | null;
    target?.focus?.();
    target?.setPointerCapture?.(event.pointerId);
    gesture.value = {
      handle,
      id,
      mode,
      moved: false,
      pointerId: event.pointerId,
      startRect: rect,
      startX: event.clientX,
      startY: event.clientY,
      x: event.clientX,
      y: event.clientY,
    };
    preview.value = { ...rect, id };
    if (typeof window !== "undefined") {
      window.addEventListener("pointermove", onPointerMove);
      window.addEventListener("pointerup", onPointerUp);
      window.addEventListener("pointercancel", onPointerCancel);
      window.addEventListener("keydown", onKeydown);
    }
  }

  function beginMove(id: string, rect: XlsxImageRect, event: PointerEvent) {
    begin(id, rect, "move", null, event);
  }

  function beginResize(id: string, rect: XlsxImageRect, handle: XlsxImageResizeHandlePosition, event: PointerEvent) {
    begin(id, rect, "resize", handle, event);
  }

  function previewFor(id: string): XlsxImageRect | null {
    return preview.value?.id === id ? preview.value : null;
  }

  function isActive(id: string): boolean {
    return gesture.value?.id === id;
  }

  onUnmounted(reset);

  return {
    beginMove,
    beginResize,
    isActive,
    previewFor,
  };
}
