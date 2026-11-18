import { clampSurfaceZoom } from "./gesture-zoom";

/**
 * Touch interaction defaults shared by every Office surface.
 *
 * The values are deliberately aligned with platform conventions: 500 ms for a
 * long press, 320 ms for the second tap of a double tap, and 44 px as the
 * minimum hit target recommended by both Material and Apple HIG guidance.
 */
export const TOUCH_LONG_PRESS_MS = 500;
export const TOUCH_LONG_PRESS_MOVE_TOLERANCE_PX = 10;
export const TOUCH_DOUBLE_TAP_MS = 320;
export const TOUCH_DOUBLE_TAP_DISTANCE_PX = 30;
export const TOUCH_CONTEXT_MENU_SUPPRESSION_MS = 1_000;
export const TOUCH_CONTEXT_MENU_SUPPRESSION_PX = 24;
export const TOUCH_MIN_TARGET_PX = 44;
export const SURFACE_DOUBLE_TAP_ZOOM = 2;

export interface TouchGesturePoint {
  clientX: number;
  clientY: number;
}

export interface TouchGesturePointer extends TouchGesturePoint {
  pointerId: number;
  pointerType: string;
  target?: unknown;
}

export interface TouchGesturePointerEvent extends TouchGesturePointer {
  preventDefault?: () => void;
}

export interface TouchGestureTrackerOptions {
  /** Called when a finger stays still for `longPressMs`. */
  onLongPress?: (pointer: TouchGesturePointer) => void;
  /**
   * Called when two fingers pinch. Only fires when `getZoom` returns a
   * controlled zoom factor; surfaces without explicit zoom keep native pinch.
   */
  onZoom?: (zoom: number, center: TouchGesturePoint) => void;
  /** Called for a double tap; surfaces usually toggle zoom here. */
  onDoubleTap?: (pointer: TouchGesturePointer) => void;
  /** Current controlled zoom factor, or `undefined` when zoom is unavailable. */
  getZoom?: () => number | undefined;
  /** Gesture handling is skipped while this returns false (e.g. region mode). */
  isEnabled?: () => boolean;
  /** Double tap is skipped while this returns false (e.g. editable grids). */
  isDoubleTapEnabled?: () => boolean;
  longPressMs?: number;
  longPressMoveTolerancePx?: number;
  doubleTapMs?: number;
  doubleTapDistancePx?: number;
  /** Injectable clock, used by tests to make tap timing deterministic. */
  now?: () => number;
}

export interface TouchGestureTracker {
  pointerDown(event: TouchGesturePointerEvent): void;
  pointerMove(event: TouchGesturePointerEvent): void;
  pointerUp(event: TouchGesturePointerEvent): void;
  pointerCancel(event: TouchGesturePointerEvent): void;
  /** Native `contextmenu` must be ignored right after a long press. */
  shouldSuppressContextMenu(clientX: number, clientY: number): boolean;
  reset(): void;
  readonly isPinching: boolean;
}

interface ActivePointer extends TouchGesturePointer {
  startX: number;
  startY: number;
  moved: boolean;
}

interface PinchState {
  idA: number;
  idB: number;
  startDistance: number;
  startZoom: number;
  lastZoom: number;
}

interface TapState extends TouchGesturePoint {
  time: number;
}

export function pointerDistance(a: TouchGesturePoint, b: TouchGesturePoint): number {
  return Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
}

export function pinchSurfaceZoom(
  startZoom: number,
  startDistance: number,
  currentDistance: number,
): number {
  if (
    !Number.isFinite(startZoom) ||
    !Number.isFinite(startDistance) ||
    !Number.isFinite(currentDistance) ||
    startDistance <= 0 ||
    currentDistance <= 0
  ) {
    return clampSurfaceZoom(startZoom);
  }
  return clampSurfaceZoom(startZoom * (currentDistance / startDistance));
}

export function nextDoubleTapZoom(currentZoom: number): number {
  const current = clampSurfaceZoom(currentZoom);
  return current > 1 ? 1 : clampSurfaceZoom(SURFACE_DOUBLE_TAP_ZOOM);
}

function isTouchPointer(pointerType: string): boolean {
  return pointerType === "touch" || pointerType === "pen";
}

/**
 * Framework-free touch gesture state machine.
 *
 * The owning component forwards pointer events from its scroll container. The
 * tracker recognises long press, double tap and two-finger pinch, and keeps
 * those gestures out of text/cell selection by only reacting to touch/pen
 * pointers. It never captures the pointer, so native panning and selection
 * keep working.
 */
export function createTouchGestureTracker(
  options: TouchGestureTrackerOptions = {},
): TouchGestureTracker {
  const longPressMs = options.longPressMs ?? TOUCH_LONG_PRESS_MS;
  const longPressMoveTolerancePx =
    options.longPressMoveTolerancePx ?? TOUCH_LONG_PRESS_MOVE_TOLERANCE_PX;
  const doubleTapMs = options.doubleTapMs ?? TOUCH_DOUBLE_TAP_MS;
  const doubleTapDistancePx =
    options.doubleTapDistancePx ?? TOUCH_DOUBLE_TAP_DISTANCE_PX;
  const now = options.now ?? Date.now;

  const active = new Map<number, ActivePointer>();
  let longPressTimer: ReturnType<typeof setTimeout> | undefined;
  let longPressPointer: ActivePointer | undefined;
  let longPressFired = false;
  let pinch: PinchState | undefined;
  let lastTap: TapState | undefined;
  let lastLongPress: TapState | undefined;
  let suppressTap = false;

  function enabled(): boolean {
    return options.isEnabled ? options.isEnabled() : true;
  }

  function cancelLongPress(): void {
    if (longPressTimer !== undefined) {
      clearTimeout(longPressTimer);
      longPressTimer = undefined;
    }
    longPressPointer = undefined;
    longPressFired = false;
  }

  function beginLongPress(pointer: ActivePointer): void {
    if (!options.onLongPress || longPressMs < 0) return;
    longPressPointer = pointer;
    longPressFired = false;
    if (longPressMs === 0) {
      fireLongPress(pointer);
      return;
    }
    longPressTimer = setTimeout(() => {
      longPressTimer = undefined;
      fireLongPress(pointer);
    }, longPressMs);
  }

  function fireLongPress(pointer: ActivePointer): void {
    if (longPressPointer !== pointer || pointer.moved || active.size !== 1) return;
    longPressFired = true;
    lastLongPress = { clientX: pointer.clientX, clientY: pointer.clientY, time: now() };
    options.onLongPress?.({
      pointerId: pointer.pointerId,
      pointerType: pointer.pointerType,
      clientX: pointer.clientX,
      clientY: pointer.clientY,
      target: pointer.target,
    });
  }

  function beginPinch(): void {
    const [a, b] = [...active.values()];
    if (!a || !b || !options.onZoom) return;
    const startZoom = options.getZoom?.();
    if (startZoom === undefined || !Number.isFinite(startZoom)) return;
    const startDistance = pointerDistance(a, b);
    if (startDistance <= 0) return;
    pinch = {
      idA: a.pointerId,
      idB: b.pointerId,
      startDistance,
      startZoom: clampSurfaceZoom(startZoom),
      lastZoom: clampSurfaceZoom(startZoom),
    };
  }

  function endPinchIfReleased(pointerId?: number): void {
    if (!pinch) return;
    if (
      pointerId === undefined ||
      pointerId === pinch.idA ||
      pointerId === pinch.idB ||
      !active.has(pinch.idA) ||
      !active.has(pinch.idB)
    ) {
      pinch = undefined;
      suppressTap = true;
    }
  }

  return {
    get isPinching() {
      return pinch !== undefined;
    },

    pointerDown(event) {
      if (!isTouchPointer(event.pointerType) || !enabled()) return;
      const pointer: ActivePointer = {
        pointerId: event.pointerId,
        pointerType: event.pointerType,
        clientX: event.clientX,
        clientY: event.clientY,
        target: event.target,
        startX: event.clientX,
        startY: event.clientY,
        moved: false,
      };
      active.set(pointer.pointerId, pointer);

      if (active.size === 1) {
        suppressTap = false;
        beginLongPress(pointer);
        return;
      }

      // A second finger turns the gesture into a pinch: cancel both the
      // pending long press and any tap that could otherwise fire on release.
      cancelLongPress();
      lastTap = undefined;
      suppressTap = true;
      if (!pinch) beginPinch();
    },

    pointerMove(event) {
      const pointer = active.get(event.pointerId);
      if (!pointer) return;
      pointer.clientX = event.clientX;
      pointer.clientY = event.clientY;
      if (
        !pointer.moved &&
        pointerDistance(pointer, { clientX: pointer.startX, clientY: pointer.startY }) >
          longPressMoveTolerancePx
      ) {
        pointer.moved = true;
      }
      if (pointer === longPressPointer && pointer.moved) cancelLongPress();
      if (!pinch || !options.onZoom) return;

      const a = active.get(pinch.idA);
      const b = active.get(pinch.idB);
      if (!a || !b) return;
      const distance = pointerDistance(a, b);
      if (distance <= 0) return;
      const nextZoom = pinchSurfaceZoom(pinch.startZoom, pinch.startDistance, distance);
      if (nextZoom === pinch.lastZoom) return;
      pinch.lastZoom = nextZoom;
      options.onZoom(nextZoom, {
        clientX: (a.clientX + b.clientX) / 2,
        clientY: (a.clientY + b.clientY) / 2,
      });
    },

    pointerUp(event) {
      const pointer = active.get(event.pointerId);
      if (!pointer) return;
      const wasLongPress =
        longPressFired && longPressPointer?.pointerId === event.pointerId;
      active.delete(event.pointerId);
      cancelLongPress();

      if (pinch) endPinchIfReleased(event.pointerId);
      if (active.size > 0) return;

      const wasSuppressed = suppressTap;
      suppressTap = false;
      if (wasLongPress || wasSuppressed || pointer.moved) {
        // A long press must not turn into a click that changes the selection
        // (or lands on the freshly opened menu) when the finger lifts.
        if (wasLongPress) event.preventDefault?.();
        lastTap = undefined;
        return;
      }

      if (options.isDoubleTapEnabled && !options.isDoubleTapEnabled()) {
        lastTap = undefined;
        return;
      }

      const time = now();
      if (
        lastTap &&
        time - lastTap.time <= doubleTapMs &&
        pointerDistance(lastTap, pointer) <= doubleTapDistancePx
      ) {
        lastTap = undefined;
        event.preventDefault?.();
        options.onDoubleTap?.({
          pointerId: pointer.pointerId,
          pointerType: pointer.pointerType,
          clientX: pointer.clientX,
          clientY: pointer.clientY,
          target: pointer.target,
        });
        return;
      }

      lastTap = { clientX: pointer.clientX, clientY: pointer.clientY, time };
    },

    pointerCancel(event) {
      if (!active.has(event.pointerId)) return;
      active.delete(event.pointerId);
      cancelLongPress();
      if (pinch) endPinchIfReleased(event.pointerId);
      if (active.size > 0) return;
      pinch = undefined;
      lastTap = undefined;
      suppressTap = false;
    },

    shouldSuppressContextMenu(clientX, clientY) {
      if (!lastLongPress) return false;
      const elapsed = now() - lastLongPress.time;
      return (
        elapsed >= 0 &&
        elapsed <= TOUCH_CONTEXT_MENU_SUPPRESSION_MS &&
        pointerDistance(lastLongPress, { clientX, clientY }) <=
          TOUCH_CONTEXT_MENU_SUPPRESSION_PX
      );
    },

    reset() {
      cancelLongPress();
      active.clear();
      pinch = undefined;
      lastTap = undefined;
      lastLongPress = undefined;
      suppressTap = false;
    },
  };
}
