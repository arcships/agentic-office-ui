# Touch Interaction Spec

Issue: [#106](https://github.com/arcships/agentic-office-ui/issues/106) — touch/mobile
interactions are partial and untested. This document is the behaviour contract the
surfaces implement.

## Reference device profile

| Property | Value |
| --- | --- |
| Viewport | `390 × 844` (portrait phone) |
| Input | coarse pointer, touch events, no hover |
| Browser zoom | must not trigger while a gesture belongs to a surface |

The same rules apply to tablets in both orientations and to touchscreen laptops
(`pointer: coarse` media query).

## Gestures

| Gesture | Behaviour |
| --- | --- |
| Tap | Unchanged. Selects text/cell/object exactly like a mouse click. |
| Long press (`500 ms`, `10 px` tolerance) | Opens the surface context menu at the finger position. The native `contextmenu` that follows is suppressed for `1 s` within `24 px` so a menu never opens twice. |
| Double tap (`320 ms`, `30 px`) | Toggles zoom: `100% → 200%`, any zoom above `100% → 100%`. Suppressed when zoom is not controlled, when the surface is editable, or on an editable spreadsheet (there double tap keeps entering cell edit mode). |
| Pinch | Two-finger distance scales zoom between `50%` and `200%` around the midpoint between the fingers. Applies to pointer events (Android/Chrome/Windows), in addition to the existing WebKit `gesture*` events and `Ctrl+wheel`. |
| Pan | Native scrolling. Surfaces set `touch-action: pan-x pan-y` so one-finger pan and the browser's momentum scrolling keep working; custom pinch takes over multi-finger input. |
| Pointer cancel | Any active long press, pinch or tap sequence is dropped without emitting selection or context-menu events. |

Long press and double tap only react to `pointerType === "touch"` or `"pen"`.
Mouse and trackpad behaviour is unchanged.

## Per-surface matrix

| Surface | Long press | Double tap zoom | Pinch zoom | Pan |
| --- | --- | --- | --- | --- |
| `PdfSurface` | page context menu | controlled `zoom` only | controlled `zoom` only | native |
| `PdfViewer` | — | always | always | native |
| `PptxStage` | slide/object context menu | controlled `zoom` only, disabled while presenting | controlled `zoom` only, disabled while presenting | native |
| `DocxDocumentSurface` / `DocxViewerRoot` | page context menu | read-only viewing mode | read-only viewing mode (editable documents keep native text selection) | native |
| `XlsxSheetSurface` | cell context menu | read-only spreadsheets | controlled `zoom` only | native |
| `XlsxViewer` | cell context menu | read-only spreadsheets | controller zoom scale | native |

## Hit targets

On coarse pointers every toolbar button, menu item, zoom selector and page input
has a minimum target of `44 × 44 px`. Desktop sizing is unchanged.

## Context menus

- Open on long press and on the native `contextmenu` event.
- Close on any outside `pointerdown`; not on the trailing `click` of the
  long-press gesture that opened them.
- Positioned at the touch coordinates and clamped by the host menu component.

## Tests

- `tests/unit/office-runtime.test.mjs` covers pinch math, double-tap toggling,
  long-press timing/tolerance, native-menu suppression and pointer cancellation.
- `tests/component/touch-interactions.test.mjs` verifies every surface wires the
  shared tracker, keeps `pointercancel` handling, sets `touch-action`, closes
  menus on outside pointer down and ships 44 px coarse-pointer targets.
- `tests/blackbox/touch_mobile.py` drives Chromium touch emulation on the
  `390 × 844` / DPR 3 mobile profile across DOCX, XLSX, PPTX and PDF: real CDP
  long press, double tap, pinch, `touchCancel` and 44 px toolbar targets. It is
  registered as `formal-touch-mobile` in the `blackbox` suite.
