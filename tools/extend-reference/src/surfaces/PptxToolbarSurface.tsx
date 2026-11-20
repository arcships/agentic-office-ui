/**
 * Issue #161 — PPTX toolbar parity surface.
 *
 * The reference is the pinned Extend `PptxToolbar` row (vendored in
 * ../vendor/extend/pptx-toolbar) with the fixture presentation: slide 3 of 12,
 * 100% zoom, thumbnail rail closed, download/upload available. Both sides
 * render the same resting row; hover/focus/press target the first control.
 */
import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { PptxToolbar } from "../vendor/extend/pptx-toolbar";

const toolbar = fixtures.pptxToolbar;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "active", matrixState: "active", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

export function PptxToolbarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <PptxToolbar
            activeSlideIndex={toolbar.activeSlideIndex}
            controlsDisabled={entry.disabled}
            showDownloadButton={toolbar.showDownload}
            showUploadButton={toolbar.showUpload}
            slideCount={entry.disabled ? 0 : toolbar.slideCount}
            resolvedZoom={toolbar.zoomPercent}
          />
        </ParityCell>
      ))}
    </>
  );
}
