/**
 * Issue #159 — DOCX toolbar parity surface.
 *
 * The reference is the pinned Extend `DocxToolbar` from
 * `apps/v4/components/extend/docx-viewer.tsx` (vendored in
 * ../vendor/extend/docx-toolbar): toggle thumbnails, Page N of M, zoom out /
 * level / zoom in, separator and the closed DOCX actions trigger. Every other
 * DOCX command the product offers lives in the actions menu, exactly like the
 * upstream toolbar.
 */
import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { DocxToolbar } from "../vendor/extend/docx-toolbar";

const toolbar = fixtures.docxToolbar;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

export function DocxToolbarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <DocxToolbar
            activePage={toolbar.activePage}
            controlsDisabled={entry.disabled}
            pageCount={entry.disabled ? 0 : toolbar.pageCount}
            zoomPercent={toolbar.zoomPercent}
          />
        </ParityCell>
      ))}
    </>
  );
}
