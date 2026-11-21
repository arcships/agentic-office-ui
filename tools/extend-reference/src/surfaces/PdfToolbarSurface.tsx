import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { PdfToolbar } from "../vendor/extend/pdf-toolbar";

const toolbar = fixtures.pdfToolbar;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "active", matrixState: "active", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

export function PdfToolbarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <PdfToolbar
            activePage={toolbar.activePage}
            controlsDisabled={entry.disabled}
            numPages={entry.disabled ? 0 : toolbar.numPages}
            showDownload={toolbar.showDownload}
            showRotateControls={toolbar.showRotateControls}
            showUpload={toolbar.showUpload}
            sidebarOpen={toolbar.sidebarOpen}
            zoomLevel={toolbar.zoomPercent / 100}
          />
        </ParityCell>
      ))}
    </>
  );
}
