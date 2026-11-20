import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { XlsxToolbar } from "../vendor/extend/xlsx-toolbar";

const toolbar = fixtures.xlsxToolbar;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

export function XlsxToolbarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <XlsxToolbar
            controlsDisabled={entry.disabled}
            zoomPercent={toolbar.zoomPercent}
          />
        </ParityCell>
      ))}
    </>
  );
}
