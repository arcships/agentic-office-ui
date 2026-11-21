import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { XlsxStatusStrip } from "../vendor/extend/xlsx-status-bar";

const bar = fixtures.xlsxStatusBar;

const states = [{ state: "default", matrixState: "default" }];

export function XlsxStatusBarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <XlsxStatusStrip segments={bar.segments} />
        </ParityCell>
      ))}
    </>
  );
}
