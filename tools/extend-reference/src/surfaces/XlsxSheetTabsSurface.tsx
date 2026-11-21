import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { XlsxSheetTabs } from "../vendor/extend/xlsx-sheet-tabs";

const tabs = fixtures.xlsxSheetTabs;

const states = [
  { state: "default", matrixState: "default" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
];

export function XlsxSheetTabsSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <XlsxSheetTabs activeIndex={tabs.activeIndex} sheets={tabs.sheets} />
        </ParityCell>
      ))}
    </>
  );
}
