import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { XlsxRibbon } from "../vendor/extend/xlsx-ribbon";

const ribbon = fixtures.xlsxRibbon;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

export function XlsxRibbonSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <XlsxRibbon
            borderColor={ribbon.borderColor}
            canMerge={ribbon.canMerge}
            canRedo={ribbon.canRedo}
            canUndo={ribbon.canUndo}
            controlsDisabled={entry.disabled}
            fillColor={ribbon.fillColor}
            fontFamily={ribbon.fontFamily}
            fontSize={ribbon.fontSize}
            fontWeight={{
              bold: false,
              italic: false,
              strikethrough: false,
              underline: false,
            }}
            horizontalAlignment={
              ribbon.horizontalAlignment as "left" | "center" | "right"
            }
            isMerged={ribbon.isMerged}
            numberFormat={
              ribbon.numberFormat as
                | "general"
                | "number"
                | "currency"
                | "percent"
                | "date"
            }
            textColor={ribbon.textColor}
            verticalAlignment={
              ribbon.verticalAlignment as "top" | "center" | "bottom"
            }
          />
        </ParityCell>
      ))}
    </>
  );
}
