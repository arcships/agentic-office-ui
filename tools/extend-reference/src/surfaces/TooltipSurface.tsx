import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { TooltipPopup } from "../vendor/ui/tooltip";

const copy = fixtures.copy.tooltip;

export function TooltipSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <ParityCell matrixState="hover" state="open" width={cellWidth}>
      <div className="relative inline-flex">
        <button
          className="text-sm"
          data-parity-tooltip-trigger
          type="button"
        >
          {copy.trigger}
        </button>
        <div className="absolute left-full top-1/2 ml-1 -translate-y-1/2">
          <TooltipPopup className="whitespace-nowrap [transform:translateZ(0)]">
            {copy.content}
          </TooltipPopup>
        </div>
      </div>
    </ParityCell>
  );
}
