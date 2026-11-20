/**
 * Issue #158 — `OfficeErrorState` parity surface.
 *
 * The reference is the pinned viewer error card (apps/v4/components/extend/
 * xlsx-viewer.tsx `loadError` fallback) inside the upstream centered host
 * (`grid place-items-center bg-muted/30 p-4`): `max-w-md rounded-lg border
 * bg-background p-4 text-sm`, a `font-medium` title, an `mt-1
 * text-muted-foreground` message and the `mt-4` outline action.
 *
 * The target is the card itself, so the card metrics (padding, 1px border,
 * 10px radius, 14/20 type) are compared geometrically as well as by pixels.
 * `message` drops the action row (`showRetry=false` on the subject).
 */
import { ParityCell } from "../parity/ParitySurface";
import { Button } from "../vendor/ui/button";
import { WorkbookErrorCard } from "../vendor/extend/viewer-states";

const copy = {
  title: "Unable to open this document",
  description: "Check the file and try again.",
};

const states = [
  { state: "default", matrixState: "error", showAction: true },
  { state: "message", matrixState: "error", showAction: false },
];

export function OfficeErrorStateSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <div className="grid w-full place-items-center bg-muted/30 p-4">
            <WorkbookErrorCard
              message={copy.description}
              showAction={entry.showAction}
              title={copy.title}
            />
          </div>
        </ParityCell>
      ))}
    </>
  );
}
