/**
 * Issue #158 — `OfficeLoadingState` parity surface.
 *
 * The reference is the pinned `ViewerLoadingSurface` (apps/v4/components/extend/
 * docx-viewer.tsx and xlsx-viewer.tsx): `grid h-full min-h-52
 * place-items-center bg-transparent` with the 16px `Spinner`. Each cell gets a
 * 320×208 frame so `h-full` and `min-h-52` resolve to the same box, and the
 * target is the state root itself, so the 208px min-height, the centered
 * placement and the spinner size are compared.
 *
 * `default` is the upstream surface verbatim. Our subject reaches it by passing
 * no title/description text; `hidden` covers the upstream `showSpinner={false}`
 * branch (an empty centered grid).
 */
import { ParityCell } from "../parity/ParitySurface";
import { ViewerLoadingSurface } from "../vendor/extend/viewer-states";

const frameHeight = 208;

const states = [
  { state: "default", matrixState: "loading", showSpinner: true },
  { state: "hidden", matrixState: "empty", showSpinner: false },
];

export function OfficeLoadingStateSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <div style={{ height: frameHeight }}>
            <ViewerLoadingSurface showSpinner={entry.showSpinner} />
          </div>
        </ParityCell>
      ))}
    </>
  );
}
