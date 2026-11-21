/**
 * Issue #158 — `OfficeEmptyState` parity surface.
 *
 * The reference is the pinned `PdfEditorEmptyState`
 * (apps/v4/components/extend/pdf-editor-shared.tsx), the only reusable empty
 * state in the Extend tree: centered column, `gap-2 px-4 py-10`, optional 40px
 * `rounded-full bg-muted` glyph, `text-sm font-medium` title, `max-w-56
 * text-xs` description and an `mt-2` action row.
 *
 * The `default` cell renders the catalog copy (no props) so the default titles
 * are compared too; `action` adds an `outline`/`sm` button, which is the same
 * upstream Button variant OfficeEmptyState hosts in its `actions` slot.
 */
import { ParityCell } from "../parity/ParitySurface";
import { Button } from "../vendor/ui/button";
import { ViewerEmptyState } from "../vendor/extend/viewer-states";

const copy = {
  title: "No document open",
  description: "Open a file to get started.",
  action: "Open",
};

const states = [
  { state: "default", matrixState: "default", action: false },
  { state: "action", matrixState: "default", action: true },
];

export function OfficeEmptyStateSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <ViewerEmptyState
            action={
              entry.action ? (
                <Button size="sm" type="button" variant="outline">
                  {copy.action}
                </Button>
              ) : undefined
            }
            description={copy.description}
            title={copy.title}
          />
        </ParityCell>
      ))}
    </>
  );
}
