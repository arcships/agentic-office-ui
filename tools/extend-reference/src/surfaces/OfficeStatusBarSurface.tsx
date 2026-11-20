/**
 * Issue #158 — `OfficeStatusBar` parity surface.
 *
 * Extend ships no standalone status bar; every viewer appends a footer strip.
 * The reference composes the pinned footer contract (`border-t`, `bg-muted/40`,
 * `px-3 py-2`, `text-xs text-muted-foreground`, `gap-1`/`gap-2`, `aria-live`
 * footer from file-system.tsx; sheet-tab strip from xlsx-viewer.tsx; hint rows
 * from pdf-editor-panels.tsx) with the page/status/file content OfficeStatusBar
 * renders, so the diff pins the shared footer metrics (33px row, 12/16 type,
 * muted palette, left/right split) rather than one upstream class list.
 */
import { ParityCell } from "../parity/ParitySurface";
import { ViewerStatusFooter } from "../vendor/extend/viewer-states";

const copy = {
  page: "Page 2 of 9",
  status: "Saved",
  fileName: "report.pdf",
};

const states = [{ state: "default", matrixState: "default" }];

export function OfficeStatusBarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <ViewerStatusFooter
            fileName={copy.fileName}
            page={copy.page}
            status={copy.status}
          />
        </ParityCell>
      ))}
    </>
  );
}
