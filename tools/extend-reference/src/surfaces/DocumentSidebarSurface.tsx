/**
 * Issue #158 — the shared `DocumentViewerSidebar` parity surface.
 *
 * `apps/v4/components/extend/document-viewer-sidebar.tsx` is the only upstream
 * source for the rail geometry: a 160px `w-40` `bg-sidebar` box with a `border-r`,
 * `-ml-40` when an inline sidebar closes, `-translate-x-full` + `shadow-lg` when
 * an overlay one closes, and a `p-4` one-card loading skeleton.
 *
 * The cells therefore cover both layout strategies:
 *   - inline: the aside is a 160px flex child of a 160px frame (1440/1280
 *     effectively; the frame width is viewport independent, like `PdfSidebar`).
 *   - overlay: the aside is absolutely positioned inside a 196px frame that
 *     stands in for the document surface, so the 160px rail, the closed
 *     translate and the `shadow-lg` edge are all part of the diff.
 *
 * Children are fixtures, not viewer state: the canonical sidebar hosts whatever
 * thumbnails its host renders (#159 DOCX, #162 PDF keep their own strips).
 */
import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import {
  DocumentViewerSidebarSkeleton,
  DocumentViewerThumbnailSidebar,
} from "../vendor/extend/document-viewer-sidebar";
import { FileThumbnail } from "../vendor/extend/file-thumbnail";

const sidebar = fixtures.documentSidebar;
const pages = fixtures.documents.filter((entry) => entry.pages);
const frameHeight = 240;

function ThumbnailStrip() {
  return (
    <div className="flex h-full flex-col gap-3 overflow-hidden p-2">
      {pages.map((page, index) => (
        <div key={page.id} className="flex flex-col items-center gap-1">
          <FileThumbnail
            file={{ name: page.name, type: page.type }}
            className="h-28 w-20 rounded-md"
          />
          <span className="text-xs text-muted-foreground">{index + 1}</span>
        </div>
      ))}
    </div>
  );
}

const states = [
  { state: "open", matrixState: "default", open: true, overlay: false, skeleton: false },
  { state: "closed", matrixState: "default", open: false, overlay: false, skeleton: false },
  { state: "skeleton", matrixState: "loading", open: true, overlay: false, skeleton: true },
  { state: "overlay-open", matrixState: "default", open: true, overlay: true, skeleton: false },
  { state: "overlay-closed", matrixState: "default", open: false, overlay: true, skeleton: false },
];

export function DocumentSidebarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <div
            data-document-sidebar-frame=""
            style={{
              width: entry.overlay ? cellWidth : sidebar.widthPx,
              height: frameHeight,
              overflow: "hidden",
              position: "relative",
              background: entry.skeleton
                ? undefined
                : entry.overlay
                  ? "var(--muted)"
                  : "var(--sidebar)",
              display: entry.skeleton || entry.overlay ? "block" : "flex",
              alignItems: "stretch",
            }}
          >
            {entry.skeleton ? (
              <DocumentViewerSidebarSkeleton inline />
            ) : (
              <DocumentViewerThumbnailSidebar
                inline={!entry.overlay}
                open={entry.open}
              >
                <ThumbnailStrip />
              </DocumentViewerThumbnailSidebar>
            )}
          </div>
        </ParityCell>
      ))}
    </>
  );
}
