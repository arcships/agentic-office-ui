/**
 * Issue #162 — PDF thumbnail/outline sidebar parity surface.
 *
 * The reference is the pinned Extend `DocumentViewerThumbnailSidebar` with the
 * real `PDFViewerThumbnails` markup (vendored in ../vendor/extend/pdf-sidebar).
 * Both sides use four 600×800 pages, one shared PNG thumbnail, numeric page
 * labels, active page 1, a 160px rail and the same 768px inline breakpoint.
 */
import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import {
  DocumentViewerSidebarSkeleton,
  DocumentViewerThumbnailSidebar,
} from "../vendor/extend/document-viewer-sidebar";
import { PdfThumbnailSidebarContent } from "../vendor/extend/pdf-sidebar";
import pdfThumbnailUrl from "../../../parity-fixtures/assets/pdf-thumbnail.png";

const frameHeight = 720;
const pageCount = 4;
const pageWidth = 600;
const pageHeight = 800;
const activePage = 1;
const sidebar = fixtures.documentSidebar;

const states = [
  { state: "open", matrixState: "default", open: true, skeleton: false },
  { state: "closed", matrixState: "default", open: false, skeleton: false },
  { state: "skeleton", matrixState: "loading", open: true, skeleton: true },
];

function inlineSidebar(): boolean {
  return window.innerWidth >= 768;
}

export function PdfSidebarSurface({ cellWidth }: { cellWidth: number }) {
  const inline = inlineSidebar();
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
            data-pdf-sidebar-frame=""
            style={{
              width: sidebar.widthPx,
              height: frameHeight,
              overflow: "hidden",
              position: "relative",
              display: entry.skeleton ? "block" : "flex",
              alignItems: "stretch",
            }}
          >
            {entry.skeleton ? (
              <DocumentViewerSidebarSkeleton inline />
            ) : (
              <DocumentViewerThumbnailSidebar
                inline={inline}
                open={entry.open}
                widthClassName="w-40"
                closedInlineClassName="-ml-40"
              >
                {entry.open ? (
                  <PdfThumbnailSidebarContent
                    activePage={activePage}
                    pageCount={pageCount}
                    pageHeight={pageHeight}
                    pageWidth={pageWidth}
                    thumbnailUrl={pdfThumbnailUrl}
                  />
                ) : null}
              </DocumentViewerThumbnailSidebar>
            )}
          </div>
        </ParityCell>
      ))}
    </>
  );
}
