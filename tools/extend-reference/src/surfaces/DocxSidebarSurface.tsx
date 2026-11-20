/**
 * Issue #159 — DOCX thumbnail sidebar parity surface.
 *
 * The reference is the pinned Extend `DocumentViewerThumbnailSidebar` (160px
 * `w-40` rail) with the real `DocxThumbnailSidebarContent` markup (vendored in
 * ../vendor/extend/docx-sidebar): no header, one 92px page card with a 2:3
 * page image, a page-number label, the active `bg-sidebar-accent` state and the
 * upstream `p-4` loading skeleton.
 *
 * Both sides use the same single-page fixture, the same page aspect ratio and
 * a 160×292 frame.
 */
import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { DocumentViewerThumbnailSidebar } from "../vendor/extend/document-viewer-sidebar";
import { DocxThumbnailSidebarList } from "../vendor/extend/docx-sidebar";
import pageImageUrl from "../../../parity-fixtures/assets/docx-thumbnail.png";

const pageImageAspectRatio = 200 / 300;
const displayFileName = "legal-contract.docx";
const sidebar = fixtures.documentSidebar;

const states = [
  { state: "open", matrixState: "default", open: true, loading: false },
  { state: "closed", matrixState: "default", open: false, loading: false },
  { state: "skeleton", matrixState: "loading", open: true, loading: true },
];

function inlineSidebar(): boolean {
  return window.innerWidth >= 768;
}

export function DocxSidebarSurface({ cellWidth }: { cellWidth: number }) {
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
            style={{
              width: sidebar.widthPx,
              height: entry.loading ? 240 : 292,
              overflow: "hidden",
              position: "relative",
              display: "flex",
              alignItems: "stretch",
              background: "var(--sidebar)",
            }}
          >
            <DocumentViewerThumbnailSidebar
              inline={inline}
              open={entry.open}
              widthClassName="w-40"
              closedInlineClassName="-ml-40"
            >
              <DocxThumbnailSidebarList
                activePage={1}
                displayFileName={displayFileName}
                isLoadingDocument={entry.loading}
                pageCount={1}
                pageImageAspectRatio={pageImageAspectRatio}
                pageImageUrl={pageImageUrl}
              />
            </DocumentViewerThumbnailSidebar>
          </div>
        </ParityCell>
      ))}
    </>
  );
}
