import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import {
  DocumentViewerSidebarSkeleton,
  DocumentViewerThumbnailSidebar,
} from "../vendor/extend/document-viewer-sidebar";
import { FileThumbnail } from "../vendor/extend/file-thumbnail";

const sidebar = fixtures.documentSidebar;
const pages = fixtures.documents.filter((entry) => entry.pages);

// Fixture children only: the reference sidebar hosts the PDF/DOCX thumbnail
// strip, which is viewer state rather than a reusable sidebar primitive.
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
  { state: "open", matrixState: "default" },
  { state: "closed", matrixState: "default" },
  { state: "skeleton", matrixState: "loading" },
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
          {entry.state === "skeleton" ? (
            <div
              style={{ width: sidebar.widthPx, height: 240, overflow: "hidden" }}
              data-document-thumbnail-sidebar=""
              data-sidebar-mode="inline"
              data-sidebar-open="true"
            >
              <DocumentViewerSidebarSkeleton inline />
            </div>
          ) : (
            <div
              style={{
                width: sidebar.widthPx,
                height: 240,
                overflow: "hidden",
                position: "relative",
                background: "var(--sidebar)",
              }}
            >
              <DocumentViewerThumbnailSidebar
                inline={sidebar.inline}
                open={entry.state === "open"}
              >
                <ThumbnailStrip />
              </DocumentViewerThumbnailSidebar>
            </div>
          )}
        </ParityCell>
      ))}
    </>
  );
}
