import { ParityCell } from "../parity/ParitySurface";
import { DocumentViewerThumbnailSidebar } from "../vendor/extend/document-viewer-sidebar";

const widthPx = 160;
const pageImage = "/docx-sidebar-page-1.png";

const mutedText = "var(--muted-foreground)";

function SidebarHeader({ count }: { count: number }) {
  return (
    <header
      style={{
        display: "flex",
        alignItems: "center",
        gap: 4,
        padding: "8px 12px",
        borderBottom: "1px solid var(--docx-header-border)",
        boxSizing: "border-box",
        height: 37,
        flex: "0 0 auto",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          width: "100%",
          gap: 8,
        }}
      >
        <span
          style={{
            color: "var(--docx-title)",
            fontSize: 12,
            fontWeight: 600,
            lineHeight: "18px",
          }}
        >
          Pages
        </span>
        <span
          style={{
            color: mutedText,
            fontSize: 12,
            lineHeight: "18px",
            padding: "1px 8px",
            borderRadius: 18,
            background: "var(--docx-skeleton-bg)",
          }}
        >
          {count}
        </span>
      </div>
    </header>
  );
}

function PageThumbnail() {
  return (
    <div
      style={{
        padding: 4,
        marginBottom: 12,
        borderRadius: 4,
        background: "var(--docx-thumbnail-active)",
        outline: "2px solid var(--docx-thumbnail-ring)",
        outlineOffset: -2,
        boxSizing: "border-box",
      }}
    >
      <div
        style={{
          width: 140,
          height: 210,
          overflow: "hidden",
          borderRadius: 2,
          background: "#fff",
          boxShadow: "0 1px 4px rgba(0, 0, 0, 0.08)",
        }}
      >
        <img
          alt=""
          src={pageImage}
          style={{ display: "block", width: 140, height: 210 }}
        />
      </div>
      <div
        style={{
          marginTop: 2,
          color: "var(--docx-thumbnail-label)",
          fontSize: 11,
          lineHeight: "14px",
          textAlign: "center",
        }}
      >
        1
      </div>
    </div>
  );
}

function SkeletonBody() {
  return (
    <div
      role="status"
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 12,
        padding: 8,
        color: mutedText,
      }}
    >
      <span
        style={{
          color: mutedText,
          fontSize: 12,
          lineHeight: "18px",
          textAlign: "center",
        }}
      >
        Opening document
      </span>
      {[0, 1, 2].map((index) => (
        <div
          key={index}
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: 4,
          }}
        >
          <div
            style={{
              width: 80,
              height: 112,
              borderRadius: 8,
              background: "var(--docx-skeleton-bg)",
            }}
          />
          <div
            style={{
              width: 32,
              height: 8,
              borderRadius: 6,
              background: "var(--docx-skeleton-bg)",
            }}
          />
        </div>
      ))}
    </div>
  );
}

const states = [
  { state: "open", matrixState: "default", open: true, loading: false },
  { state: "closed", matrixState: "default", open: false, loading: false },
  { state: "skeleton", matrixState: "loading", open: true, loading: true },
];

export function DocxSidebarSurface({ cellWidth }: { cellWidth: number }) {
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
              width: widthPx,
              height: entry.loading ? 240 : 292,
              overflow: "hidden",
              position: "relative",
              background: "var(--sidebar)",
            }}
          >
            <DocumentViewerThumbnailSidebar
              inline={false}
              open={entry.open}
              className={entry.loading ? "border-r-0 shadow-sm" : "shadow-sm"}
            >
              <div
                style={{
                  display: "flex",
                  flexDirection: "column",
                  height: "100%",
                  background: "var(--docx-sidebar-bg)",
                  borderRightWidth: entry.loading ? 0 : undefined,
                }}
              >
                <SidebarHeader count={entry.loading ? 0 : 1} />
                <div style={{ flex: "1 1 auto", overflow: "hidden" }}>
                  {entry.loading ? (
                    <SkeletonBody />
                  ) : (
                    <div style={{ padding: 4 }}>
                      <PageThumbnail />
                    </div>
                  )}
                </div>
              </div>
            </DocumentViewerThumbnailSidebar>
          </div>
        </ParityCell>
      ))}
    </>
  );
}
