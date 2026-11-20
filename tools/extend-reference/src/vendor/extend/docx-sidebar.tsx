/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Sources: apps/v4/components/extend/docx-viewer.tsx (DocxSidebarThumbnail,
 * DocxThumbnailSidebarList, DocxThumbnailSidebarContent),
 * apps/v4/components/extend/document-viewer-sidebar.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the virtualizer, the thumbnail render window and the ScrollArea
 * primitive are replaced by a deterministic static layout fed by the parity
 * fixture (page count, one page image, active page). The rendered classes,
 * option roles, docs-width card, label geometry, skeleton and active state are
 * copied verbatim from the pinned source.
 */
import * as React from "react";

import { cn } from "../../lib/cn";
import { FileThumbnail } from "./file-thumbnail";

const DOCX_THUMBNAIL_LIST_PADDING = 16;
const DOCX_THUMBNAIL_FOCUS_RING_CLASS =
  "group-focus-visible/docx-thumbnail-sidebar:ring-2 group-focus-visible/docx-thumbnail-sidebar:ring-ring group-focus-visible/docx-thumbnail-sidebar:ring-offset-1 group-focus-visible/docx-thumbnail-sidebar:ring-offset-background";

/**
 * Upstream renders the page through `previewContent={<canvas/>}`. The harness
 * draws the same 2x page bitmap into the canvas so both sides rasterize a
 * canvas at 1:1 device pixels instead of comparing a canvas with an `<img>`.
 */
function PageThumbnailCanvas({
  pageImageUrl,
  pixelWidthPx,
  pixelHeightPx,
}: {
  pageImageUrl: string;
  pixelWidthPx: number;
  pixelHeightPx: number;
}) {
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);

  React.useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const context = canvas.getContext("2d");
    if (!context) return;
    let cancelled = false;
    const image = new Image();
    image.onload = () => {
      if (cancelled) return;
      context.clearRect(0, 0, canvas.width, canvas.height);
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
    };
    image.src = pageImageUrl;
    return () => {
      cancelled = true;
    };
  }, [pageImageUrl]);

  return (
    <canvas
      ref={canvasRef}
      width={pixelWidthPx}
      height={pixelHeightPx}
      className="!size-full bg-white object-cover object-top"
    />
  );
}

export function DocxThumbnailSidebarList({
  activePage,
  displayFileName,
  isLoadingDocument,
  pageCount,
  pageImageUrl,
  pageImageAspectRatio,
}: {
  activePage: number;
  displayFileName: string;
  isLoadingDocument: boolean;
  pageCount: number;
  pageImageUrl: string;
  pageImageAspectRatio: number;
}) {
  const thumbnailListboxId = React.useId();
  const activeDescendantId =
    pageCount > 0 ? `${thumbnailListboxId}-page-${activePage}` : undefined;
  const label = (pageNumber: number) => String(pageNumber);
  // Upstream sizes the list with `virtualizer.getTotalSize() + PADDING * 2`.
  // This static harness derives the same row box from the card geometry:
  // p-2 + card + gap-2 + one text-xs line + p-2, then `pb-3` on the row.
  const cardHeight = 92 / pageImageAspectRatio;
  const rowHeight = 8 + cardHeight + 8 + 16 + 8 + 12;
  const listHeight = rowHeight + DOCX_THUMBNAIL_LIST_PADDING * 2;

  return (
    <div
      aria-activedescendant={activeDescendantId}
      aria-busy={isLoadingDocument || undefined}
      aria-label="DOCX pages"
      className="group/docx-thumbnail-sidebar h-full w-full overflow-y-auto focus-visible:ring-0 focus-visible:ring-offset-0"
      role="listbox"
      tabIndex={0}
    >
      {isLoadingDocument ? (
        <div className="p-4">
          <div className="mx-auto h-28 w-20 overflow-hidden rounded-md bg-background shadow-xs">
            <div className="h-full animate-pulse bg-muted" />
          </div>
          <div className="mx-auto mt-3 h-3 w-10 rounded-full bg-muted" />
        </div>
      ) : pageCount > 0 ? (
        <div
          className="relative"
          style={{ height: listHeight }}
        >
          <div
            className="absolute top-0 right-3 left-3 pb-3 [contain:layout]"
            style={{ transform: `translateY(${DOCX_THUMBNAIL_LIST_PADDING}px)` }}
          >
            <div
              id={`${thumbnailListboxId}-page-${activePage}`}
              role="option"
              aria-current="page"
              aria-label={`Page ${activePage}`}
              aria-posinset={activePage}
              aria-selected
              aria-setsize={pageCount}
              data-docx-viewer-thumbnail-option={activePage}
              className={cn(
                "flex h-auto w-full cursor-default flex-col items-center gap-2 rounded-md p-2 text-xs transition-shadow outline-none select-none hover:bg-sidebar-accent",
                "bg-sidebar-accent text-foreground",
                DOCX_THUMBNAIL_FOCUS_RING_CLASS,
              )}
            >
              <FileThumbnail
                file={{
                  name: `${displayFileName} page ${activePage}`,
                  type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                }}
                previewAspectRatio={pageImageAspectRatio}
                previewClassName="rounded-md bg-white"
                previewContent={
                  <PageThumbnailCanvas
                    pageImageUrl={pageImageUrl}
                    pixelWidthPx={Math.round(92 * 2)}
                    pixelHeightPx={Math.round((92 / pageImageAspectRatio) * 2)}
                  />
                }
                className={cn(
                  "w-[92px] rounded-md border-0 shadow-xs ring-0 transition-shadow duration-150",
                  "shadow-sm",
                )}
              />
              {label(activePage)}
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
