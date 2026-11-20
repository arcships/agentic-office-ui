/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Sources: apps/v4/components/extend/pdf-viewer.tsx (PDFViewerThumbnails +
 * PDFViewerThumbnailScrollArea) and apps/v4/components/extend/document-viewer-sidebar.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the EmbedPDF thumbnail plugin is replaced by a deterministic
 * static layout fed by the parity fixture (page count, 600×800 page boxes,
 * one thumbnail image). The rendered classes, thumbnail geometry, listbox
 * roles and active/focus states are copied verbatim from the pinned source.
 */
import * as React from "react";

import { cn } from "../../lib/cn";

const THUMBNAIL_PAGE_WIDTH = 92;
const THUMBNAIL_IMAGE_PADDING = 8;
const THUMBNAIL_WIDTH = THUMBNAIL_PAGE_WIDTH + THUMBNAIL_IMAGE_PADDING * 2;
const THUMBNAIL_LABEL_HEIGHT = 24;
const THUMBNAIL_GAP = 12;
const THUMBNAIL_PANE_PADDING_Y = 16;
const THUMBNAIL_FOCUS_RING_CLASS =
  "group-focus-visible/pdf-thumbnail-sidebar:ring-2 group-focus-visible/pdf-thumbnail-sidebar:ring-ring group-focus-visible/pdf-thumbnail-sidebar:ring-offset-1 group-focus-visible/pdf-thumbnail-sidebar:ring-offset-background";

interface ThumbMeta {
  pageIndex: number;
  width: number;
  height: number;
  wrapperHeight: number;
  top: number;
  labelHeight: number;
  padding: number;
}

function buildLayout(pageCount: number, pageWidth: number, pageHeight: number): {
  items: ThumbMeta[];
  totalHeight: number;
} {
  const innerWidth = Math.max(1, THUMBNAIL_WIDTH - THUMBNAIL_IMAGE_PADDING * 2);
  const imageHeight = Math.round(innerWidth * (pageHeight / pageWidth));
  const wrapperHeight =
    THUMBNAIL_IMAGE_PADDING + imageHeight + THUMBNAIL_IMAGE_PADDING + THUMBNAIL_LABEL_HEIGHT;
  let top = THUMBNAIL_PANE_PADDING_Y;
  const items = Array.from({ length: pageCount }, (_, pageIndex) => {
    const meta: ThumbMeta = {
      pageIndex,
      width: innerWidth,
      height: imageHeight,
      wrapperHeight,
      top,
      labelHeight: THUMBNAIL_LABEL_HEIGHT,
      padding: THUMBNAIL_IMAGE_PADDING,
    };
    top += wrapperHeight + THUMBNAIL_GAP;
    return meta;
  });
  return {
    items,
    totalHeight: items.length
      ? top - THUMBNAIL_GAP + THUMBNAIL_PANE_PADDING_Y
      : THUMBNAIL_PANE_PADDING_Y * 2,
  };
}

export function PdfThumbnailSidebarContent({
  activePage,
  pageCount,
  pageHeight,
  pageWidth,
  thumbnailUrl,
}: {
  activePage: number;
  pageCount: number;
  pageHeight: number;
  pageWidth: number;
  thumbnailUrl: string;
}) {
  const thumbnailListboxId = React.useId();
  const activeDescendantId =
    activePage > 0 ? `${thumbnailListboxId}-page-${activePage}` : undefined;
  const layout = buildLayout(pageCount, pageWidth, pageHeight);
  const pageLabel = (pageNumber: number) => String(pageNumber);

  return (
    <div
      aria-activedescendant={activeDescendantId}
      aria-label="PDF pages"
      aria-multiselectable
      className="group/pdf-thumbnail-sidebar h-full w-full overflow-y-auto px-4 focus-visible:ring-0 focus-visible:ring-offset-0"
      role="listbox"
      style={{
        paddingBottom: THUMBNAIL_PANE_PADDING_Y,
        paddingTop: THUMBNAIL_PANE_PADDING_Y,
      }}
      tabIndex={0}
    >
      <div className="relative" style={{ height: layout.totalHeight }}>
        {layout.items.map((meta) => {
          const pageNumber = meta.pageIndex + 1;
          const isActive = pageNumber === activePage;
          return (
            <div
              key={meta.pageIndex}
              className={cn(
                "absolute right-0 left-0 flex justify-center",
                isActive && "z-10",
              )}
              data-pdf-sidebar-thumbnail={pageNumber}
              style={{ top: meta.top, height: meta.wrapperHeight }}
            >
              <div
                id={`${thumbnailListboxId}-page-${pageNumber}`}
                aria-current={isActive ? "page" : undefined}
                aria-label={`Page ${pageNumber}`}
                aria-posinset={pageNumber}
                aria-selected={false}
                aria-setsize={pageCount}
                className={cn(
                  "flex h-full w-full cursor-default flex-col items-center justify-between rounded-md px-2 py-0 text-xs transition-shadow outline-none select-none hover:bg-sidebar-accent",
                  isActive
                    ? "bg-sidebar-accent text-foreground"
                    : "text-muted-foreground",
                  isActive && THUMBNAIL_FOCUS_RING_CLASS,
                )}
                data-pdf-sidebar-page-option={pageNumber}
                role="option"
              >
                <span
                  className="mt-0 flex items-center justify-center overflow-hidden rounded-md bg-transparent"
                  style={{
                    width: meta.width + meta.padding * 2,
                    height: meta.height + meta.padding * 2,
                    padding: meta.padding,
                  }}
                >
                  <img
                    alt=""
                    className="block rounded-sm object-contain"
                    src={thumbnailUrl}
                    style={{ width: meta.width, height: meta.height }}
                  />
                </span>
                <span
                  className="flex items-center justify-center tabular-nums"
                  style={{ height: meta.labelHeight }}
                >
                  <span className="flex min-w-5 items-center justify-center px-1.5 text-center leading-5">
                    {pageLabel(pageNumber)}
                  </span>
                </span>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
