/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/pptx-viewer.tsx (PptxSidebarThumbnail + the
 * `role="option"` item wrapper from PptxThumbnailSidebarList) and
 * apps/v4/components/extend/file-thumbnail.tsx (rendered by ../extend/file-thumbnail)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the slide canvas inside `previewContent` is a static fixture
 * image instead of the `@extend-ai/react-pptx` render window. The item, label,
 * frame and active/hover class strings are the upstream ones.
 */
import type * as React from "react";

import { cn } from "../../lib/cn";
import { FileThumbnail } from "./file-thumbnail";

const PPTX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export interface PptxSlideCardProps {
  slideNumber: number;
  active: boolean;
  /** Preview aspect ratio (width / height). */
  aspectRatio: number;
  imageUrl: string;
  fileName?: string;
}

export function PptxSlideCard({
  slideNumber,
  active,
  aspectRatio,
  imageUrl,
  fileName = "roadmap-review.pptx",
}: PptxSlideCardProps): React.ReactElement {
  return (
    <div
      role="option"
      aria-current={active ? "page" : undefined}
      aria-label={`Slide ${slideNumber}`}
      aria-selected={active}
      data-slide-index={slideNumber - 1}
      data-active={active ? "true" : "false"}
      className={cn(
        "flex h-auto w-full cursor-default flex-col items-center gap-2 rounded-md p-2 text-xs transition-colors outline-none select-none hover:bg-sidebar-accent",
        active
          ? "bg-sidebar-accent text-foreground"
          : "text-muted-foreground",
      )}
    >
      <FileThumbnail
        file={{
          name: `${fileName} slide ${slideNumber}`,
          type: PPTX_MIME_TYPE,
        }}
        previewAspectRatio={aspectRatio}
        previewClassName="rounded-sm bg-white"
        previewContent={
          <img
            src={imageUrl}
            alt=""
            draggable={false}
            className="size-full object-cover"
          />
        }
        className={cn(
          "w-full rounded-sm border-0 shadow-xs ring-0 transition-shadow duration-150",
          active && "shadow-sm",
        )}
      />
      {slideNumber}
    </div>
  );
}
