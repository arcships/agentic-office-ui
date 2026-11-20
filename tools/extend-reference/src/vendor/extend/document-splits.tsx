/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/document-splits.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (parity harness only):
 * - `@dnd-kit` is not vendored. The sortable wrappers render the same card
 *   markup and classes; drag/drop and keyboard reordering are covered by the
 *   Vue subject's interaction tests.
 * - The Base UI ScrollArea is replaced by plain scroll viewports.
 * - Loading / error / empty states are harness extensions with the same class
 *   contract as the Vue subject.
 */
import * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  Add01Icon,
  Delete02Icon,
  DragDropVerticalIcon,
} from "@hugeicons/core-free-icons";

import { cn } from "../../lib/cn";
import { Button } from "../ui/button";
import { FileThumbnail } from "./file-thumbnail";

export type DocumentSplit = {
  id: string;
  title: string;
  pages: string[];
};

const THUMBNAIL_WIDTH = 72;
const THUMBNAIL_HEIGHT = 92;
const THUMBNAIL_SURFACE_CLASS_NAME =
  "relative overflow-hidden rounded-lg border border-border bg-background shadow-xs";
const THUMBNAIL_BUTTON_CLASS_NAME = cn(
  THUMBNAIL_SURFACE_CLASS_NAME,
  "cursor-grab text-left transition-[border-color,box-shadow,opacity] outline-none hover:border-ring/40 hover:ring-2 hover:ring-ring focus-visible:border-ring/40 focus-visible:ring-2 focus-visible:ring-ring active:cursor-grabbing"
);

function pageNumber(pageId: string) {
  return Number(pageId.replace("page-", ""));
}

function pageRangeLabel(group: DocumentSplit) {
  if (group.pages.length === 0) return "No pages";
  return `Pages ${group.pages.map(pageNumber).join(", ")}`;
}

function SplitGroupCard({
  canRemove,
  group,
  thumbnailImages,
}: {
  canRemove: boolean;
  group: DocumentSplit;
  thumbnailImages: Record<string, string>;
}) {
  return (
    <section
      className="w-full rounded-lg border bg-background"
      data-split-id={group.id}
    >
      <div className="flex items-center justify-between gap-3 border-b p-3">
        <div className="flex min-w-0 items-center gap-2">
          <button
            type="button"
            aria-label={`Reorder ${group.title}`}
            className="inline-flex size-7 shrink-0 cursor-grab items-center justify-center rounded-md text-muted-foreground hover:bg-muted active:cursor-grabbing"
          >
            <HugeiconsIcon icon={DragDropVerticalIcon} className="size-4" />
          </button>
          <div className="min-w-0">
            <div className="truncate text-sm font-medium">{group.title}</div>
            <div className="mt-1 text-xs text-muted-foreground">
              {pageRangeLabel(group)}
            </div>
          </div>
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <div className="rounded-full bg-muted px-2 py-0.5 text-xs text-muted-foreground">
            {group.pages.length}
          </div>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label={`Remove ${group.title}`}
            disabled={!canRemove}
          >
            <HugeiconsIcon icon={Delete02Icon} className="size-4" />
          </Button>
        </div>
      </div>
      <div className="min-h-[116px] rounded-lg p-2">
        {group.pages.length ? (
          <div className="h-[110px] w-full overflow-hidden overflow-x-auto overflow-y-hidden pb-3.5">
            <div className="flex w-max gap-2 overflow-y-hidden py-1 pr-6">
              {group.pages.map((pageId) => (
                <button
                  key={pageId}
                  type="button"
                  data-page-id={pageId}
                  className={cn(THUMBNAIL_BUTTON_CLASS_NAME, "shrink-0")}
                  style={{
                    width: THUMBNAIL_WIDTH,
                    height: THUMBNAIL_HEIGHT,
                  }}
                >
                  <FileThumbnail
                    file={{
                      name: `page-${pageNumber(pageId)}.pdf`,
                      type: "application/pdf",
                    }}
                    previewImageUrl={thumbnailImages[pageId]}
                    isLoading={!thumbnailImages[pageId]}
                    previewClassName="h-full aspect-auto"
                    className="size-full rounded-[inherit] border-0 bg-transparent"
                  />
                  <span className="absolute right-1 bottom-1 rounded bg-background/95 px-1.5 py-0.5 text-[10px] font-semibold text-foreground shadow-xs ring-1 ring-border/80">
                    {pageNumber(pageId)}
                  </span>
                </button>
              ))}
            </div>
          </div>
        ) : (
          <div className="grid h-[104px] place-items-center rounded-lg bg-muted/35 text-xs text-muted-foreground">
            Drop pages here
          </div>
        )}
      </div>
    </section>
  );
}

export function DocumentSplitsPanel({
  splits,
  thumbnailImages,
  loading = false,
  error,
  emptyMessage = "No document splits yet.",
  hideAddSplit = false,
}: {
  splits: DocumentSplit[];
  thumbnailImages: Record<string, string>;
  loading?: boolean;
  error?: string;
  emptyMessage?: string;
  hideAddSplit?: boolean;
}) {
  return (
    <aside
      className="flex h-full min-h-0 flex-col bg-muted/20"
      data-document-splits
    >
      {hideAddSplit ? null : (
        <div className="flex min-h-12 items-center justify-end gap-3 border-b bg-background px-3">
          <Button type="button" variant="outline" size="sm">
            <HugeiconsIcon icon={Add01Icon} className="size-4" />
            Add split
          </Button>
        </div>
      )}
      <div className="min-h-0 flex-1 overflow-y-auto">
        {error ? (
          <div className="p-3">
            <div
              role="alert"
              className="rounded-lg border border-destructive/45 bg-destructive/8 p-3 text-sm text-destructive"
            >
              {error}
            </div>
          </div>
        ) : loading ? (
          <div className="p-3">
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              Loading document splits…
            </div>
          </div>
        ) : splits.length === 0 ? (
          <div className="p-3">
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              {emptyMessage}
            </div>
          </div>
        ) : (
          <div className="space-y-3 p-3">
            {splits.map((group) => (
              <SplitGroupCard
                key={group.id}
                canRemove={splits.length > 1}
                group={group}
                thumbnailImages={thumbnailImages}
              />
            ))}
          </div>
        )}
      </div>
    </aside>
  );
}
