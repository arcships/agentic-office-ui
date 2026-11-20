/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/docx-viewer.tsx (DocxToolbar,
 * DocxPageNumberControl, DocxFileActionsMenu)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the Radix/shadcn dropdown, tooltip and select popups are reduced
 * to their static triggers; the captured surface compares the closed toolbar.
 * Class strings, spacing, page-control markup and icon sizes are copied from
 * the pinned source so the reference pixels stay representative.
 */
import * as React from "react";
import {
  MinusSignCircleIcon,
  MoreHorizontalIcon,
  PlusSignCircleIcon,
  SidebarLeftIcon,
} from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";

import { Button } from "../ui/button";
import { SelectTrigger, SelectValue } from "../ui/select";
import { Separator } from "../ui/separator";

export function DocxToolbar({
  activePage,
  controlsDisabled,
  pageCount,
  zoomPercent,
}: {
  activePage: number;
  controlsDisabled: boolean;
  pageCount: number;
  zoomPercent: number;
}) {
  return (
    <div
      data-parity-docx-toolbar=""
      className="flex min-h-12 flex-wrap items-center justify-between gap-2 border-b bg-background px-3 py-2"
    >
      <div className="flex min-w-0 flex-wrap items-center gap-2">
        <Button
          type="button"
          variant="ghost"
          size="icon-sm"
          aria-label="Toggle thumbnails"
          disabled={controlsDisabled}
        >
          <HugeiconsIcon icon={SidebarLeftIcon} className="size-4" />
        </Button>
        <div className="flex items-center text-sm whitespace-nowrap text-primary">
          <span>Page</span>
          <Button
            type="button"
            variant="ghost"
            size="sm"
            className="font-normal"
            aria-label={`Current page ${activePage}. Edit page number`}
            disabled={controlsDisabled || !pageCount}
          >
            {activePage}
          </Button>
          <span>of {pageCount || "-"}</span>
        </div>
      </div>
      <div className="ml-auto flex min-w-0 flex-wrap items-center justify-end gap-1">
        <div className="flex flex-none items-center gap-1">
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            disabled={controlsDisabled || zoomPercent <= 50}
            aria-label="Zoom out"
          >
            <HugeiconsIcon icon={MinusSignCircleIcon} className="size-4" />
          </Button>
          <SelectTrigger
            size="sm"
            className="w-[104px] min-w-[104px]"
            aria-label="Zoom level"
            disabled={controlsDisabled}
          >
            <SelectValue>{Math.round(zoomPercent)}%</SelectValue>
          </SelectTrigger>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            disabled={controlsDisabled || zoomPercent >= 200}
            aria-label="Zoom in"
          >
            <HugeiconsIcon icon={PlusSignCircleIcon} className="size-4" />
          </Button>
        </div>
        <Separator orientation="vertical" className="mx-1 h-4 self-center" />
        <Button
          type="button"
          variant="ghost"
          size="icon-sm"
          aria-label="Document actions"
          disabled={controlsDisabled}
        >
          <HugeiconsIcon icon={MoreHorizontalIcon} className="size-4" />
        </Button>
      </div>
    </div>
  );
}
