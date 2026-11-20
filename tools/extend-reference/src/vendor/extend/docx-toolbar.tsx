/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/docx-viewer.tsx (DocxToolbar,
 * DocxPageNumberControl, DocxFileActionsMenu)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptations (all drop interactive popups, none change resting-state layout):
 * - The hover-only tooltip primitive renders its trigger wrapper unchanged
 *   (`<span class="inline-flex">`) and no tooltip content.
 * - The zoom Select becomes the static SelectTrigger/SelectValue from the
 *   vendored registry select, with the same `w-[104px] min-w-[104px]` classes.
 * - The DOCX actions DropdownMenu trigger becomes a plain ghost icon button
 *   with the same classes; the menu is not rendered because the parity matrix
 *   never opens it.
 * - The page-number control keeps the display (button) branch. The edit branch
 *   and its Input are not vendored.
 */
import type * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  MinusSignCircleIcon,
  MoreHorizontalIcon,
  PlusSignCircleIcon,
  SidebarLeftIcon,
} from "@hugeicons/core-free-icons";

import { Button } from "../ui/button";
import { SelectTrigger, SelectValue } from "../ui/select";
import { Separator } from "../ui/separator";

function TooltipTriggerWrapper({ children }: { children: React.ReactNode }) {
  return <span className="inline-flex">{children}</span>;
}

export interface DocxToolbarProps {
  activePage: number;
  controlsDisabled: boolean;
  pageCount: number;
  zoomPercent: number;
}

export function DocxToolbar({
  activePage,
  controlsDisabled,
  pageCount,
  zoomPercent,
}: DocxToolbarProps): React.ReactElement {
  return (
    <div
      data-parity-docx-toolbar=""
      className="flex min-h-12 flex-wrap items-center justify-between gap-2 border-b bg-background px-3 py-2"
    >
      <div className="flex min-w-0 flex-wrap items-center gap-2">
        <TooltipTriggerWrapper>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Toggle thumbnails"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon icon={SidebarLeftIcon} className="size-4" />
          </Button>
        </TooltipTriggerWrapper>
        <div className="flex items-center text-sm whitespace-nowrap text-primary">
          <span>Page</span>
          <Button
            type="button"
            variant="ghost"
            size="sm"
            className="font-normal"
            aria-label={`Current page ${pageCount ? activePage : 1}. Edit page number`}
            disabled={controlsDisabled || !pageCount}
          >
            {pageCount ? activePage : 1}
          </Button>
          <span>of {pageCount || "-"}</span>
        </div>
      </div>
      <div className="ml-auto flex min-w-0 flex-wrap items-center justify-end gap-1">
        <div className="flex flex-none items-center gap-1">
          <TooltipTriggerWrapper>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              disabled={controlsDisabled || zoomPercent <= 50}
              aria-label="Zoom out"
            >
              <HugeiconsIcon icon={MinusSignCircleIcon} className="size-4" />
            </Button>
          </TooltipTriggerWrapper>
          <SelectTrigger
            size="sm"
            className="w-[104px] min-w-[104px]"
            aria-label="Zoom level"
            disabled={controlsDisabled}
          >
            <SelectValue>{Math.round(zoomPercent)}%</SelectValue>
          </SelectTrigger>
          <TooltipTriggerWrapper>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              disabled={controlsDisabled || zoomPercent >= 200}
              aria-label="Zoom in"
            >
              <HugeiconsIcon icon={PlusSignCircleIcon} className="size-4" />
            </Button>
          </TooltipTriggerWrapper>
        </div>
        <Separator orientation="vertical" className="mx-1 h-4 self-center" />
        <Button
          type="button"
          variant="ghost"
          size="icon-sm"
          aria-label="Open DOCX actions"
          disabled={controlsDisabled}
        >
          <HugeiconsIcon icon={MoreHorizontalIcon} className="size-4" />
        </Button>
      </div>
    </div>
  );
}
