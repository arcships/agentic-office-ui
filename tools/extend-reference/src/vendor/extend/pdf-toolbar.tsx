/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/pdf-viewer.tsx (PDFViewerToolbar, ToolbarTooltip,
 * PDFViewerPageNumberControl, PDFViewerFileActionsMenu)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptations (all drop interactive popups, none change resting-state layout):
 * - The hover-only tooltip primitive renders its trigger wrapper unchanged
 *   (`<span class="inline-flex">`) and no tooltip content.
 * - The zoom Select becomes the static SelectTrigger/SelectValue from the
 *   vendored registry select, with the same `w-[104px] min-w-[104px]` classes.
 * - The search Popover trigger and the PDF actions DropdownMenu trigger become
 *   plain ghost icon buttons with the same classes; their popups are not
 *   rendered because the parity matrix never opens them.
 * - The page-number control keeps the display (button) branch. The edit branch
 *   and its Input are not vendored.
 */
import type * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  CircleMinusIcon,
  CirclePlusIcon,
  MoreHorizontalIcon,
  RotateClockwiseIcon,
  Search01Icon,
  SidebarLeftIcon,
} from "@hugeicons/core-free-icons";

import { cn } from "../../lib/cn";
import { Button } from "../ui/button";
import { SelectTrigger, SelectValue } from "../ui/select";
import { Separator } from "../ui/separator";

function TooltipTriggerWrapper({ children }: { children: React.ReactNode }) {
  return <span className="inline-flex">{children}</span>;
}

const ZOOM_OPTIONS = [0.1, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 2];

export interface PdfToolbarProps {
  activePage: number;
  controlsDisabled: boolean;
  numPages: number;
  sidebarOpen: boolean;
  showDownload: boolean;
  showRotateControls: boolean;
  showUpload: boolean;
  zoomLevel: number;
}

export function PdfToolbar({
  activePage,
  controlsDisabled,
  numPages,
  sidebarOpen,
  showDownload,
  showRotateControls,
  showUpload,
  zoomLevel,
}: PdfToolbarProps): React.ReactElement {
  return (
    <div className="flex min-h-12 flex-wrap items-center justify-between gap-2 border-b bg-background px-3 py-2">
      <div className="flex min-w-0 flex-wrap items-center gap-2">
        <TooltipTriggerWrapper>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Pages sidebar"
            aria-pressed={sidebarOpen}
            data-active={sidebarOpen ? "" : undefined}
            className={cn(
              sidebarOpen &&
                "bg-accent text-accent-foreground ring-1 ring-ring/40 ring-inset",
            )}
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
            aria-label={`Current page ${numPages ? activePage : 1}. Edit page number`}
            disabled={controlsDisabled || !numPages}
          >
            {numPages ? activePage : 1}
          </Button>
          <span>of {numPages || "–"}</span>
        </div>
      </div>
      <div className="flex min-w-0 flex-wrap items-center justify-end gap-1">
        {showRotateControls ? (
          <>
            <div className="flex flex-none items-center gap-1">
              <TooltipTriggerWrapper>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon-sm"
                  aria-label="Rotate counterclockwise"
                  disabled={controlsDisabled}
                >
                  <HugeiconsIcon icon={RotateClockwiseIcon} className="size-4" />
                </Button>
              </TooltipTriggerWrapper>
              <TooltipTriggerWrapper>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon-sm"
                  aria-label="Rotate clockwise"
                  disabled={controlsDisabled}
                >
                  <HugeiconsIcon
                    icon={RotateClockwiseIcon}
                    className="size-4 -scale-x-100"
                  />
                </Button>
              </TooltipTriggerWrapper>
            </div>
            <Separator orientation="vertical" className="mx-1 h-4 self-center" />
          </>
        ) : null}
        <div className="flex flex-none items-center gap-1">
          <TooltipTriggerWrapper>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Zoom out"
              disabled={controlsDisabled || zoomLevel <= ZOOM_OPTIONS[0]}
            >
              <HugeiconsIcon icon={CircleMinusIcon} className="size-4" />
            </Button>
          </TooltipTriggerWrapper>
          <SelectTrigger
            size="sm"
            className="w-[104px] min-w-[104px]"
            aria-label="Zoom level"
            disabled={controlsDisabled}
            icon="hugeicons"
          >
            <SelectValue placeholder="Zoom">
              {Math.round(zoomLevel * 100)}%
            </SelectValue>
          </SelectTrigger>
          <TooltipTriggerWrapper>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Zoom in"
              disabled={
                controlsDisabled ||
                zoomLevel >= ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1]
              }
            >
              <HugeiconsIcon icon={CirclePlusIcon} className="size-4" />
            </Button>
          </TooltipTriggerWrapper>
        </div>
        <Separator orientation="vertical" className="mx-1 h-4 self-center" />
        <TooltipTriggerWrapper>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Search text"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon icon={Search01Icon} className="size-4" />
          </Button>
        </TooltipTriggerWrapper>
        {showDownload || showUpload ? (
          <>
            <Separator orientation="vertical" className="mx-1 h-4 self-center" />
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Open PDF actions"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon icon={MoreHorizontalIcon} className="size-4" />
            </Button>
          </>
        ) : null}
      </div>
    </div>
  );
}
