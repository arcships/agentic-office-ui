/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/xlsx-viewer.tsx (WorkbookToolbar,
 * ToolbarTooltip, WorkbookSearchPopover trigger, WorkbookFileActionsMenu
 * trigger)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (#160): the workbook engine is not vendored, so the two
 * `@extend-ai/react-xlsx` hooks are replaced by props:
 *   - `useXlsxViewerZoom()` -> `zoomPercent` + `controlsDisabled`
 *     (`canZoomOut` / `canZoomIn` are the ZOOM_OPTIONS bounds, as upstream).
 *   - `useXlsxViewer()` (search/actions state) -> nothing; the search popover
 *     and the workbook actions dropdown are never opened by the parity matrix,
 *     so only their triggers are rendered.
 * Tooltip/popover/dropdown popups are dropped; the trigger wrappers keep the
 * upstream `<span class="inline-flex">` and button class strings, and the zoom
 * Select renders the vendored static `SelectTrigger`/`SelectValue` with the
 * upstream `w-[84px] min-w-[84px]`.
 */
import type * as React from "react";
import {
  CircleMinusIcon,
  CirclePlusIcon,
  MoreHorizontalIcon,
  Search01Icon,
} from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";

import { Button } from "../ui/button";
import { SelectTrigger, SelectValue } from "../ui/select";
import { Separator } from "../ui/separator";

/** Upstream `ZOOM_OPTIONS` (xlsx-viewer.tsx). */
export const XLSX_ZOOM_OPTIONS = [10, 25, 50, 75, 100, 125, 150, 175, 200, 400] as const;

function ToolbarTooltip({ children }: { children: React.ReactNode }) {
  return <span className="inline-flex">{children}</span>;
}

export interface XlsxToolbarProps {
  controlsDisabled: boolean;
  zoomPercent: number;
}

export function XlsxToolbar({
  controlsDisabled,
  zoomPercent,
}: XlsxToolbarProps): React.ReactElement {
  const canZoomOut = zoomPercent > XLSX_ZOOM_OPTIONS[0];
  const canZoomIn = zoomPercent < XLSX_ZOOM_OPTIONS[XLSX_ZOOM_OPTIONS.length - 1];

  return (
    <div
      className="flex min-h-12 flex-wrap items-center justify-end gap-2 border-b bg-background px-3 py-2"
      data-parity-xlsx-toolbar=""
    >
      <div className="ml-auto flex min-w-0 flex-wrap items-center justify-end gap-1">
        <div className="flex flex-none items-center gap-1">
          <ToolbarTooltip>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              disabled={controlsDisabled || !canZoomOut}
              aria-label="Zoom out"
            >
              <HugeiconsIcon icon={CircleMinusIcon} className="size-4" />
            </Button>
          </ToolbarTooltip>
          <SelectTrigger
            size="sm"
            className="w-[84px] min-w-[84px]"
            aria-label="Zoom level"
          >
            <SelectValue>{zoomPercent}%</SelectValue>
          </SelectTrigger>
          <ToolbarTooltip>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              disabled={controlsDisabled || !canZoomIn}
              aria-label="Zoom in"
            >
              <HugeiconsIcon icon={CirclePlusIcon} className="size-4" />
            </Button>
          </ToolbarTooltip>
        </div>
        <Separator orientation="vertical" className="mx-1 h-4 self-center" />
        <ToolbarTooltip>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Search workbook"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon icon={Search01Icon} className="size-4" />
          </Button>
        </ToolbarTooltip>
        <Separator orientation="vertical" className="mx-1 h-4 self-center" />
        <ToolbarTooltip>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Open workbook actions"
          >
            <HugeiconsIcon icon={MoreHorizontalIcon} className="size-4" />
          </Button>
        </ToolbarTooltip>
      </div>
    </div>
  );
}
