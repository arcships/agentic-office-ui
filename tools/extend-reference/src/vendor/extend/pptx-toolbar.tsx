/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/pptx-viewer.tsx (PptxToolbar,
 * ToolbarTooltip, PptxSlideNumberControl display branch, PptxFileActionsMenu
 * trigger)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptations (all drop interactive popups, none change resting-state layout):
 * - The hover-only tooltip primitive renders its trigger wrapper unchanged
 *   (`<span class="inline-flex">`) and no tooltip content.
 * - The zoom Select becomes the static SelectTrigger/SelectValue pair from the
 *   vendored registry select, keeping `size="sm"`, `w-[104px] min-w-[104px]`.
 * - The file actions DropdownMenu trigger becomes a plain ghost icon button
 *   with the upstream classes and Hugeicons `MoreHorizontalIcon`; the menu
 *   popup is not rendered because the parity matrix never opens it.
 * - The slide-number display (button) branch is vendored. The inline edit
 *   branch and its Input are not.
 */
import type * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  ArrowLeft01Icon,
  ArrowRight01Icon,
  CircleMinusIcon,
  CirclePlusIcon,
  MoreHorizontalIcon,
  SidebarLeftIcon,
} from "@hugeicons/core-free-icons";

import { Button } from "../ui/button";
import { SelectTrigger, SelectValue } from "../ui/select";
import { Separator } from "../ui/separator";

/** Upstream `ToolbarTooltip` trigger wrapper; the popup itself is not vendored. */
function TooltipTriggerWrapper({ children }: { children: React.ReactNode }) {
  return <span className="inline-flex">{children}</span>;
}

export interface PptxToolbarProps {
  activeSlideIndex: number;
  controlsDisabled: boolean;
  isPreparingDownload?: boolean;
  showDownloadButton: boolean;
  showUploadButton: boolean;
  slideCount: number;
  resolvedZoom: number;
}

/** Upstream `PptxSlideNumberControl` display branch. */
function PptxSlideNumberControl({
  activeSlideIndex,
  controlsDisabled,
  slideCount,
}: {
  activeSlideIndex: number;
  controlsDisabled: boolean;
  slideCount: number;
}) {
  const displaySlide = slideCount ? activeSlideIndex + 1 : 1;

  return (
    <div className="flex items-center text-sm whitespace-nowrap text-primary">
      <span>Slide</span>
      <Button
        type="button"
        variant="ghost"
        size="sm"
        className="font-normal"
        aria-label={`Current slide ${displaySlide}. Edit slide number`}
        disabled={controlsDisabled || !slideCount}
      >
        {displaySlide}
      </Button>
      <span>of {slideCount || "-"}</span>
    </div>
  );
}

export function PptxToolbar({
  activeSlideIndex,
  controlsDisabled,
  isPreparingDownload = false,
  showDownloadButton,
  showUploadButton,
  slideCount,
  resolvedZoom,
}: PptxToolbarProps): React.ReactElement {
  const canGoPrevious = !controlsDisabled && activeSlideIndex > 0;
  const canGoNext =
    !controlsDisabled && slideCount > 0 && activeSlideIndex < slideCount - 1;
  const canZoomOut = !controlsDisabled && resolvedZoom > 25;
  const canZoomIn = !controlsDisabled && resolvedZoom < 400;

  return (
    <div
      className="flex min-h-12 flex-wrap items-center justify-between gap-2 border-b bg-background px-3 py-2"
      data-pptx-toolbar=""
      role="toolbar"
      aria-label="Presentation toolbar"
    >
      <div className="flex min-w-0 flex-wrap items-center gap-1">
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
        <Separator orientation="vertical" className="mx-1 h-4 self-center" />
        <TooltipTriggerWrapper>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Previous slide"
            disabled={!canGoPrevious}
          >
            <HugeiconsIcon icon={ArrowLeft01Icon} className="size-4" />
          </Button>
        </TooltipTriggerWrapper>
        <PptxSlideNumberControl
          activeSlideIndex={activeSlideIndex}
          controlsDisabled={controlsDisabled}
          slideCount={slideCount}
        />
        <TooltipTriggerWrapper>
          <Button
            type="button"
            variant="ghost"
            size="icon-sm"
            aria-label="Next slide"
            disabled={!canGoNext}
          >
            <HugeiconsIcon icon={ArrowRight01Icon} className="size-4" />
          </Button>
        </TooltipTriggerWrapper>
      </div>
      <div className="ml-auto flex min-w-0 flex-wrap items-center justify-end gap-1">
        <div className="flex flex-none items-center gap-1">
          <TooltipTriggerWrapper>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Zoom out"
              disabled={!canZoomOut}
            >
              <HugeiconsIcon icon={CircleMinusIcon} className="size-4" />
            </Button>
          </TooltipTriggerWrapper>
          <SelectTrigger
            size="sm"
            className="w-[104px] min-w-[104px]"
            aria-label="Zoom level"
            disabled={controlsDisabled}
          >
            <SelectValue>{Math.round(resolvedZoom)}%</SelectValue>
          </SelectTrigger>
          <TooltipTriggerWrapper>
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Zoom in"
              disabled={!canZoomIn}
            >
              <HugeiconsIcon icon={CirclePlusIcon} className="size-4" />
            </Button>
          </TooltipTriggerWrapper>
        </div>
        {showDownloadButton || showUploadButton ? (
          <>
            <Separator
              orientation="vertical"
              className="mx-1 h-4 self-center"
            />
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              aria-label="Open PowerPoint actions"
              data-preparing-download={isPreparingDownload ? "" : undefined}
            >
              <HugeiconsIcon icon={MoreHorizontalIcon} className="size-4" />
            </Button>
          </>
        ) : null}
      </div>
    </div>
  );
}
