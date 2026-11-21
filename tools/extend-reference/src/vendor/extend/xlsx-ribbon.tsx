/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/xlsx-editor.tsx (`EditorToolbar` format
 * row, ~L1840-L2270: undo/redo, font family, font size, B/I/U/S toggle group,
 * text/fill colour pickers, horizontal/vertical alignment toggle groups, wrap
 * text, number format, cell borders, merge cells)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (#170): the workbook engine is not vendored, so the private
 * `useXlsxViewer*` hooks become props and the row renders one *resting* state:
 *
 * - `useXlsxViewer()` (`sheets`, `activeSheet`, …) -> `canUndo` / `canRedo`.
 * - `useXlsxViewerSelection()` + `useXlsxViewerEditing()` -> `controlsDisabled`
 *   (upstream `!canStyleSelection`) and the active-cell style props
 *   (`fontFamily`, `fontSize`, `fontWeight`, `textColor`, `fillColor`,
 *   `horizontalAlignment`, `verticalAlignment`, `numberFormat`,
 *   `borderColor`, `canMerge`, `isMerged`).
 * - Tooltip/popover/dropdown *popups* are not rendered (the parity matrix never
 *   opens them); each trigger keeps the upstream wrapper (`ToolbarTooltip` ->
 *   `<span class="inline-flex">`) and its class string.
 * - Select dropdowns become the static `SelectTrigger`/`SelectValue` pairs from
 *   the vendored registry select (`w-[132px]`, `w-[78px]`,
 *   `XLSX_EDITOR_SELECT_CHROME_CLASS`).
 * - ColorPicker/ColorPickerPanel keep the swatch trigger only.
 * - Not vendored because they have no counterpart in this repository's ribbon:
 *   pivot insert/fields (product gap), add/remove sheet + active-sheet select
 *   (covered by the `xlsx-sheet-tabs` surface), zoom cluster (covered by the
 *   `xlsx-toolbar` surface), the formula row (separate surface), and the
 *   toolbar's theme/upload controls (the editor surface shell, not the ribbon).
 * - Also omitted from this subset because the Vue ribbon has no such command:
 *   "No fill" and the "More alignment" dropdown.
 */
import type * as React from "react";
import {
  AlignBoxBottomCenterIcon,
  AlignBoxMiddleCenterIcon,
  AlignBoxTopCenterIcon,
  BorderAll01Icon,
  Calendar03Icon,
  ChevronDownIcon,
  DollarCircleIcon,
  PathfinderMergeIcon,
  PaintBucketIcon,
  PercentIcon,
  Redo02Icon,
  TextAlignCenterIcon,
  TextAlignLeft01Icon,
  TextAlignRight01Icon,
  TextBoldIcon,
  TextColorIcon,
  TextItalicIcon,
  TextNumberSignIcon,
  TextStrikethroughIcon,
  TextUnderlineIcon,
  TextWrapIcon,
  Undo02Icon,
} from "@hugeicons/core-free-icons";
import { HugeiconsIcon, type IconSvgElement } from "@hugeicons/react";

import { ColorPickerTrigger } from "../ui/color-picker";
import { Button } from "../ui/button";
import { SelectTrigger, SelectValue } from "../ui/select";
import { Separator } from "../ui/separator";
import { ToggleGroup, ToggleGroupItem } from "../ui/toggle";

/** Upstream `XLSX_EDITOR_SELECT_CHROME_CLASS`. */
const XLSX_EDITOR_SELECT_CHROME_CLASS =
  "shadow-none before:shadow-none not-data-disabled:not-focus-visible:not-aria-invalid:not-data-pressed:before:shadow-none dark:not-data-disabled:not-focus-visible:not-aria-invalid:not-data-pressed:before:shadow-none";

type HorizontalAlignment = "left" | "center" | "right";
type VerticalAlignment = "top" | "center" | "bottom";
type NumberFormatId = "general" | "number" | "currency" | "percent" | "date";

/** Upstream `NUMBER_FORMAT_OPTIONS` glyph + label pairs. */
const NUMBER_FORMATS: Record<
  NumberFormatId,
  { icon: IconSvgElement; label: string }
> = {
  currency: { icon: DollarCircleIcon, label: "Currency" },
  date: { icon: Calendar03Icon, label: "Date" },
  general: { icon: TextNumberSignIcon, label: "General" },
  number: { icon: TextNumberSignIcon, label: "Number" },
  percent: { icon: PercentIcon, label: "Percent" },
};

function ToolbarTooltip({ children }: { children: React.ReactNode }) {
  return <span className="inline-flex">{children}</span>;
}

/** Upstream `ToolbarSeparator`: `mx-1 h-4 self-center`. */
function ToolbarSeparator() {
  return <Separator className="mx-1 h-4 self-center" orientation="vertical" />;
}

export interface XlsxRibbonProps {
  borderColor: string;
  canMerge: boolean;
  canRedo: boolean;
  canUndo: boolean;
  /** Upstream `!canStyleSelection`: every style control is disabled. */
  controlsDisabled: boolean;
  fillColor: string;
  fontFamily: string;
  fontSize: number;
  fontWeight: {
    bold: boolean;
    italic: boolean;
    strikethrough: boolean;
    underline: boolean;
  };
  horizontalAlignment: HorizontalAlignment;
  isMerged: boolean;
  numberFormat: NumberFormatId;
  textColor: string;
  verticalAlignment: VerticalAlignment;
}

export function XlsxRibbon({
  borderColor,
  canMerge,
  canRedo,
  canUndo,
  controlsDisabled,
  fillColor,
  fontFamily,
  fontSize,
  fontWeight,
  horizontalAlignment,
  isMerged,
  numberFormat,
  textColor,
  verticalAlignment,
}: XlsxRibbonProps): React.ReactElement {
  const numberFormatOption = NUMBER_FORMATS[numberFormat];
  const fontToggleValues = [
    ...(fontWeight.bold ? ["bold"] : []),
    ...(fontWeight.italic ? ["italic"] : []),
    ...(fontWeight.underline ? ["underline"] : []),
    ...(fontWeight.strikethrough ? ["strikethrough"] : []),
  ];

  return (
    <div
      className="flex min-h-12 flex-wrap items-center gap-2 border-b bg-background px-3 py-2"
      data-parity-xlsx-ribbon=""
    >
      <div className="flex shrink-0 items-center gap-1">
        <ToolbarTooltip>
          <Button
            aria-label="Undo"
            disabled={!canUndo}
            size="icon-sm"
            type="button"
            variant="ghost"
          >
            <HugeiconsIcon className="size-4" icon={Undo02Icon} />
          </Button>
        </ToolbarTooltip>
        <ToolbarTooltip>
          <Button
            aria-label="Redo"
            disabled={!canRedo}
            size="icon-sm"
            type="button"
            variant="ghost"
          >
            <HugeiconsIcon className="size-4" icon={Redo02Icon} />
          </Button>
        </ToolbarTooltip>
      </div>
      <ToolbarSeparator />
      <div className="flex shrink-0 items-center gap-1">
        <SelectTrigger
          aria-label="Font family"
          className={`w-[132px] min-w-[132px] ${XLSX_EDITOR_SELECT_CHROME_CLASS}`}
          disabled={controlsDisabled}
          size="sm"
        >
          <SelectValue placeholder="Font">{fontFamily}</SelectValue>
        </SelectTrigger>
        <SelectTrigger
          aria-label="Font size"
          className={`w-[78px] min-w-[78px] ${XLSX_EDITOR_SELECT_CHROME_CLASS}`}
          disabled={controlsDisabled}
          size="sm"
        >
          <SelectValue placeholder="Size">{`${fontSize} pt`}</SelectValue>
        </SelectTrigger>
      </div>
      <ToolbarSeparator />
      <ToggleGroup
        className="shrink-0"
        spacing="default"
        value={fontToggleValues}
      >
        <ToolbarTooltip>
          <ToggleGroupItem
            aria-label="Bold"
            pressed={fontWeight.bold}
            size="sm"
            value="bold"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon className="size-4" icon={TextBoldIcon} />
          </ToggleGroupItem>
        </ToolbarTooltip>
        <ToolbarTooltip>
          <ToggleGroupItem
            aria-label="Italic"
            pressed={fontWeight.italic}
            size="sm"
            value="italic"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon className="size-4" icon={TextItalicIcon} />
          </ToggleGroupItem>
        </ToolbarTooltip>
        <ToolbarTooltip>
          <ToggleGroupItem
            aria-label="Underline"
            pressed={fontWeight.underline}
            size="sm"
            value="underline"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon className="size-4" icon={TextUnderlineIcon} />
          </ToggleGroupItem>
        </ToolbarTooltip>
        <ToolbarTooltip>
          <ToggleGroupItem
            aria-label="Strikethrough"
            pressed={fontWeight.strikethrough}
            size="sm"
            value="strikethrough"
            disabled={controlsDisabled}
          >
            <HugeiconsIcon className="size-4" icon={TextStrikethroughIcon} />
          </ToggleGroupItem>
        </ToolbarTooltip>
      </ToggleGroup>
      <ToolbarSeparator />
      <div className="flex shrink-0 items-center gap-1">
        <ColorPickerTrigger
          color={textColor}
          disabled={controlsDisabled}
          icon={TextColorIcon}
          label="Text color"
        />
        <ColorPickerTrigger
          color={fillColor}
          disabled={controlsDisabled}
          icon={PaintBucketIcon}
          label="Fill color"
        />
      </div>
      <ToolbarSeparator />
      <div className="flex shrink-0 items-center gap-1">
        <ToggleGroup
          className="shrink-0"
          spacing="default"
          value={horizontalAlignment}
        >
          <ToolbarTooltip>
            <ToggleGroupItem
              aria-label="Align left"
              pressed={horizontalAlignment === "left"}
              size="sm"
              value="left"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon className="size-4" icon={TextAlignLeft01Icon} />
            </ToggleGroupItem>
          </ToolbarTooltip>
          <ToolbarTooltip>
            <ToggleGroupItem
              aria-label="Align center"
              pressed={horizontalAlignment === "center"}
              size="sm"
              value="center"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon
                className="size-4"
                icon={TextAlignCenterIcon}
              />
            </ToggleGroupItem>
          </ToolbarTooltip>
          <ToolbarTooltip>
            <ToggleGroupItem
              aria-label="Align right"
              pressed={horizontalAlignment === "right"}
              size="sm"
              value="right"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon className="size-4" icon={TextAlignRight01Icon} />
            </ToggleGroupItem>
          </ToolbarTooltip>
        </ToggleGroup>
        <ToggleGroup
          className="shrink-0"
          spacing="default"
          value={verticalAlignment}
        >
          <ToolbarTooltip>
            <ToggleGroupItem
              aria-label="Align top"
              pressed={verticalAlignment === "top"}
              size="sm"
              value="top"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon className="size-4" icon={AlignBoxTopCenterIcon} />
            </ToggleGroupItem>
          </ToolbarTooltip>
          <ToolbarTooltip>
            <ToggleGroupItem
              aria-label="Align middle"
              pressed={verticalAlignment === "center"}
              size="sm"
              value="center"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon
                className="size-4"
                icon={AlignBoxMiddleCenterIcon}
              />
            </ToggleGroupItem>
          </ToolbarTooltip>
          <ToolbarTooltip>
            <ToggleGroupItem
              aria-label="Align bottom"
              pressed={verticalAlignment === "bottom"}
              size="sm"
              value="bottom"
              disabled={controlsDisabled}
            >
              <HugeiconsIcon
                className="size-4"
                icon={AlignBoxBottomCenterIcon}
              />
            </ToggleGroupItem>
          </ToolbarTooltip>
        </ToggleGroup>
        <ToolbarTooltip>
          <Button
            aria-label="Wrap text"
            disabled={controlsDisabled}
            size="icon-sm"
            type="button"
            variant="ghost"
          >
            <HugeiconsIcon className="size-4" icon={TextWrapIcon} />
          </Button>
        </ToolbarTooltip>
      </div>
      <ToolbarSeparator />
      <div className="flex shrink-0 items-center gap-1">
        <SelectTrigger
          aria-label="Number format"
          className={`w-[132px] min-w-[132px] ${XLSX_EDITOR_SELECT_CHROME_CLASS}`}
          disabled={controlsDisabled}
          size="sm"
        >
          <HugeiconsIcon className="size-4" icon={numberFormatOption.icon} />
          <SelectValue placeholder="General">
            {numberFormatOption.label}
          </SelectValue>
        </SelectTrigger>
        <Button
          aria-label="Cell borders"
          className="relative shrink-0"
          disabled={controlsDisabled}
          size="icon-sm"
          type="button"
          variant="ghost"
        >
          <HugeiconsIcon className="size-4" icon={BorderAll01Icon} />
          <span
            aria-hidden="true"
            className="absolute right-1 bottom-1 h-1 w-4 rounded-full border border-background"
            style={{ backgroundColor: borderColor }}
          />
        </Button>
      </div>
      <ToolbarSeparator />
      <div className="inline-flex shrink-0 overflow-hidden rounded-lg border border-transparent">
        <ToolbarTooltip>
          <Button
            aria-label="Merge & Center"
            className="h-7 rounded-r-none border-0 px-2.5 before:rounded-r-none"
            data-pressed={isMerged ? "" : undefined}
            disabled={controlsDisabled || !canMerge}
            size="sm"
            type="button"
            variant={isMerged ? "secondary" : "ghost"}
          >
            <HugeiconsIcon className="size-4" icon={PathfinderMergeIcon} />
            <span className="hidden text-xs font-medium lg:inline">
              Merge &amp; Center
            </span>
          </Button>
        </ToolbarTooltip>
        <ToolbarTooltip>
          <Button
            aria-label="Merge cells options"
            className="h-7 w-6 rounded-l-none border-0 px-0 before:rounded-l-none"
            data-pressed={isMerged ? "" : undefined}
            disabled={controlsDisabled}
            size="icon-sm"
            type="button"
            variant={isMerged ? "secondary" : "ghost"}
          >
            <HugeiconsIcon
              className="size-3.5 opacity-100"
              icon={ChevronDownIcon}
            />
          </Button>
        </ToolbarTooltip>
      </div>
    </div>
  );
}
