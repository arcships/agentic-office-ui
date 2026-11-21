/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/ui/color-picker.tsx (ColorPicker, swatch trigger
 * branch)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the popover panel, the OKLCH/HSL/RGB/hex editors, the eyedropper
 * and `@/lib/color-picker-utils` are not vendored — the parity matrix only
 * captures the closed toolbar trigger. `ColorPickerTrigger` keeps the swatch
 * trigger's markup and class strings (ghost `icon-sm` button, `size-4` glyph and
 * the 16x4 colour bar at the bottom-right), with the colour as a plain prop and
 * the Hugeicons glyph data passed straight to the icon renderer.
 */
import type * as React from "react";
import { HugeiconsIcon, type IconSvgElement } from "@hugeicons/react";

import { cn } from "../../lib/cn";
import { Button } from "./button";

export interface ColorPickerTriggerProps {
  /** Current colour, as a CSS colour string. */
  color: string;
  disabled?: boolean;
  className?: string;
  /** Hugeicons glyph data rendered inside the button. */
  icon: IconSvgElement;
  /** Accessible name of the trigger. */
  label: string;
}

export function ColorPickerTrigger({
  color,
  disabled = false,
  className,
  icon,
  label,
}: ColorPickerTriggerProps): React.ReactElement {
  return (
    <Button
      aria-label={label}
      className={cn("relative shrink-0", className)}
      data-color-picker-trigger=""
      disabled={disabled}
      size="icon-sm"
      type="button"
      variant="ghost"
    >
      <HugeiconsIcon className="size-4" icon={icon} />
      <span
        aria-hidden="true"
        className="absolute right-1 bottom-1 h-1 w-4 rounded-full border border-background"
        style={{ backgroundColor: color }}
      />
    </Button>
  );
}
