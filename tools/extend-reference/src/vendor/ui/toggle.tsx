/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/ui/toggle.tsx (toggleVariants,
 * toggleGroupVariants, ToggleGroup, ToggleGroupItem)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the Base UI `Toggle` / `ToggleGroup` primitives are replaced by
 * static `<button>` / `<div>` markup with the same class strings and the
 * `data-slot` / `data-spacing` / `data-pressed` attributes the class list keys
 * off. The parity matrix captures the resting toolbar row, so no press or
 * keyboard handling is vendored; `pressed` and `disabled` are props.
 */
import type * as React from "react";
import { cva, type VariantProps } from "class-variance-authority";

import { cn } from "../../lib/cn";
import { groupVariants } from "./group";

export const toggleVariants = cva(
  "relative inline-flex shrink-0 cursor-pointer items-center justify-center gap-2 rounded-lg border text-base font-medium whitespace-nowrap text-foreground transition-shadow outline-none select-none before:pointer-events-none before:absolute before:inset-0 before:rounded-[calc(var(--radius-lg)-1px)] hover:bg-accent focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-1 focus-visible:ring-offset-background disabled:pointer-events-none disabled:opacity-64 data-pressed:bg-input/64 data-pressed:text-accent-foreground sm:text-sm pointer-coarse:after:absolute pointer-coarse:after:size-full pointer-coarse:after:min-h-11 pointer-coarse:after:min-w-11 [&_svg]:pointer-events-none [&_svg]:-mx-0.5 [&_svg]:shrink-0 [&_svg:not([class*='opacity-'])]:opacity-80 [&_svg:not([class*='size-'])]:size-4.5 sm:[&_svg:not([class*='size-'])]:size-4",
  {
    defaultVariants: {
      size: "default",
      variant: "default",
    },
    variants: {
      size: {
        default: "h-9 min-w-9 px-[calc(--spacing(2)-1px)] sm:h-8 sm:min-w-8",
        lg: "h-10 min-w-10 px-[calc(--spacing(2.5)-1px)] sm:h-9 sm:min-w-9",
        sm: "h-8 min-w-8 px-[calc(--spacing(1.5)-1px)] sm:h-7 sm:min-w-7",
      },
      variant: {
        default: "border-transparent",
        outline:
          "border-input bg-background shadow-xs/5 not-dark:bg-clip-padding not-disabled:not-active:not-data-pressed:before:shadow-[0_1px_--theme(--color-black/4%)] dark:bg-input/32 dark:not-disabled:before:shadow-[0_-1px_--theme(--color-white/2%)] dark:not-disabled:not-active:not-data-pressed:before:shadow-[0_-1px_--theme(--color-white/6%)] dark:hover:bg-input/64 dark:data-pressed:bg-input [:disabled,:active,[data-pressed]]:shadow-none",
      },
    },
  },
);

export const toggleGroupVariants = cva("inline-flex w-fit", {
  defaultVariants: {
    orientation: "horizontal",
    spacing: "default",
  },
  variants: {
    orientation: {
      horizontal: "flex-row",
      vertical: "flex-col",
    },
    spacing: {
      default: "gap-1",
      none: "",
    },
  },
});

export function ToggleGroup({
  className,
  orientation = "horizontal",
  spacing,
  children,
}: {
  className?: string;
  orientation?: "horizontal" | "vertical";
  spacing?: VariantProps<typeof toggleGroupVariants>["spacing"];
  children: React.ReactNode;
}): React.ReactElement {
  const resolvedSpacing = spacing ?? "default";

  return (
    <div
      className={cn(
        toggleGroupVariants({ orientation, spacing: resolvedSpacing }),
        resolvedSpacing === "none" && groupVariants({ orientation }),
        "inline-flex",
        className,
      )}
      data-orientation={orientation}
      data-slot="toggle-group"
      data-spacing={resolvedSpacing}
      role="group"
    >
      {children}
    </div>
  );
}

export function ToggleGroupItem({
  className,
  variant,
  size,
  pressed = false,
  ...props
}: React.ComponentProps<"button"> &
  VariantProps<typeof toggleVariants> & {
    pressed?: boolean;
  }): React.ReactElement {
  return (
    <button
      className={cn(toggleVariants({ size, variant }), className)}
      data-pressed={pressed ? "" : undefined}
      data-slot="toggle-group-item"
      data-state={pressed ? "on" : "off"}
      type="button"
      {...props}
    />
  );
}
