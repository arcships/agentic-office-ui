/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/ui/group.tsx (groupVariants, Group,
 * GroupText, GroupSeparator)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the Base UI `useRender` / `mergeProps` render-prop plumbing is
 * replaced by plain elements; every class string is copied verbatim. Base UI
 * `Separator` is replaced by the vendored static separator, which carries the
 * same `data-slot="separator"` hook the group class list keys off.
 */
import type * as React from "react";
import { cva, type VariantProps } from "class-variance-authority";

import { cn } from "../../lib/cn";
import { Separator } from "./separator";

export const groupVariants = cva(
  "flex w-fit *:focus-visible:z-1 *:has-focus-visible:z-1 has-[>[data-slot=group]]:gap-2 dark:*:[[data-slot=separator]:has(~button:hover):not(:has(~[data-slot=separator]~[data-slot]:hover)),[data-slot=separator]:has(~[data-slot][data-pressed]):not(:has(~[data-slot=separator]~[data-slot][data-pressed]))]:before:bg-input/64 dark:*:[button:hover~[data-slot=separator]:not([data-slot]:hover~[data-slot=separator]~[data-slot=separator]),[data-slot][data-pressed]~[data-slot=separator]:not([data-slot][data-pressed]~[data-slot=separator]~[data-slot=separator])]:before:bg-input/64",
  {
    defaultVariants: {
      orientation: "horizontal",
    },
    variants: {
      orientation: {
        horizontal:
          "*:data-slot:has-[~[data-slot]]:rounded-e-none *:data-slot:has-[~[data-slot]]:border-e-0 *:data-slot:has-[~[data-slot]]:before:rounded-e-none *:data-slot:not-data-[slot=separator]:has-[~[data-slot]]:before:-end-[0.5px] *:pointer-coarse:after:min-w-auto *:[[data-slot]~[data-slot]]:rounded-s-none *:[[data-slot]~[data-slot]]:border-s-0 *:[[data-slot]~[data-slot]]:before:rounded-s-none *:[[data-slot]~[data-slot]:not([data-slot=separator])]:before:-start-[0.5px]",
        vertical:
          "flex-col *:data-slot:has-[~[data-slot]]:rounded-b-none *:data-slot:has-[~[data-slot]]:border-b-0 *:data-slot:has-[~[data-slot]]:before:rounded-b-none *:data-slot:not-data-[slot=separator]:has-[~[data-slot]]:before:-bottom-[0.5px] *:data-slot:not-data-[slot=separator]:has-[~[data-slot]]:before:hidden dark:*:first:before:block dark:*:last:before:hidden *:pointer-coarse:after:min-h-auto *:[[data-slot]~[data-slot]]:rounded-t-none *:[[data-slot]~[data-slot]]:border-t-0 *:[[data-slot]~[data-slot]]:before:rounded-t-none *:[[data-slot]~[data-slot]:not([data-slot=separator])]:before:-top-[0.5px]",
      },
    },
  },
);

export function Group({
  className,
  orientation,
  children,
  ...props
}: {
  className?: string;
  orientation?: VariantProps<typeof groupVariants>["orientation"];
  children: React.ReactNode;
} & React.ComponentProps<"div">): React.ReactElement {
  return (
    <div
      className={cn(groupVariants({ orientation }), className)}
      data-orientation={orientation}
      data-slot="group"
      role="group"
      {...props}
    >
      {children}
    </div>
  );
}

export function GroupText({
  className,
  children,
  ...props
}: React.ComponentProps<"div">): React.ReactElement {
  return (
    <div
      className={cn(
        "relative inline-flex items-center gap-2 rounded-lg border border-input bg-muted px-[calc(--spacing(3)-1px)] text-base whitespace-nowrap text-muted-foreground shadow-xs/5 transition-shadow outline-none not-dark:bg-clip-padding before:pointer-events-none before:absolute before:inset-0 before:rounded-[calc(var(--radius-lg)-1px)] before:shadow-[0_1px_--theme(--color-black/6%)] sm:text-sm dark:bg-input/64 dark:before:shadow-[0_-1px_--theme(--color-white/6%)] [&_svg]:-mx-0.5 [&_svg]:shrink-0 [&_svg:not([class*='size-'])]:size-4.5 sm:[&_svg:not([class*='size-'])]:size-4",
        className,
      )}
      data-slot="group-text"
      {...props}
    >
      {children}
    </div>
  );
}

export function GroupSeparator({
  className,
  orientation = "vertical",
  ...props
}: {
  className?: string;
} & React.ComponentProps<typeof Separator>): React.ReactElement {
  return (
    <Separator
      className={cn(
        "pointer-events-none relative z-2 bg-input before:absolute before:inset-0 has-[+[data-slot=input-control]:focus-within,+[data-slot=input-group]:focus-within,+[data-slot=select-trigger]:focus-visible+*,+[data-slot=number-field]:focus-within]:translate-x-px has-[+[data-slot=input-control]:focus-within,+[data-slot=input-group]:focus-within,+[data-slot=select-trigger]:focus-visible+*,+[data-slot=number-field]:focus-within]:bg-ring dark:before:bg-input/32 [[data-slot=input-control]:focus-within+&,[data-slot=input-group]:focus-within+&,[data-slot=select-trigger]:focus-visible+*+&,[data-slot=number-field]:focus-within+&,[data-slot=number-field]:focus-within+input+&]:bg-ring [[data-slot=input-control]:focus-within+&,[data-slot=input-group]:focus-within+&,[data-slot=select-trigger]:focus-visible+*+&,[data-slot=number-field]:focus-within+&,[data-slot=number-field]:focus-within+input+&]:-translate-x-px",
        className,
      )}
      orientation={orientation}
      {...props}
    />
  );
}

export {
  Group as ButtonGroup,
  GroupSeparator as ButtonGroupSeparator,
  GroupText as ButtonGroupText,
};
