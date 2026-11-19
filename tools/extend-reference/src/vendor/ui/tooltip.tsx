/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/registry/new-york-v4/ui/tooltip.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the Base UI Tooltip primitive is replaced by a statically
 * rendered popup. The harness only needs the open state, so the popup class
 * strings (positioner-independent) and the viewport padding are copied
 * verbatim; positioning classes belong to the harness surface.
 */
import type * as React from "react";

import { cn } from "../../lib/cn";

export function TooltipPopup({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn(
        "relative flex h-(--popup-height,auto) w-(--popup-width,auto) origin-(--transform-origin) rounded-md border bg-popover text-xs text-balance text-popover-foreground shadow-md/5 transition-[width,height,scale,opacity] not-dark:bg-clip-padding before:pointer-events-none before:absolute before:inset-0 before:rounded-[calc(var(--radius-md)-1px)] before:shadow-[0_1px_--theme(--color-black/4%)] dark:before:shadow-[0_-1px_--theme(--color-white/6%)]",
        className,
      )}
      data-parity-tooltip-content
      data-slot="tooltip-popup"
    >
      <div
        className="relative size-full overflow-clip px-(--viewport-inline-padding) py-1 [--viewport-inline-padding:--spacing(2)]"
        data-slot="tooltip-viewport"
      >
        {children}
      </div>
    </div>
  );
}
