/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/registry/new-york-v4/ui/dialog.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: the Base UI Dialog primitives are replaced by static `<div>`s.
 * The harness only captures an already-open dialog, so focus management and
 * portal behaviour are not exercised. Every rendered class string (backdrop,
 * viewport, popup, header, panel, footer, title, description, close button) is
 * copied verbatim from the pinned source.
 */
import type * as React from "react";

import { cn } from "../../lib/cn";
import { Button } from "./button";

export function DialogBackdrop(): React.ReactElement {
  return (
    <div
      className="fixed inset-0 z-50 bg-black/32 backdrop-blur-sm transition-all duration-200"
      data-slot="dialog-backdrop"
    />
  );
}

export function DialogViewport({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn(
        "fixed inset-0 z-50 flex items-center justify-center p-4",
        className,
      )}
      data-slot="dialog-viewport"
    >
      {children}
    </div>
  );
}

export function DialogPopup({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn(
        "relative row-start-2 flex max-h-full min-h-0 w-full max-w-lg min-w-0 origin-center flex-col rounded-2xl border bg-popover text-popover-foreground shadow-lg/5 transition-[scale,opacity,translate] duration-200 ease-in-out will-change-transform outline-none not-dark:bg-clip-padding before:pointer-events-none before:absolute before:inset-0 before:rounded-[calc(var(--radius-2xl)-1px)] before:shadow-[0_1px_--theme(--color-black/4%)] dark:before:shadow-[0_-1px_--theme(--color-white/6%)] max-sm:max-w-none max-sm:origin-bottom max-sm:rounded-none max-sm:border-x-0 max-sm:border-t max-sm:border-b-0 max-sm:before:hidden max-sm:before:rounded-none",
        className,
      )}
      data-slot="dialog-popup"
    >
      {children}
    </div>
  );
}

export function DialogHeader({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn("flex flex-col gap-2 p-6 max-sm:pb-4", className)}
      data-slot="dialog-header"
    >
      {children}
    </div>
  );
}

export function DialogPanel({
  className,
  children,
}: {
  className?: string;
  children?: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn("p-6", className)}
      data-slot="dialog-panel"
    >
      {children}
    </div>
  );
}

export function DialogFooter({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn(
        "flex flex-col-reverse gap-2 border-t bg-muted/72 px-6 py-4 sm:flex-row sm:justify-end sm:rounded-b-[calc(var(--radius-2xl)-1px)]",
        className,
      )}
      data-slot="dialog-footer"
    >
      {children}
    </div>
  );
}

export function DialogTitle({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <div
      className={cn("font-heading text-xl leading-none font-semibold", className)}
      data-slot="dialog-title"
    >
      {children}
    </div>
  );
}

export function DialogDescription({
  className,
  children,
}: {
  className?: string;
  children: React.ReactNode;
}): React.ReactElement {
  return (
    <p
      className={cn("text-sm text-muted-foreground", className)}
      data-slot="dialog-description"
    >
      {children}
    </p>
  );
}

export function DialogClose({
  children,
  className,
}: {
  children: React.ReactNode;
  className?: string;
}): React.ReactElement {
  return (
    <div className={cn("absolute end-2 top-2", className)} data-slot="dialog-close">
      {children}
    </div>
  );
}

export { Button };
