/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/registry/new-york-v4/ui/spinner.tsx
 *   (apps/v4/components/ui/spinner.tsx re-exports the registry entry)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: Extend's `IconPlaceholder` supports Lucide; the harness uses the
 * Lucide `Loader2` glyph so the spinner matches `@arcships/vue-ui`'s
 * `OfficeIcon` family. The `animate-spin` class and role/aria contract are
 * unchanged.
 */
import type * as React from "react";

import { cn } from "../../lib/cn";

export function Spinner({
  className,
  ...props
}: React.ComponentProps<"svg">): React.ReactElement {
  return (
    <svg
      aria-label="Loading"
      className={cn("animate-spin", className)}
      fill="none"
      role="status"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      viewBox="0 0 24 24"
      {...props}
    >
      <path d="M21 12a9 9 0 1 1-6.219-8.56" />
    </svg>
  );
}
