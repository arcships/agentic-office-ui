/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/registry/new-york-v4/ui/spinner.tsx
 *   (apps/v4/components/ui/spinner.tsx re-exports the registry entry)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation: Extend's `IconPlaceholder` resolves Hugeicons by default; the
 * resolved `Loading03Icon` is imported directly so the harness does not need
 * the registry icon-switching machinery.
 */
import type * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import { Loading03Icon } from "@hugeicons/core-free-icons";

import { cn } from "../../lib/cn";

export function Spinner({
  className,
  ...props
}: React.ComponentProps<"svg">): React.ReactElement {
  return (
    <HugeiconsIcon
      icon={Loading03Icon}
      aria-label="Loading"
      className={cn("animate-spin", className)}
      role="status"
      {...(props as Record<string, unknown>)}
    />
  );
}
