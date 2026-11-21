// Copied from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn.
import type React from "react"

import { cn } from "@/lib/utils"
import { IconPlaceholder } from "@/components/icon-placeholder"

export function Spinner({
  className,
  ...props
}: React.ComponentProps<"svg">): React.ReactElement {
  return (
    <IconPlaceholder
      lucide="Loader2"
      tabler="IconLoader"
      hugeicons="Loading03Icon"
      phosphor="SpinnerIcon"
      remixicon="RiLoaderLine"
      aria-label="Loading"
      className={cn("animate-spin", className)}
      role="status"
      {...props}
    />
  )
}
