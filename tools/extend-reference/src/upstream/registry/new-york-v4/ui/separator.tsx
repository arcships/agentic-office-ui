// Copied from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn.
import type React from "react"
import { Separator as SeparatorPrimitive } from "@base-ui/react/separator"

import { cn } from "@/lib/utils"

export function Separator({
  className,
  orientation = "horizontal",
  ...props
}: SeparatorPrimitive.Props): React.ReactElement {
  return (
    <SeparatorPrimitive
      className={cn(
        "shrink-0 bg-border data-[orientation=horizontal]:h-px data-[orientation=horizontal]:w-full data-[orientation=vertical]:w-px data-[orientation=vertical]:not-[[class^='h-']]:not-[[class*='_h-']]:self-stretch",
        className
      )}
      data-slot="separator"
      orientation={orientation}
      {...props}
    />
  )
}

export { SeparatorPrimitive }
