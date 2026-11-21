/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/lib/registry-icon-props.ts
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 */
import type * as React from "react";

export type RegistryIconProps = Omit<
  React.ComponentProps<"svg">,
  "children" | "strokeWidth"
> & { strokeWidth?: number };
