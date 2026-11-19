/**
 * Adapted from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT).
 * Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
 * Source: apps/v4/lib/utils.ts
 * Licensed under the MIT License; see ../vendor/NOTICE.md.
 */
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

// Same helper as extend-hq/ui apps/v4/lib/utils.ts (siteConfig/absoluteUrl
// dropped: the harness never builds absolute URLs).
export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}
