import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

// Same helper as extend-hq/ui apps/v4/lib/utils.ts (siteConfig/absoluteUrl
// dropped: the harness never builds absolute URLs).
export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}
