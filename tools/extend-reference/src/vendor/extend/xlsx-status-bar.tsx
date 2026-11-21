/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/file-system.tsx (the `aria-live` status
 * strip at the bottom of the file-system block)
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (#160): the pinned Extend XLSX viewer/editor has no status bar
 * (see docs/design/xlsx-office-chrome-migration.md §4.4), so this surface uses
 * Extend's own status-strip markup — the 28px `aria-live` footer that ships
 * with the file-system block — as the reference for the XLSX status strip.
 * The segments are passed in so both sides render the same fixture strings;
 * only geometry, typography and colour are compared.
 */
import type * as React from "react";

export function XlsxStatusStrip({ segments }: { segments: readonly string[] }): React.ReactElement {
  return (
    <div
      aria-live="polite"
      className="flex h-7 shrink-0 items-center justify-center gap-1 border-t bg-muted/40 px-3 text-xs text-muted-foreground"
      data-parity-xlsx-status-bar=""
    >
      {segments.map((segment, index) => (
        <span key={`${index}-${segment}`}>
          {index === 0 ? segment : `· ${segment}`}
        </span>
      ))}
    </div>
  );
}
