/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Sources (class strings copied verbatim, no adaptation to the rendered
 * markup):
 *
 *   ViewerEmptyState   apps/v4/components/extend/pdf-editor-shared.tsx:2227
 *                      (`PdfEditorEmptyState`: glyph slot, title, description,
 *                      action). The glyph is a harness prop instead of an
 *                      upstream `PdfEditorGlyph` icon set so the parity cells
 *                      compare layout, not icon artwork.
 *   ViewerLoadingSurface apps/v4/components/extend/docx-viewer.tsx:302 and
 *                      xlsx-viewer.tsx:469 (`ViewerLoadingSurface`)
 *   WorkbookErrorCard  apps/v4/components/extend/xlsx-viewer.tsx:1675 (the
 *                      `loadError` fallback card; its upload action is the same
 *                      `Button variant="outline" size="sm"` this card renders as
 *                      "Retry", minus the upstream upload glyph)
 *   ViewerStatusFooter apps/v4/components/extend/file-system.tsx:2387 (the
 *                      `aria-live` footer), xlsx-viewer.tsx:1307 and
 *                      pdf-editor-panels.tsx:1951 (the `border-t` footer rows).
 *                      Extend has no standalone status bar; the composition
 *                      below keeps the upstream footer classes and reproduces
 *                      the page/status/file content `OfficeStatusBar` renders.
 */
import * as React from "react";

import { Button } from "../ui/button";
import { Spinner } from "../ui/spinner";

export function ViewerEmptyState({
  action,
  description,
  glyph,
  title,
}: {
  action?: React.ReactNode;
  description?: React.ReactNode;
  glyph?: React.ReactNode;
  title: React.ReactNode;
}) {
  return (
    <div
      className="flex flex-col items-center gap-2 px-4 py-10 text-center"
      data-viewer-empty-state=""
    >
      {glyph ? (
        <div className="grid size-10 place-items-center rounded-full bg-muted text-muted-foreground">
          {glyph}
        </div>
      ) : null}
      <div className="text-sm font-medium text-foreground">{title}</div>
      {description ? (
        <div className="max-w-56 text-xs text-muted-foreground">
          {description}
        </div>
      ) : null}
      {action ? <div className="mt-2">{action}</div> : null}
    </div>
  );
}

export function ViewerLoadingSurface({ showSpinner = true }: { showSpinner?: boolean }) {
  return (
    <div
      className="grid h-full min-h-52 place-items-center bg-transparent"
      data-viewer-loading-surface=""
    >
      {showSpinner ? <Spinner className="size-4" /> : null}
    </div>
  );
}

export function WorkbookErrorCard({
  message,
  showAction = true,
  title,
}: {
  message: React.ReactNode;
  showAction?: boolean;
  title: React.ReactNode;
}) {
  return (
    <div
      className="max-w-md rounded-lg border bg-background p-4 text-sm"
      data-viewer-error-card=""
    >
      <p className="font-medium">{title}</p>
      <p className="mt-1 text-muted-foreground">{message}</p>
      {showAction ? (
        <Button className="mt-4" size="sm" type="button" variant="outline">
          Try again
        </Button>
      ) : null}
    </div>
  );
}

export function ViewerStatusFooter({
  fileName,
  page,
  status,
}: {
  fileName: string;
  page: string;
  status: string;
}) {
  return (
    <footer
      aria-label="Document status"
      className="flex flex-wrap items-center justify-between gap-x-2 gap-y-1 border-t bg-muted/40 px-3 py-2 text-xs text-muted-foreground"
      data-status-bar-frame=""
    >
      <div className="flex min-w-0 grow flex-wrap items-center gap-x-2 gap-y-1">
        <span className="whitespace-nowrap text-foreground tabular-nums">
          {page}
        </span>
        <span className="inline-flex min-w-0 items-center gap-1 whitespace-nowrap">
          {status}
        </span>
      </div>
      <div className="flex min-w-0 flex-wrap items-center justify-end gap-x-2 gap-y-1">
        <span className="inline-flex max-w-full min-w-0 items-center gap-1">
          <span className="shrink-0 text-muted-foreground">File</span>
          <span className="overflow-hidden text-ellipsis whitespace-nowrap text-foreground">
            {fileName}
          </span>
        </span>
      </div>
    </footer>
  );
}
