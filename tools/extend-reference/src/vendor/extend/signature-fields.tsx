/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/registry/new-york-v4/blocks/e-signature.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (parity harness only):
 * - Only the right-hand `SignatureFieldsPanel` is vendored. The PDF viewer and
 *   the `signature_pad` dialog are not reproducible as static markup; the Vue
 *   subject covers them with interaction tests instead.
 * - The Base UI ScrollArea is replaced by a plain scroll viewport.
 * - `showPreview` and the loading / error states are harness extensions with
 *   the same class contract as the Vue subject.
 */
import * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  Download01Icon,
  FilePenIcon,
  Pen01Icon,
} from "@hugeicons/core-free-icons";

import { cn } from "../../lib/cn";
import { Button } from "../ui/button";

export type SignatureField = {
  id: string;
  label: string;
  page: number;
  bbox: { x: number; y: number; width: number; height: number };
  imageDataUrl?: string;
};

export function SignatureFieldsPanel({
  fields,
  canExport,
  isDownloading,
  showPreview = false,
  loading = false,
  error,
  className,
}: {
  fields: SignatureField[];
  canExport: boolean;
  isDownloading: boolean;
  showPreview?: boolean;
  loading?: boolean;
  error?: string;
  className?: string;
}) {
  return (
    <aside
      className={cn("flex min-h-0 flex-col bg-background", className)}
      data-signature-fields
    >
      <div className="min-h-0 flex-1 overflow-y-auto">
        <div className="space-y-4 p-4">
          <div className="space-y-1">
            <h3 className="text-sm font-medium">Signature fields</h3>
            <p className="text-xs text-muted-foreground">
              Review fields, collect signatures, and export a signed PDF.
            </p>
          </div>
          {error ? (
            <div
              role="alert"
              className="rounded-lg border border-destructive/45 bg-destructive/8 p-3 text-sm text-destructive"
            >
              {error}
            </div>
          ) : loading ? (
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              Loading signature fields…
            </div>
          ) : fields.length === 0 ? (
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              Unsigned
            </div>
          ) : (
            fields.map((field) => (
              <div
                key={field.id}
                className="rounded-lg border bg-background p-3"
                data-signature-field-id={field.id}
              >
                <div className="flex items-start gap-3">
                  <div className="grid size-9 shrink-0 place-items-center rounded-md bg-blue-500/10 text-blue-600 dark:text-blue-300">
                    <HugeiconsIcon icon={FilePenIcon} className="size-4" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center justify-between gap-2">
                      <div className="text-sm font-medium">{field.label}</div>
                      <div
                        className={cn(
                          "rounded-full px-2 py-0.5 text-xs",
                          field.imageDataUrl
                            ? "bg-emerald-500/10 text-emerald-600 dark:text-emerald-300"
                            : "bg-muted text-muted-foreground"
                        )}
                      >
                        {field.imageDataUrl ? "Signed" : "Unsigned"}
                      </div>
                    </div>
                    <div className="mt-2 text-xs text-muted-foreground">
                      {Math.round(field.bbox.width)} x{" "}
                      {Math.round(field.bbox.height)} on page {field.page}
                    </div>
                    {showPreview && field.imageDataUrl ? (
                      <div className="mt-3 overflow-hidden rounded-md border bg-white">
                        <img
                          src={field.imageDataUrl}
                          alt="Signature preview"
                          className="block h-[58px] w-full object-contain"
                        />
                      </div>
                    ) : null}
                    <div className="mt-3 flex gap-2">
                      <Button
                        type="button"
                        size="sm"
                        variant={field.imageDataUrl ? "outline" : "default"}
                        className="flex-1"
                      >
                        <HugeiconsIcon icon={Pen01Icon} className="size-4" />
                        {field.imageDataUrl ? "Edit" : "Sign"}
                      </Button>
                      {field.imageDataUrl ? (
                        <Button type="button" size="sm" variant="outline">
                          Clear
                        </Button>
                      ) : null}
                    </div>
                  </div>
                </div>
              </div>
            ))
          )}
          <Button
            type="button"
            className="w-full"
            disabled={!canExport || isDownloading}
          >
            <HugeiconsIcon icon={Download01Icon} className="size-4" />
            {isDownloading ? "Exporting..." : "Export signed PDF"}
          </Button>
        </div>
      </div>
    </aside>
  );
}
