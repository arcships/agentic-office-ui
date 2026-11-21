import * as React from "react";

import { PDFViewer } from "@/components/extend/pdf-viewer";

/**
 * Upstream docs composition for `/docs/components/pdf-viewer`.
 *
 * Source of truth: `apps/v4/components/pdf-viewer-docs.tsx:207`
 * (`PdfViewerDemo`) -> `pdf-viewer-preview-client.tsx` (`PdfViewerPreviewClient`).
 *
 * Deliberate differences from the docs page, both owned by the harness:
 * - the docs page's code block (`DocsViewCodeBlock`, `[data-slot="code"]`) is
 *   not rendered: it is page scaffolding, excluded from the product region;
 * - the page shell (`padding: 24px; max-width: 1200px`) reproduces the demo
 *   route's host box so both sides hand the viewer the same available width
 *   (see `docs/testing/ui-parity-routes.md`).
 *
 * The fixture, zoom and theme come from the query, the same contract the demo
 * route reads, so a capture can pair both sides on one state.
 */
function readParam(params: URLSearchParams, name: string): string | null {
  const value = params.get(name);
  return value === null || value.trim() === "" ? null : value;
}

export function PdfViewerRoute(): React.ReactElement {
  const params = new URLSearchParams(window.location.search);
  const doc = readParam(params, "doc") ?? "/fixtures/sample.pdf";
  const zoomParam = Number(readParam(params, "zoom") ?? "1");
  const defaultZoom = Number.isFinite(zoomParam) ? zoomParam : 1;

  return (
    <div className="route-shell">
      <div
        data-slot="component-preview"
        className="group relative flex flex-col overflow-hidden rounded-xl border"
      >
        <PDFViewer
          src={doc}
          className="h-[560px]"
          defaultZoom={defaultZoom}
          showRotateControls
        />
      </div>
    </div>
  );
}
