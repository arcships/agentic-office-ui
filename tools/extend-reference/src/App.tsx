import manifest from "../../parity-fixtures/manifest.json";
import type { ParityTheme } from "./parity/ParitySurface";
import { ParitySurface } from "./parity/ParitySurface";
import { ButtonSurface } from "./surfaces/ButtonSurface";
import { BoundingBoxCitationsSurface } from "./surfaces/BoundingBoxCitationsSurface";
import { DialogSurface } from "./surfaces/DialogSurface";
import { DocumentSplitsSurface } from "./surfaces/DocumentSplitsSurface";
import { ESignatureSurface } from "./surfaces/ESignatureSurface";
import { LayoutBlocksSurface } from "./surfaces/LayoutBlocksSurface";
import { DocxSidebarSurface } from "./surfaces/DocxSidebarSurface";
import { DocxToolbarSurface } from "./surfaces/DocxToolbarSurface";
import { DocumentSidebarSurface } from "./surfaces/DocumentSidebarSurface";
import { FileThumbnailSurface } from "./surfaces/FileThumbnailSurface";
import { FileUploadSurface } from "./surfaces/FileUploadSurface";
import { InputSurface } from "./surfaces/InputSurface";
import { PdfSidebarSurface } from "./surfaces/PdfSidebarSurface";
import { PdfToolbarSurface } from "./surfaces/PdfToolbarSurface";
import { PptxSorterSurface } from "./surfaces/PptxSorterSurface";
import { PptxToolbarSurface } from "./surfaces/PptxToolbarSurface";
import { SelectSurface } from "./surfaces/SelectSurface";
import { TooltipSurface } from "./surfaces/TooltipSurface";
import { XlsxSheetTabsSurface } from "./surfaces/XlsxSheetTabsSurface";
import { XlsxStatusBarSurface } from "./surfaces/XlsxStatusBarSurface";
import { XlsxToolbarSurface } from "./surfaces/XlsxToolbarSurface";

type Surface = (typeof manifest.surfaces)[number];

const renderers: Record<
  string,
  (props: { cellWidth: number }) => React.ReactElement
> = {
  button: ButtonSurface,
  input: InputSurface,
  select: SelectSurface,
  dialog: DialogSurface,
  tooltip: TooltipSurface,
  "file-upload": FileUploadSurface,
  "file-thumbnail": FileThumbnailSurface,
  "pdf-toolbar": PdfToolbarSurface,
  "pdf-sidebar": PdfSidebarSurface,
  "pptx-toolbar": PptxToolbarSurface,
  "pptx-sorter": PptxSorterSurface,
  "docx-toolbar": DocxToolbarSurface,
  "docx-sidebar": DocxSidebarSurface,
  "document-sidebar": DocumentSidebarSurface,
  "bounding-box-citations": BoundingBoxCitationsSurface,
  "layout-blocks": LayoutBlocksSurface,
  "e-signature": ESignatureSurface,
  "document-splits": DocumentSplitsSurface,
  "xlsx-toolbar": XlsxToolbarSurface,
  "xlsx-sheet-tabs": XlsxSheetTabsSurface,
  "xlsx-status-bar": XlsxStatusBarSurface,
};

function hasReference(surface: Surface): boolean {
  const status = (surface as { referenceStatus?: string }).referenceStatus;
  return status !== "fixtures-only" && status !== "unreachable";
}

export function App(): React.ReactElement {
  const params = new URLSearchParams(window.location.search);
  const theme: ParityTheme = params.get("theme") === "dark" ? "dark" : "light";
  const surfaceParam = params.get("surface");
  const selected = manifest.surfaces.filter(
    (surface) =>
      hasReference(surface) &&
      (!surfaceParam || surface.id === surfaceParam),
  );

  if (selected.length === 0) {
    return (
      <div style={{ padding: 24, fontFamily: "var(--font-sans)" }}>
        No reference surface matches “{surfaceParam ?? ""}”.
      </div>
    );
  }

  return (
    <>
      {selected.map((surface) => {
        const renderer = renderers[surface.id];
        return (
          <ParitySurface
            key={surface.id}
            surface={surface.id}
            theme={theme}
          >
            {renderer ? (
              renderer({ cellWidth: surface.cellWidth })
            ) : (
              <p>No reference renderer for {surface.id}.</p>
            )}
          </ParitySurface>
        );
      })}
    </>
  );
}
