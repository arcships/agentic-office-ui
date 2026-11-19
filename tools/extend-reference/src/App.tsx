import manifest from "../../parity-fixtures/manifest.json";
import type { ParityTheme } from "./parity/ParitySurface";
import { ParitySurface } from "./parity/ParitySurface";
import { ButtonSurface } from "./surfaces/ButtonSurface";
import { DocumentSidebarSurface } from "./surfaces/DocumentSidebarSurface";
import { FileThumbnailSurface } from "./surfaces/FileThumbnailSurface";
import { FileUploadSurface } from "./surfaces/FileUploadSurface";
import { PdfToolbarSurface } from "./surfaces/PdfToolbarSurface";

type Surface = (typeof manifest.surfaces)[number];

const renderers: Record<
  string,
  (props: { cellWidth: number }) => React.ReactElement
> = {
  button: ButtonSurface,
  "file-upload": FileUploadSurface,
  "file-thumbnail": FileThumbnailSurface,
  "pdf-toolbar": PdfToolbarSurface,
  "document-sidebar": DocumentSidebarSurface,
};

function hasReference(surface: Surface): boolean {
  return (surface as { referenceStatus?: string }).referenceStatus !== "fixtures-only";
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
