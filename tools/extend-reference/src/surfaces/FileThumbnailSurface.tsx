import { fixtureDocument } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { FileThumbnail } from "../vendor/extend/file-thumbnail";

const pdf = fixtureDocument("pdf");
const png = fixtureDocument("png");

const states = [
  { state: "default", matrixState: "empty" },
  { state: "loading", matrixState: "loading" },
  { state: "error", matrixState: "error" },
  { state: "image", matrixState: "default" },
];

export function FileThumbnailSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <FileThumbnail
            key={entry.state}
            file={{ name: pdf.name, type: pdf.type }}
            className="size-20"
            hasError={entry.state === "error"}
            isLoading={entry.state === "loading"}
            previewImageUrl={entry.state === "image" ? png.asset : null}
          />
        </ParityCell>
      ))}
    </>
  );
}
