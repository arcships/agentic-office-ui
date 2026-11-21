import { ParityCell } from "../parity/ParitySurface";
import { FileUpload } from "../vendor/extend/file-upload";

const states = [
  { state: "default", matrixState: "empty" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
  { state: "dragging", matrixState: "active" },
  { state: "rejected", matrixState: "error" },
  { state: "file-selected", matrixState: "default" },
];

export function FileUploadSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          {/* No props: the cell measures the component's shipped defaults. */}
          <FileUpload />
        </ParityCell>
      ))}
    </>
  );
}
