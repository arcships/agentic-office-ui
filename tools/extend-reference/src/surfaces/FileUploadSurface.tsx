import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { FileUpload } from "../vendor/extend/file-upload";

const uploadCopy = fixtures.copy.fileUpload;
const uploadFixture = fixtures.fileUpload;

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
          <FileUpload
            acceptedFileTypes={uploadFixture.acceptedFileTypes}
            accept={uploadFixture.accept}
            browseLabel={uploadCopy.browseLabel}
            description={uploadCopy.description}
            draggingLabel={uploadCopy.draggingLabel}
            multiple={uploadFixture.multiple}
            showFileList
            title={uploadCopy.title}
          />
        </ParityCell>
      ))}
    </>
  );
}
