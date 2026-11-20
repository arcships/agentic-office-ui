import fixtures from "../../../parity-fixtures/fixtures.json";
import {
  SignatureFieldsPanel,
  type SignatureField,
} from "../vendor/extend/signature-fields";
import { ParityCell } from "../parity/ParitySurface";

const signature = fixtures.blocks.signature;
const unsignedField: SignatureField = {
  id: "signature-1",
  label: signature.label,
  page: signature.page,
  bbox: signature.bbox,
};
const signedField: SignatureField = {
  ...unsignedField,
  imageDataUrl: signature.imageDataUrl,
};

const states = [
  { state: "unsigned", matrixState: "default" },
  { state: "signed", matrixState: "default" },
  { state: "exporting", matrixState: "loading" },
  { state: "preview", matrixState: "default" },
  { state: "loading", matrixState: "loading" },
  { state: "error", matrixState: "error" },
];

export function ESignatureSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => {
        const signed = entry.state === "signed" || entry.state === "exporting" || entry.state === "preview";
        return (
          <ParityCell
            key={entry.state}
            matrixState={entry.matrixState}
            state={entry.state}
            width={cellWidth}
          >
            <SignatureFieldsPanel
              className="w-full"
              fields={signed ? [signedField] : [unsignedField]}
              canExport={signed}
              isDownloading={entry.state === "exporting"}
              showPreview={entry.state === "preview"}
              loading={entry.state === "loading"}
              error={
                entry.state === "error"
                  ? "Signature fields could not be loaded."
                  : undefined
              }
            />
          </ParityCell>
        );
      })}
    </>
  );
}
