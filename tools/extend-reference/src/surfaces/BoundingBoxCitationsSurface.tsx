import fixtures from "../../../parity-fixtures/fixtures.json";
import { HumanReviewPanel, type ReviewField } from "../vendor/extend/bounding-box-citations";
import { ParityCell } from "../parity/ParitySurface";

const review = fixtures.blocks.boundingBoxCitations;
const fields = review.fields as ReviewField[];
const modifiedExpected = review.modifiedExpected;

const states = [
  { state: "default", matrixState: "default" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
  { state: "selected", matrixState: "default" },
  { state: "modified", matrixState: "default" },
  { state: "empty", matrixState: "empty" },
  { state: "loading", matrixState: "loading" },
  { state: "error", matrixState: "error" },
];

export function BoundingBoxCitationsSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <HumanReviewPanel
            className="w-full"
            fields={entry.state === "empty" ? [] : fields}
            activeFieldKey={
              entry.state === "selected" ? "invoice_total" : undefined
            }
            loading={entry.state === "loading"}
            error={
              entry.state === "error"
                ? "Review fields could not be loaded."
                : undefined
            }
            modifiedFieldKey={
              entry.state === "modified" ? review.modifiedFieldKey : undefined
            }
            modifiedExpected={
              entry.state === "modified" ? modifiedExpected : undefined
            }
          />
        </ParityCell>
      ))}
    </>
  );
}
