import fixtures from "../../../parity-fixtures/fixtures.json";
import {
  DocumentSplitsPanel,
  type DocumentSplit,
} from "../vendor/extend/document-splits";
import { ParityCell } from "../parity/ParitySurface";

const splits = fixtures.blocks.documentSplits;
// Warm the shared page image before thumbnails mount (deterministic capture).
if (typeof window !== "undefined") {
  const preload = new Image();
  preload.src = splits.thumbnailImage;
}

const thumbnailImages = Object.fromEntries(
  splits.splits.flatMap((group) =>
    group.pages.map((pageId) => [pageId, splits.thumbnailImage])
  )
);

const states = [
  { state: "default", matrixState: "default", splits: splits.splits as DocumentSplit[] },
  { state: "hover", matrixState: "hover", splits: splits.splits as DocumentSplit[] },
  { state: "focus", matrixState: "focus", splits: splits.splits as DocumentSplit[] },
  { state: "single", matrixState: "default", splits: splits.singleSplit as DocumentSplit[] },
  { state: "empty-pages", matrixState: "empty", splits: splits.emptyPagesSplit as DocumentSplit[] },
  { state: "empty", matrixState: "empty", splits: [] as DocumentSplit[] },
  { state: "loading", matrixState: "loading", splits: [] as DocumentSplit[] },
  { state: "error", matrixState: "error", splits: [] as DocumentSplit[] },
];

export function DocumentSplitsSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <div style={{ height: 640 }}>
            <DocumentSplitsPanel
              splits={entry.splits}
              thumbnailImages={thumbnailImages}
              loading={entry.state === "loading"}
              error={
                entry.state === "error"
                  ? "Document splits could not be loaded."
                  : undefined
              }
            />
          </div>
        </ParityCell>
      ))}
    </>
  );
}
