import fixtures from "../../../parity-fixtures/fixtures.json";
import {
  LayoutBlocksPanel,
  type LayoutBlockItem,
} from "../vendor/extend/layout-blocks-panel";
import { ParityCell } from "../parity/ParitySurface";

const blocks = fixtures.blocks.layoutBlocks as LayoutBlockItem[];

const states = [
  { state: "default", matrixState: "default", activeBlockId: undefined },
  { state: "hover", matrixState: "hover", activeBlockId: undefined },
  { state: "focus", matrixState: "focus", activeBlockId: undefined },
  { state: "active", matrixState: "default", activeBlockId: "block-2" },
  { state: "empty", matrixState: "empty", activeBlockId: undefined },
  { state: "loading", matrixState: "loading", activeBlockId: undefined },
  { state: "error", matrixState: "error", activeBlockId: undefined },
];

export function LayoutBlocksSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <LayoutBlocksPanel
            className="w-full"
            blocks={entry.state === "empty" ? [] : blocks}
            activeBlockId={entry.activeBlockId}
            loading={entry.state === "loading"}
            error={
              entry.state === "error"
                ? "Layout blocks could not be loaded."
                : undefined
            }
          />
        </ParityCell>
      ))}
    </>
  );
}
