/**
 * Issue #161 — PPTX slide-sorter card parity surface.
 *
 * The pinned Extend PPTX viewer has no slide-sorter grid; its only slide card
 * markup is the thumbnail-rail item (`PptxSidebarThumbnail` inside the
 * `role="option"` wrapper). This surface renders that vendored upstream card in
 * a harness grid (3 columns, 12px gap — the production sorter's own grid
 * metrics) so the card padding, radius, shadow, label, hover and active states
 * are compared pixel for pixel. The grid composition itself is a harness
 * fixture, not an upstream layout.
 */
import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { PptxSlideCard } from "../vendor/extend/pptx-slide-card";
import slideUrl from "../../../parity-fixtures/assets/pptx-slide.png";

const sorter = fixtures.pptxSorter;
const cardCount = 3;
const columns = 3;

const states = [
  { state: "default", matrixState: "default", activeSlide: 1 },
  { state: "hover", matrixState: "hover", activeSlide: 1 },
  { state: "selected", matrixState: "active", activeSlide: 3 },
];

function gridStyle(): React.CSSProperties {
  return {
    display: "grid",
    gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`,
    gap: "12px",
    width: "100%",
  };
}

export function PptxSorterSurface({ cellWidth }: { cellWidth: number }) {
  void cellWidth;
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <div data-pptx-sorter-grid="" style={gridStyle()}>
            {Array.from({ length: cardCount }, (_, index) => (
              <PptxSlideCard
                key={index}
                slideNumber={index + 1}
                active={index + 1 === entry.activeSlide}
                aspectRatio={sorter.aspectRatio}
                imageUrl={slideUrl}
              />
            ))}
          </div>
        </ParityCell>
      ))}
    </>
  );
}
