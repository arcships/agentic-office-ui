/** @jsxImportSource vue */
import type { XlsxChart } from "@arcships/xlsx-core";
import { chartImageFillId, collectChartImageFills } from "./chart-image-fills";

/**
 * `<defs>` for every picture fill the chart paints.
 *
 * `url(#…)` fills are only resolvable inside the SVG that defines them, so the
 * patterns live in the chart's own root `<svg>`. Each pattern stretches its
 * image over the filled shape's bounding box (`objectBoundingBox`), which is how
 * Excel's default "stretch" picture fill behaves for bars, slices and areas.
 */
export function renderChartImageFillDefs(chart: XlsxChart) {
  const fills = collectChartImageFills(chart);
  if (fills.length === 0) {
    return null;
  }
  return (
    <defs data-xlsx-chart-image-defs="">
      {fills.map((fill) => (
        <pattern
          height="1"
          id={chartImageFillId(chart, fill.key)}
          key={fill.key}
          patternContentUnits="objectBoundingBox"
          width="1"
        >
          <image height="1" href={fill.url} preserveAspectRatio="none" width="1" />
        </pattern>
      ))}
    </defs>
  );
}
