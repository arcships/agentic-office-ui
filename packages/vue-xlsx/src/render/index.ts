// Barrel exports for the xlsx-render module.
// Chart rendering (SVG-based, adapted from upstream @extend-ai/react-xlsx chart-renderer.tsx).
export { MemoChartSvg } from "./chart-renderer";

// Surface/WebGL 3D rendering (regl-based, adapted from upstream surface-regl.tsx).
export { MemoSurfaceChartComposite } from "./surface-regl";

// Types used by the render module.
export type {
  ChartRendererPalette,
  ChartSvgProps,
  ChartLayout,
  LegendItem,
} from "./chart-renderer";
