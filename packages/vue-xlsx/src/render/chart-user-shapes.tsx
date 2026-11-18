/** @jsxImportSource vue */
import type {
  XlsxChart,
  XlsxChartUserShape,
  XlsxShapeParagraph,
  XlsxShapeTextRun,
} from "@arcships/xlsx-core";
import type { ChartLayout, ChartRendererPalette } from "./chart-types";
import { estimateSvgTextWidth } from "./chart-shared";
import { sanitizeChartImageSource } from "./chart-image-fills";

/**
 * Preset-geometry and text renderer for `chartUserShapes` overlays.
 *
 * A chart can carry a `chartUserShapes` part with text boxes, callouts and
 * pictures painted on top of the plot. The parser normalises their anchors to
 * chart-relative fractions; this module turns those shapes back into SVG using
 * DrawingML preset geometries so a "source: …" note or a logo lands where Excel
 * draws it, at any zoom level.
 */

interface ShapeBox {
  height: number;
  width: number;
  x: number;
  y: number;
}

function shapeAdjustment(shape: XlsxChartUserShape, name: string, fallback: number) {
  const value = shape.geometryAdjustments?.[name];
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function polygonPath(box: ShapeBox, sides: number, rotationDeg: number) {
  const centerX = box.x + box.width / 2;
  const centerY = box.y + box.height / 2;
  const radiusX = box.width / 2;
  const radiusY = box.height / 2;
  const points: string[] = [];
  for (let index = 0; index < sides; index += 1) {
    const angle = (rotationDeg + (index * 360) / sides) * (Math.PI / 180);
    points.push(`${centerX + radiusX * Math.cos(angle)} ${centerY + radiusY * Math.sin(angle)}`);
  }
  return `M ${points.join(" L ")} Z`;
}

function starPath(box: ShapeBox, points: number, innerRatio: number) {
  const centerX = box.x + box.width / 2;
  const centerY = box.y + box.height / 2;
  const radiusX = box.width / 2;
  const radiusY = box.height / 2;
  const path: string[] = [];
  const total = points * 2;
  for (let index = 0; index < total; index += 1) {
    const angle = (-90 + (index * 180) / points) * (Math.PI / 180);
    const ratio = index % 2 === 0 ? 1 : innerRatio;
    const px = centerX + radiusX * ratio * Math.cos(angle);
    const py = centerY + radiusY * ratio * Math.sin(angle);
    path.push(index === 0 ? `M ${px} ${py}` : `L ${px} ${py}`);
  }
  return `${path.join(" ")} Z`;
}

/**
 * SVG path for one DrawingML preset geometry.
 *
 * Coordinates are absolute chart pixels so stroke widths stay uniform. Unknown
 * presets and `custom` fall back to the shape's bounding rectangle: the overlay
 * keeps its fill and text instead of disappearing.
 */
export function chartPresetGeometryPath(shape: XlsxChartUserShape, box: ShapeBox): string {
  const { x, y, width, height } = box;
  const right = x + width;
  const bottom = y + height;
  const centerX = x + width / 2;
  const centerY = y + height / 2;
  const inset = Math.min(width, height);
  const cornerRadius = (inset * shapeAdjustment(shape, "adj", 16667)) / 100000;
  const cornerRadiusX = Math.min(width / 2, cornerRadius);
  const cornerRadiusY = Math.min(height / 2, cornerRadius);
  const arrowThickness = height * (shapeAdjustment(shape, "adj1", 50000) / 100000) / 2;
  const arrowHead = Math.min(width, height) * (shapeAdjustment(shape, "adj2", 50000) / 100000);

  switch (shape.geometry) {
    case "rect":
    case "flowChartProcess":
    case "textBox":
    case "custom":
      return `M ${x} ${y} H ${right} V ${bottom} H ${x} Z`;
    case "roundRect":
    case "flowChartAlternateProcess":
      return `M ${x + cornerRadiusX} ${y} H ${right - cornerRadiusX} A ${cornerRadiusX} ${cornerRadiusY} 0 0 1 ${right} ${y + cornerRadiusY} V ${bottom - cornerRadiusY} A ${cornerRadiusX} ${cornerRadiusY} 0 0 1 ${right - cornerRadiusX} ${bottom} H ${x + cornerRadiusX} A ${cornerRadiusX} ${cornerRadiusY} 0 0 1 ${x} ${bottom - cornerRadiusY} V ${y + cornerRadiusY} A ${cornerRadiusX} ${cornerRadiusY} 0 0 1 ${x + cornerRadiusX} ${y} Z`;
    case "ellipse":
      return `M ${x} ${centerY} A ${width / 2} ${height / 2} 0 1 0 ${right} ${centerY} A ${width / 2} ${height / 2} 0 1 0 ${x} ${centerY} Z`;
    case "triangle":
      return `M ${centerX} ${y} L ${right} ${bottom} L ${x} ${bottom} Z`;
    case "rtTriangle":
      return `M ${x} ${y} L ${x} ${bottom} L ${right} ${bottom} Z`;
    case "diamond":
      return `M ${centerX} ${y} L ${right} ${centerY} L ${centerX} ${bottom} L ${x} ${centerY} Z`;
    case "parallelogram": {
      const offset = (width * shapeAdjustment(shape, "adj", 25000)) / 100000;
      return `M ${x + offset} ${y} H ${right} L ${right - offset} ${bottom} H ${x} Z`;
    }
    case "trapezoid": {
      const offset = (width * shapeAdjustment(shape, "adj", 25000)) / 100000;
      return `M ${x + offset} ${y} H ${right - offset} L ${right} ${bottom} H ${x} Z`;
    }
    case "pentagon":
      return polygonPath(box, 5, -90);
    case "hexagon": {
      const offset = (Math.min(width, height) * shapeAdjustment(shape, "adj", 25000)) / 100000;
      return `M ${x + offset} ${y} H ${right - offset} L ${right} ${centerY} L ${right - offset} ${bottom} H ${x + offset} L ${x} ${centerY} Z`;
    }
    case "heptagon":
      return polygonPath(box, 7, -90);
    case "octagon":
      return polygonPath(box, 8, -112.5);
    case "homePlate": {
      const offset = Math.min(width, height * 0.5);
      return `M ${x} ${y} H ${right - offset} L ${right} ${centerY} L ${right - offset} ${bottom} H ${x} Z`;
    }
    case "chevron":
      return `M ${x} ${y} H ${centerX} L ${right} ${centerY} L ${centerX} ${bottom} H ${x} L ${centerX} ${centerY} Z`;
    case "rightArrow":
      return `M ${x} ${centerY - arrowThickness} H ${right - arrowHead} V ${y} L ${right} ${centerY} L ${right - arrowHead} ${bottom} V ${centerY + arrowThickness} H ${x} Z`;
    case "leftArrow":
      return `M ${right} ${centerY - arrowThickness} H ${x + arrowHead} V ${y} L ${x} ${centerY} L ${x + arrowHead} ${bottom} V ${centerY + arrowThickness} H ${right} Z`;
    case "upArrow":
      return `M ${centerX - arrowThickness} ${bottom} V ${y + arrowHead} H ${x} L ${centerX} ${y} L ${right} ${y + arrowHead} H ${centerX + arrowThickness} V ${bottom} Z`;
    case "downArrow":
      return `M ${centerX - arrowThickness} ${y} V ${bottom - arrowHead} H ${x} L ${centerX} ${bottom} L ${right} ${bottom - arrowHead} H ${centerX + arrowThickness} V ${y} Z`;
    case "leftRightArrow":
      return `M ${x} ${centerY} L ${x + arrowHead} ${y} V ${centerY - arrowThickness} H ${right - arrowHead} V ${y} L ${right} ${centerY} L ${right - arrowHead} ${bottom} V ${centerY + arrowThickness} H ${x + arrowHead} V ${bottom} Z`;
    case "upDownArrow":
      return `M ${centerX} ${y} L ${right} ${y + arrowHead} H ${centerX + arrowThickness} V ${bottom - arrowHead} H ${right} L ${centerX} ${bottom} L ${x} ${bottom - arrowHead} H ${centerX - arrowThickness} V ${y + arrowHead} H ${x} Z`;
    case "star4":
      return starPath(box, 4, 0.28);
    case "star5":
      return starPath(box, 5, 0.38);
    case "star6":
      return starPath(box, 6, 0.4);
    case "star8":
      return starPath(box, 8, 0.5);
    case "plus":
    case "mathPlus": {
      const thickness = (Math.min(width, height) * shapeAdjustment(shape, "adj", 25000)) / 100000;
      const left = centerX - thickness;
      const top = centerY - thickness;
      return `M ${left} ${y} H ${centerX + thickness} V ${top} H ${right} V ${centerY + thickness} H ${centerX + thickness} V ${bottom} H ${left} V ${centerY + thickness} H ${x} V ${top} H ${left} Z`;
    }
    case "plaque": {
      const offset = Math.min(cornerRadiusX, cornerRadiusY);
      return `M ${x + offset} ${y} H ${right - offset} Q ${right} ${y} ${right} ${y + offset} V ${bottom - offset} Q ${right} ${bottom} ${right - offset} ${bottom} H ${x + offset} Q ${x} ${bottom} ${x} ${bottom - offset} V ${y + offset} Q ${x} ${y} ${x + offset} ${y} Z`;
    }
    case "line":
    case "straightConnector1":
      return `M ${x} ${y} L ${right} ${bottom}`;
    case "none":
      return "";
    default:
      // Unknown or animated presets still get a visible body and their text.
      return `M ${x} ${y} H ${right} V ${bottom} H ${x} Z`;
  }
}

// ── Text layout ──────────────────────────────────────────────────────────

const DEFAULT_SHAPE_FONT_SIZE = 10;

interface ShapeTextSegment {
  fontSize: number;
  run: XlsxShapeTextRun;
  text: string;
}

interface ShapeTextLine {
  fontSize: number;
  segments: ShapeTextSegment[];
  width: number;
}

interface ShapeTextBlock {
  align: XlsxShapeParagraph["align"];
  fontSize: number;
  lines: ShapeTextLine[];
}

function runFontSize(run: XlsxShapeTextRun) {
  return typeof run.fontSizePt === "number" && Number.isFinite(run.fontSizePt) && run.fontSizePt > 0
    ? run.fontSizePt
    : DEFAULT_SHAPE_FONT_SIZE;
}

function splitRunTokens(runs: XlsxShapeTextRun[]): ShapeTextSegment[][] {
  // Each paragraph becomes token groups separated by explicit line breaks.
  const groups: ShapeTextSegment[][] = [[]];
  for (const run of runs) {
    const fontSize = runFontSize(run);
    const parts = run.text.split("\n");
    parts.forEach((part, partIndex) => {
      if (partIndex > 0) {
        groups.push([]);
      }
      if (!part) {
        return;
      }
      // Keep trailing spaces glued to the word so wrapping matches the visible text.
      const tokens = part.match(/\s+|\S+\s*/g) ?? [part];
      for (const token of tokens) {
        groups[groups.length - 1].push({ fontSize, run, text: token });
      }
    });
  }
  return groups;
}

function wrapParagraph(paragraph: XlsxShapeParagraph, maxWidth: number): ShapeTextLine[] {
  const lines: ShapeTextLine[] = [];
  for (const tokens of splitRunTokens(paragraph.runs)) {
    let current: ShapeTextLine = { fontSize: DEFAULT_SHAPE_FONT_SIZE, segments: [], width: 0 };
    const flush = () => {
      if (current.segments.length > 0) {
        lines.push(current);
      }
      current = { fontSize: DEFAULT_SHAPE_FONT_SIZE, segments: [], width: 0 };
    };
    for (const token of tokens) {
      const width = estimateSvgTextWidth(token.text, token.fontSize);
      if (current.width + width > maxWidth && current.segments.length > 0) {
        flush();
      }
      const trimmed = current.segments.length === 0 ? token.text.replace(/^\s+/, "") : token.text;
      if (!trimmed) {
        continue;
      }
      current.segments.push({ ...token, text: trimmed });
      current.fontSize = Math.max(current.fontSize, token.fontSize);
      current.width += estimateSvgTextWidth(trimmed, token.fontSize);
    }
    flush();
  }
  return lines;
}

function buildTextBlocks(paragraphs: XlsxShapeParagraph[], maxWidth: number): ShapeTextBlock[] {
  return paragraphs.map((paragraph) => {
    const lines = wrapParagraph(paragraph, Math.max(8, maxWidth));
    return {
      align: paragraph.align,
      fontSize: lines.reduce((size, line) => Math.max(size, line.fontSize), DEFAULT_SHAPE_FONT_SIZE),
      lines,
    };
  }).filter((block) => block.lines.length > 0);
}

function renderTextLine(
  block: Pick<ShapeTextBlock, "align" | "fontSize">,
  line: ShapeTextLine,
  lineIndex: number,
  baseline: number,
  box: ShapeBox,
  textColor: string,
) {
  const anchor = block.align === "center" ? "middle" : block.align === "right" ? "end" : "start";
  const x = block.align === "center"
    ? box.x + box.width / 2
    : block.align === "right"
      ? box.x + box.width
      : box.x;
  return (
    <text
      font-size={line.fontSize}
      fill={textColor}
      key={`shape-text-${lineIndex}`}
      text-anchor={anchor}
      x={x}
      y={baseline}
    >
      {line.segments.map((segment, segmentIndex) => (
        <tspan
          fill={segment.run.color ?? textColor}
          font-size={segment.fontSize}
          font-style={segment.run.italic ? "italic" : undefined}
          font-weight={segment.run.bold ? "bold" : undefined}
          key={`shape-run-${lineIndex}-${segmentIndex}`}
          text-decoration={segment.run.underline ? "underline" : undefined}
        >{segment.text}</tspan>
      ))}
    </text>
  );
}

function renderShapeText(
  shape: XlsxChartUserShape,
  box: ShapeBox,
  palette: ChartRendererPalette,
  chartId: string,
  shapeIndex: number,
) {
  const inset = shape.textBox?.insetPx ?? { bottom: 4.8, left: 9.6, right: 9.6, top: 4.8 };
  const textBox: ShapeBox = {
    height: Math.max(0, box.height - inset.top - inset.bottom),
    width: Math.max(0, box.width - inset.left - inset.right),
    x: box.x + inset.left,
    y: box.y + inset.top,
  };
  if (textBox.width <= 0 || textBox.height <= 0) {
    return null;
  }
  const blocks = buildTextBlocks(shape.paragraphs, textBox.width);
  if (blocks.length === 0) {
    return null;
  }
  const lineHeights = blocks.map((block) => block.lines.map((line) => Math.max(line.fontSize, block.fontSize) * 1.2));
  const totalHeight = lineHeights.reduce((sum, heights) => sum + heights.reduce((inner, value) => inner + value, 0), 0);
  const verticalAlign = shape.textBox?.verticalAlign ?? "top";
  let cursor = verticalAlign === "middle"
    ? textBox.y + (textBox.height - totalHeight) / 2
    : verticalAlign === "bottom"
      ? textBox.y + textBox.height - totalHeight
      : textBox.y;
  const nodes: unknown[] = [];
  let lineIndex = 0;
  blocks.forEach((block, blockIndex) => {
    block.lines.forEach((line, indexInBlock) => {
      const height = lineHeights[blockIndex]?.[indexInBlock] ?? line.fontSize * 1.2;
      // SVG text sits on its baseline, roughly 80 % down the line box.
      const baseline = cursor + height * 0.8;
      nodes.push(renderTextLine(
        block,
        line,
        lineIndex,
        baseline,
        textBox,
        line.segments[0]?.run.color ?? palette.text,
      ));
      cursor += height;
      lineIndex += 1;
    });
  });
  if (nodes.length === 0) {
    return null;
  }
  return (
    <g
      clip-path={`url(#${shapeClipPathId(chartId, shapeIndex)})`}
      data-xlsx-chart-shape-text=""
      key="shape-text"
    >
      {nodes}
    </g>
  );
}

function shapeClipPathId(chartId: string, shapeIndex: number) {
  const safeChartId = chartId.replace(/[^A-Za-z0-9_-]+/g, "-");
  return `xlsx-chart-shape-clip-${safeChartId}-${shapeIndex}`;
}

/**
 * Renders every user shape of a chart's `chartUserShapes` parts.
 *
 * Shapes paint above the plot, in part order, matching Excel's z-order. Returns
 * `null` when the chart carries no shapes so the memoised chart tree stays
 * unchanged for workbooks without overlays.
 */
export function renderChartUserShapes(
  chart: Pick<XlsxChart, "drawings" | "id">,
  layout: Pick<ChartLayout, "height" | "width">,
  palette: ChartRendererPalette,
) {
  const shapes = collectChartUserShapes(chart);
  if (shapes.length === 0) {
    return null;
  }
  const nodes = shapes.map((shape, shapeIndex) => {
    const box: ShapeBox = {
      height: layout.height * shape.rect.height,
      width: layout.width * shape.rect.width,
      x: layout.width * shape.rect.left,
      y: layout.height * shape.rect.top,
    };
    if (box.width <= 0 || box.height <= 0) {
      return null;
    }
    const path = chartPresetGeometryPath(shape, box);
    const image = sanitizeChartImageSource(shape.fill?.image);
    const strokeColor = shape.stroke?.hidden ? "none" : shape.stroke?.color ?? "none";
    const strokeWidth = shape.stroke?.widthPx ?? 1;
    const fillColor = shape.fill?.none ? "none" : shape.fill?.color ?? "none";
    return (
      <g data-xlsx-chart-shape={shape.geometry} data-xlsx-chart-shape-index={shapeIndex} key={`shape-${shapeIndex}`}>
        <defs>
          <clipPath id={shapeClipPathId(chart.id, shapeIndex)}>
            <path d={path} />
          </clipPath>
        </defs>
        {image
          ? <image
              clip-path={`url(#${shapeClipPathId(chart.id, shapeIndex)})`}
              height={box.height}
              href={image}
              preserveAspectRatio="none"
              width={box.width}
              x={box.x}
              y={box.y}
            />
          : null}
        <path
          d={path}
          fill={image ? "none" : fillColor}
          stroke={strokeColor}
          stroke-dasharray={shape.stroke?.dash && shape.stroke.dash !== "solid" ? "4 3" : undefined}
          stroke-width={strokeColor === "none" ? undefined : strokeWidth}
        />
        {renderShapeText(shape, box, palette, chart.id, shapeIndex)}
      </g>
    );
  }).filter(Boolean);
  return nodes.length > 0
    ? <g data-xlsx-chart-user-shapes="" pointer-events="none">{nodes}</g>
    : null;
}

/** Every parsed user shape of the chart, in part order. */
export function collectChartUserShapes(chart: Pick<XlsxChart, "drawings">): XlsxChartUserShape[] {
  if (!chart.drawings) {
    return [];
  }
  return Object.values(chart.drawings).flatMap((drawing) => drawing.shapes ?? []);
}
