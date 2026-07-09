/** @jsxImportSource vue */
import type { XlsxChart, XlsxChartElementSelection } from "@extend-ai/xlsx-core";
import { clamp, lightenColor, darkenColor, resolveChartAxisTextColor } from "./chart-shared";
import { chartElementDataProps, isSelectedChartPointOrSeries, renderSelectionPointHandles } from "./chart-element";
import { chartSeriesColor, chartSeriesStrokeColor, normalizeChartMarkerSymbol } from "./chart-data";
import { markerSymbolPath, toSvgNumber } from "./chart-svg-utils";
import { resolveAxisDomainWithChartOverrides, formatTickValue } from "./chart-axis";
import type { ChartRendererPalette, ChartLayout, BarRect } from "./chart-types";

export function resolve3dFrameOffsets(chart: XlsxChart, baseDepthX = 11, baseDepthY = 8) {
  const depthRatio = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.35, 3.2);
  const rotX = clamp(chart.view3d?.rotX ?? 20, -80, 80);
  const rotY = clamp(chart.view3d?.rotY ?? 20, -80, 80);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 26 : 0)) / 100, 0, 1);
  const projectedOrigin = projectCartesian3dPoint(
    0,
    0,
    0,
    rotX * (Math.PI / 180),
    rotY * (Math.PI / 180),
    usePerspective,
    perspectiveStrength
  );
  const projectedDepthPoint = projectCartesian3dPoint(
    0,
    0,
    1,
    rotX * (Math.PI / 180),
    rotY * (Math.PI / 180),
    usePerspective,
    perspectiveStrength
  );
  const depthDirectionX = projectedDepthPoint.x - projectedOrigin.x;
  const depthDirectionY = projectedDepthPoint.y - projectedOrigin.y;
  const horizontalBias = clamp(1.42 + Math.abs(rotY) / 48 + depthRatio * 0.06, 1.5, 1.95);
  const verticalBias = clamp(1.18 + Math.abs(rotX) / 78, 1.22, 1.5);
  const biasedDepthDirectionX = depthDirectionX * horizontalBias;
  const biasedDepthDirectionY = depthDirectionY * verticalBias;
  const depthDirectionMagnitude = Math.hypot(biasedDepthDirectionX, biasedDepthDirectionY) || 1;
  const horizontalFactor = clamp(Math.abs(rotY) / 22, 0.55, 2.2);
  const verticalFactor = clamp(Math.abs(rotX) / 18, 0.45, 1.9);
  const legacyDepthX = Math.max(6, baseDepthX * depthRatio * horizontalFactor) * (rotY < 0 ? -1 : 1);
  const legacyDepthY = -Math.max(4, baseDepthY * depthRatio * verticalFactor);
  const depthMagnitude = Math.max(10, Math.hypot(legacyDepthX, legacyDepthY) * 1.34);
  const depthX = biasedDepthDirectionX / depthDirectionMagnitude * depthMagnitude;
  const depthY = biasedDepthDirectionY / depthDirectionMagnitude * depthMagnitude;
  return {
    depthRatio,
    depthX,
    depthY,
    insetBottom: Math.max(6, Math.abs(depthY) + 5),
    insetLeft: depthX < 0 ? Math.abs(depthX) + 4 : 0,
    insetRight: depthX > 0 ? depthX + 4 : 0,
    insetTop: Math.max(4, Math.abs(depthY) + 2)
  };
}

export function scaleProjectedVector(vector: { x: number; y: number }, scale: number) {
  return {
    x: vector.x * scale,
    y: vector.y * scale
  };
}

export function sampleProjectedEllipseArc(
  center: { x: number; y: number },
  axisA: { x: number; y: number },
  axisB: { x: number; y: number },
  startAngle: number,
  endAngle: number
) {
  const steps = Math.max(12, Math.ceil(Math.abs(endAngle - startAngle) / (Math.PI / 14)));
  const points: Array<{ x: number; y: number }> = [];
  for (let step = 0; step <= steps; step += 1) {
    const ratio = step / steps;
    const angle = startAngle + ((endAngle - startAngle) * ratio);
    points.push({
      x: center.x + (axisA.x * Math.cos(angle)) + (axisB.x * Math.sin(angle)),
      y: center.y + (axisA.y * Math.cos(angle)) + (axisB.y * Math.sin(angle))
    });
  }
  return points;
}

export function buildRibbonSvgPath(startArc: Array<{ x: number; y: number }>, endArc: Array<{ x: number; y: number }>) {
  if (startArc.length === 0 || endArc.length === 0) {
    return "";
  }
  return buildLinearSvgPath([...startArc, ...endArc.slice().reverse()], true);
}

export function renderRadialFrustum(
  bar: BarRect,
  normalizedShape: "cone" | "cylinder",
  frontFill: string,
  sideFill: string,
  topFill: string
) {
  const depthX = bar.depthX ?? (bar.isHorizontal ? 10 : 9);
  const depthY = bar.depthY ?? -7;
  const frontX = bar.left + (bar.depthOffsetX ?? 0);
  const frontY = bar.top + (bar.depthOffsetY ?? 0);
  const frontW = bar.width;
  const frontH = bar.height;
  const centerX = frontX + frontW / 2;
  const centerY = frontY + frontH / 2;
  const showStartCap = bar.capStart !== false;
  const showEndCap = bar.capEnd !== false;
  const startScale = normalizedShape === "cone" ? clamp(bar.bottomScale ?? 1, 0, 1) : 1;
  const endScale = normalizedShape === "cone" ? clamp(bar.topScale ?? 1, 0, 1) : 1;
  const radialDepthScale = normalizedShape === "cylinder"
    ? (bar.isHorizontal ? 0.52 : 0.48)
    : (bar.isHorizontal ? 0.7 : 0.58);
  const depthAxis = scaleProjectedVector({ x: depthX / 2, y: depthY / 2 }, radialDepthScale);
  const secondaryStrokeWidth = Math.max(0.6, bar.strokeWidth * 0.65);

  if (bar.isHorizontal) {
    const startCenter = { x: frontX + depthAxis.x, y: centerY + depthAxis.y };
    const endCenter = { x: frontX + frontW + depthAxis.x, y: centerY + depthAxis.y };
    const startAxisA = { x: 0, y: (frontH * startScale) / 2 };
    const endAxisA = { x: 0, y: (frontH * endScale) / 2 };
    const startAxisB = scaleProjectedVector(depthAxis, startScale);
    const endAxisB = scaleProjectedVector(depthAxis, endScale);
    const backStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI);
    const backEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI);
    const frontStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, Math.PI, Math.PI * 2);
    const frontEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, Math.PI, Math.PI * 2);
    const backWallPath = buildRibbonSvgPath(backEndArc, backStartArc);
    const frontWallPath = buildRibbonSvgPath(frontEndArc, frontStartArc);
    const endCapPath = endScale > 0.018
      ? buildLinearSvgPath(sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI * 2), true)
      : "";
    const startCapPath = startScale > 0.028
      ? buildLinearSvgPath(sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI * 2), true)
      : "";
    return (
      <g key={`${bar.key}-3d-horizontal-${normalizedShape}`}>
        {backWallPath ? <path d={backWallPath} fill={sideFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
        {showEndCap && endCapPath ? <path d={endCapPath} fill={topFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
        {frontWallPath ? <path d={frontWallPath} fill={frontFill} stroke={bar.stroke} stroke-width={bar.strokeWidth} /> : null}
        {showStartCap && startCapPath ? <path d={startCapPath} fill={frontFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
      </g>
    );
  }

  const startCenter = { x: centerX + depthAxis.x, y: frontY + frontH + depthAxis.y };
  const endCenter = { x: centerX + depthAxis.x, y: frontY + depthAxis.y };
  const startAxisA = { x: (frontW * startScale) / 2, y: 0 };
  const endAxisA = { x: (frontW * endScale) / 2, y: 0 };
  const startAxisB = scaleProjectedVector(depthAxis, startScale);
  const endAxisB = scaleProjectedVector(depthAxis, endScale);
  const backStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI);
  const backEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI);
  const frontStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, Math.PI, Math.PI * 2);
  const frontEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, Math.PI, Math.PI * 2);
  const backWallPath = buildRibbonSvgPath(backEndArc, backStartArc);
  const frontWallPath = buildRibbonSvgPath(frontEndArc, frontStartArc);
  const endCapPath = endScale > 0.018
    ? buildLinearSvgPath(sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI * 2), true)
    : "";
  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      {backWallPath ? <path d={backWallPath} fill={sideFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
      {showEndCap && endCapPath ? <path d={endCapPath} fill={topFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
      {frontWallPath ? <path d={frontWallPath} fill={frontFill} stroke={bar.stroke} stroke-width={bar.strokeWidth} /> : null}
    </g>
  );
}

export function renderExtrudedRect(bar: BarRect) {
  const normalizedShape = (() => {
    switch (bar.shape3d) {
      case "cone":
      case "coneToMax":
        return "cone";
      case "cylinder":
        return "cylinder";
      case "pyramid":
      case "pyramidToMax":
        return "pyramid";
      default:
        return "box";
    }
  })();
  const depthX = bar.depthX ?? (bar.isHorizontal ? 10 : 9);
  const depthY = bar.depthY ?? -7;
  const frontX = bar.left + (bar.depthOffsetX ?? 0);
  const frontY = bar.top + (bar.depthOffsetY ?? 0);
  const frontW = bar.width;
  const frontH = bar.height;
  const frontX2 = frontX + frontW;
  const frontY2 = frontY + frontH;

  const sideAnchorX = frontX2;
  const sideDepthX = depthX;
  const centerX = frontX + frontW / 2;
  const bottomScale = clamp(bar.bottomScale ?? 1, 0.04, 1);
  const topScale = clamp(bar.topScale ?? 1, 0.04, 1);
  const bottomHalfWidth = (frontW * bottomScale) / 2;
  const topHalfWidth = (frontW * topScale) / 2;
  const bottomLeft = centerX - bottomHalfWidth;
  const bottomRight = centerX + bottomHalfWidth;
  const topLeft = centerX - topHalfWidth;
  const topRight = centerX + topHalfWidth;
  const showStartCap = bar.capStart !== false;
  const showEndCap = bar.capEnd !== false;

  const topFace = `${frontX},${frontY} ${frontX2},${frontY} ${frontX2 + sideDepthX},${frontY + depthY} ${frontX + sideDepthX},${frontY + depthY}`;
  const sideFace = `${sideAnchorX},${frontY} ${sideAnchorX},${frontY2} ${sideAnchorX + sideDepthX},${frontY2 + depthY} ${sideAnchorX + sideDepthX},${frontY + depthY}`;
  const frontFill = bar.gradientId ? `url(#${bar.gradientId})` : bar.color;
  const sideFill = bar.invertedNegative ? bar.color : darkenColor(bar.color, 0.22);
  const topFill = bar.invertedNegative ? lightenColor(bar.color, 0.04) : lightenColor(bar.color, 0.24);

  if (normalizedShape === "cylinder" || normalizedShape === "cone") {
    return renderRadialFrustum(bar, normalizedShape, frontFill, sideFill, topFill);
  }

  if (normalizedShape === "box") {
    return (
      <g key={`${bar.key}-3d`}>
        <polygon fill={sideFill} points={sideFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <rect fill={frontFill} height={frontH} stroke={bar.stroke} stroke-width={bar.strokeWidth} width={frontW} x={frontX} y={frontY} />
      </g>
    );
  }

  if (bar.isHorizontal) {
    const centerY = frontY + frontH / 2;
    const startHalfHeight = (frontH * clamp(bar.bottomScale ?? 1, 0.04, 1)) / 2;
    const endHalfHeight = (frontH * clamp(bar.topScale ?? 1, 0.04, 1)) / 2;
    const startTop = centerY - startHalfHeight;
    const startBottom = centerY + startHalfHeight;
    const endTop = centerY - endHalfHeight;
    const endBottom = centerY + endHalfHeight;
    const topFacePoints = `${frontX},${startTop} ${frontX2},${endTop} ${frontX2 + sideDepthX},${endTop + depthY} ${frontX + sideDepthX},${startTop + depthY}`;
    const farSidePoints = `${frontX2},${endTop} ${frontX2},${endBottom} ${frontX2 + sideDepthX},${endBottom + depthY} ${frontX2 + sideDepthX},${endTop + depthY}`;

    const taperedFarFace = `${frontX2},${endTop} ${frontX2 + sideDepthX},${endTop + depthY} ${frontX2 + sideDepthX},${endBottom + depthY} ${frontX2},${endBottom}`;
    const frontPolygon = `${frontX},${startTop} ${frontX2},${endTop} ${frontX2},${endBottom} ${frontX},${startBottom}`;

    return (
      <g key={`${bar.key}-3d-horizontal-${normalizedShape}`}>
        <polygon fill={sideFill} points={taperedFarFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFacePoints} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={frontFill} points={frontPolygon} stroke={bar.stroke} stroke-width={bar.strokeWidth} />
      </g>
    );
  }

  const taperedTopFace = `${topLeft},${frontY} ${topRight},${frontY} ${topRight + sideDepthX},${frontY + depthY} ${topLeft + sideDepthX},${frontY + depthY}`;
  const taperedSideFace = `${topRight},${frontY} ${bottomRight},${frontY2} ${bottomRight + sideDepthX},${frontY2 + depthY} ${topRight + sideDepthX},${frontY + depthY}`;
  const pyramidFrontFace = `${topLeft},${frontY} ${topRight},${frontY} ${bottomRight},${frontY2} ${bottomLeft},${frontY2}`;

  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      <polygon fill={sideFill} points={taperedSideFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
      {showEndCap ? (
        <polygon fill={topFill} points={taperedTopFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
      ) : null}
      <polygon fill={frontFill} points={pyramidFrontFace} stroke={bar.stroke} stroke-width={bar.strokeWidth} />
    </g>
  );
}

export function buildLinearSvgPath(points: Array<{ x: number; y: number }>, close = false) {
  if (points.length === 0) {
    return "";
  }
  const commands = points.map((point, index) => (
    `${index === 0 ? "M" : "L"} ${toSvgNumber(point.x)} ${toSvgNumber(point.y)}`
  ));
  if (close) {
    commands.push("Z");
  }
  return commands.join(" ");
}

export function projectCartesian3dPoint(
  x: number,
  y: number,
  z: number,
  rotXRad: number,
  rotYRad: number,
  usePerspective: boolean,
  perspectiveStrength: number
) {
  const cosX = Math.cos(rotXRad);
  const sinX = Math.sin(rotXRad);
  const cosY = Math.cos(rotYRad);
  const sinY = Math.sin(rotYRad);

  const x1 = x * cosY + z * sinY;
  const z1 = -x * sinY + z * cosY;
  const y1 = y * cosX - z1 * sinX;
  const z2 = y * sinX + z1 * cosX;
  const perspective = usePerspective
    ? 1 / Math.max(0.22, 1 + z2 * (0.26 + perspectiveStrength * 0.54))
    : 1;

  return {
    depth: z2,
    x: x1 * perspective,
    y: y1 * perspective
  };
}

export function renderLineOrAreaChart3d(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  categories: string[],
  stackedPointsBySeries: Array<Array<{ defined: boolean; y: number | null; y0: number | null; y1: number | null }>>,
  minValue: number,
  maxValue: number,
  isAreaChart: boolean,
  isStackedSeries: boolean,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  const plot = layout.plot;
  const valueDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minValue,
    maxValue,
    isAreaChart || chart.valueAxis?.crosses === "autoZero"
  );
  minValue = valueDomain.min;
  maxValue = valueDomain.max;
  const valueSpan = Math.max(1e-6, maxValue - minValue);
  const seriesCount = Math.max(1, chart.series.length);
  const categoryCount = Math.max(1, categories.length);
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.5, 4);
  const halfDepth = (seriesCount <= 1 ? 0.82 : 1.02) * depthScale;
  const frontZ = halfDepth;
  const backZ = -halfDepth;
  const baseValue = isAreaChart ? Math.max(minValue, Math.min(maxValue, 0)) : minValue;
  const rotXRad = clamp(chart.view3d?.rotX ?? 18, -80, 80) * (Math.PI / 180);
  const rotYRad = clamp(chart.view3d?.rotY ?? 24, -80, 80) * (Math.PI / 180);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 26 : 0)) / 100, 0, 1);

  const normalizeX = (categoryIndex: number) => (
    categoryCount <= 1 ? 0 : ((categoryIndex / (categoryCount - 1)) - 0.5) * 2
  );
  const normalizeY = (value: number) => (
    -((((value - minValue) / valueSpan) - 0.5) * 2)
  );
  const stackedAreaDepthSpan = isAreaChart && isStackedSeries
    ? Math.max(0.18, (frontZ - backZ) * 0.56)
    : 0;
  const stackedAreaBackZ = frontZ - stackedAreaDepthSpan;
  const normalizeZ = (seriesIndex: number) => (
    seriesCount <= 1 || (isAreaChart && isStackedSeries)
      ? frontZ * 0.94
      : (((seriesIndex / (seriesCount - 1)) - 0.5) * (frontZ - backZ))
  );
  const projectSeriesPoint = (x: number, value: number, z: number) => projectCartesian3dPoint(
    x,
    normalizeY(value),
    z,
    rotXRad,
    rotYRad,
    usePerspective,
    perspectiveStrength
  );

  const cubeCorners = [
    { x: -1, y: -1, z: backZ },
    { x: 1, y: -1, z: backZ },
    { x: 1, y: 1, z: backZ },
    { x: -1, y: 1, z: backZ },
    { x: -1, y: -1, z: frontZ },
    { x: 1, y: -1, z: frontZ },
    { x: 1, y: 1, z: frontZ },
    { x: -1, y: 1, z: frontZ }
  ].map((corner) => projectCartesian3dPoint(
    corner.x,
    corner.y,
    corner.z,
    rotXRad,
    rotYRad,
    usePerspective,
    perspectiveStrength
  ));

  const projectedSeries = stackedPointsBySeries.map((seriesPoints, seriesIndex) => {
    const z = normalizeZ(seriesIndex);
    return seriesPoints.map((point, categoryIndex) => {
      const x = normalizeX(categoryIndex);
      const topValue = isStackedSeries ? (point.y1 ?? point.y ?? baseValue) : (point.y ?? baseValue);
      const bottomValue = isAreaChart
        ? (isStackedSeries ? (point.y0 ?? baseValue) : baseValue)
        : baseValue;
      const top = projectSeriesPoint(x, topValue, z);
      const bottom = projectSeriesPoint(x, bottomValue, z);
      const topBack = isAreaChart && isStackedSeries
        ? projectSeriesPoint(x, topValue, stackedAreaBackZ)
        : null;
      const bottomBack = isAreaChart && isStackedSeries
        ? projectSeriesPoint(x, bottomValue, stackedAreaBackZ)
        : null;
      return {
        bottom,
        bottomBack,
        categoryIndex,
        defined: point.defined,
        depth: top.depth,
        depthBack: topBack?.depth ?? top.depth,
        top,
        topBack
      };
    });
  });

  const bounds = [
    ...cubeCorners,
    ...projectedSeries.flatMap((series) => series.flatMap((point) => point.defined
      ? [point.top, point.bottom, ...(point.topBack ? [point.topBack] : []), ...(point.bottomBack ? [point.bottomBack] : [])]
      : []))
  ];
  const minX = Math.min(...bounds.map((point) => point.x));
  const maxX = Math.max(...bounds.map((point) => point.x));
  const minY = Math.min(...bounds.map((point) => point.y));
  const maxY = Math.max(...bounds.map((point) => point.y));
  const scale = Math.min(
    plot.width / Math.max(0.4, maxX - minX),
    plot.height / Math.max(0.4, maxY - minY)
  ) * 0.82;
  const centerRawX = (minX + maxX) / 2;
  const centerRawY = (minY + maxY) / 2;
  const centerX = plot.left + plot.width / 2;
  const centerY = plot.top + plot.height / 2;
  const toScreenPoint = (point: { depth: number; x: number; y: number }) => ({
    depth: point.depth,
    x: centerX + (point.x - centerRawX) * scale,
    y: centerY + (point.y - centerRawY) * scale
  });

  const screenCorners = cubeCorners.map(toScreenPoint);
  const seriesGeometry = projectedSeries.map((series, seriesIndex) => ({
    averageDepth: series.reduce((sum, point) => sum + ((point.depth + point.depthBack) / 2), 0) / Math.max(1, series.length),
    points: series.map((point) => ({
      bottom: toScreenPoint(point.bottom),
      bottomBack: point.bottomBack ? toScreenPoint(point.bottomBack) : null,
      categoryIndex: point.categoryIndex,
      defined: point.defined,
      top: toScreenPoint(point.top),
      topBack: point.topBack ? toScreenPoint(point.topBack) : null
    })),
    seriesIndex
  })).sort((left, right) => left.averageDepth - right.averageDepth);

  const edgeColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const gridColor = lightenColor(edgeColor, 0.56);
  const labelColor = resolveChartAxisTextColor(chart);
  const boxEdges: Array<[number, number, boolean]> = [
    [0, 1, false], [1, 2, false], [2, 3, false], [3, 0, false],
    [4, 5, true], [5, 6, true], [6, 7, true], [7, 4, true],
    [0, 4, false], [1, 5, true], [2, 6, true], [3, 7, false]
  ];
  const yTicks = valueDomain.ticks;
  const xLabelPoints = categories.map((category, categoryIndex) => ({
    label: category,
    point: toScreenPoint(projectCartesian3dPoint(
      normalizeX(categoryIndex),
      -1.06,
      frontZ,
      rotXRad,
      rotYRad,
      usePerspective,
      perspectiveStrength
    ))
  }));
  const yLabelPoints = yTicks.map((tick) => ({
    point: toScreenPoint(projectCartesian3dPoint(
      -1.04,
      normalizeY(tick),
      frontZ,
      rotXRad,
      rotYRad,
      usePerspective,
      perspectiveStrength
    )),
    tick
  }));

  return (
    <g>
      {yTicks.map((tick) => {
        const yNorm = normalizeY(tick);
        const leftBack = toScreenPoint(projectCartesian3dPoint(-1, yNorm, backZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        const rightBack = toScreenPoint(projectCartesian3dPoint(1, yNorm, backZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        const leftFront = toScreenPoint(projectCartesian3dPoint(-1, yNorm, frontZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        return (
          <g key={`line3d-grid-${tick}`}>
            <line stroke={gridColor} stroke-width={1} x1={leftBack.x} x2={rightBack.x} y1={leftBack.y} y2={rightBack.y} />
            <line stroke={gridColor} stroke-width={0.9} x1={leftBack.x} x2={leftFront.x} y1={leftBack.y} y2={leftFront.y} />
          </g>
        );
      })}
      {boxEdges.map(([startIndex, endIndex, emphasized], index) => {
        const start = screenCorners[startIndex];
        const end = screenCorners[endIndex];
        return (
          <line
            key={`line3d-box-${index}`}
            stroke={emphasized ? edgeColor : lightenColor(edgeColor, 0.34)}
            stroke-width={emphasized ? 1.35 : 1}
            x1={start.x}
            x2={end.x}
            y1={start.y}
            y2={end.y}
          />
        );
      })}
      {seriesGeometry.map(({ points, seriesIndex }) => {
        const definedPoints = points.filter((point) => point.defined);
        if (definedPoints.length === 0) {
          return null;
        }
        const linePoints = definedPoints.map((point) => point.top);
        const areaPoints = isAreaChart
          ? [
              ...definedPoints.map((point) => point.top),
              ...definedPoints.slice().reverse().map((point) => point.bottom)
            ]
          : [];
        const areaBackPoints = isAreaChart && isStackedSeries
          ? [
              ...definedPoints.map((point) => point.topBack ?? point.top),
              ...definedPoints.slice().reverse().map((point) => point.bottomBack ?? point.bottom)
            ]
          : [];
        const strokeColor = chartSeriesStrokeColor(chart, seriesIndex);
        const fillColor = chartSeriesColor(chart, seriesIndex);
        const markerSymbol = normalizeChartMarkerSymbol(chart.series[seriesIndex]?.markerSymbol);
        const markerPath = markerSymbolPath(markerSymbol, Math.max(4, chart.series[seriesIndex]?.markerSize ?? 6) * 0.52);
        const slabFaces = isAreaChart && isStackedSeries
          ? definedPoints.slice(1).map((point, pointIndex) => {
              const previous = definedPoints[pointIndex];
              if (!previous?.topBack || !point.topBack || !previous.bottomBack || !point.bottomBack) {
                return null;
              }
              const topFace = buildLinearSvgPath([previous.top, point.top, point.topBack, previous.topBack], true);
              const bottomFace = buildLinearSvgPath([previous.bottom, point.bottom, point.bottomBack, previous.bottomBack], true);
              return (
                <>
                  <path
                    {...chartElementDataProps(seriesIndex)}
                    d={topFace}
                    fill={lightenColor(fillColor, 0.08)}
                    fill-opacity={0.8}
                    stroke={darkenColor(fillColor, 0.16)}
                    stroke-width={0.8}
                  />
                  <path
                    {...chartElementDataProps(seriesIndex)}
                    d={bottomFace}
                    fill={darkenColor(fillColor, 0.2)}
                    fill-opacity={0.34}
                    stroke={darkenColor(fillColor, 0.24)}
                    stroke-width={0.6}
                  />
                </>
              );
            })
          : [];
        const firstDefinedPoint = definedPoints[0] ?? null;
        const lastDefinedPoint = definedPoints[definedPoints.length - 1] ?? null;
        const startCap = isAreaChart && isStackedSeries && firstDefinedPoint?.topBack && firstDefinedPoint?.bottomBack
          ? buildLinearSvgPath([
              firstDefinedPoint.top,
              firstDefinedPoint.bottom,
              firstDefinedPoint.bottomBack,
              firstDefinedPoint.topBack
            ], true)
          : "";
        const endCap = isAreaChart && isStackedSeries && lastDefinedPoint?.topBack && lastDefinedPoint?.bottomBack
          ? buildLinearSvgPath([
              lastDefinedPoint.top,
              lastDefinedPoint.bottom,
              lastDefinedPoint.bottomBack,
              lastDefinedPoint.topBack
            ], true)
          : "";

        return (
          <g key={`line3d-series-${seriesIndex}`}>
            {isAreaChart && isStackedSeries && areaBackPoints.length >= 3 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={buildLinearSvgPath(areaBackPoints, true)}
                fill={darkenColor(fillColor, 0.18)}
                fill-opacity={0.44}
                stroke={darkenColor(fillColor, 0.26)}
                stroke-width={0.8}
              />
            ) : null}
            {slabFaces}
            {startCap ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={startCap}
                fill={darkenColor(fillColor, 0.24)}
                fill-opacity={0.54}
                stroke={darkenColor(fillColor, 0.3)}
                stroke-width={0.7}
              />
            ) : null}
            {endCap ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={endCap}
                fill={darkenColor(fillColor, 0.14)}
                fill-opacity={0.6}
                stroke={darkenColor(fillColor, 0.24)}
                stroke-width={0.7}
              />
            ) : null}
            {isAreaChart && areaPoints.length >= 3 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={buildLinearSvgPath(areaPoints, true)}
                fill={fillColor}
                fill-opacity={0.74}
                stroke={darkenColor(fillColor, 0.12)}
                stroke-width={0.9}
              />
            ) : null}
            {!isAreaChart
              ? definedPoints.map((point, pointIndex) => (
                  <line
                    key={`line3d-drop-${seriesIndex}-${pointIndex}`}
                    stroke={lightenColor(strokeColor, 0.44)}
                    stroke-dasharray="2 2"
                    stroke-width={0.9}
                    x1={point.top.x}
                    x2={point.bottom.x}
                    y1={point.top.y}
                    y2={point.bottom.y}
                  />
                ))
              : null}
            <path
              {...chartElementDataProps(seriesIndex)}
              d={buildLinearSvgPath(linePoints)}
              fill="none"
              stroke={strokeColor}
              stroke-linejoin="round"
              stroke-width={Math.max(1.8, chart.series[seriesIndex]?.lineWidthPx ?? 2)}
            />
            {markerPath.length > 0
              ? definedPoints.map((point, pointIndex) => (
                  <g
                    key={`line3d-marker-${seriesIndex}-${pointIndex}`}
                    transform={`translate(${toSvgNumber(point.top.x)}, ${toSvgNumber(point.top.y)})`}
                  >
                    <path
                      {...chartElementDataProps(seriesIndex, point.categoryIndex)}
                      d={markerPath}
                      fill={chart.series[seriesIndex]?.markerColor ?? fillColor}
                      stroke={chart.series[seriesIndex]?.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
                      stroke-width={1}
                    />
                  </g>
                ))
              : null}
            {definedPoints.map((point) => (
              <circle
                {...chartElementDataProps(seriesIndex, point.categoryIndex)}
                cx={point.top.x}
                cy={point.top.y}
                fill="transparent"
                key={`line3d-hit-${seriesIndex}-${point.categoryIndex}`}
                r={7}
                stroke="none"
              />
            ))}
            {definedPoints.map((point) => (
              isSelectedChartPointOrSeries(selectedChartElement, chart.id, seriesIndex, point.categoryIndex)
                ? renderSelectionPointHandles(
                    `line3d-selection-${seriesIndex}-${point.categoryIndex}`,
                    [{ x: point.top.x, y: point.top.y }]
                  )
                : null
            ))}
          </g>
        );
      })}
      {yLabelPoints.map(({ point, tick }) => (
        <text
          key={`line3d-y-label-${tick}`}
          fill={labelColor}
          font-size={10}
          text-anchor="end"
          x={point.x - 6}
          y={point.y + 3}
        >
          {formatTickValue(tick)}
        </text>
      ))}
      {xLabelPoints.map(({ label, point }, index) => (
        <text
          key={`line3d-x-label-${index}`}
          fill={labelColor}
          font-size={10}
          text-anchor="middle"
          x={point.x}
          y={point.y + 14}
        >
          {label}
        </text>
      ))}
    </g>
  );
}
