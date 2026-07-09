import { geoIdentity, geoMercator, geoNaturalEarth1, geoPath } from "d3-geo";
import { scaleLinear } from "d3-scale";
import { feature as topojsonFeature } from "topojson-client";
import countiesAlbers10m from "us-atlas/counties-albers-10m.json";
import countries50m from "world-atlas/countries-50m.json";
import type { Feature, FeatureCollection, Geometry } from "geojson";
import type { XlsxChart, XlsxChartElementSelection, XlsxChartSeries, XlsxImageRect } from "@extend-ai/xlsx-core";
import { chartSeriesColor, chartSeriesStrokeColor, chartPointColor, safeNumber, clamp, darkenColor, lightenColor, coerceLooseNumber, resolveRenderableSeriesValue, resolveChartTextColor, resolveChartAxisTextColor, resolveChartMutedTextColor, buildChartFontFamily, formatTickValue, formatPercentTickValue, chartElementDataProps, barChartElementDataProps, resolveChartElementTarget, resolveChartSelectionFromTarget, isSelectedChartSeries, isSelectedChartPoint, isSelectedChartPointOrSeries, renderSelectionRectHandles, renderSelectionPointHandles, normalizeRendererHexColor, toSvgNumber, normalizeChartMarkerSymbol, markerSymbolPath, normalizeLegendPosition, parseRgbColor, mixRgbColor, escapeCssFontFamilyToken, DEFAULT_CHART_FONT_STACK, DEFAULT_CHART_TEXT_COLOR, DEFAULT_CHART_MUTED_TEXT_COLOR, truncateSvgText, getCategoryLabels, normalizeCategoryLabel, estimateReferencePointCount, isHistogramLikeSeries, isHistogramLikeChart, resolve3dFrameOffsets, resolveAxisDomainWithChartOverrides, resolveCategoryBandPadding, chartSeriesBarColors, buildNumericTickValues, buildNiceStep, resolveNumericAxisDomain, renderCartesianAxes, renderExtrudedRect, buildChartStages, resolveChartStageSubtotal, computeBoxWhiskerStats, resolveBoxWhiskerVisibility, resolveBoxWhiskerQuartileMethod, computePercentile, renderLineOrAreaChart3d, projectCartesian3dPoint, scaleProjectedVector, buildLinearSvgPath, buildRibbonSvgPath, renderRadialFrustum, sampleProjectedEllipseArc, isComboChart, getComboLegendSeries, findAxisForGroup, buildComboGroups, resolveStockRoleIndices, resolveStockPalette, normalizeBuiltinPieStyleId, getBuiltinPiePalette, resolvePiePointColor, selectPrimaryPieSeriesIndex, buildPieEntries, buildHierarchyData, resolveHierarchyNodeColor, resolveTreemapNodeColor, excelTreemapTile, normalizePieArc, resolvePieFrontSegments, pieEllipsePoint, buildPieOuterWallPath, isPieFrontFacingAngle, buildPieRadialWallPath, TWO_PI, normalizeRegionMapKey, resolveRegionMapFeature, resolveRegionMapBaseColor, resolveRegionMapDataColor, resolveRegionMapValueColors, resolveRegionMapColorStrings, resolveRegionMapLayoutProperties, resolveRegionMapFeatureSet, getRegionMapBaseFeatures, resolveRegionMapValueColorFromStops, resolveRegionMapValueColor, resolveRegionMapNoDataColor, resolveSurfaceBaseColor, normalizeBuiltinSurfaceStyleId, hasExplicitSurfaceBaseColor, buildMonochromeSurfacePalette, getBuiltinSurfacePalette, shouldPreferBuiltinSurfacePalette, getSurfaceBandCount, getSurfaceColorStops, resolveSurfaceBandPaletteColor, getSurfaceDomain, resolveSurfaceColor, resolveSurfaceBandColor, resolveSurfaceBandIndex, buildSurfaceSmoothPath, splitSurfacePointRuns, connectSurfaceContourSegments, getSurfaceWireframePalette, resolveSurfaceWireframeColor, resolveSurfacePlotRect, buildSurfaceLegendItems, isContourSurfaceChart, getLegendItems, buildLayout, normalizeRenderableChartType, renderSurfaceAxes, renderTitle, renderLegend, buildRegionMapLegendItems, buildStockLegendItems, excelSerialToDate, isLikelyDateFormatCode, formatExcelDateSerial, formatCategoryLabel } from "./chart-renderer";
import type { ChartRendererPalette, ChartSvgProps, ChartLayout, LegendItem, BarRect, PlotRect, ChartStage, ChartHierarchyDatum, ComboRenderableGroup, BoxWhiskerStats, SurfaceDomain, RegionMapFeature, SurfacePathPoint, SurfaceContourSegment, ChartElementDataOptions } from "./chart-renderer";


export function renderSurfaceChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const plot = resolveSurfacePlotRect(chart, layout);
  const categories = getCategoryLabels(chart);
  const rows = chart.series.length;
  const cols = Math.max(
    categories.length,
    chart.series.reduce((max, series) => Math.max(max, series.values.length), 0)
  );
  if (rows === 0 || cols === 0) {
    return null;
  }
  const matrix = chart.series.map((series) => (
    Array.from({ length: cols }, (_, columnIndex) => safeNumber(series.values[columnIndex]))
  ));
  const domain = getSurfaceDomain(chart);
  if (!domain) {
    return null;
  }
  const isContour = isContourSurfaceChart(chart);
  const minValue = domain.minValue;
  const safeMax = domain.safeMax;
  const wallFill = chart.backWall?.fillColor ?? "#d9d9df";
  const wallLineColor = chart.backWall?.lineColor ?? chart.sideWall?.lineColor ?? chart.floor?.lineColor ?? (chart.axisLineColor ?? lightenColor(resolveSurfaceBaseColor(chart, palette), 0.4));

  if (isContour) {
    const thresholds = domain.ticks.slice(1, -1);
    const quads: import("vue").VNode[] = [];
    const contourSegmentsByThreshold = new Map<number, SurfaceContourSegment[]>();

    for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
        const p00 = matrix[rowIndex]?.[columnIndex];
        const p10 = matrix[rowIndex]?.[columnIndex + 1];
        const p01 = matrix[rowIndex + 1]?.[columnIndex];
        const p11 = matrix[rowIndex + 1]?.[columnIndex + 1];
        if (p00 == null || p10 == null || p01 == null || p11 == null) {
          continue;
        }

        const x0 = plot.left + (columnIndex / Math.max(1, cols - 1)) * plot.width;
        const x1 = plot.left + ((columnIndex + 1) / Math.max(1, cols - 1)) * plot.width;
        const y0 = plot.top + plot.height - (rowIndex / Math.max(1, rows - 1)) * plot.height;
        const y1 = plot.top + plot.height - ((rowIndex + 1) / Math.max(1, rows - 1)) * plot.height;
        const averageValue = (p00 + p10 + p01 + p11) / 4;

        if (!chart.wireframe) {
          const primaryAvgA = (p00 + p10 + p11) / 3;
          const primaryAvgB = (p00 + p11 + p01) / 3;
          const secondaryAvgA = (p00 + p10 + p01) / 3;
          const secondaryAvgB = (p10 + p11 + p01) / 3;
          const primaryRange = (Math.max(p00, p10, p11) - Math.min(p00, p10, p11))
            + (Math.max(p00, p11, p01) - Math.min(p00, p11, p01));
          const secondaryRange = (Math.max(p00, p10, p01) - Math.min(p00, p10, p01))
            + (Math.max(p10, p11, p01) - Math.min(p10, p11, p01));
          const usePrimaryDiagonal = primaryRange <= secondaryRange;
          const triangles = usePrimaryDiagonal
            ? [
                {
                  bandColor: resolveSurfaceBandColor(chart, palette, domain, primaryAvgA),
                  bandIndex: resolveSurfaceBandIndex(domain, primaryAvgA),
                  points: `${x0},${y0} ${x1},${y0} ${x1},${y1}`
                },
                {
                  bandColor: resolveSurfaceBandColor(chart, palette, domain, primaryAvgB),
                  bandIndex: resolveSurfaceBandIndex(domain, primaryAvgB),
                  points: `${x0},${y0} ${x1},${y1} ${x0},${y1}`
                }
              ]
            : [
                {
                  bandColor: resolveSurfaceBandColor(chart, palette, domain, secondaryAvgA),
                  bandIndex: resolveSurfaceBandIndex(domain, secondaryAvgA),
                  points: `${x0},${y0} ${x1},${y0} ${x0},${y1}`
                },
                {
                  bandColor: resolveSurfaceBandColor(chart, palette, domain, secondaryAvgB),
                  bandIndex: resolveSurfaceBandIndex(domain, secondaryAvgB),
                  points: `${x1},${y0} ${x1},${y1} ${x0},${y1}`
                }
              ];
          const splitLine = usePrimaryDiagonal
            ? { x1: x0, y1: y0, x2: x1, y2: y1 }
            : { x1: x1, y1: y0, x2: x0, y2: y1 };
          const splitBands = triangles[0]?.bandIndex !== triangles[1]?.bandIndex;
          quads.push(
            <g key={`surface-contour-cell-${rowIndex}-${columnIndex}`}>
              {splitBands ? (
                <>
                  <polygon fill={triangles[0]?.bandColor} points={triangles[0]?.points} stroke="none" />
                  <polygon fill={triangles[1]?.bandColor} points={triangles[1]?.points} stroke="none" />
                  <line
                    stroke={darkenColor(resolveSurfaceBandPaletteColor(chart, palette, domain, averageValue), 0.22)}
                    stroke-width={0.8}
                    x1={splitLine.x1}
                    x2={splitLine.x2}
                    y1={splitLine.y1}
                    y2={splitLine.y2}
                  />
                </>
              ) : (
                <rect
                  fill={resolveSurfaceBandColor(chart, palette, domain, averageValue)}
                  height={Math.abs(y1 - y0)}
                  stroke="none"
                  width={Math.abs(x1 - x0)}
                  x={Math.min(x0, x1)}
                  y={Math.min(y0, y1)}
                />
              )}
            </g>
          );
        }

        const corners = [
          { value: p00, x: x0, y: y0 },
          { value: p10, x: x1, y: y0 },
          { value: p11, x: x1, y: y1 },
          { value: p01, x: x0, y: y1 }
        ];
        const edges: Array<[typeof corners[number], typeof corners[number]]> = [
          [corners[0], corners[1]],
          [corners[1], corners[2]],
          [corners[2], corners[3]],
          [corners[3], corners[0]]
        ];
        thresholds.forEach((threshold) => {
          const intersections: Array<{ x: number; y: number }> = [];
          edges.forEach(([start, end]) => {
            const delta = end.value - start.value;
            if (delta === 0) {
              return;
            }
            const crosses = (start.value < threshold && end.value > threshold) || (start.value > threshold && end.value < threshold);
            if (!crosses) {
              return;
            }
            const mix = (threshold - start.value) / delta;
            intersections.push({
              x: start.x + (end.x - start.x) * mix,
              y: start.y + (end.y - start.y) * mix
            });
          });
          if (intersections.length === 2) {
            const segments = contourSegmentsByThreshold.get(threshold) ?? [];
            segments.push({
              end: intersections[1] ?? intersections[0] ?? { x: x1, y: y1 },
              start: intersections[0] ?? { x: x0, y: y0 }
            });
            contourSegmentsByThreshold.set(threshold, segments);
          } else if (intersections.length === 4) {
            const center = averageValue;
            const pairings = center >= threshold
              ? [[0, 1], [2, 3]]
              : [[0, 3], [1, 2]];
            pairings.forEach(([startIndex, endIndex], pairingIndex) => {
              const leftPoint = intersections[startIndex] ?? intersections[0];
              const rightPoint = intersections[endIndex] ?? intersections[intersections.length - 1];
              if (!leftPoint || !rightPoint) {
                return;
              }
              const segments = contourSegmentsByThreshold.get(threshold) ?? [];
              segments.push({
                end: rightPoint,
                start: leftPoint
              });
              contourSegmentsByThreshold.set(threshold, segments);
            });
          }
        });
      }
    }

    const contourLines = Array.from(contourSegmentsByThreshold.entries()).flatMap(([threshold, segments], thresholdIndex) => (
      connectSurfaceContourSegments(segments).map((points, pathIndex) => {
        const path = buildSurfaceSmoothPath(points, true);
        if (!path) {
          return null;
        }
        return (
          <path
            key={`surface-contour-line-${thresholdIndex}-${pathIndex}`}
            d={path}
            fill="none"
            stroke={chart.wireframe
              ? resolveSurfaceWireframeColor(chart, palette, domain, threshold)
              : darkenColor(resolveSurfaceBandColor(chart, palette, domain, threshold), 0.18)}
            stroke-linecap="round"
            stroke-linejoin="round"
            stroke-width={chart.wireframe ? 1.45 : 1.2}
          />
        );
      })
    ));

    const columnAverages = Array.from({ length: cols }, (_, columnIndex) => {
      const values = Array.from({ length: rows }, (_, rowIndex) => matrix[rowIndex]?.[columnIndex])
        .filter((value): value is number => value != null);
      if (values.length === 0) {
        return domain.minValue;
      }
      return values.reduce((sum, value) => sum + value, 0) / values.length;
    });
    const rowAverages = Array.from({ length: rows }, (_, rowIndex) => {
      const values = matrix[rowIndex]?.filter((value): value is number => value != null) ?? [];
      if (values.length === 0) {
        return domain.minValue;
      }
      return values.reduce((sum, value) => sum + value, 0) / values.length;
    });

    return (
      <g>
        <rect
          fill={wallFill}
          height={plot.height}
          stroke={lightenColor(wallLineColor, 0.14)}
          stroke-width={0.8}
          width={plot.width}
          x={plot.left}
          y={plot.top}
        />
        {Array.from({ length: cols }, (_, columnIndex) => {
          const x = plot.left + (cols <= 1 ? plot.width / 2 : (columnIndex / (cols - 1)) * plot.width);
          return (
            <line
              key={`surface-contour-grid-col-${columnIndex}`}
              stroke={chart.wireframe
                ? resolveSurfaceWireframeColor(chart, palette, domain, columnAverages[columnIndex] ?? domain.minValue)
                : lightenColor(wallLineColor, 0.18)}
              stroke-linecap="round"
              stroke-linejoin="round"
              stroke-width={chart.wireframe ? 1.15 : 0.8}
              x1={x}
              x2={x}
              y1={plot.top}
              y2={plot.top + plot.height}
            />
          );
        })}
        {Array.from({ length: rows }, (_, rowIndex) => {
          const y = plot.top + plot.height - (rows <= 1 ? plot.height / 2 : (rowIndex / (rows - 1)) * plot.height);
          const rowStroke = chart.series[rowIndex]?.lineColor
            ?? chart.series[rowIndex]?.color
            ?? resolveSurfaceWireframeColor(chart, palette, domain, rowAverages[rowIndex] ?? domain.minValue);
          return (
            <line
              key={`surface-contour-grid-row-${rowIndex}`}
              stroke={chart.wireframe ? rowStroke : lightenColor(wallLineColor, 0.18)}
              stroke-linecap="round"
              stroke-linejoin="round"
              stroke-width={chart.wireframe ? 1.15 : 0.8}
              x1={plot.left}
              x2={plot.left + plot.width}
              y1={y}
              y2={y}
            />
          );
        })}
        {quads}
        {contourLines}
        {chart.wireframe ? null : (
          <rect
            fill="none"
            height={plot.height}
            stroke={lightenColor(wallLineColor, 0.1)}
            stroke-width={0.8}
            width={plot.width}
            x={plot.left}
            y={plot.top}
          />
        )}
      </g>
    );
  }

  const rotX = clamp(chart.view3d?.rotX ?? (chart.wireframe ? 90 : 25), -88, 88) * (Math.PI / 180);
  const rotY = clamp(chart.view3d?.rotY ?? (chart.wireframe ? 0 : 30), -88, 88) * (Math.PI / 180);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp(
    (chart.view3d?.perspective ?? (usePerspective ? 30 : 0)) / 100,
    0,
    1
  );
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.2, 4);
  const cosX = Math.cos(rotX);
  const sinX = Math.sin(rotX);
  const cosY = Math.cos(rotY);
  const sinY = Math.sin(rotY);

  type SurfacePoint = { depth: number; hasValue: boolean; value: number; x: number; y: number };
  const rawPoints: SurfacePoint[][] = Array.from({ length: rows }, (_, rowIndex) => (
    Array.from({ length: cols }, (_, columnIndex) => {
      const value = matrix[rowIndex][columnIndex];
      const hasValue = value != null;
      const normalizedX = cols <= 1 ? 0 : ((columnIndex / (cols - 1)) - 0.5) * 2;
      const normalizedY = hasValue ? (-((((value - minValue) / (safeMax - minValue)) - 0.5) * 1.8) * depthScale) : (0.9 * depthScale);
      const normalizedZ = rows <= 1 ? 0 : ((rowIndex / (rows - 1)) - 0.5) * 2;

      const x1 = normalizedX * cosY + normalizedZ * sinY;
      const z1 = -normalizedX * sinY + normalizedZ * cosY;
      const y1 = normalizedY * cosX - z1 * sinX;
      const z2 = normalizedY * sinX + z1 * cosX;
      const perspective = usePerspective
        ? 1 / Math.max(0.18, 1 + z2 * (0.24 + perspectiveStrength * 0.5))
        : 1;

      return {
        depth: z2,
        hasValue,
        value: value ?? minValue,
        x: x1 * perspective,
        y: y1 * perspective
      };
    })
  ));

  const bounds = rawPoints.flat();
  const minX = Math.min(...bounds.map((point) => point.x));
  const maxX = Math.max(...bounds.map((point) => point.x));
  const minY = Math.min(...bounds.map((point) => point.y));
  const maxY = Math.max(...bounds.map((point) => point.y));
  const scale = Math.min(
    layout.plot.width / Math.max(0.25, maxX - minX),
    layout.plot.height / Math.max(0.25, maxY - minY)
  ) * 0.82;
  const centerX = layout.plot.left + layout.plot.width / 2;
  const centerY = layout.plot.top + layout.plot.height / 2;
  const centerRawX = (minX + maxX) / 2;
  const centerRawY = (minY + maxY) / 2;

  const points = rawPoints.map((row) => row.map((point) => ({
    ...point,
    x: centerX + (point.x - centerRawX) * scale,
    y: centerY + (point.y - centerRawY) * scale
  })));

  const baseColor = resolveSurfaceBaseColor(chart, palette);
  const axisColor = chart.axisLineColor ?? lightenColor(baseColor, 0.4);
  type Quad = {
    color: string;
    depth: number;
    key: string;
    points: string;
    stroke: string;
  };
  const quads: Quad[] = [];
  for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
      const p00 = points[rowIndex][columnIndex];
      const p10 = points[rowIndex][columnIndex + 1];
      const p11 = points[rowIndex + 1][columnIndex + 1];
      const p01 = points[rowIndex + 1][columnIndex];
      if (!p00.hasValue || !p10.hasValue || !p11.hasValue || !p01.hasValue) {
        continue;
      }
      const averageValue = (p00.value + p10.value + p11.value + p01.value) / 4;
      const ratio = clamp((averageValue - minValue) / (safeMax - minValue), 0, 1);
      quads.push({
        color: resolveSurfaceColor(chart, palette, ratio),
        depth: (p00.depth + p10.depth + p11.depth + p01.depth) / 4,
        key: `surface-quad-${rowIndex}-${columnIndex}`,
        points: `${p00.x},${p00.y} ${p10.x},${p10.y} ${p11.x},${p11.y} ${p01.x},${p01.y}`,
        stroke: lightenColor(baseColor, 0.18)
      });
    }
  }
  quads.sort((left, right) => left.depth - right.depth);

  const surfaceContourLines = !isContour
    ? (() => {
        const contourSegmentsByThreshold = new Map<number, SurfaceContourSegment[]>();
        const thresholds = domain.ticks.slice(1, -1);
        for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
          for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
            const topLeft = points[rowIndex][columnIndex];
            const topRight = points[rowIndex][columnIndex + 1];
            const bottomRight = points[rowIndex + 1][columnIndex + 1];
            const bottomLeft = points[rowIndex + 1][columnIndex];
            if (!topLeft.hasValue || !topRight.hasValue || !bottomRight.hasValue || !bottomLeft.hasValue) {
              continue;
            }
            const corners = [
              { point: topLeft, value: topLeft.value },
              { point: topRight, value: topRight.value },
              { point: bottomRight, value: bottomRight.value },
              { point: bottomLeft, value: bottomLeft.value }
            ];
            const edges: Array<[typeof corners[number], typeof corners[number]]> = [
              [corners[0], corners[1]],
              [corners[1], corners[2]],
              [corners[2], corners[3]],
              [corners[3], corners[0]]
            ];
            thresholds.forEach((threshold) => {
              const intersections: SurfacePathPoint[] = [];
              edges.forEach(([start, end]) => {
                const delta = end.value - start.value;
                if (delta === 0) {
                  return;
                }
                const crosses = (start.value < threshold && end.value > threshold) || (start.value > threshold && end.value < threshold);
                if (!crosses) {
                  return;
                }
                const ratio = (threshold - start.value) / delta;
                intersections.push({
                  x: start.point.x + (end.point.x - start.point.x) * ratio,
                  y: start.point.y + (end.point.y - start.point.y) * ratio
                });
              });
              if (intersections.length === 2) {
                const segments = contourSegmentsByThreshold.get(threshold) ?? [];
                segments.push({
                  end: intersections[1] ?? intersections[0] ?? { x: topRight.x, y: topRight.y },
                  start: intersections[0] ?? { x: topLeft.x, y: topLeft.y }
                });
                contourSegmentsByThreshold.set(threshold, segments);
              } else if (intersections.length === 4) {
                const center = (topLeft.value + topRight.value + bottomRight.value + bottomLeft.value) / 4;
                const pairings = center >= threshold
                  ? [[0, 1], [2, 3]]
                  : [[0, 3], [1, 2]];
                pairings.forEach(([startIndex, endIndex]) => {
                  const startPoint = intersections[startIndex] ?? intersections[0];
                  const endPoint = intersections[endIndex] ?? intersections[intersections.length - 1];
                  if (!startPoint || !endPoint) {
                    return;
                  }
                  const segments = contourSegmentsByThreshold.get(threshold) ?? [];
                  segments.push({
                    end: endPoint,
                    start: startPoint
                  });
                  contourSegmentsByThreshold.set(threshold, segments);
                });
              }
            });
          }
        }
        return Array.from(contourSegmentsByThreshold.entries()).flatMap(([threshold, segments], thresholdIndex) => (
          connectSurfaceContourSegments(segments).map((pathPoints, pathIndex) => {
            const path = buildSurfaceSmoothPath(pathPoints, true);
            if (!path) {
              return null;
            }
            return (
              <path
                key={`surface-level-line-${thresholdIndex}-${pathIndex}`}
                d={path}
                fill="none"
                stroke={darkenColor(resolveSurfaceBandPaletteColor(chart, palette, domain, threshold), 0.26)}
                stroke-linecap="round"
                stroke-linejoin="round"
                stroke-width={1.55}
              />
            );
          })
        ));
      })()
    : [];

  const buildSurfacePathRuns = (
    segmentPoints: Array<{ depth: number; hasValue: boolean; value: number; x: number; y: number }>,
    keyPrefix: string,
    strokeColor: string,
    strokeWidth: number
  ) => (
    splitSurfacePointRuns(segmentPoints).map((run, runIndex) => {
      const path = buildSurfaceSmoothPath(run, chart.wireframe === true);
      if (!path) {
        return null;
      }
      return (
        <path
          key={`${keyPrefix}-${runIndex}`}
          d={path}
          fill="none"
          stroke={strokeColor}
          stroke-linecap="round"
          stroke-linejoin="round"
          stroke-width={strokeWidth}
        />
      );
    })
  );

  return (
    <g>
      {isContour ? (
        <>
          <rect
            fill={wallFill}
            height={layout.plot.height}
            stroke={lightenColor(axisColor, 0.18)}
            stroke-width={0.8}
            width={layout.plot.width}
            x={layout.plot.left}
            y={layout.plot.top}
          />
          {Array.from({ length: cols }, (_, columnIndex) => {
            const x = layout.plot.left + (cols <= 1 ? layout.plot.width / 2 : (columnIndex / (cols - 1)) * layout.plot.width);
            return (
              <line
                key={`surface-fallback-column-grid-${columnIndex}`}
                stroke={lightenColor(wallLineColor, 0.12)}
                stroke-width={0.8}
                x1={x}
                x2={x}
                y1={layout.plot.top}
                y2={layout.plot.top + layout.plot.height}
              />
            );
          })}
          {Array.from({ length: rows }, (_, rowIndex) => {
            const y = layout.plot.top + layout.plot.height - (rows <= 1 ? layout.plot.height / 2 : (rowIndex / (rows - 1)) * layout.plot.height);
            return (
              <line
                key={`surface-fallback-row-grid-${rowIndex}`}
                stroke={lightenColor(wallLineColor, 0.12)}
                stroke-width={0.8}
                x1={layout.plot.left}
                x2={layout.plot.left + layout.plot.width}
                y1={y}
                y2={y}
              />
            );
          })}
        </>
      ) : null}
      {chart.wireframe ? null : quads.map((quad) => (
        <polygon
          key={quad.key}
          fill={quad.color}
          fill-opacity={0.95}
          points={quad.points}
          stroke={quad.stroke}
          stroke-width={0.7}
        />
      ))}
      {Array.from({ length: rows }, (_, rowIndex) => {
        const rowPoints = points[rowIndex];
        const averageValue = matrix[rowIndex]
          ?.filter((value): value is number => value != null)
          .reduce((sum, value, _, values) => sum + value / Math.max(1, values.length), 0) ?? domain.minValue;
        const rowStroke = chart.wireframe
          ? (chart.series[rowIndex]?.lineColor
              ?? chart.series[rowIndex]?.color
              ?? resolveSurfaceWireframeColor(chart, palette, domain, averageValue))
          : darkenColor(resolveSurfaceBandColor(chart, palette, domain, averageValue), 0.1);
        return (
          <g key={`surface-row-${rowIndex}`}>
            {buildSurfacePathRuns(
              rowPoints,
              `surface-row-${rowIndex}`,
              rowStroke,
              chart.wireframe ? Math.max(1.45, chart.series[rowIndex]?.lineWidthPx ?? 1.8) : 0.8
            )}
          </g>
        );
      })}
      {Array.from({ length: cols }, (_, columnIndex) => {
        const columnPoints = Array.from({ length: rows }, (_, rowIndex) => points[rowIndex][columnIndex]);
        const columnValues = Array.from({ length: rows }, (_, rowIndex) => matrix[rowIndex]?.[columnIndex])
          .filter((value): value is number => value != null);
        const columnAverage = columnValues.length > 0
          ? columnValues.reduce((sum, value) => sum + value, 0) / columnValues.length
          : domain.minValue;
        const columnStroke = chart.wireframe
          ? resolveSurfaceWireframeColor(chart, palette, domain, columnAverage)
          : darkenColor(resolveSurfaceBandColor(chart, palette, domain, columnAverage), 0.1);
        return (
          <g key={`surface-column-${columnIndex}`}>
            {buildSurfacePathRuns(
              columnPoints,
              `surface-column-${columnIndex}`,
              columnStroke,
              chart.wireframe ? 1.35 : 0.8
            )}
          </g>
        );
      })}
      {surfaceContourLines}
      {chart.wireframe ? (
        <rect
          fill="none"
          height={layout.plot.height}
          stroke={lightenColor(axisColor, 0.18)}
          stroke-width={0.8}
          width={layout.plot.width}
          x={layout.plot.left}
          y={layout.plot.top}
        />
      ) : null}
    </g>
  );
}

export function renderSurfaceHitOverlay(
  chart: XlsxChart,
  layout: ChartLayout,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  const plot = resolveSurfacePlotRect(chart, layout);
  const categories = getCategoryLabels(chart);
  const rows = chart.series.length;
  const cols = Math.max(
    categories.length,
    chart.series.reduce((max, series) => Math.max(max, series.values.length), 0)
  );
  if (rows === 0 || cols === 0) {
    return null;
  }

  const cellWidth = plot.width / cols;
  const cellHeight = plot.height / rows;
  return (
    <g>
      {chart.series.flatMap((_, seriesIndex) => (
        Array.from({ length: cols }, (_, pointIndex) => {
          const left = plot.left + pointIndex * cellWidth;
          const top = plot.top + seriesIndex * cellHeight;
          const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, seriesIndex, pointIndex);
          return (
            <>
              <rect
                {...chartElementDataProps(seriesIndex, pointIndex)}
                fill="transparent"
                height={cellHeight}
                stroke="none"
                width={cellWidth}
                x={left}
                y={top}
              />
              {selected
                ? renderSelectionRectHandles(
                    `surface-selection-${seriesIndex}-${pointIndex}`,
                    left,
                    top,
                    cellWidth,
                    cellHeight
                  )
                : null}
            </>
          );
        })
      ))}
    </g>
  );
}

export function renderRegionMapChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const primarySeriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
  const primarySeries = chart.series[primarySeriesIndex] ?? null;
  if (!primarySeries) {
    return null;
  }

  const labels = getCategoryLabels(chart);
  const layoutProperties = resolveRegionMapLayoutProperties(primarySeries);
  const geography = layoutProperties?.geography && typeof layoutProperties.geography === "object"
    ? layoutProperties.geography as Record<string, unknown>
    : null;
  const featureSet = resolveRegionMapFeatureSet(labels, geography);
  const viewedRegionType = typeof geography?.viewedRegionType === "string"
    ? geography.viewedRegionType
    : null;
  const projectionType = typeof geography?.projectionType === "string"
    ? geography.projectionType
    : null;
  const projection = featureSet === "us-state"
    ? geoIdentity()
    : projectionType === "miller"
      ? geoMercator()
      : geoNaturalEarth1();
  const availableFeatures = getRegionMapBaseFeatures(featureSet);
  const categoricalValues = resolveRegionMapColorStrings(primarySeries);
  const categoricalColorByLabel = categoricalValues
    ? new Map(
        Array.from(new Set(categoricalValues)).map((value, index) => [
          value,
          chartPointColor(chart, index, primarySeriesIndex)
        ])
      )
    : null;
  const entries = labels.map((label, index) => {
    const feature = resolveRegionMapFeature(label, featureSet);
    const value = safeNumber(primarySeries.values[index]);
    const colorLabel = categoricalValues?.[index] ?? null;
    return {
      colorLabel,
      feature,
      key: normalizeCategoryLabel(label),
      label: normalizeCategoryLabel(label),
      value
    };
  }).filter((entry) => entry.feature != null && (entry.value != null || entry.colorLabel != null)) as Array<{
    colorLabel: string | null;
    feature: RegionMapFeature;
    key: string;
    label: string;
    value: number | null;
  }>;
  const fitToMatchedData = entries.length > 0 && (viewedRegionType === "dataOnly" || entries.length === 1);
  const fitFeatures = fitToMatchedData
    ? entries.map((entry) => entry.feature)
    : availableFeatures;
  projection.fitExtent(
    [
      [layout.plot.left + 8, layout.plot.top + 8],
      [layout.plot.left + layout.plot.width - 8, layout.plot.top + layout.plot.height - 8]
    ],
    {
      type: "FeatureCollection",
      features: fitFeatures
    } satisfies FeatureCollection<Geometry, { name?: string }>
  );
  const path = geoPath(projection);
  const baseFeatures = fitToMatchedData
    ? fitFeatures
    : availableFeatures;
  const noDataFill = resolveRegionMapNoDataColor(chart, primarySeriesIndex);
  const outlineColor = chart.chartAreaBorderColor && chart.chartAreaBorderColor !== "transparent"
    ? chart.chartAreaBorderColor
    : darkenColor(noDataFill, 0.22);
  const numericEntryValues = entries
    .map((entry) => entry.value)
    .filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  const minValue = numericEntryValues.length > 0 ? Math.min(...numericEntryValues) : 0;
  const maxValue = numericEntryValues.length > 0 ? Math.max(...numericEntryValues) : 1;
  const showRegionLabels = layoutProperties?.regionLabelLayout !== "none";

  return (
    <g>
      {baseFeatures.map((feature, index) => {
        const d = path(feature);
        if (!d) {
          return null;
        }
        return (
          <path
            d={d}
            fill={noDataFill}
            key={`region-map-base-${feature.id ?? index}`}
            stroke={outlineColor}
            stroke-linejoin="round"
            stroke-width={0.6}
          />
        );
      })}
      {entries.map((entry, index) => {
        const d = path(entry.feature);
        if (!d) {
          return null;
        }
        const fill = entry.value != null
          ? (() => {
              const ratio = maxValue <= minValue ? 1 : (entry.value - minValue) / Math.max(1e-6, maxValue - minValue);
              return resolveRegionMapValueColor(chart, primarySeriesIndex, ratio);
            })()
          : categoricalColorByLabel?.get(entry.colorLabel ?? "") ?? resolveRegionMapDataColor(chart, primarySeriesIndex);
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, primarySeriesIndex, index);
        const centroid = path.centroid(entry.feature);
        const canShowSelectionDot = Number.isFinite(centroid[0]) && Number.isFinite(centroid[1]);
        return (
          <>
            <path
              {...chartElementDataProps(primarySeriesIndex, index)}
              d={d}
              fill={fill}
              stroke={selected ? "#64748b" : darkenColor(fill, 0.18)}
              stroke-linejoin="round"
              stroke-width={selected ? 1.8 : 0.85}
            />
            {selected && canShowSelectionDot
              ? renderSelectionPointHandles(`region-map-selection-${index}`, [{ x: centroid[0], y: centroid[1] }])
              : null}
          </>
        );
      })}
      {showRegionLabels
        ? entries.map((entry, index) => {
            const bounds = path.bounds(entry.feature);
            const [[x0, y0], [x1, y1]] = bounds;
            const width = x1 - x0;
            const height = y1 - y0;
            if (!Number.isFinite(width) || !Number.isFinite(height) || width < 26 || height < 12) {
              return null;
            }
            const centroid = path.centroid(entry.feature);
            if (!Number.isFinite(centroid[0]) || !Number.isFinite(centroid[1])) {
              return null;
            }
            return (
              <text
                fill={resolveChartTextColor(chart)}
                font-size={9}
                font-weight={600}
                key={`region-map-label-${entry.key || index}`}
                text-anchor="middle"
                x={centroid[0]}
                y={centroid[1]}
              >
                {truncateSvgText(entry.label, Math.max(28, width - 4), 9)}
              </text>
            );
          })
        : null}
    </g>
  );
}
