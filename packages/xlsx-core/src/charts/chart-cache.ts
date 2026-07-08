import {
  getLocalChildren,
  getFirstLocalChild,
  readChartNumericAttribute,
} from "./chart-xml-utils";

import {
  cellValueToNumber,
} from "./chart-series";

export function parseChartCacheValues(
  parentNode: Element | null,
  cacheName: "numCache" | "strCache",
  mode: "category" | "value"
): Array<number | string | null> | null {
  if (!parentNode) {
    return null;
  }

  const referenceNode = getFirstLocalChild(parentNode, "numRef")
    ?? getFirstLocalChild(parentNode, "strRef")
    ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, cacheName);
  if (!cacheNode) {
    return null;
  }

  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const pointNodes = getLocalChildren(cacheNode, "pt")
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
      };
    })
    .sort((left, right) => left.index - right.index);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
    } else {
      values[point.index] = point.value.length > 0 ? point.value : null;
    }
  }
  return values;
}

export function parseChartMultiLevelCacheValues(
  parentNode: Element | null,
  mode: "category" | "value"
): Array<number | string | null> | null {
  if (!parentNode) {
    return null;
  }

  const referenceNode = getFirstLocalChild(parentNode, "multiLvlStrRef") ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, "multiLvlStrCache");
  if (!cacheNode) {
    return null;
  }

  const levelNodes = getLocalChildren(cacheNode, "lvl");
  if (levelNodes.length === 0) {
    return null;
  }

  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const primaryLevelNode = mode === "category"
    ? levelNodes[levelNodes.length - 1] ?? levelNodes[0]
    : levelNodes[0];
  const pointNodes = getLocalChildren(primaryLevelNode, "pt")
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
      };
    })
    .sort((left, right) => left.index - right.index);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
      continue;
    }
    values[point.index] = point.value.length > 0 ? point.value : null;
  }
  return values;
}
