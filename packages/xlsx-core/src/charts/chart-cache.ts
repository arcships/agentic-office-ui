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
  const readLevelPoints = (levelNode: Element) => (
    getLocalChildren(levelNode, "pt")
      .map((pointNode) => {
        const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
        return {
          index: Number.isFinite(rawIndex) ? rawIndex : 0,
          value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
        };
      })
      .sort((left, right) => left.index - right.index)
  );
  const levelPointNodes = levelNodes.map(readLevelPoints).filter((points) => points.length > 0);
  if (levelPointNodes.length === 0) {
    return null;
  }
  const pointNodes = mode === "value"
    ? levelPointNodes[0] ?? []
    : levelPointNodes.flatMap((points) => points);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    ...levelPointNodes.map((points) => points.length),
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  if (mode === "value") {
    for (const point of pointNodes) {
      values[point.index] = cellValueToNumber(point.value);
    }
    return values;
  }

  const levelMaps = levelPointNodes.map((points) => new Map(points.map((point) => [point.index, point.value])));
  for (let index = 0; index < targetLength; index += 1) {
    const labels = levelMaps
      .map((level) => level.get(index) ?? "")
      .filter((value) => value.length > 0);
    values[index] = labels.length > 0 ? labels.join(" / ") : null;
  }
  return values;
}
