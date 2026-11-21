// Stage 2–4 object index for XLSX.
//
// Providers are registered, not hard-coded: a host can add its own object
// family (a pivot table, a slicer) without touching the reference contract.
// The defaults cover the product matrix rows this batch owns — tables,
// charts, images, shapes, form controls, comments, chart internals and the
// invisible workbook behaviors.

import type {
  OfficeObjectCandidate,
  OfficeObjectDescriptor,
  OfficeObjectKind,
  OfficeObjectProvider,
  OfficeObjectRegistry,
} from "@arcships/office-interaction";
import {
  createOfficeObjectDescriptor,
  createOfficeObjectRegistry,
  officeObjectReliability,
  rankOfficeObjectCandidates,
} from "@arcships/office-interaction";
import type {
  XlsxChart,
} from "./types";
import type {
  XlsxBehaviorRecord,
  XlsxBehaviorSource,
} from "./xlsx-behaviors";
import { createXlsxBehaviorRecords } from "./xlsx-behaviors";
import type {
  XlsxChartElementTarget,
  XlsxReferenceComment,
  XlsxReferenceContext,
  XlsxReferenceDrawing,
  XlsxReferenceSheet,
  XlsxReferenceTable,
} from "./reference-adapter";
import {
  describeXlsxBehavior,
  describeXlsxChart,
  describeXlsxChartElement,
  describeXlsxComment,
  describeXlsxDrawing,
  describeXlsxTable,
  describeXlsxTableCell,
  describeXlsxTableRow,
} from "./reference-adapter";
import { parseA1RangeReference } from "./images/image-parser";

export const XLSX_STAGE2_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "table",
  "chart",
  "image",
  "shape",
  "form-control",
  "comment",
]);

export const XLSX_STAGE3_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "table-row",
  "table-cell",
  "chart-series",
  "chart-point",
  "chart-legend-entry",
]);

export const XLSX_STAGE4_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "formula",
  "conditional-format",
  "data-validation",
  "named-range",
  "protection",
]);

export function xlsxChartObjectId(sheet: XlsxReferenceSheet, chart: XlsxChart): string {
  return `xlsx:chart:${sheet.id ?? sheet.name}:${chart.id}`;
}

function chartElementObjectId(
  sheet: XlsxReferenceSheet,
  chart: XlsxChart,
  target: XlsxChartElementTarget,
): string {
  const key = sheet.id ?? sheet.name;
  if (target.element === "series") {
    return `xlsx:chart-element:${key}:${chart.id}:series:${target.seriesIndex}`;
  }
  if (target.element === "point") {
    return `xlsx:chart-element:${key}:${chart.id}:point:${target.pointIndex}`;
  }
  return `xlsx:chart-element:${key}:${chart.id}:legend-entry:${target.legendIndex}`;
}

function chartElementParent(
  descriptor: OfficeObjectDescriptor,
  chartObjectId: string,
): OfficeObjectDescriptor {
  return createOfficeObjectDescriptor({
    objectId: descriptor.objectId,
    draft: descriptor.draft,
    label: descriptor.label,
    path: descriptor.path,
    parentObjectId: chartObjectId,
    childrenState: descriptor.childrenState,
    ...(descriptor.content === undefined ? {} : { content: descriptor.content }),
    ...(descriptor.visual === undefined ? {} : { visual: descriptor.visual }),
    ...(descriptor.facets === undefined ? {} : { facets: descriptor.facets }),
  });
}

function chartElementTargets(chart: XlsxChart): readonly XlsxChartElementTarget[] {
  const targets: XlsxChartElementTarget[] = [];
  chart.series.forEach((series, seriesIndex) => {
    targets.push({ element: "series", seriesIndex });
    const points = Math.min(series.values.length, 50);
    for (let pointIndex = 0; pointIndex < points; pointIndex += 1) {
      targets.push({ element: "point", seriesIndex, pointIndex });
    }
  });
  return targets;
}

export const XLSX_TABLE_OBJECT_PROVIDER: OfficeObjectProvider<XlsxReferenceContext> = {
  id: "xlsx.tables",
  format: "xlsx",
  kinds: ["table", "table-row", "table-cell"],
  list(context, options) {
    const descriptors: OfficeObjectDescriptor[] = [];
    const maxObjects = options?.maxObjects ?? Number.POSITIVE_INFINITY;
    for (const table of context.tables ?? []) {
      if (descriptors.length >= maxObjects) break;
      try {
        descriptors.push(describeXlsxTable(context, table));
      } catch {
        continue;
      }
      const range = parseA1RangeReference(table.a1);
      if (!range) continue;
      const rowCount = range.end.row - range.start.row + 1;
      const columnCount = range.end.col - range.start.col + 1;
      for (let rowIndex = 0; rowIndex < rowCount && descriptors.length < maxObjects; rowIndex += 1) {
        try {
          descriptors.push(describeXlsxTableRow(context, table, rowIndex));
        } catch {
          continue;
        }
        for (let cellIndex = 0; cellIndex < columnCount && descriptors.length < maxObjects; cellIndex += 1) {
          try {
            descriptors.push(describeXlsxTableCell(context, table, rowIndex, cellIndex));
          } catch {
            // A cell outside the current table range is simply not an object;
            // one bad index must not drop the rest of the table.
          }
        }
      }
    }
    return descriptors;
  },
};

export const XLSX_DRAWING_OBJECT_PROVIDER: OfficeObjectProvider<XlsxReferenceContext> = {
  id: "xlsx.drawings",
  format: "xlsx",
  kinds: ["image", "shape", "form-control"],
  list(context) {
    return (context.drawings ?? []).flatMap((drawing) => {
      try {
        return [describeXlsxDrawing(context, drawing)];
      } catch {
        return [];
      }
    });
  },
};

export const XLSX_COMMENT_OBJECT_PROVIDER: OfficeObjectProvider<XlsxReferenceContext> = {
  id: "xlsx.comments",
  format: "xlsx",
  kinds: ["comment"],
  list(context) {
    return (context.comments ?? []).flatMap((comment) => {
      try {
        return [describeXlsxComment(context, comment)];
      } catch {
        return [];
      }
    });
  },
};

export const XLSX_CHART_OBJECT_PROVIDER: OfficeObjectProvider<XlsxReferenceContext> = {
  id: "xlsx.charts",
  format: "xlsx",
  kinds: ["chart", "chart-series", "chart-point", "chart-legend-entry"],
  list(context) {
    const descriptors: OfficeObjectDescriptor[] = [];
    for (const chart of context.charts ?? []) {
      const sheet = context.sheets.find((candidate) =>
        candidate.workbookSheetIndex === chart.workbookSheetIndex
        || (candidate.workbookSheetIndex === undefined && candidate.index === chart.sheetIndex));
      if (!sheet) continue;
      let chartDescriptor: OfficeObjectDescriptor;
      try {
        chartDescriptor = describeXlsxChart(context, sheet, chart);
      } catch {
        continue;
      }
      descriptors.push(chartDescriptor);
      const objectId = xlsxChartObjectId(sheet, chart);
      for (const target of chartElementTargets(chart)) {
        try {
          const parentId = target.element === "point"
            ? chartElementObjectId(sheet, chart, { element: "series", seriesIndex: target.seriesIndex })
            : objectId;
          descriptors.push(chartElementParent(
            describeXlsxChartElement(context, sheet, chart, target),
            parentId,
          ));
        } catch {
          // A point index outside the parsed values or an unregistered chart is
          // simply not an object; one bad element must not drop the chart.
        }
      }
    }
    return descriptors;
  },
};

export const XLSX_BEHAVIOR_OBJECT_PROVIDER: OfficeObjectProvider<XlsxReferenceContext> = {
  id: "xlsx.behaviors",
  format: "xlsx",
  kinds: XLSX_STAGE4_KINDS,
  list(context) {
    return (context.behaviors ?? []).flatMap((record) => {
      try {
        return [describeXlsxBehavior(context, record)];
      } catch {
        return [];
      }
    });
  },
};

/** Registry with the built-in XLSX object families registered. */
export function createXlsxObjectRegistry(): OfficeObjectRegistry<XlsxReferenceContext> {
  return createOfficeObjectRegistry<XlsxReferenceContext>([
    XLSX_TABLE_OBJECT_PROVIDER,
    XLSX_DRAWING_OBJECT_PROVIDER,
    XLSX_COMMENT_OBJECT_PROVIDER,
    XLSX_CHART_OBJECT_PROVIDER,
    XLSX_BEHAVIOR_OBJECT_PROVIDER,
  ]);
}

/**
 * Collect stage 4 behaviors from the parsed sheet data.
 *
 * The caller owns the worksheet read: the viewer passes the already-parsed
 * conditional formats, validations, defined names and protection, plus the
 * formula cells it wants to expose.
 */
export function xlsxBehaviorRecordsFromSource(
  source: XlsxBehaviorSource,
): readonly XlsxBehaviorRecord[] {
  return createXlsxBehaviorRecords(source);
}

/** Convenience: attach collected behaviors to a reference context. */
export function withXlsxBehaviors(
  context: XlsxReferenceContext,
  behaviors: readonly XlsxBehaviorRecord[],
): XlsxReferenceContext {
  return { ...context, behaviors };
}

/** Exported for tests and hosts that build table descriptors directly. */
export function xlsxTableDescriptor(
  context: XlsxReferenceContext,
  table: XlsxReferenceTable,
): OfficeObjectDescriptor {
  return describeXlsxTable(context, table);
}

/** Exported for tests and hosts that build drawing descriptors directly. */
export function xlsxDrawingDescriptor(
  context: XlsxReferenceContext,
  drawing: XlsxReferenceDrawing,
): OfficeObjectDescriptor {
  return describeXlsxDrawing(context, drawing);
}

/** Exported for tests and hosts that build comment descriptors directly. */
export function xlsxCommentDescriptor(
  context: XlsxReferenceContext,
  comment: XlsxReferenceComment,
): OfficeObjectDescriptor {
  return describeXlsxComment(context, comment);
}

/**
 * Format-specific paint order used when candidates overlap.
 *
 * Excel paints later series above earlier series and drawing objects in their
 * worksheet `zIndex` order. The generic candidate ranker only knows the
 * optional `zIndex` field, so XLSX projects its own paint order into it here
 * instead of relying on descriptor id or reliability as a tie-breaker.
 */
export function xlsxObjectCandidateZIndex(descriptor: OfficeObjectDescriptor): number | undefined {
  const facet = descriptor.facets?.zIndex;
  if (typeof facet === "number" && Number.isFinite(facet)) return facet;
  for (const fragment of descriptor.visual?.fragments ?? []) {
    if (fragment.zIndex !== undefined && Number.isFinite(fragment.zIndex)) return fragment.zIndex;
  }
  return undefined;
}

/** Fill missing XLSX z-order evidence and rank with the shared candidate sort. */
export function rankXlsxObjectCandidates(
  candidates: readonly OfficeObjectCandidate[],
): readonly OfficeObjectCandidate[] {
  return rankOfficeObjectCandidates(candidates.map((candidate) => {
    if (candidate.zIndex !== undefined) return candidate;
    const zIndex = xlsxObjectCandidateZIndex(candidate.descriptor);
    return zIndex === undefined ? candidate : { ...candidate, zIndex };
  }));
}

/** Shared reliability copy for host-defined XLSX object families. */
export function xlsxStructuralObjectReliability() {
  return officeObjectReliability(
    { level: "likely", reasonCodes: ["xlsx.structural-object"] },
    { level: "likely", reasonCodes: ["xlsx.derived-geometry"] },
    { level: "likely", reasonCodes: ["xlsx.sheet-identity"] },
    { level: "uncertain", reasonCodes: ["xlsx.structural-object"] },
  );
}
