// Stage 2–4 object index for PPTX.
//
// Slide objects, groups, table cells, chart internals and stage 4 behaviors all
// become descriptors here. The index reads the playback identities (native
// shape id, source, group path) and the two optional host providers for table
// grids and chart snapshots; it never touches the DOM.

import type {
  OfficeObjectDescriptor,
  OfficeObjectKind,
  OfficeObjectProvider,
  OfficeObjectRegistry,
} from "@arcships/office-interaction";
import {
  createOfficeObjectDescriptor,
  createOfficeObjectRegistry,
} from "@arcships/office-interaction";
import type {
  PptxObjectIdentity,
  PptxPlaybackSlide,
} from "./playback/types";
import { createPptxObjectKey } from "./playback/identity";
import type {
  PptxBehaviorRecord,
} from "./pptx-behaviors";
import { PPTX_BEHAVIOR_REFERENCE_KIND, enumeratePptxSlideBehaviors } from "./pptx-behaviors";
import type {
  PptxChartSnapshot,
  PptxReferenceContext,
  PptxSubObjectTarget,
  PptxTableSnapshot,
} from "./reference-adapter";
import {
  createPptxBehaviorDraftFromRecord,
  createPptxSubObjectReferenceDraft,
  pptxReferenceKindForObject,
} from "./reference-adapter";

export const PPTX_STAGE2_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "text-box",
  "shape",
  "image",
  "table",
  "chart",
  "group",
]);

export const PPTX_STAGE3_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "table-cell",
  "chart-series",
  "chart-point",
  "chart-legend-entry",
]);

export const PPTX_STAGE4_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "animation",
  "transition",
  "action",
  "master",
  "note",
]);

function slideKey(context: PptxReferenceContext, slideIndex: number): string {
  return context.slideIds?.[slideIndex]?.trim() || `slide:${slideIndex}`;
}

function slideDescriptor(
  context: PptxReferenceContext,
  slideIndex: number,
): OfficeObjectDescriptor {
  return createOfficeObjectDescriptor({
    objectId: `pptx:slide:${slideIndex}`,
    draft: {
      schemaVersion: 1,
      document: context.revision,
      kind: "slide",
      source: "native",
      locator: {
        type: "format",
        format: "pptx",
        value: {
          kind: "slide",
          slide: {
            ...(slideKey(context, slideIndex) ? { slideId: slideKey(context, slideIndex) } : {}),
            index: slideIndex,
          },
        },
      },
      fingerprint: {
        ancestorKeys: [`pptx.slide-id:${slideKey(context, slideIndex)}`],
      },
      reliability: {
        semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-slide"] },
        boundary: { level: "exact", score: 1, reasonCodes: ["pptx.slide-boundary"] },
        hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.presentation-slide"] },
        relocation: { level: "likely", reasonCodes: ["pptx.slide-id"] },
      },
    },
    label: `Slide ${slideIndex + 1}`,
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
    ],
    childrenState: "available",
    facets: { slideIndex },
  });
}

function enclosingGroupObjectId(
  slide: PptxPlaybackSlide,
  identity: PptxObjectIdentity,
): string | undefined {
  const groupPath = identity.groupPath;
  if (groupPath.length === 0) return undefined;
  const parentGroupPath = groupPath.slice(0, -1);
  const groupShapeId = groupPath[groupPath.length - 1]!;
  const candidates = slide.objects.filter((object) =>
    object.shapeId === groupShapeId
    && object.groupPath.length === parentGroupPath.length
    && object.groupPath.every((id, index) => id === parentGroupPath[index]));
  if (candidates.length === 0) {
    // A group node without its own rendered identity still keeps the shape id
    // stable across revisions.
    const key = createPptxObjectKey({
      slidePath: identity.slidePath,
      source: identity.source,
      shapeId: groupShapeId,
      groupPath: parentGroupPath,
    });
    return `pptx:object:${key}`;
  }
  return `pptx:object:${candidates[0]!.key}`;
}

function objectDescriptor(
  context: PptxReferenceContext,
  slideIndex: number,
  identity: PptxObjectIdentity,
  parentObjectId: string,
): OfficeObjectDescriptor {
  const kind = pptxReferenceKindForObject(identity) ?? "shape";
  const label = identity.name?.trim() || `${kind} ${identity.shapeId}`;
  return createOfficeObjectDescriptor({
    objectId: `pptx:object:${identity.key}`,
    draft: {
      schemaVersion: 1,
      document: context.revision,
      kind,
      source: "native",
      locator: {
        type: "format",
        format: "pptx",
        value: {
          kind: "object",
          slide: { slideId: slideKey(context, slideIndex), index: slideIndex },
          object: {
            objectKey: identity.key,
            shapeId: identity.shapeId,
            source: identity.source,
            groupPath: [...identity.groupPath],
          },
        },
      },
      fingerprint: {
        ...(identity.name ? { objectName: identity.name } : {}),
        ancestorKeys: [
          `pptx.slide-id:${slideKey(context, slideIndex)}`,
          `pptx.node-type:${identity.nodeType}`,
          `pptx.shape-id:${identity.shapeId}`,
          ...identity.groupPath.map((id) => `pptx.group:${id}`),
        ],
      },
      reliability: {
        semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-object-type"] },
        boundary: { level: "exact", score: 1, reasonCodes: ["pptx.shape-id"] },
        hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.group-path"] },
        relocation: { level: "likely", reasonCodes: ["pptx.object-key-or-name"] },
      },
    },
    label,
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      ...identity.groupPath.map((id) => ({ kind: "group" as const, label: `Group ${id}` })),
      { kind, label },
    ],
    parentObjectId,
    childrenState: kind === "group" || kind === "table" || kind === "chart" ? "available" : "none",
    ...(identity.text ? { content: { text: identity.text } } : {}),
    facets: {
      slideIndex,
      objectKey: identity.key,
      shapeId: identity.shapeId,
      source: identity.source,
      nodeType: identity.nodeType,
      groupPath: [...identity.groupPath],
    },
  });
}

function subObjectTargets(
  slideIndex: number,
  identity: PptxObjectIdentity,
  context: PptxReferenceContext,
): readonly PptxSubObjectTarget[] {
  const nodeType = identity.nodeType.toLowerCase();
  const targets: PptxSubObjectTarget[] = [];
  if (nodeType === "table") {
    const table: PptxTableSnapshot | undefined = context.getObjectTable?.(slideIndex, identity.key);
    table?.rows.forEach((row, rowIndex) => {
      row.forEach((_cell, cellIndex) => {
        targets.push({ kind: "table-cell", rowIndex, cellIndex });
      });
    });
    return targets;
  }
  if (nodeType === "chart") {
    const chart: PptxChartSnapshot | undefined = context.getObjectChart?.(slideIndex, identity.key);
    chart?.series.forEach((series, seriesIndex) => {
      targets.push({ kind: "chart-series", seriesIndex });
      (series.points ?? []).forEach((_point, pointIndex) => {
        targets.push({ kind: "chart-point", seriesIndex, pointIndex });
      });
      targets.push({ kind: "chart-legend-entry", legendIndex: seriesIndex, seriesIndex });
    });
  }
  return targets;
}

function subObjectObjectId(identity: PptxObjectIdentity, target: PptxSubObjectTarget): string {
  return `pptx:sub-object:${identity.key}:${target.kind}:${JSON.stringify(target)}`;
}

function seriesObjectId(identity: PptxObjectIdentity, seriesIndex: number): string {
  return subObjectObjectId(identity, { kind: "chart-series", seriesIndex });
}

function subObjectDescriptors(
  context: PptxReferenceContext,
  slideIndex: number,
  identity: PptxObjectIdentity,
  parentObjectId: string,
): readonly OfficeObjectDescriptor[] {
  return subObjectTargets(slideIndex, identity, context).flatMap((target) => {
    try {
      const draft = createPptxSubObjectReferenceDraft(context, slideIndex, identity.key, target);
      const label = target.kind === "table-cell"
        ? `Cell ${target.rowIndex + 1}, ${target.cellIndex + 1}`
        : target.kind === "chart-series"
          ? `Series: ${context.getObjectChart?.(slideIndex, identity.key)?.series[target.seriesIndex]?.name ?? `Series ${target.seriesIndex + 1}`}`
          : target.kind === "chart-point"
            ? `Point ${target.pointIndex + 1}`
            : `Legend entry ${target.legendIndex + 1}`;
      return [createOfficeObjectDescriptor({
        objectId: subObjectObjectId(identity, target),
        draft,
        label,
        path: [
          { kind: "document", label: "Presentation" },
          { kind: "slide", label: `Slide ${slideIndex + 1}` },
          ...identity.groupPath.map((id) => ({ kind: "group" as const, label: `Group ${id}` })),
          { kind: identity.nodeType === "table" ? "table" : "chart", label: identity.name ?? identity.shapeId },
          { kind: draft.kind, label },
        ],
        parentObjectId: target.kind === "chart-point"
          ? seriesObjectId(identity, target.seriesIndex)
          : parentObjectId,
        childrenState: "none",
        facets: { slideIndex, objectKey: identity.key, target },
      })];
    } catch {
      return [];
    }
  });
}

function behaviorDescriptor(
  context: PptxReferenceContext,
  slideIndex: number,
  record: PptxBehaviorRecord,
  parentObjectId: string,
): OfficeObjectDescriptor {
  const draft = createPptxBehaviorDraftFromRecord(context, slideIndex, record);
  return createOfficeObjectDescriptor({
    objectId: `pptx:behavior:${record.instanceId}`,
    draft,
    label: record.label,
    path: record.path,
    parentObjectId,
    childrenState: "none",
    ...(record.content === undefined ? {} : { content: record.content }),
    facets: {
      ...record.facets,
      behavior: record.behavior,
      instanceId: record.instanceId,
      referenceKind: PPTX_BEHAVIOR_REFERENCE_KIND[record.behavior],
    },
  });
}

export interface PptxObjectIndexOptions {
  /** Include stage 3 sub-objects (table cells, chart internals). Defaults to true. */
  includeSubObjects?: boolean;
  /** Include stage 4 behaviors. Defaults to true. */
  includeBehaviors?: boolean;
}

export interface PptxObjectIndex {
  descriptors: readonly OfficeObjectDescriptor[]
  byObjectId: ReadonlyMap<string, OfficeObjectDescriptor>
}

function buildPptxObjectIndex(
  context: PptxReferenceContext,
  options: PptxObjectIndexOptions = {},
): PptxObjectIndex {
  const descriptors: OfficeObjectDescriptor[] = [];
  context.playbackDocument.slides.forEach((slide: PptxPlaybackSlide, slideIndex) => {
    let slideObject: OfficeObjectDescriptor;
    try {
      slideObject = slideDescriptor(context, slideIndex);
    } catch {
      return;
    }
    descriptors.push(slideObject);
    for (const identity of slide.objects) {
      let object: OfficeObjectDescriptor;
      try {
        object = objectDescriptor(
          context,
          slideIndex,
          identity,
          enclosingGroupObjectId(slide, identity) ?? slideObject.objectId,
        );
      } catch {
        continue;
      }
      descriptors.push(object);
      if (options.includeSubObjects !== false) {
        descriptors.push(...subObjectDescriptors(context, slideIndex, identity, object.objectId));
      }
    }
    if (options.includeBehaviors !== false) {
      for (const record of enumeratePptxSlideBehaviors(context, slideIndex)) {
        try {
          descriptors.push(behaviorDescriptor(context, slideIndex, record, slideObject.objectId));
        } catch {
          // A behavior that cannot build a valid draft is not an object; one
          // bad animation must not drop the rest of the slide.
        }
      }
    }
  });
  return {
    descriptors: Object.freeze(descriptors),
    byObjectId: new Map(descriptors.map((descriptor) => [descriptor.objectId, descriptor])),
  };
}

const indexCache = new WeakMap<object, Map<string, PptxObjectIndex>>();

/** Build (and memoize per playback document + options) the PPTX object index. */
export function createPptxObjectIndex(
  context: PptxReferenceContext,
  options: PptxObjectIndexOptions = {},
): PptxObjectIndex {
  const key = `${context.revision.documentId}:${context.revision.revision}:${options.includeSubObjects !== false}:${options.includeBehaviors !== false}`;
  const cacheKey = context.playbackDocument as object;
  let entries = indexCache.get(cacheKey);
  if (!entries) {
    entries = new Map();
    indexCache.set(cacheKey, entries);
  }
  const cached = entries.get(key);
  if (cached) return cached;
  const index = buildPptxObjectIndex(context, options);
  if (entries.size > 4) entries.clear();
  entries.set(key, index);
  return index;
}

export const PPTX_OBJECT_PROVIDER: OfficeObjectProvider<PptxReferenceContext> = {
  id: "pptx.objects",
  format: "pptx",
  kinds: [...PPTX_STAGE2_KINDS, ...PPTX_STAGE3_KINDS],
  list(context) {
    return createPptxObjectIndex(context).descriptors
      .filter((descriptor) => descriptor.draft.kind !== "slide");
  },
};

export const PPTX_BEHAVIOR_OBJECT_PROVIDER: OfficeObjectProvider<PptxReferenceContext> = {
  id: "pptx.behaviors",
  format: "pptx",
  kinds: PPTX_STAGE4_KINDS,
  list(context) {
    return createPptxObjectIndex(context).descriptors
      .filter((descriptor) => PPTX_STAGE4_KINDS.includes(descriptor.draft.kind));
  },
};

/** Registry with the built-in PPTX object families registered. */
export function createPptxObjectRegistry(): OfficeObjectRegistry<PptxReferenceContext> {
  return createOfficeObjectRegistry<PptxReferenceContext>([
    PPTX_OBJECT_PROVIDER,
    PPTX_BEHAVIOR_OBJECT_PROVIDER,
  ]);
}
