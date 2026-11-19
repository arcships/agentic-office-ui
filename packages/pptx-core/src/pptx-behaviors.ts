// Stage 4 invisible behavior records for PPTX.
//
// Animations, transitions, actions, masters and speaker notes all live in the
// slide playback document. They are references with a slide scope and a native
// id; a host renders them in a timeline, a transition panel or a structure
// tree, never as a canvas outline.

import type { JsonValue, OfficeObjectKind } from "@arcships/office-interaction";
import type {
  PptxPlaybackDocument,
  PptxPlaybackSlide,
  PptxSlideAction,
  PptxTimeNode,
} from "./playback/types";
import type { PptxPreviewDocument } from "./types";

export type PptxBehaviorKind = "animation" | "transition" | "action" | "master" | "note";

export const PPTX_BEHAVIOR_REFERENCE_KIND: Record<PptxBehaviorKind, OfficeObjectKind> = {
  animation: "animation",
  transition: "transition",
  action: "action",
  master: "master",
  note: "note",
};

export interface PptxBehaviorRecord {
  behavior: PptxBehaviorKind
  /** Native id the locator carries; unique inside the slide scope. */
  instanceId: string
  /** Object the behavior belongs to, when it targets one. */
  ownerObjectKey?: string
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  content?: { text?: string }
  facets: Readonly<Record<string, JsonValue>>
}

export interface PptxBehaviorContext {
  document: PptxPreviewDocument
  playbackDocument: PptxPlaybackDocument
  slideIds?: readonly (string | undefined)[]
}

function slideKey(context: PptxBehaviorContext, slideIndex: number): string {
  return context.slideIds?.[slideIndex]?.trim() || `slide:${slideIndex}`;
}

function nodes(slide: PptxPlaybackSlide): readonly PptxTimeNode[] {
  return Object.values(slide.nodes ?? {});
}

function effectLabel(node: PptxTimeNode): string {
  const effect = node.effect;
  const preset = effect?.presetId ?? effect?.presetClass ?? effect?.kind ?? "effect";
  return `Animation ${preset}`;
}

function animationRecords(
  context: PptxBehaviorContext,
  slide: PptxPlaybackSlide,
  slideIndex: number,
): PptxBehaviorRecord[] {
  const slideId = slideKey(context, slideIndex);
  return nodes(slide)
    .filter((node) => node.effect !== undefined || node.kind === "effect" || node.kind === "media")
    .map((node) => {
      const effect = node.effect;
      const label = node.kind === "media" ? "Media playback" : effectLabel(node);
      return {
        behavior: "animation",
        instanceId: `pptx.animation:${slideId}:${node.id}`,
        ...(effect?.targetObjectKey ? { ownerObjectKey: effect.targetObjectKey } : {}),
        label,
        path: [
          { kind: "document", label: "Presentation" },
          { kind: "slide", label: `Slide ${slideIndex + 1}` },
          { kind: "animation", label },
        ],
        facets: {
          nodeId: node.id,
          nodeKind: node.kind,
          delayMs: node.delayMs,
          durationMs: node.durationMs === "indefinite" ? -1 : node.durationMs,
          nodeType: node.nodeType ?? "",
          ...(effect?.id ? { effectId: effect.id } : {}),
          ...(effect?.presetId ? { presetId: effect.presetId } : {}),
          ...(effect?.targetShapeId ? { targetShapeId: effect.targetShapeId } : {}),
        },
      } satisfies PptxBehaviorRecord;
    });
}

function transitionRecord(
  context: PptxBehaviorContext,
  slide: PptxPlaybackSlide,
  slideIndex: number,
): PptxBehaviorRecord[] {
  const transition = slide.transition;
  if (!transition) return [];
  return [{
    behavior: "transition",
    instanceId: `pptx.transition:${slideKey(context, slideIndex)}:${transition.kind}`,
    label: "Slide transition",
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      { kind: "transition", label: "Slide transition" },
    ],
    facets: {
      kind: transition.kind,
      durationMs: transition.durationMs,
      advanceOnClick: transition.advanceOnClick,
      ...(transition.direction ? { direction: transition.direction } : {}),
      ...(transition.orientation ? { orientation: transition.orientation } : {}),
      ...(transition.advanceAfterMs === undefined ? {} : { advanceAfterMs: transition.advanceAfterMs }),
    },
  }];
}

function actionLabel(action: PptxSlideAction): string {
  if (action.kind === "go-to-slide") return `Action: slide ${(action.targetSlideIndex ?? 0) + 1}`;
  if (action.kind === "open-url") return `Action: ${action.url ?? "URL"}`;
  if (action.kind === "mailto") return `Action: ${action.url ?? "mail"}`;
  return "Action";
}

function actionRecords(
  context: PptxBehaviorContext,
  slide: PptxPlaybackSlide,
  slideIndex: number,
): PptxBehaviorRecord[] {
  const slideId = slideKey(context, slideIndex);
  return (slide.actions ?? []).map((action) => ({
    behavior: "action",
    instanceId: `pptx.action:${slideId}:${action.id}`,
    ownerObjectKey: action.sourceObjectKey,
    label: actionLabel(action),
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      { kind: "action", label: actionLabel(action) },
    ],
    facets: {
      actionKind: action.kind,
      trigger: action.trigger,
      sourceObjectKey: action.sourceObjectKey,
      ...(action.targetSlideIndex === undefined ? {} : { targetSlideIndex: action.targetSlideIndex }),
      ...(action.url ? { url: action.url } : {}),
      ...(action.tooltip ? { tooltip: action.tooltip } : {}),
    },
  }));
}

function noteRecord(
  context: PptxBehaviorContext,
  slideIndex: number,
): PptxBehaviorRecord[] {
  const notes = context.document.slides[slideIndex]?.notes?.trim();
  if (!notes) return [];
  return [{
    behavior: "note",
    instanceId: `pptx.note:${slideKey(context, slideIndex)}`,
    label: "Speaker notes",
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      { kind: "note", label: "Speaker notes" },
    ],
    content: { text: notes },
    facets: {
      length: notes.length,
      paragraphs: notes.split(/\n/u).length,
    },
  }];
}

function masterRecords(
  context: PptxBehaviorContext,
  slide: PptxPlaybackSlide,
  slideIndex: number,
): PptxBehaviorRecord[] {
  const masters = new Map<string, string>();
  for (const identity of slide.objects) {
    if (identity.source !== "master") continue;
    const path = identity.slidePath;
    if (!path || masters.has(path)) continue;
    masters.set(path, identity.name ?? path);
  }
  return [...masters.entries()].map(([path, name]) => ({
    behavior: "master",
    instanceId: `pptx.master:${path}`,
    label: "Slide master",
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      { kind: "master", label: "Slide master" },
    ],
    facets: { masterPath: path, masterName: name },
  }));
}

/** Enumerate every stage 4 behavior on one slide. */
export function enumeratePptxSlideBehaviors(
  context: PptxBehaviorContext,
  slideIndex: number,
): readonly PptxBehaviorRecord[] {
  const slide = context.playbackDocument.slides[slideIndex];
  if (!slide) return Object.freeze([]);
  return Object.freeze([
    ...animationRecords(context, slide, slideIndex),
    ...transitionRecord(context, slide, slideIndex),
    ...actionRecords(context, slide, slideIndex),
    ...masterRecords(context, slide, slideIndex),
    ...noteRecord(context, slideIndex),
  ]);
}
