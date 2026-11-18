// Tracked-change decisions: accept / reject revisions at the model level.
//
// Imported OOXML reaches the serializer unchanged through `sourceXml`, while
// revisions are only half-materialised in the model: inserted content already
// exists as ordinary children/cells, deleted content does not. Accepting or
// rejecting a revision therefore edits both representations in lockstep:
//
//   * `sourceXml` loses (or unwraps) the revision envelope, so the decision
//     survives export and any remaining revisions stay visible;
//   * model children, table rows or table cells gain (reject a deletion) or
//     lose (reject an insertion) the content the revision covers.
//
// Paragraph revisions map their XML spans onto rendered children with
// `collectParagraphVisibleUnits` / `mapVisibleUnitsToParagraphChildren`
// (tracked-changes-inline.ts), the same cursor space the run renderer and the
// gutter use, so merged runs, notes and form fields cannot drift.
//
// Formatting revisions are reverted where the model can represent the previous
// formatting: `w:rPrChange` restores the previous run style and `w:pPrChange`
// restores an explicit alignment. Table/row/cell property changes have no
// previous-formatting counterpart in the model, so rejecting those only clears
// the revision markup; callers receive their ids in `clearedOnlyChangeIds`.
//
// Everything here is pure with respect to its input: the model is cloned before
// mutation and returned in the result.

import type {
  DocModel,
  ParagraphAlignment,
  ParagraphChildNode,
  ParagraphNode,
  TextRunNode
} from "../../engine/types";
import type {
  DocxTextRangeLocation,
  DocxTrackedChange
} from "./editor-types";
import { cloneDocModel } from "../../engine/clone";
import {
  decodeXmlText,
  extractBalancedTagRanges,
  parseParagraphAlignmentFromXml,
  parseRunStyleFromRunXml,
  type XmlBalancedTagRange
} from "./xml-parsing";
import {
  collectTrackedChangesFromModel,
  resolveParagraphTrackedMarkup
} from "./tracked-changes";
import {
  collectParagraphVisibleUnits,
  mapVisibleUnitsToParagraphChildren,
  paragraphVisibleCursorBefore,
  type ParagraphVisibleUnit
} from "./tracked-changes-inline";
import {
  applyXmlEdits,
  collectContentRevisionRanges,
  collectPropertyChangeRanges,
  deletedContentToVisibleXml,
  innerXmlOfRange,
  matchRevisionRange,
  type ContentRevisionRange,
  type PropertyChangeRange,
  type XmlEdit
} from "./tracked-change-xml";
import {
  applyTableRevisionDecisions,
  tableRevisionStableId
} from "./tracked-change-table-decisions";
import {
  resolveTableTrackedMarkup,
  type TableTrackedRevision
} from "./tracked-changes-table";
import { getParagraphAtLocation } from "./selection-helpers";
import { paragraphLocationKey } from "./text-mutation";

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export type DocxTrackedChangeDecision = "accept" | "reject";

/** Narrows which revisions a decision applies to. */
export interface DocxTrackedChangeDecisionScope {
  /** Only these `DocxTrackedChange.id` values. */
  changeIds?: readonly string[];
  /** Only changes anchored in this document node (paragraph or table). */
  nodeIndex?: number;
  /** Only table-cell changes whose row falls inside the half-open range. */
  tableRowRange?: {
    startRowIndex: number;
    endRowIndex: number;
  };
}

export interface DocxTrackedChangeDecisionResult {
  /** Model after the decision. Identical to the input when nothing matched. */
  model: DocModel;
  /** Revisions that were accepted or rejected. */
  resolvedChangeIds: string[];
  /** Revisions whose source markup was already gone (edited paragraph). */
  unresolvedChangeIds: string[];
  /** Revisions cleared without reverting the recorded formatting. */
  clearedOnlyChangeIds: string[];
}

// ---------------------------------------------------------------------------
// Paragraph child cursor
// ---------------------------------------------------------------------------

function visibleChildIndexes(paragraph: ParagraphNode): number[] {
  const indexes: number[] = [];
  paragraph.children.forEach((child, index) => {
    if (child.type !== "form-field") {
      indexes.push(index);
    }
  });
  return indexes;
}

function paragraphVisibleUnits(paragraph: ParagraphNode): ParagraphVisibleUnit[] {
  const units = collectParagraphVisibleUnits(paragraph.sourceXml ?? "");
  mapVisibleUnitsToParagraphChildren(units, paragraph.children);
  return units;
}

function visibleOrdinalsWithinRange(
  units: readonly ParagraphVisibleUnit[],
  range: { start: number; end: number }
): number[] {
  const ordinals = new Set<number>();
  for (const unit of units) {
    if (unit.visibleChildIndex === undefined) {
      continue;
    }
    if (unit.start >= range.start && unit.end <= range.end) {
      ordinals.add(unit.visibleChildIndex);
    }
  }
  return [...ordinals].sort((left, right) => left - right);
}

interface DeletedSegmentRestore {
  ordinal: number;
  text: string;
  style?: TextRunNode["style"];
}

function deletedTextWithinRange(xml: string, range: XmlBalancedTagRange): string {
  const slice = xml.slice(range.start, range.end);
  let text = "";
  for (const match of slice.matchAll(
    /<w:(?:delText|instrText|delInstrText)\b[^>]*>([\s\S]*?)<\/w:(?:delText|instrText|delInstrText)>/giu
  )) {
    text += decodeXmlText(match[1] ?? "");
  }
  return text;
}

function firstRunStyleWithinRange(
  xml: string,
  range: XmlBalancedTagRange
): TextRunNode["style"] | undefined {
  const slice = xml.slice(range.start, range.end);
  const runRange = extractBalancedTagRanges(slice, "w:r")[0];
  return runRange ? parseRunStyleFromRunXml(slice.slice(runRange.start, runRange.end)) : undefined;
}

interface ParagraphChildRebuild {
  removeOrdinals: ReadonlySet<number>;
  restores: readonly DeletedSegmentRestore[];
}

function rebuildParagraphChildren(
  paragraph: ParagraphNode,
  rebuild: ParagraphChildRebuild
): ParagraphChildNode[] | undefined {
  if (rebuild.removeOrdinals.size === 0 && rebuild.restores.length === 0) {
    return undefined;
  }

  const indexes = visibleChildIndexes(paragraph);
  const removeChildIndexes = new Set<number>();
  rebuild.removeOrdinals.forEach((ordinal) => {
    const childIndex = indexes[ordinal];
    if (childIndex !== undefined) {
      removeChildIndexes.add(childIndex);
    }
  });

  const restoresByOrdinal = new Map<number, ParagraphChildNode[]>();
  for (const restore of rebuild.restores) {
    const bucket = restoresByOrdinal.get(restore.ordinal) ?? [];
    bucket.push({
      type: "text",
      text: restore.text,
      style: restore.style ? { ...restore.style } : undefined
    });
    restoresByOrdinal.set(restore.ordinal, bucket);
  }

  const nextChildren: ParagraphChildNode[] = [];
  let ordinal = 0;
  paragraph.children.forEach((child, childIndex) => {
    if (child.type !== "form-field") {
      const pending = restoresByOrdinal.get(ordinal);
      if (pending) {
        nextChildren.push(...pending);
        restoresByOrdinal.delete(ordinal);
      }
      ordinal += 1;
    }
    if (!removeChildIndexes.has(childIndex)) {
      nextChildren.push(child);
    }
  });

  const trailing = restoresByOrdinal.get(ordinal);
  if (trailing) {
    nextChildren.push(...trailing);
    restoresByOrdinal.delete(ordinal);
  }

  return nextChildren;
}

function revertRunStyleAtOrdinals(
  paragraph: ParagraphNode,
  ordinals: readonly number[],
  style: TextRunNode["style"]
): number {
  const visibleChildren = paragraph.children.filter(
    (child) => child.type !== "form-field"
  );
  let applied = 0;
  for (const ordinal of ordinals) {
    const child = visibleChildren[ordinal];
    if (!child || child.type !== "text") {
      continue;
    }
    child.style = { ...style };
    applied += 1;
  }
  return applied;
}

function hasMeaningfulRunStyle(
  style: TextRunNode["style"] | undefined
): style is TextRunNode["style"] {
  return Boolean(
    style && Object.values(style).some((value) => value !== undefined)
  );
}

// ---------------------------------------------------------------------------
// Paragraph resolution
// ---------------------------------------------------------------------------

/**
 * Resolves the paragraph a change is anchored to, including paragraphs inside
 * nested tables (their `paragraphIndex` follows `tableCellParagraphEntries`).
 */
export function resolveTrackedChangeParagraph(
  model: DocModel,
  location: DocxTextRangeLocation
): ParagraphNode | undefined {
  return getParagraphAtLocation(model, location).paragraph;
}

// ---------------------------------------------------------------------------
// Per-paragraph decision planning
// ---------------------------------------------------------------------------

interface DecisionAccumulator {
  resolvedChangeIds: string[];
  unresolvedChangeIds: string[];
  clearedOnlyChangeIds: string[];
}

interface ParagraphDecisionPlan {
  /** `w:rPrChange` / `w:pPrChange` removals, applied before content edits. */
  propertyEdits: XmlEdit[];
  /** Content revisions still to be unwrapped or dropped. */
  contentChanges: DocxTrackedChange[];
  runStyleReverts: Array<{ ordinals: number[]; style: TextRunNode["style"] }>;
  alignRevert?: ParagraphAlignment;
  removeOrdinals: number[];
  restores: DeletedSegmentRestore[];
}

function runRangeContaining(
  sourceXml: string,
  range: XmlBalancedTagRange
): XmlBalancedTagRange | undefined {
  return extractBalancedTagRanges(sourceXml, "w:r").find(
    (run) => run.start <= range.start && run.end >= range.end
  );
}

function planParagraphDecision(
  paragraph: ParagraphNode,
  changes: readonly DocxTrackedChange[],
  decision: DocxTrackedChangeDecision,
  accumulator: DecisionAccumulator
): ParagraphDecisionPlan | undefined {
  const sourceXml = paragraph.sourceXml ?? "";
  const markup = sourceXml ? resolveParagraphTrackedMarkup(paragraph) : undefined;
  if (!sourceXml || !markup) {
    accumulator.unresolvedChangeIds.push(...changes.map((change) => change.id));
    return undefined;
  }

  const contentRanges = collectContentRevisionRanges(sourceXml);
  const runPropertyRanges = collectPropertyChangeRanges(sourceXml, "w:rPrChange");
  const paragraphPropertyRanges = collectPropertyChangeRanges(
    sourceXml,
    "w:pPrChange"
  );
  const units = paragraphVisibleUnits(paragraph);

  const claimedContentRanges = new Set<ContentRevisionRange>();
  const claimedRunPropertyRanges = new Set<PropertyChangeRange>();
  const claimedParagraphPropertyRanges = new Set<PropertyChangeRange>();

  const plan: ParagraphDecisionPlan = {
    propertyEdits: [],
    contentChanges: [],
    runStyleReverts: [],
    removeOrdinals: [],
    restores: []
  };
  const removeOrdinals = new Set<number>();

  for (const change of changes) {
    if (change.kind === "format-change") {
      const range = matchRevisionRange(
        change,
        runPropertyRanges,
        claimedRunPropertyRanges
      );
      if (!range) {
        accumulator.unresolvedChangeIds.push(change.id);
        continue;
      }
      claimedRunPropertyRanges.add(range);
      plan.propertyEdits.push({ start: range.start, end: range.end, text: "" });
      accumulator.resolvedChangeIds.push(change.id);

      if (decision === "accept") {
        continue;
      }
      const previousStyle = parseRunStyleFromRunXml(
        sourceXml.slice(range.start, range.end)
      );
      const runRange = runRangeContaining(sourceXml, range);
      const ordinals = runRange
        ? visibleOrdinalsWithinRange(units, runRange)
        : [];
      if (hasMeaningfulRunStyle(previousStyle) && ordinals.length > 0) {
        plan.runStyleReverts.push({ ordinals, style: previousStyle });
      } else {
        accumulator.clearedOnlyChangeIds.push(change.id);
      }
      continue;
    }

    if (change.kind === "paragraph-format-change") {
      const range = matchRevisionRange(
        change,
        paragraphPropertyRanges,
        claimedParagraphPropertyRanges
      );
      if (!range) {
        accumulator.unresolvedChangeIds.push(change.id);
        continue;
      }
      claimedParagraphPropertyRanges.add(range);
      plan.propertyEdits.push({ start: range.start, end: range.end, text: "" });
      accumulator.resolvedChangeIds.push(change.id);

      if (decision === "accept") {
        continue;
      }
      const previousAlign = parseParagraphAlignmentFromXml(
        innerXmlOfRange(sourceXml, range)
      );
      if (previousAlign) {
        plan.alignRevert = previousAlign;
      } else {
        accumulator.clearedOnlyChangeIds.push(change.id);
      }
      continue;
    }

    const range = matchRevisionRange(change, contentRanges, claimedContentRanges);
    if (!range) {
      accumulator.unresolvedChangeIds.push(change.id);
      continue;
    }
    claimedContentRanges.add(range);
    accumulator.resolvedChangeIds.push(change.id);
    plan.contentChanges.push(change);

    const isInsertionLike =
      change.kind === "insertion" || change.kind === "move-to";
    if (isInsertionLike) {
      if (decision === "accept") {
        continue;
      }
      visibleOrdinalsWithinRange(units, range).forEach((ordinal) =>
        removeOrdinals.add(ordinal)
      );
      continue;
    }

    if (decision === "accept") {
      continue;
    }
    const restoredText = deletedTextWithinRange(sourceXml, range);
    if (restoredText) {
      plan.restores.push({
        ordinal: paragraphVisibleCursorBefore(units, range.start),
        text: restoredText,
        style: firstRunStyleWithinRange(sourceXml, range)
      });
    }
  }

  plan.removeOrdinals = [...removeOrdinals];
  return plan;
}

// ---------------------------------------------------------------------------
// Content-revision rewrite (second pass, after property edits)
// ---------------------------------------------------------------------------

function rewriteContentRevisions(
  paragraph: ParagraphNode,
  plan: ParagraphDecisionPlan,
  decision: DocxTrackedChangeDecision
): boolean {
  if (plan.contentChanges.length === 0) {
    return false;
  }

  const sourceXml = paragraph.sourceXml ?? "";
  const ranges = collectContentRevisionRanges(sourceXml);
  const claimed = new Set<ContentRevisionRange>();
  const edits: XmlEdit[] = [];
  for (const change of plan.contentChanges) {
    const range = matchRevisionRange(change, ranges, claimed);
    if (!range) {
      continue;
    }
    claimed.add(range);
    const innerXml = innerXmlOfRange(sourceXml, range);
    const isInsertionLike =
      change.kind === "insertion" || change.kind === "move-to";
    let replacement: string;
    if (decision === "accept") {
      replacement = isInsertionLike ? innerXml : "";
    } else {
      replacement = isInsertionLike ? "" : deletedContentToVisibleXml(innerXml);
    }
    edits.push({ start: range.start, end: range.end, text: replacement });
  }

  if (edits.length === 0) {
    return false;
  }
  paragraph.sourceXml = applyXmlEdits(sourceXml, edits);
  return true;
}

// ---------------------------------------------------------------------------
// Public entry points
// ---------------------------------------------------------------------------

/**
 * True when the revision is already drawn inline and must not repeat as a
 * gutter card. Inserted and moved-to runs are highlighted in the text itself;
 * table row/cell revisions have no inline representation of their own and keep
 * their card even though they carry the `insertion` kind.
 */
export function trackedChangeDuplicatesInlineMarkup(
  change: DocxTrackedChange
): boolean {
  if (!change.inlineAnchorId) {
    return false;
  }
  return change.kind === "insertion" || change.kind === "move-to";
}

function changeMatchesScope(
  change: DocxTrackedChange,
  scope: DocxTrackedChangeDecisionScope | undefined
): boolean {
  if (!scope) {
    return true;
  }
  if (scope.changeIds && !scope.changeIds.includes(change.id)) {
    return false;
  }
  if (scope.nodeIndex !== undefined && change.nodeIndex !== scope.nodeIndex) {
    return false;
  }
  if (scope.tableRowRange && change.location.kind === "table-cell") {
    const { startRowIndex, endRowIndex } = scope.tableRowRange;
    return (
      change.location.rowIndex >= startRowIndex &&
      change.location.rowIndex < endRowIndex
    );
  }
  return true;
}

/**
 * Accepts or rejects the tracked changes selected by `scope` (every change in
 * the document when omitted) and returns the rewritten model.
 */
export function applyTrackedChangeDecision(
  model: DocModel,
  decision: DocxTrackedChangeDecision,
  scope?: DocxTrackedChangeDecisionScope
): DocxTrackedChangeDecisionResult {
  const selected = collectTrackedChangesFromModel(model).filter((change) =>
    changeMatchesScope(change, scope)
  );
  if (selected.length === 0) {
    return {
      model,
      resolvedChangeIds: [],
      unresolvedChangeIds: [],
      clearedOnlyChangeIds: []
    };
  }

  const accumulator: DecisionAccumulator = {
    resolvedChangeIds: [],
    unresolvedChangeIds: [],
    clearedOnlyChangeIds: []
  };

  const next = cloneDocModel(model);
  let mutated = false;

  const tableChanges = new Map<number, DocxTrackedChange[]>();
  const paragraphChanges = new Map<string, DocxTrackedChange[]>();
  for (const change of selected) {
    if (tableRevisionStableId(change)) {
      const bucket = tableChanges.get(change.nodeIndex) ?? [];
      bucket.push(change);
      tableChanges.set(change.nodeIndex, bucket);
      continue;
    }
    const key = paragraphLocationKey(change.location);
    const bucket = paragraphChanges.get(key) ?? [];
    bucket.push(change);
    paragraphChanges.set(key, bucket);
  }

  // ── Table revisions: property changes, inserted/deleted rows and cells ──
  for (const [nodeIndex, changes] of tableChanges) {
    const outcome = applyTableRevisionDecisions(next, nodeIndex, changes, decision);
    accumulator.resolvedChangeIds.push(...outcome.resolvedChangeIds);
    accumulator.unresolvedChangeIds.push(...outcome.unresolvedChangeIds);
    accumulator.clearedOnlyChangeIds.push(...outcome.clearedOnlyChangeIds);
    if (outcome.changed) {
      mutated = true;
    }
  }

  // ── Paragraph revisions ─────────────────────────────────────────────
  for (const changes of paragraphChanges.values()) {
    const first = changes[0];
    const paragraph = first
      ? resolveTrackedChangeParagraph(next, first.location)
      : undefined;
    if (!paragraph) {
      accumulator.unresolvedChangeIds.push(...changes.map((change) => change.id));
      continue;
    }

    const plan = planParagraphDecision(paragraph, changes, decision, accumulator);
    if (!plan) {
      continue;
    }

    const sourceXml = paragraph.sourceXml ?? "";
    if (plan.propertyEdits.length > 0) {
      const nextXml = applyXmlEdits(sourceXml, plan.propertyEdits);
      if (nextXml !== sourceXml) {
        paragraph.sourceXml = nextXml;
        mutated = true;
      }
    }

    if (rewriteContentRevisions(paragraph, plan, decision)) {
      mutated = true;
    }

    const rebuiltChildren = rebuildParagraphChildren(paragraph, {
      removeOrdinals: new Set(plan.removeOrdinals),
      restores: plan.restores
    });
    if (rebuiltChildren) {
      paragraph.children = rebuiltChildren;
      mutated = true;
    }

    for (const revert of plan.runStyleReverts) {
      if (revertRunStyleAtOrdinals(paragraph, revert.ordinals, revert.style) > 0) {
        mutated = true;
      }
    }

    if (plan.alignRevert) {
      paragraph.style = { ...(paragraph.style ?? {}), align: plan.alignRevert };
      mutated = true;
    }
  }

  return {
    model: mutated ? next : model,
    resolvedChangeIds: accumulator.resolvedChangeIds,
    unresolvedChangeIds: accumulator.unresolvedChangeIds,
    clearedOnlyChangeIds: accumulator.clearedOnlyChangeIds
  };
}

export function acceptTrackedChangesInModel(
  model: DocModel,
  scope?: DocxTrackedChangeDecisionScope
): DocxTrackedChangeDecisionResult {
  return applyTrackedChangeDecision(model, "accept", scope);
}

export function rejectTrackedChangesInModel(
  model: DocModel,
  scope?: DocxTrackedChangeDecisionScope
): DocxTrackedChangeDecisionResult {
  return applyTrackedChangeDecision(model, "reject", scope);
}
