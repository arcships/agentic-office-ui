// Paragraph editing operations — core helpers, paragraph-level mutation ops,
// and clipboard helpers. Framework-agnostic, pure functions operating on DocModel.
//
// Split from editor-ops.ts to stay within the ≤1000-line constraint.

import type {
  DocModel,
  FormFieldRunNode,
  HeadingLevel,
  ImageRunNode,
  ParagraphAlignment,
  ParagraphNode,
  TableCellContentNode,
  TableNode,
  ParagraphStyle,
  TextRunNode,
  TextStyle
} from "../engine/types";
import { cloneDocModel } from "../engine/clone";

export interface InsertParagraphOptions {
  paragraphStyle?: ParagraphStyle;
  runStyle?: TextStyle;
}

export interface UpdateTextOptions {
  insertedStyle?: TextStyle;
}

export function paragraphFromText(text: string, options?: InsertParagraphOptions): ParagraphNode {
  return {
    type: "paragraph",
    style: options?.paragraphStyle,
    children: [{ type: "text", text, style: options?.runStyle }]
  };
}

export function getParagraph(model: DocModel, nodeIndex: number): ParagraphNode | undefined {
  const node = model.nodes[nodeIndex];
  if (!node || node.type !== "paragraph") {
    return undefined;
  }
  return node;
}

export function ensureTextRun(paragraph: ParagraphNode, runIndex: number): TextRunNode {
  let textRunCount = -1;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }

    textRunCount += 1;
    if (textRunCount === runIndex) {
      return child;
    }
  }

  const created: TextRunNode = {
    type: "text",
    text: "",
    style: {}
  };

  paragraph.children.push(created);
  return created;
}

function textRuns(paragraph: ParagraphNode): TextRunNode[] {
  return paragraph.children.filter((child): child is TextRunNode => child.type === "text");
}

function cloneTextRun(run: TextRunNode): TextRunNode {
  return {
    type: "text",
    text: run.text,
    style: run.style ? { ...run.style } : undefined,
    link: run.link,
    noteReference: run.noteReference ? { ...run.noteReference } : undefined
  };
}

function noteReferencesEqual(left?: TextRunNode["noteReference"], right?: TextRunNode["noteReference"]): boolean {
  if (!left || !right) {
    return left === right;
  }
  return left.kind === right.kind && left.id === right.id;
}

function cloneFormFieldRun(run: FormFieldRunNode): FormFieldRunNode {
  return {
    type: "form-field",
    fieldType: run.fieldType,
    sourceKind: run.sourceKind,
    id: run.id,
    tag: run.tag,
    title: run.title,
    placeholder: run.placeholder,
    checked: run.checked,
    value: run.value,
    options: run.options?.map((option) => ({
      displayText: option.displayText,
      value: option.value
    })),
    widget: run.widget
      ? {
          name: run.widget.name,
          enabled: run.widget.enabled,
          calcOnExit: run.widget.calcOnExit,
          text: run.widget.text
            ? {
                inputType: run.widget.text.inputType,
                defaultText: run.widget.text.defaultText,
                maxLength: run.widget.text.maxLength,
                textFormat: run.widget.text.textFormat
              }
            : undefined,
          checkbox: run.widget.checkbox
            ? {
                defaultChecked: run.widget.checkbox.defaultChecked,
                sizeMode: run.widget.checkbox.sizeMode,
                sizePt: run.widget.checkbox.sizePt
              }
            : undefined,
          dropdown: run.widget.dropdown
            ? {
                defaultValue: run.widget.dropdown.defaultValue
              }
            : undefined
        }
      : undefined,
    checkedSymbol: run.checkedSymbol,
    uncheckedSymbol: run.uncheckedSymbol,
    style: run.style ? { ...run.style } : undefined,
    link: run.link,
    sourceXml: run.sourceXml
  };
}

function cloneImageRun(run: ImageRunNode): ImageRunNode {
  return {
    type: "image",
    src: run.src,
    alt: run.alt,
    widthPx: run.widthPx,
    heightPx: run.heightPx,
    partName: run.partName,
    contentType: run.contentType,
    data: run.data ? new Uint8Array(run.data) : undefined,
    floating: run.floating ? { ...run.floating } : undefined,
    syntheticTextBox: run.syntheticTextBox,
    textBoxText: run.textBoxText
  };
}

function cloneParagraphChildRun(
  run: ParagraphNode["children"][number]
): ParagraphNode["children"][number] {
  if (run.type === "text") {
    return cloneTextRun(run);
  }

  if (run.type === "form-field") {
    return cloneFormFieldRun(run);
  }

  return cloneImageRun(run);
}

function cloneTextStyle(style?: TextStyle): TextStyle | undefined {
  return style ? { ...style } : undefined;
}

function textStylesEqual(left?: TextStyle, right?: TextStyle): boolean {
  return JSON.stringify(left ?? {}) === JSON.stringify(right ?? {});
}

function mergeAdjacentRuns(runs: TextRunNode[]): TextRunNode[] {
  const merged: TextRunNode[] = [];

  for (const run of runs) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      textStylesEqual(previous.style, run.style) &&
      previous.link === run.link &&
      noteReferencesEqual(previous.noteReference, run.noteReference)
    ) {
      previous.text += run.text;
      continue;
    }

    merged.push(cloneTextRun(run));
  }

  return merged;
}

function formFieldDisplayValue(field: FormFieldRunNode): string {
  switch (field.fieldType) {
    case "checkbox":
      return field.checked ?? field.widget?.checkbox?.defaultChecked
        ? field.checkedSymbol ?? "☒"
        : field.uncheckedSymbol ?? "☐";
    case "dropdown":
      return field.value ?? field.options?.[0]?.displayText ?? "";
    case "date":
    case "text":
      return field.value ?? field.widget?.text?.defaultText ?? "";
    default:
      return "";
  }
}

function splitRunsAtOffset(
  runs: TextRunNode[],
  offset: number
): {
  left: TextRunNode[];
  right: TextRunNode[];
} {
  const safeOffset = Math.max(0, offset);
  const left: TextRunNode[] = [];
  const right: TextRunNode[] = [];
  let cursor = 0;

  for (const run of runs) {
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (runEnd <= safeOffset) {
      left.push(cloneTextRun(run));
      continue;
    }

    if (runStart >= safeOffset) {
      right.push(cloneTextRun(run));
      continue;
    }

    const localSplit = Math.max(0, Math.min(runLength, safeOffset - runStart));
    const before = run.text.slice(0, localSplit);
    const after = run.text.slice(localSplit);
    if (before.length > 0) {
      left.push({
        type: "text",
        text: before,
        style: cloneTextStyle(run.style),
        link: run.link,
        noteReference: run.noteReference ? { ...run.noteReference } : undefined
      });
    }
    if (after.length > 0) {
      right.push({
        type: "text",
        text: after,
        style: cloneTextStyle(run.style),
        link: run.link,
        noteReference: run.noteReference ? { ...run.noteReference } : undefined
      });
    }
  }

  return { left, right };
}

function commonPrefixLength(left: string, right: string): number {
  const limit = Math.min(left.length, right.length);
  let index = 0;
  while (index < limit && left[index] === right[index]) {
    index += 1;
  }
  return index;
}

function commonSuffixLength(left: string, right: string, prefixLength: number): number {
  const leftRemaining = left.length - prefixLength;
  const rightRemaining = right.length - prefixLength;
  const limit = Math.min(leftRemaining, rightRemaining);
  let matched = 0;

  while (
    matched < limit &&
    left[left.length - 1 - matched] === right[right.length - 1 - matched]
  ) {
    matched += 1;
  }

  return matched;
}

function distributeTextAcrossRuns(
  text: string,
  templateRuns: TextRunNode[],
  options?: UpdateTextOptions
): TextRunNode[] {
  const normalizedText = text ?? "";
  if (templateRuns.length === 0) {
    return [
      {
        type: "text",
        text: normalizedText,
        style: cloneTextStyle(options?.insertedStyle)
      }
    ];
  }

  const originalRuns = templateRuns.map(cloneTextRun);
  const originalText = originalRuns.map((run) => run.text).join("");
  if (originalText === normalizedText) {
    return originalRuns;
  }

  const prefixLength = commonPrefixLength(originalText, normalizedText);
  const suffixLength = commonSuffixLength(originalText, normalizedText, prefixLength);
  const removeStart = prefixLength;
  const removeEnd = Math.max(removeStart, originalText.length - suffixLength);
  const insertedText = normalizedText.slice(prefixLength, normalizedText.length - suffixLength);

  const splitBefore = splitRunsAtOffset(originalRuns, removeStart);
  const splitAfter = splitRunsAtOffset(splitBefore.right, removeEnd - removeStart);
  const leftRuns = splitBefore.left;
  const rightRuns = splitAfter.right;

  const insertedRuns: TextRunNode[] = [];
  if (insertedText.length > 0) {
    const previousRun = leftRuns[leftRuns.length - 1];
    const nextRun = rightRuns[0];
    const inferredLink =
      previousRun?.link && nextRun?.link && previousRun.link === nextRun.link
        ? previousRun.link
        : previousRun?.link ?? nextRun?.link;

    insertedRuns.push({
      type: "text",
      text: insertedText,
      style: cloneTextStyle(options?.insertedStyle ?? previousRun?.style ?? nextRun?.style),
      link: inferredLink
    });
  }

  const merged = mergeAdjacentRuns([...leftRuns, ...insertedRuns, ...rightRuns]);
  if (merged.length > 0) {
    return merged;
  }

  return [
    {
      type: "text",
      text: "",
      style: cloneTextStyle(options?.insertedStyle ?? templateRuns[0]?.style),
      link: templateRuns[0]?.link
    }
  ];
}

export function distributeTextAcrossParagraphChildren(
  paragraph: ParagraphNode,
  text: string,
  options?: UpdateTextOptions
): ParagraphNode["children"] {
  const hasNonTextRuns = paragraph.children.some((child) => child.type !== "text");
  if (!hasNonTextRuns) {
    return distributeTextAcrossRuns(text, textRuns(paragraph), options);
  }

  const textGroups: TextRunNode[][] = [];
  const anchors: Array<Exclude<ParagraphNode["children"][number], TextRunNode>> = [];
  let currentGroup: TextRunNode[] = [];

  for (const child of paragraph.children) {
    if (child.type === "text") {
      currentGroup.push(cloneTextRun(child));
      continue;
    }

    textGroups.push(currentGroup);
    currentGroup = [];
    anchors.push(
      child.type === "form-field"
        ? cloneFormFieldRun(child)
        : cloneImageRun(child)
    );
  }
  textGroups.push(currentGroup);

  const allAnchorsAreImages = anchors.length > 0 && anchors.every((anchor) => anchor.type === "image");
  if (allAnchorsAreImages) {
    const originalSegmentTexts = textGroups.map((group) => group.map((run) => run.text).join(""));
    const originalText = originalSegmentTexts.join("");
    const originalAnchorOffsets: number[] = [];
    let originalOffsetCursor = 0;
    for (let index = 0; index < anchors.length; index += 1) {
      originalOffsetCursor += originalSegmentTexts[index]?.length ?? 0;
      originalAnchorOffsets.push(originalOffsetCursor);
    }

    const prefixLength = (() => {
      const limit = Math.min(originalText.length, text.length);
      let index = 0;
      while (index < limit && originalText[index] === text[index]) {
        index += 1;
      }
      return index;
    })();
    const suffixLength = (() => {
      const remainingOriginal = originalText.length - prefixLength;
      const remainingNext = text.length - prefixLength;
      const limit = Math.min(remainingOriginal, remainingNext);
      let index = 0;
      while (
        index < limit &&
        originalText[originalText.length - 1 - index] === text[text.length - 1 - index]
      ) {
        index += 1;
      }
      return index;
    })();
    const replacedOriginalStart = prefixLength;
    const replacedOriginalEnd = Math.max(replacedOriginalStart, originalText.length - suffixLength);
    const replacedNextEnd = Math.max(replacedOriginalStart, text.length - suffixLength);
    const delta = replacedNextEnd - replacedOriginalEnd;
    const remappedAnchorOffsets = originalAnchorOffsets.map((anchorOffset) => {
      if (anchorOffset < replacedOriginalStart) {
        return anchorOffset;
      }
      if (anchorOffset >= replacedOriginalEnd) {
        return anchorOffset + delta;
      }
      return replacedOriginalStart;
    });

    const segments: string[] = [];
    let cursor = 0;
    remappedAnchorOffsets.forEach((anchorOffset) => {
      const safeAnchorOffset = Math.max(cursor, Math.min(anchorOffset, text.length));
      segments.push(text.slice(cursor, safeAnchorOffset));
      cursor = safeAnchorOffset;
    });
    segments.push(text.slice(cursor));

    const nextChildren: ParagraphNode["children"] = [];
    for (let index = 0; index < textGroups.length; index += 1) {
      const templateRuns = textGroups[index];
      const segmentText = segments[index] ?? "";

      if (templateRuns.length > 0) {
        nextChildren.push(...distributeTextAcrossRuns(segmentText, templateRuns, options));
      } else if (segmentText.length > 0) {
        nextChildren.push({
          type: "text",
          text: segmentText,
          style: cloneTextStyle(options?.insertedStyle)
        });
      }

      if (index < anchors.length) {
        nextChildren.push(cloneParagraphChildRun(anchors[index]));
      }
    }

    if (nextChildren.length > 0) {
      return nextChildren;
    }
  }

  const segments: string[] = [];
  let cursor = 0;
  for (let index = 0; index < anchors.length; index += 1) {
    const anchor = anchors[index];
    const anchorText = anchor.type === "form-field" ? formFieldDisplayValue(anchor) : "";
    if (!anchorText) {
      return distributeTextAcrossRuns(text, textRuns(paragraph), options);
    }

    const anchorIndex = text.indexOf(anchorText, cursor);
    if (anchorIndex < 0) {
      segments.push(text.slice(cursor));
      cursor = text.length;
      break;
    }

    segments.push(text.slice(cursor, anchorIndex));
    cursor = anchorIndex + anchorText.length;
  }

  segments.push(text.slice(cursor));
  while (segments.length < textGroups.length) {
    segments.push("");
  }

  const nextChildren: ParagraphNode["children"] = [];
  for (let index = 0; index < textGroups.length; index += 1) {
    const templateRuns = textGroups[index];
    const segmentText = segments[index] ?? "";

    if (templateRuns.length > 0) {
      nextChildren.push(...distributeTextAcrossRuns(segmentText, templateRuns, options));
    } else if (segmentText.length > 0) {
      nextChildren.push({
        type: "text",
        text: segmentText,
        style: cloneTextStyle(options?.insertedStyle)
      });
    }

    if (index < anchors.length) {
      nextChildren.push(cloneParagraphChildRun(anchors[index]));
    }
  }

  if (nextChildren.length > 0) {
    return nextChildren;
  }

  return [
    {
      type: "text",
      text: "",
      style: cloneTextStyle(options?.insertedStyle)
    }
  ];
}

export function splitParagraphChildrenAtTextOffsets(
  paragraph: ParagraphNode,
  text: string,
  startOffset: number,
  endOffset: number,
  options?: {
    beforeInsertedStyle?: TextStyle;
    afterInsertedStyle?: TextStyle;
  }
): {
  beforeChildren: ParagraphNode["children"];
  afterChildren: ParagraphNode["children"];
} {
  const normalizedText = text ?? "";
  const safeStart = Math.max(0, Math.min(Math.round(startOffset), normalizedText.length));
  const safeEnd = Math.max(safeStart, Math.min(Math.round(endOffset), normalizedText.length));
  const hasNonTextRuns = paragraph.children.some((child) => child.type !== "text");

  if (!hasNonTextRuns) {
    return {
      beforeChildren: distributeTextAcrossRuns(
        normalizedText.slice(0, safeStart),
        textRuns(paragraph),
        { insertedStyle: options?.beforeInsertedStyle }
      ),
      afterChildren: distributeTextAcrossRuns(
        normalizedText.slice(safeEnd),
        textRuns(paragraph),
        { insertedStyle: options?.afterInsertedStyle }
      )
    };
  }

  const textGroups: TextRunNode[][] = [];
  const anchors: Array<Exclude<ParagraphNode["children"][number], TextRunNode>> = [];
  let currentGroup: TextRunNode[] = [];

  for (const child of paragraph.children) {
    if (child.type === "text") {
      currentGroup.push(cloneTextRun(child));
      continue;
    }

    textGroups.push(currentGroup);
    currentGroup = [];
    anchors.push(
      child.type === "form-field"
        ? cloneFormFieldRun(child)
        : cloneImageRun(child)
    );
  }
  textGroups.push(currentGroup);

  let segments: string[] | undefined;
  let anchorOffsets: number[] | undefined;

  const allAnchorsAreImages = anchors.length > 0 && anchors.every((anchor) => anchor.type === "image");
  if (allAnchorsAreImages) {
    const originalSegmentTexts = textGroups.map((group) => group.map((run) => run.text).join(""));
    const originalText = originalSegmentTexts.join("");
    const originalAnchorOffsets: number[] = [];
    let originalOffsetCursor = 0;
    for (let index = 0; index < anchors.length; index += 1) {
      originalOffsetCursor += originalSegmentTexts[index]?.length ?? 0;
      originalAnchorOffsets.push(originalOffsetCursor);
    }

    const prefixLength = commonPrefixLength(originalText, normalizedText);
    const suffixLength = commonSuffixLength(originalText, normalizedText, prefixLength);
    const replacedOriginalStart = prefixLength;
    const replacedOriginalEnd = Math.max(replacedOriginalStart, originalText.length - suffixLength);
    const replacedNextEnd = Math.max(replacedOriginalStart, normalizedText.length - suffixLength);
    const delta = replacedNextEnd - replacedOriginalEnd;
    const remappedAnchorOffsets = originalAnchorOffsets.map((anchorOffset) => {
      if (anchorOffset < replacedOriginalStart) {
        return anchorOffset;
      }
      if (anchorOffset >= replacedOriginalEnd) {
        return anchorOffset + delta;
      }
      return replacedOriginalStart;
    });

    const nextSegments: string[] = [];
    let cursor = 0;
    remappedAnchorOffsets.forEach((anchorOffset) => {
      const safeAnchorOffset = Math.max(cursor, Math.min(anchorOffset, normalizedText.length));
      nextSegments.push(normalizedText.slice(cursor, safeAnchorOffset));
      cursor = safeAnchorOffset;
    });
    nextSegments.push(normalizedText.slice(cursor));
    segments = nextSegments;
    anchorOffsets = remappedAnchorOffsets;
  } else {
    const nextSegments: string[] = [];
    const nextAnchorOffsets: number[] = [];
    let cursor = 0;

    for (let index = 0; index < anchors.length; index += 1) {
      const anchor = anchors[index];
      const anchorText = anchor.type === "form-field" ? formFieldDisplayValue(anchor) : "";
      if (!anchorText) {
        segments = undefined;
        anchorOffsets = undefined;
        break;
      }

      const anchorIndex = normalizedText.indexOf(anchorText, cursor);
      if (anchorIndex < 0) {
        nextSegments.push(normalizedText.slice(cursor));
        cursor = normalizedText.length;
        break;
      }

      nextSegments.push(normalizedText.slice(cursor, anchorIndex));
      cursor = anchorIndex + anchorText.length;
      nextAnchorOffsets.push(anchorIndex);
    }

    if (nextSegments.length > 0 || anchors.length === 0) {
      nextSegments.push(normalizedText.slice(cursor));
      while (nextSegments.length < textGroups.length) {
        nextSegments.push("");
      }
      segments = nextSegments;
      anchorOffsets = nextAnchorOffsets;
    }
  }

  if (!segments || !anchorOffsets) {
    return {
      beforeChildren: distributeTextAcrossRuns(
        normalizedText.slice(0, safeStart),
        textRuns(paragraph),
        { insertedStyle: options?.beforeInsertedStyle }
      ),
      afterChildren: distributeTextAcrossRuns(
        normalizedText.slice(safeEnd),
        textRuns(paragraph),
        { insertedStyle: options?.afterInsertedStyle }
      )
    };
  }

  const beforeChildren: ParagraphNode["children"] = [];
  const afterChildren: ParagraphNode["children"] = [];
  let cursor = 0;

  for (let index = 0; index < textGroups.length; index += 1) {
    const templateRuns = textGroups[index];
    const segmentText = segments[index] ?? "";
    const segmentStart = cursor;
    const segmentEnd = segmentStart + segmentText.length;
    cursor = segmentEnd;

    const beforePart =
      safeStart <= segmentStart
        ? ""
        : segmentText.slice(0, Math.max(0, Math.min(segmentText.length, safeStart - segmentStart)));
    const afterPart =
      safeEnd >= segmentEnd
        ? ""
        : segmentText.slice(Math.max(0, Math.min(segmentText.length, safeEnd - segmentStart)));

    if (templateRuns.length > 0) {
      if (beforePart.length > 0) {
        beforeChildren.push(
          ...distributeTextAcrossRuns(beforePart, templateRuns, {
            insertedStyle: options?.beforeInsertedStyle
          })
        );
      }
      if (afterPart.length > 0) {
        afterChildren.push(
          ...distributeTextAcrossRuns(afterPart, templateRuns, {
            insertedStyle: options?.afterInsertedStyle
          })
        );
      }
    } else {
      if (beforePart.length > 0) {
        beforeChildren.push({
          type: "text",
          text: beforePart,
          style: cloneTextStyle(options?.beforeInsertedStyle)
        });
      }
      if (afterPart.length > 0) {
        afterChildren.push({
          type: "text",
          text: afterPart,
          style: cloneTextStyle(options?.afterInsertedStyle)
        });
      }
    }

    if (index < anchors.length) {
      const anchor = cloneParagraphChildRun(anchors[index]);
      const anchorOffset = anchorOffsets[index] ?? segmentEnd;
      if (anchorOffset <= safeStart) {
        beforeChildren.push(anchor);
      } else if (anchorOffset >= safeEnd) {
        afterChildren.push(anchor);
      } else {
        beforeChildren.push(anchor);
      }
    }
  }

  if (beforeChildren.length === 0) {
    beforeChildren.push({
      type: "text",
      text: "",
      style: cloneTextStyle(options?.beforeInsertedStyle)
    });
  }

  if (afterChildren.length === 0) {
    afterChildren.push({
      type: "text",
      text: "",
      style: cloneTextStyle(options?.afterInsertedStyle)
    });
  }

  return {
    beforeChildren,
    afterChildren
  };
}

export function cloneParagraph(paragraph: ParagraphNode): ParagraphNode {
  return {
    type: "paragraph",
    style: paragraph.style ? { ...paragraph.style } : undefined,
    sourceXml: paragraph.sourceXml,
    children: paragraph.children.map(cloneParagraphChildRun)
  };
}

export function insertParagraph(
  model: DocModel,
  text: string,
  index = model.nodes.length,
  options?: InsertParagraphOptions
): DocModel {
  const next = cloneDocModel(model);
  const safeIndex = Math.max(0, Math.min(index, next.nodes.length));
  next.nodes.splice(safeIndex, 0, paragraphFromText(text, options));
  return next;
}

export function removeParagraph(model: DocModel, index: number): DocModel {
  const next = cloneDocModel(model);
  const node = getParagraph(next, index);
  if (!node) {
    return next;
  }

  next.nodes.splice(index, 1);

  if (!next.nodes.some((candidate) => candidate.type === "paragraph")) {
    next.nodes.push(paragraphFromText(""));
  }

  return next;
}

export function duplicateParagraph(model: DocModel, index: number): DocModel {
  const next = cloneDocModel(model);
  const node = getParagraph(next, index);
  if (!node) {
    return next;
  }

  next.nodes.splice(index + 1, 0, cloneParagraph(node));
  return next;
}

export function updateParagraphText(
  model: DocModel,
  index: number,
  text: string,
  options?: UpdateTextOptions
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, index);
  if (!paragraph) {
    return next;
  }

  paragraph.children = distributeTextAcrossParagraphChildren(paragraph, text, options);
  paragraph.sourceXml = undefined;

  return next;
}

