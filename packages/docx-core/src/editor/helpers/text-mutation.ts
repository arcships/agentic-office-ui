// Text style/link cloning and range mutation, form-field widget cloning,
// text-run splitting, and image drop-target parsing.
// Upstream editor.tsx: lines 21915-22673.

import type {
  DocModel,
  FormFieldRunNode,
  ImageRunNode,
  ParagraphNode,
  TableNode,
  TextRunNode,
  TextStyle
} from "../../engine/types";
import type {
  DocxEditorSelection,
  DocxTextRange,
  DocxTextRangeLocation,
  DocxImageDropTarget,
  DocxImageLocation,
  ParagraphLocation
} from "./editor-types";

// Minimal ambient declaration for the Node.js `Buffer` global. `toBase64` below
// uses `typeof Buffer !== "undefined"` to pick the Node fast path at runtime;
// this keeps typechecking portable without pulling in @types/node.
declare const Buffer: {
  from(data: Uint8Array): { toString(encoding: string): string };
} | undefined;

export function cloneTextStyle(
  style?: TextRunNode["style"]
): TextRunNode["style"] | undefined {
  return style ? { ...style } : undefined;
}

export function cloneFormFieldWidget(
  widget?: FormFieldRunNode["widget"]
): FormFieldRunNode["widget"] | undefined {
  if (!widget) {
    return undefined;
  }

  return {
    name: widget.name,
    enabled: widget.enabled,
    calcOnExit: widget.calcOnExit,
    text: widget.text
      ? {
          inputType: widget.text.inputType,
          defaultText: widget.text.defaultText,
          maxLength: widget.text.maxLength,
          textFormat: widget.text.textFormat,
        }
      : undefined,
    checkbox: widget.checkbox
      ? {
          defaultChecked: widget.checkbox.defaultChecked,
          sizeMode: widget.checkbox.sizeMode,
          sizePt: widget.checkbox.sizePt,
        }
      : undefined,
    dropdown: widget.dropdown
      ? {
          defaultValue: widget.dropdown.defaultValue,
        }
      : undefined,
  };
}

export function mergeFormFieldWidgetPatch(
  current: FormFieldRunNode["widget"] | undefined,
  patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
): FormFieldRunNode["widget"] | undefined {
  const hasPatch =
    patch.name !== undefined ||
    patch.enabled !== undefined ||
    patch.calcOnExit !== undefined ||
    patch.text !== undefined ||
    patch.checkbox !== undefined ||
    patch.dropdown !== undefined;
  if (!hasPatch) {
    return cloneFormFieldWidget(current);
  }

  const mergedText =
    patch.text === undefined
      ? cloneFormFieldWidget(current)?.text
      : {
          ...(current?.text ?? {}),
          ...patch.text,
        };
  const mergedCheckbox =
    patch.checkbox === undefined
      ? cloneFormFieldWidget(current)?.checkbox
      : {
          ...(current?.checkbox ?? {}),
          ...patch.checkbox,
        };
  const mergedDropdown =
    patch.dropdown === undefined
      ? cloneFormFieldWidget(current)?.dropdown
      : {
          ...(current?.dropdown ?? {}),
          ...patch.dropdown,
        };

  return {
    ...(current ?? {}),
    ...(patch.name !== undefined ? { name: patch.name } : undefined),
    ...(patch.enabled !== undefined ? { enabled: patch.enabled } : undefined),
    ...(patch.calcOnExit !== undefined
      ? { calcOnExit: patch.calcOnExit }
      : undefined),
    ...(mergedText ? { text: mergedText } : undefined),
    ...(mergedCheckbox ? { checkbox: mergedCheckbox } : undefined),
    ...(mergedDropdown ? { dropdown: mergedDropdown } : undefined),
  };
}

export function cloneFormFieldRun(field: FormFieldRunNode): FormFieldRunNode {
  return {
    type: "form-field",
    fieldType: field.fieldType,
    sourceKind: field.sourceKind,
    id: field.id,
    tag: field.tag,
    title: field.title,
    placeholder: field.placeholder,
    checked: field.checked,
    value: field.value,
    options: field.options?.map((option) => ({
      displayText: option.displayText,
      value: option.value,
    })),
    widget: cloneFormFieldWidget(field.widget),
    checkedSymbol: field.checkedSymbol,
    uncheckedSymbol: field.uncheckedSymbol,
    style: cloneTextStyle(field.style),
    link: field.link,
    sourceXml: field.sourceXml,
  };
}

export function textStylesEqual(
  a?: TextRunNode["style"],
  b?: TextRunNode["style"]
): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});
}

export function mergeAdjacentTextRuns(
  children: ParagraphNode["children"]
): ParagraphNode["children"] {
  const merged: ParagraphNode["children"] = [];

  for (const child of children) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      previous.type === "text" &&
      child.type === "text" &&
      textStylesEqual(previous.style, child.style) &&
      previous.link === child.link
    ) {
      previous.text += child.text;
      continue;
    }

    merged.push(
      child.type === "text"
        ? {
            type: "text",
            text: child.text,
            style: cloneTextStyle(child.style),
            link: child.link,
          }
        : child.type === "form-field"
        ? cloneFormFieldRun(child)
        : {
            type: "image",
            src: child.src,
            alt: child.alt,
            widthPx: child.widthPx,
            heightPx: child.heightPx,
            partName: child.partName,
            contentType: child.contentType,
            data: child.data ? new Uint8Array(child.data) : undefined,
            floating: child.floating ? { ...child.floating } : undefined,
            syntheticTextBox: child.syntheticTextBox,
            textBoxText: child.textBoxText,
          }
    );
  }

  return merged;
}

export function paragraphHasOnlyTextRuns(paragraph: ParagraphNode): boolean {
  return paragraph.children.every(
    (child): child is TextRunNode => child.type === "text"
  );
}

export function cloneTextRunWithMetadata(run: TextRunNode): TextRunNode {
  return {
    type: "text",
    text: run.text,
    style: cloneTextStyle(run.style),
    link: run.link,
    noteReference: run.noteReference ? { ...run.noteReference } : undefined,
  };
}

export function splitTextRunsAtOffset(
  runs: TextRunNode[],
  offset: number
): {
  left: TextRunNode[];
  right: TextRunNode[];
} {
  const safeOffset = Math.max(0, Math.round(offset));
  const left: TextRunNode[] = [];
  const right: TextRunNode[] = [];
  let cursor = 0;

  for (const run of runs) {
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (runEnd <= safeOffset) {
      left.push(cloneTextRunWithMetadata(run));
      continue;
    }

    if (runStart >= safeOffset) {
      right.push(cloneTextRunWithMetadata(run));
      continue;
    }

    const localSplit = Math.max(0, Math.min(runLength, safeOffset - runStart));
    const before = run.text.slice(0, localSplit);
    const after = run.text.slice(localSplit);
    if (before.length > 0) {
      left.push({
        ...cloneTextRunWithMetadata(run),
        text: before,
      });
    }
    if (after.length > 0) {
      right.push({
        ...cloneTextRunWithMetadata(run),
        text: after,
      });
    }
  }

  return { left, right };
}

export function firstTextStyleAtOffset(
  paragraph: ParagraphNode,
  offset: number,
  preferPreviousAtBoundary: boolean
): TextRunNode["style"] | undefined {
  const textChildren = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textChildren.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(0, offset);
  let cursor = 0;

  for (let index = 0; index < textChildren.length; index += 1) {
    const run = textChildren[index];
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (safeOffset < runEnd) {
      return cloneTextStyle(run.style);
    }

    if (safeOffset === runEnd) {
      if (preferPreviousAtBoundary || index === textChildren.length - 1) {
        return cloneTextStyle(run.style);
      }

      return cloneTextStyle(textChildren[index + 1]?.style);
    }
  }

  return cloneTextStyle(textChildren[textChildren.length - 1]?.style);
}

export function linkAtOffset(
  paragraph: ParagraphNode,
  offset: number,
  preferPreviousAtBoundary: boolean
): string | undefined {
  const textChildren = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textChildren.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(0, offset);
  let cursor = 0;

  for (let index = 0; index < textChildren.length; index += 1) {
    const run = textChildren[index];
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (safeOffset < runEnd) {
      return run.link;
    }

    if (safeOffset === runEnd) {
      if (preferPreviousAtBoundary || index === textChildren.length - 1) {
        return run.link;
      }

      return textChildren[index + 1]?.link;
    }
  }

  return textChildren[textChildren.length - 1]?.link;
}

export function uniformLinkInRange(
  paragraph: ParagraphNode,
  startOffset: number,
  endOffset: number
): string | undefined {
  const safeStart = Math.max(0, Math.min(startOffset, endOffset));
  const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
  if (safeStart === safeEnd) {
    return linkAtOffset(paragraph, safeStart, true);
  }

  let cursor = 0;
  let candidateLink: string | undefined;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }

    const runStart = cursor;
    const runEnd = runStart + child.text.length;
    cursor = runEnd;

    if (runEnd <= safeStart || runStart >= safeEnd) {
      continue;
    }

    if (!candidateLink) {
      candidateLink = child.link;
      continue;
    }

    if (candidateLink !== child.link) {
      return undefined;
    }
  }

  return candidateLink;
}

export function linkRangeAtOffset(
  paragraph: ParagraphNode,
  offset: number
):
  | {
      start: number;
      end: number;
      link: string;
    }
  | undefined {
  const textChildren = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textChildren.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(0, offset);
  let cursor = 0;
  let runIndex = -1;

  for (let index = 0; index < textChildren.length; index += 1) {
    const run = textChildren[index];
    const runStart = cursor;
    const runEnd = runStart + run.text.length;
    cursor = runEnd;

    if (
      safeOffset < runEnd ||
      safeOffset === runEnd ||
      index === textChildren.length - 1
    ) {
      runIndex = index;
      break;
    }
  }

  if (runIndex < 0) {
    return undefined;
  }

  const currentRun = textChildren[runIndex];
  const currentLink = currentRun.link;
  if (!currentLink) {
    return undefined;
  }

  const runStarts: number[] = [];
  cursor = 0;
  for (const run of textChildren) {
    runStarts.push(cursor);
    cursor += run.text.length;
  }

  let startIndex = runIndex;
  while (startIndex > 0 && textChildren[startIndex - 1]?.link === currentLink) {
    startIndex -= 1;
  }

  let endIndex = runIndex;
  while (
    endIndex + 1 < textChildren.length &&
    textChildren[endIndex + 1]?.link === currentLink
  ) {
    endIndex += 1;
  }

  const start = runStarts[startIndex] ?? 0;
  const end =
    (runStarts[endIndex] ?? 0) + (textChildren[endIndex]?.text.length ?? 0);
  if (end <= start) {
    return undefined;
  }

  return {
    start,
    end,
    link: currentLink,
  };
}

export function mutateParagraphTextStyleInRange(
  paragraph: ParagraphNode,
  startOffset: number,
  endOffset: number,
  mutator: (
    currentStyle: TextRunNode["style"] | undefined
  ) => TextRunNode["style"] | undefined
): boolean {
  const safeStart = Math.max(0, Math.min(startOffset, endOffset));
  const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
  if (safeStart === safeEnd) {
    return false;
  }

  const nextChildren: ParagraphNode["children"] = [];
  let cursor = 0;
  let touched = false;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      if (child.type === "form-field") {
        nextChildren.push(cloneFormFieldRun(child));
      } else {
        nextChildren.push({
          type: "image",
          src: child.src,
          alt: child.alt,
          widthPx: child.widthPx,
          heightPx: child.heightPx,
          partName: child.partName,
          contentType: child.contentType,
          data: child.data ? new Uint8Array(child.data) : undefined,
          floating: child.floating ? { ...child.floating } : undefined,
          syntheticTextBox: child.syntheticTextBox,
          textBoxText: child.textBoxText,
        });
      }
      continue;
    }

    const text = child.text;
    const runStart = cursor;
    const runEnd = runStart + text.length;
    cursor = runEnd;

    if (runEnd <= safeStart || runStart >= safeEnd || text.length === 0) {
      nextChildren.push({
        type: "text",
        text,
        style: cloneTextStyle(child.style),
        link: child.link,
      });
      continue;
    }

    touched = true;
    const localStart = Math.max(0, safeStart - runStart);
    const localEnd = Math.min(text.length, safeEnd - runStart);

    if (localStart > 0) {
      nextChildren.push({
        type: "text",
        text: text.slice(0, localStart),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }

    const selectedText = text.slice(localStart, localEnd);
    if (selectedText.length > 0) {
      nextChildren.push({
        type: "text",
        text: selectedText,
        style: mutator(cloneTextStyle(child.style)),
        link: child.link,
      });
    }

    if (localEnd < text.length) {
      nextChildren.push({
        type: "text",
        text: text.slice(localEnd),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }
  }

  if (!touched) {
    return false;
  }

  paragraph.children = mergeAdjacentTextRuns(nextChildren);
  return true;
}

export function mutateParagraphLinkInRange(
  paragraph: ParagraphNode,
  startOffset: number,
  endOffset: number,
  link?: string
): boolean {
  const safeStart = Math.max(0, Math.min(startOffset, endOffset));
  const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
  if (safeStart === safeEnd) {
    return false;
  }

  const nextChildren: ParagraphNode["children"] = [];
  let cursor = 0;
  let touched = false;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      if (child.type === "form-field") {
        nextChildren.push(cloneFormFieldRun(child));
      } else {
        nextChildren.push({
          type: "image",
          src: child.src,
          alt: child.alt,
          widthPx: child.widthPx,
          heightPx: child.heightPx,
          partName: child.partName,
          contentType: child.contentType,
          data: child.data ? new Uint8Array(child.data) : undefined,
          floating: child.floating ? { ...child.floating } : undefined,
          syntheticTextBox: child.syntheticTextBox,
          textBoxText: child.textBoxText,
        });
      }
      continue;
    }

    const text = child.text;
    const runStart = cursor;
    const runEnd = runStart + text.length;
    cursor = runEnd;

    if (runEnd <= safeStart || runStart >= safeEnd || text.length === 0) {
      nextChildren.push({
        type: "text",
        text,
        style: cloneTextStyle(child.style),
        link: child.link,
      });
      continue;
    }

    touched = true;
    const localStart = Math.max(0, safeStart - runStart);
    const localEnd = Math.min(text.length, safeEnd - runStart);

    if (localStart > 0) {
      nextChildren.push({
        type: "text",
        text: text.slice(0, localStart),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }

    const selectedText = text.slice(localStart, localEnd);
    if (selectedText.length > 0) {
      nextChildren.push({
        type: "text",
        text: selectedText,
        style: cloneTextStyle(child.style),
        link,
      });
    }

    if (localEnd < text.length) {
      nextChildren.push({
        type: "text",
        text: text.slice(localEnd),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }
  }

  if (!touched) {
    return false;
  }

  paragraph.children = mergeAdjacentTextRuns(nextChildren);
  return true;
}

export function toBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  if (typeof btoa === "function") {
    return btoa(binary);
  }

  return "";
}

export function isParagraphSelected(
  selection: DocxEditorSelection,
  nodeIndex: number
): boolean {
  return selection.kind === "paragraph" && selection.nodeIndex === nodeIndex;
}

export function isCellSelected(
  selection: DocxEditorSelection,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number
): boolean {
  return (
    selection.kind === "table-cell" &&
    selection.tableIndex === tableIndex &&
    selection.rowIndex === rowIndex &&
    selection.cellIndex === cellIndex
  );
}

export function paragraphLocationKey(location: ParagraphLocation): string {
  if (location.kind === "paragraph") {
    return `p:${location.nodeIndex}`;
  }

  return `t:${location.tableIndex}:${location.rowIndex}:${location.cellIndex}:${location.paragraphIndex}`;
}

export function imageLocationKey(location: DocxImageLocation): string {
  return `${paragraphLocationKey(location)}:${location.childIndex}`;
}

export function dropTargetKey(target: DocxImageDropTarget): string {
  return `${paragraphLocationKey(target)}:${target.childIndex}`;
}

export function parseImageDropTargetFromDataset(
  dataset: DOMStringMap
): DocxImageDropTarget | undefined {
  const kind = dataset.docxTargetKind;
  const childIndex = Number.parseInt(dataset.docxChildIndex ?? "", 10);
  if (!Number.isFinite(childIndex) || childIndex < 0) {
    return undefined;
  }

  if (kind === "paragraph") {
    const nodeIndex = Number.parseInt(dataset.docxNodeIndex ?? "", 10);
    if (!Number.isFinite(nodeIndex) || nodeIndex < 0) {
      return undefined;
    }

    return {
      kind: "paragraph",
      nodeIndex,
      childIndex,
    };
  }

  if (kind === "table-cell") {
    const tableIndex = Number.parseInt(dataset.docxTableIndex ?? "", 10);
    const rowIndex = Number.parseInt(dataset.docxRowIndex ?? "", 10);
    const cellIndex = Number.parseInt(dataset.docxCellIndex ?? "", 10);
    const paragraphIndex = Number.parseInt(
      dataset.docxParagraphIndex ?? "",
      10
    );
    if (
      !Number.isFinite(tableIndex) ||
      !Number.isFinite(rowIndex) ||
      !Number.isFinite(cellIndex) ||
      !Number.isFinite(paragraphIndex) ||
      tableIndex < 0 ||
      rowIndex < 0 ||
      cellIndex < 0 ||
      paragraphIndex < 0
    ) {
      return undefined;
    }

    return {
      kind: "table-cell",
      tableIndex,
      rowIndex,
      cellIndex,
      paragraphIndex,
      childIndex,
    };
  }

  return undefined;
}

export function firstTableCellAnchorLocation(
  table: TableNode,
  tableIndex: number
): Extract<DocxTextRangeLocation, { kind: "table-cell" }> | undefined {
  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      if (!cell) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0,
      };
    }
  }

  return undefined;
}
