// Run-style editing operations + remaining paragraph mutation + clipboard
// helpers. Split from editor-ops.ts to stay within the ≤1000-line constraint.

import type {
  DocModel,
  FormFieldRunNode,
  HeadingLevel,
  ImageRunNode,
  ParagraphAlignment,
  ParagraphNode,
  TextRunNode,
  TextStyle,
} from "../engine/types";
import { cloneDocModel } from "../engine/clone";
import {
  cloneParagraph,
  ensureTextRun,
  getParagraph,
} from "./paragraph-ops";
import {
  cloneFormFieldRun,
  cloneTextStyle,
} from "./helpers/text-mutation";

function mutateParagraphTextRuns(model: DocModel, transform: (run: TextRunNode) => void): void {
  for (const node of model.nodes) {
    if (node.type === "paragraph") {
      for (const child of node.children) {
        if (child.type === "text") {
          transform(child);
        }
      }
      node.sourceXml = undefined;
      continue;
    }

      for (const row of node.rows) {
      for (const cell of row.cells) {
        for (const paragraph of cell.nodes.filter((child): child is ParagraphNode => child.type === "paragraph")) {
          for (const child of paragraph.children) {
            if (child.type === "text") {
              transform(child);
            }
          }
          paragraph.sourceXml = undefined;
        }
      }
    }
    node.sourceXml = undefined;
  }
}

export function replaceText(model: DocModel, searchValue: string | RegExp, replacement: string): DocModel {
  const next = cloneDocModel(model);

  mutateParagraphTextRuns(next, (run) => {
    run.text = run.text.replace(searchValue, replacement);
  });

  return next;
}

export function setParagraphHeading(
  model: DocModel,
  nodeIndex: number,
  headingLevel?: HeadingLevel
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, nodeIndex);
  if (!paragraph) {
    return next;
  }

  paragraph.style = {
    ...(paragraph.style ?? {}),
    headingLevel
  };
  paragraph.sourceXml = undefined;

  return next;
}

export function setParagraphAlignment(
  model: DocModel,
  nodeIndex: number,
  align?: ParagraphAlignment
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, nodeIndex);
  if (!paragraph) {
    return next;
  }

  paragraph.style = {
    ...(paragraph.style ?? {}),
    align
  };
  paragraph.sourceXml = undefined;

  return next;
}

export function applyRunStyle(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  style: Partial<TextStyle>
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, nodeIndex);
  if (!paragraph) {
    return next;
  }

  const textRun = ensureTextRun(paragraph, runIndex);
  textRun.style = {
    ...(textRun.style ?? {}),
    ...style
  };
  paragraph.sourceXml = undefined;

  return next;
}

export function toggleRunStyleFlag(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  key: "bold" | "italic" | "underline" | "strike"
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, nodeIndex);
  if (!paragraph) {
    return next;
  }

  const textRun = ensureTextRun(paragraph, runIndex);
  const current = Boolean(textRun.style?.[key]);
  textRun.style = {
    ...(textRun.style ?? {}),
    [key]: !current
  };
  paragraph.sourceXml = undefined;

  return next;
}

export function setRunHighlight(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  highlight?: string
): DocModel {
  return applyRunStyle(model, nodeIndex, runIndex, { highlight });
}

export function setRunColor(
  model: DocModel,
  nodeIndex: number,
  runIndex: number,
  color?: string
): DocModel {
  return applyRunStyle(model, nodeIndex, runIndex, { color });
}

export function copyParagraphs(model: DocModel, startIndex: number, endIndex = startIndex): ParagraphNode[] {
  const start = Math.max(0, Math.min(startIndex, model.nodes.length - 1));
  const end = Math.max(start, Math.min(endIndex, model.nodes.length - 1));

  const paragraphs: ParagraphNode[] = [];
  for (let index = start; index <= end; index += 1) {
    const node = model.nodes[index];
    if (node?.type === "paragraph") {
      paragraphs.push(cloneParagraph(node));
    }
  }

  return paragraphs;
}

export function pasteParagraphs(model: DocModel, index: number, paragraphs: ParagraphNode[]): DocModel {
  const next = cloneDocModel(model);
  const safeIndex = Math.max(0, Math.min(index, next.nodes.length));
  const copies = paragraphs.map(cloneParagraph);
  next.nodes.splice(safeIndex, 0, ...copies);
  return next;
}

export function serializeParagraphsForClipboard(paragraphs: ParagraphNode[]): string {
  return JSON.stringify(
    paragraphs.map((paragraph) => ({
      type: "paragraph",
      style: paragraph.style ?? undefined,
      sourceXml: paragraph.sourceXml,
      children: paragraph.children.map((run) =>
        run.type === "text"
          ? {
              type: "text" as const,
              text: run.text,
              style: run.style ?? undefined,
              link: run.link
            }
          : run.type === "form-field"
            ? cloneFormFieldRun(run)
          : {
              type: "image" as const,
              src: run.src,
              alt: run.alt,
              widthPx: run.widthPx,
              heightPx: run.heightPx,
              partName: run.partName,
              contentType: run.contentType,
              data: run.data ? Array.from(run.data) : undefined,
              floating: run.floating ? { ...run.floating } : undefined,
              syntheticTextBox: run.syntheticTextBox,
              textBoxText: run.textBoxText
            }
      )
    }))
  );
}

export function parseParagraphsFromClipboard(input: string): ParagraphNode[] | undefined {
  try {
    const parsed = JSON.parse(input) as unknown;
    if (!Array.isArray(parsed)) {
      return undefined;
    }

    const normalized: ParagraphNode[] = [];
    for (const item of parsed) {
      if (!item || typeof item !== "object") {
        continue;
      }

      const value = item as Partial<ParagraphNode>;
      if (value.type !== "paragraph" || !Array.isArray(value.children)) {
        continue;
      }

      const children: ParagraphNode["children"] = [];
      for (const run of value.children) {
        if (!run || typeof run !== "object") {
          continue;
        }

        if (run.type === "text" && typeof run.text === "string") {
          children.push({
            type: "text",
            text: run.text,
            style: run.style ? { ...run.style } : undefined,
            link: typeof run.link === "string" ? run.link : undefined
          });
          continue;
        }

        if (run.type === "image") {
          children.push({
            type: "image",
            src: run.src,
            alt: run.alt,
            widthPx: run.widthPx,
            heightPx: run.heightPx,
            partName: run.partName,
            contentType: run.contentType,
            data: Array.isArray(run.data) ? new Uint8Array(run.data) : undefined,
            floating: run.floating && typeof run.floating === "object" ? { ...run.floating } : undefined,
            syntheticTextBox: Boolean(run.syntheticTextBox),
            textBoxText:
              typeof run.textBoxText === "string" ? run.textBoxText : undefined
          });
          continue;
        }

        if (run.type === "form-field") {
          const options = Array.isArray(run.options)
            ? run.options.reduce<NonNullable<FormFieldRunNode["options"]>>((collected, option) => {
                if (!option || typeof option !== "object") {
                  return collected;
                }

                const displayText = typeof option.displayText === "string" ? option.displayText : undefined;
                if (!displayText) {
                  return collected;
                }

                collected.push({
                  displayText,
                  value: typeof option.value === "string" ? option.value : undefined
                });
                return collected;
              }, [])
            : undefined;
          const widget =
            run.widget && typeof run.widget === "object"
              ? {
                  name: typeof run.widget.name === "string" ? run.widget.name : undefined,
                  enabled: typeof run.widget.enabled === "boolean" ? run.widget.enabled : undefined,
                  calcOnExit:
                    typeof run.widget.calcOnExit === "boolean" ? run.widget.calcOnExit : undefined,
                  text:
                    run.widget.text && typeof run.widget.text === "object"
                      ? {
                          inputType:
                            typeof run.widget.text.inputType === "string"
                              ? run.widget.text.inputType
                              : undefined,
                          defaultText:
                            typeof run.widget.text.defaultText === "string"
                              ? run.widget.text.defaultText
                              : undefined,
                          maxLength:
                            typeof run.widget.text.maxLength === "number"
                              ? run.widget.text.maxLength
                              : undefined,
                          textFormat:
                            typeof run.widget.text.textFormat === "string"
                              ? run.widget.text.textFormat
                              : undefined
                        }
                      : undefined,
                  checkbox:
                    run.widget.checkbox && typeof run.widget.checkbox === "object"
                      ? {
                          defaultChecked:
                            typeof run.widget.checkbox.defaultChecked === "boolean"
                              ? run.widget.checkbox.defaultChecked
                              : undefined,
                          sizeMode:
                            run.widget.checkbox.sizeMode === "auto" ||
                            run.widget.checkbox.sizeMode === "exact"
                              ? run.widget.checkbox.sizeMode
                              : undefined,
                          sizePt:
                            typeof run.widget.checkbox.sizePt === "number"
                              ? run.widget.checkbox.sizePt
                              : undefined
                        }
                      : undefined,
                  dropdown:
                    run.widget.dropdown && typeof run.widget.dropdown === "object"
                      ? {
                          defaultValue:
                            typeof run.widget.dropdown.defaultValue === "string"
                              ? run.widget.dropdown.defaultValue
                              : undefined
                        }
                      : undefined
                }
              : undefined;

          children.push({
            type: "form-field",
            fieldType: run.fieldType,
            sourceKind:
              run.sourceKind === "legacy" || run.sourceKind === "sdt" ? run.sourceKind : undefined,
            id: typeof run.id === "number" ? run.id : undefined,
            tag: typeof run.tag === "string" ? run.tag : undefined,
            title: typeof run.title === "string" ? run.title : undefined,
            placeholder: typeof run.placeholder === "string" ? run.placeholder : undefined,
            checked: typeof run.checked === "boolean" ? run.checked : undefined,
            value: typeof run.value === "string" ? run.value : undefined,
            options: options && options.length > 0 ? options : undefined,
            widget,
            checkedSymbol: typeof run.checkedSymbol === "string" ? run.checkedSymbol : undefined,
            uncheckedSymbol: typeof run.uncheckedSymbol === "string" ? run.uncheckedSymbol : undefined,
            style: run.style ? { ...run.style } : undefined,
            link: typeof run.link === "string" ? run.link : undefined,
            sourceXml: typeof run.sourceXml === "string" ? run.sourceXml : undefined
          });
        }
      }

      normalized.push({
        type: "paragraph",
        style: value.style ? { ...value.style } : undefined,
        sourceXml: typeof value.sourceXml === "string" ? value.sourceXml : undefined,
        children: children.length > 0 ? children : [{ type: "text", text: "" }]
      });
    }

    return normalized.length > 0 ? normalized : undefined;
  } catch {
    return undefined;
  }
}
