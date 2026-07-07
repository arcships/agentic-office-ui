// Paragraph structure inspection (part 1): text utilities, editable-text
// extraction, run/cell helpers, form-field display value, style-definition
// normalization.
// Upstream editor.tsx: lines 478-541, 4646-4922, 4923-5007.
//
// The floating-image/cover detection (5008-5811) lives in paragraph-geometry.ts
// and the structural analysis (7936-8640) in paragraph-tracked.ts.
// renderStaticHtml is framework-specific and lives in vue-docx/render.

import type {
  DocModel,
  FooterSection,
  FormFieldRunNode,
  HeaderSection,
  HeadingLevel,
  ParagraphNode,
  ParagraphStyleDefinition,
  TableCellContentNode,
  TableNode,
  TextRunNode
} from "../../engine/types";
import { DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY, DEFAULT_WORD_HEADING_RUN_STYLES } from "./constants";
import { cssFontFamily } from "./style-to-css";

/** Whether the paragraph contains any image run (inline or floating). */
export function paragraphHasImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => child.type === "image");
}

export function headingLevelFromStyleLabel(value?: string): HeadingLevel | undefined {
  if (!value) {
    return undefined;
  }

  const match = value.match(/(?:^|[\s_-])(?:heading|h)\s*([1-6])(?:$|[\s_-])/i);
  if (!match?.[1]) {
    return undefined;
  }

  const parsed = Number(match[1]);
  return parsed >= 1 && parsed <= 6 ? (parsed as HeadingLevel) : undefined;
}

export function resolveParagraphStyleHeadingLevel(
  styleDefinition: Pick<
    ParagraphStyleDefinition,
    "headingLevel" | "id" | "name"
  >
): HeadingLevel | undefined {
  if (
    Number.isFinite(styleDefinition.headingLevel) &&
    (styleDefinition.headingLevel as number) >= 1 &&
    (styleDefinition.headingLevel as number) <= 6
  ) {
    return styleDefinition.headingLevel as HeadingLevel;
  }

  return (
    headingLevelFromStyleLabel(styleDefinition.id) ??
    headingLevelFromStyleLabel(styleDefinition.name)
  );
}

export function normalizeParagraphStyleDefinitionsForUi(
  styles: ParagraphStyleDefinition[]
): ParagraphStyleDefinition[] {
  return styles.map((styleDefinition) => {
    const headingLevel = resolveParagraphStyleHeadingLevel(styleDefinition);
    const headingRunStyle = headingLevel
      ? DEFAULT_WORD_HEADING_RUN_STYLES[headingLevel]
      : undefined;
    const mergedRunStyle = headingRunStyle
      ? {
          ...headingRunStyle,
          ...(styleDefinition.runStyle ?? {}),
        }
      : styleDefinition.runStyle;
    const runStyleChanged =
      mergedRunStyle !== styleDefinition.runStyle ||
      (headingRunStyle !== undefined && styleDefinition.runStyle === undefined);
    const headingChanged = headingLevel !== styleDefinition.headingLevel;

    if (!runStyleChanged && !headingChanged) {
      return styleDefinition;
    }

    return {
      ...styleDefinition,
      headingLevel,
      runStyle: mergedRunStyle,
    };
  });
}
export function textRuns(paragraph: ParagraphNode): TextRunNode[] {
  return paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
}

export function paragraphText(paragraph: ParagraphNode): string {
  return textRuns(paragraph)
    .map((run) => run.text)
    .join("");
}

export function nodeTreeContainsExplicitFontFamily(
  nodes: DocModel["nodes"] | HeaderSection["nodes"] | FooterSection["nodes"]
): boolean {
  return nodes.some((node) => {
    if (node.type === "paragraph") {
      return node.children.some((child) => {
        if (child.type === "text" || child.type === "form-field") {
          return Boolean(child.style?.fontFamily?.trim());
        }
        return false;
      });
    }

    if (node.type === "table") {
      return node.rows.some((row) =>
        row.cells.some((cell) => nodeTreeContainsExplicitFontFamily(cell.nodes))
      );
    }

    return false;
  });
}

export function firstExplicitFontFamilyInNodeTree(
  nodes: DocModel["nodes"] | HeaderSection["nodes"] | FooterSection["nodes"]
): string | undefined {
  for (const node of nodes) {
    if (node.type === "paragraph") {
      for (const child of node.children) {
        if (child.type !== "text" && child.type !== "form-field") {
          continue;
        }

        const fontFamily = child.style?.fontFamily?.trim();
        if (fontFamily) {
          return fontFamily;
        }
      }
      continue;
    }

    if (node.type === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) {
          const nestedFontFamily = firstExplicitFontFamilyInNodeTree(
            cell.nodes
          );
          if (nestedFontFamily) {
            return nestedFontFamily;
          }
        }
      }
    }
  }

  return undefined;
}

export function resolveDocumentInheritedFontFamily(
  model: DocModel
): string | undefined {
  const paragraphStyles = model.metadata.paragraphStyles ?? [];
  const normalizedDefaultStyleId =
    model.metadata.defaultParagraphStyleId?.trim().toLowerCase() ?? "";
  const defaultParagraphStyle =
    paragraphStyles.find(
      (style) => style.id.trim().toLowerCase() === normalizedDefaultStyleId
    ) ??
    paragraphStyles.find((style) => style.isDefault) ??
    paragraphStyles.find((style) => style.id.trim().toLowerCase() === "normal");
  const defaultStyleFontFamily =
    defaultParagraphStyle?.runStyle?.fontFamily?.trim();
  if (defaultStyleFontFamily) {
    return cssFontFamily(defaultStyleFontFamily);
  }

  // Never adopt a heading style's font as the document body default. Heading
  // styles (e.g. "Calibri Light") are not the body typeface, and using one here
  // made freshly typed body text render in the wrong font until it committed.
  const paragraphStyleFontFamily = paragraphStyles.find(
    (style) =>
      style.headingLevel === undefined &&
      Boolean(style.runStyle?.fontFamily?.trim())
  )?.runStyle?.fontFamily;
  if (paragraphStyleFontFamily) {
    return cssFontFamily(paragraphStyleFontFamily);
  }

  const explicitBodyFontFamily = firstExplicitFontFamilyInNodeTree(model.nodes);
  if (explicitBodyFontFamily) {
    return cssFontFamily(explicitBodyFontFamily);
  }

  for (const section of model.metadata.headerSections ?? []) {
    const explicitHeaderFontFamily = firstExplicitFontFamilyInNodeTree(
      section.nodes
    );
    if (explicitHeaderFontFamily) {
      return cssFontFamily(explicitHeaderFontFamily);
    }
  }

  for (const section of model.metadata.footerSections ?? []) {
    const explicitFooterFontFamily = firstExplicitFontFamilyInNodeTree(
      section.nodes
    );
    if (explicitFooterFontFamily) {
      return cssFontFamily(explicitFooterFontFamily);
    }
  }

  return nodeTreeContainsExplicitFontFamily(model.nodes) ||
    (model.metadata.headerSections ?? []).some((section) =>
      nodeTreeContainsExplicitFontFamily(section.nodes)
    ) ||
    (model.metadata.footerSections ?? []).some((section) =>
      nodeTreeContainsExplicitFontFamily(section.nodes)
    )
    ? undefined
    : cssFontFamily(DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY);
}

export function replaceTabLayoutMarkersWithTabText(root: HTMLElement): void {
  const centerLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='center']")
  );
  centerLayouts.forEach((layout) => {
    const left =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='left']")
        ?.textContent ?? "";
    const center =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='center']")
        ?.textContent ?? "";
    layout.replaceWith(`${left}\t${center}`);
  });

  const centerRightLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='center-right']")
  );
  centerRightLayouts.forEach((layout) => {
    const first =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='0']")
        ?.textContent ?? "";
    const second =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='1']")
        ?.textContent ?? "";
    const third =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='2']")
        ?.textContent ?? "";
    layout.replaceWith(`${first}\t${second}\t${third}`);
  });

  const leaderLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='leader']")
  );
  leaderLayouts.forEach((layout) => {
    const left =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='left']")
        ?.textContent ?? "";
    const right =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='right']")
        ?.textContent ?? "";
    layout.replaceWith(`${left}\t${right}`);
  });

  const rightLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='right']")
  );
  rightLayouts.forEach((layout) => {
    const left =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='left']")
        ?.textContent ?? "";
    const right =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='right']")
        ?.textContent ?? "";
    layout.replaceWith(`${left}\t${right}`);
  });

  const explicitTabMarkers = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-char='true']")
  );
  explicitTabMarkers.forEach((marker) => {
    marker.replaceWith("\t");
  });
}

export function editableTextFromElement(element: HTMLElement): string {
  const clone = element.cloneNode(true) as HTMLElement;
  replaceTabLayoutMarkersWithTabText(clone);
  clone
    .querySelectorAll("[data-docx-numbering-label='true']")
    .forEach((label) => {
      label.remove();
    });
  clone.querySelectorAll("br").forEach((lineBreak) => {
    lineBreak.replaceWith("\n");
  });
  return clone.textContent ?? "";
}

export function editableTextFromTableCellElement(element: HTMLElement): string {
  const paragraphLikeChildren = Array.from(element.children).filter(
    (child): child is HTMLElement =>
      child instanceof HTMLElement &&
      child.getAttribute("contenteditable") !== "false"
  );
  if (paragraphLikeChildren.length === 0) {
    return editableTextFromElement(element);
  }

  return paragraphLikeChildren
    .map((paragraphLikeChild) => {
      if (paragraphLikeChild.tagName.toUpperCase() === "BR") {
        return "";
      }
      const text = editableTextFromElement(paragraphLikeChild);
      return text === "\n" ? "" : text;
    })
    .join("\n");
}

export function editableTextFromDraftHtml(html: string): string {
  if (typeof document === "undefined") {
    return html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/gi, " ");
  }

  const container = document.createElement("div");
  container.innerHTML = html;
  return editableTextFromElement(container);
}

export function editableTextFromTableCellDraftHtml(html: string): string {
  if (typeof document === "undefined") {
    return editableTextFromDraftHtml(html);
  }

  const container = document.createElement("div");
  container.innerHTML = html;
  return editableTextFromTableCellElement(container);
}


export function formFieldDisplayValue(field: FormFieldRunNode): string {
  switch (field.fieldType) {
    case "checkbox":
      return field.checked ?? field.widget?.checkbox?.defaultChecked
        ? field.checkedSymbol ?? "☒"
        : field.uncheckedSymbol ?? "☐";
    case "dropdown":
      return field.value ?? field.options?.[0]?.displayText ?? "";
    case "date":
      return field.value ?? "";
    case "text":
      return field.value ?? field.widget?.text?.defaultText ?? "";
    default:
      return field.value ?? "";
  }
}

export function firstRunStyle(paragraph?: ParagraphNode): TextRunNode["style"] {
  return textRuns(
    paragraph ?? ({ type: "paragraph", children: [] } as ParagraphNode)
  )[0]?.style;
}

export function ensureTextRunNode(paragraph: ParagraphNode): TextRunNode {
  const existing = paragraph.children.find(
    (child): child is TextRunNode => child.type === "text"
  );
  if (existing) {
    return existing;
  }

  const created: TextRunNode = {
    type: "text",
    text: "",
    style: {},
  };
  paragraph.children.unshift(created);
  return created;
}

export function isParagraphCellContentNode(
  node: TableCellContentNode
): node is ParagraphNode {
  return node.type === "paragraph";
}

export function isTableCellTableContentNode(
  node: TableCellContentNode
): node is TableNode {
  return node.type === "table";
}

export function tableCellParagraphs(
  nodeContent: TableCellContentNode[]
): ParagraphNode[] {
  return nodeContent.filter(isParagraphCellContentNode);
}

export function tableCellParagraphsRecursively(
  nodeContent: TableCellContentNode[]
): ParagraphNode[] {
  const paragraphs: ParagraphNode[] = [];

  const walk = (entries: TableCellContentNode[]): void => {
    for (const entry of entries) {
      if (isParagraphCellContentNode(entry)) {
        paragraphs.push(entry);
        continue;
      }

      for (const row of entry.rows) {
        for (const nestedCell of row.cells) {
          walk(nestedCell.nodes);
        }
      }
    }
  };

  walk(nodeContent);
  return paragraphs;
}

export function tableCellHasImage(nodeContent: TableCellContentNode[]): boolean {
  for (const entry of nodeContent) {
    if (isParagraphCellContentNode(entry) && paragraphHasImage(entry)) {
      return true;
    }

    if (isTableCellTableContentNode(entry)) {
      for (const row of entry.rows) {
        for (const nestedCell of row.cells) {
          if (tableCellHasImage(nestedCell.nodes)) {
            return true;
          }
        }
      }
    }
  }

  return false;
}

