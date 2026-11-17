import type {
  DocModel,
  DocNode,
  FormFieldRunNode,
  ImageRunNode,
  NumberingDefinitionSet,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphChildNode,
  ParagraphIndent,
  ParagraphNode,
  ParagraphNumbering,
  ParagraphSpacing,
  ParagraphStyle,
  TableBorderSet,
  TableBorderStyle,
  TableBoxSpacing,
  TableCellContentNode,
  TableNode,
  TableStyle,
  TextRunNode,
} from "./types";
import {
  cloneDocumentCommentAnchor,
  getCachedDocumentCommentAnchor,
  setCachedDocumentCommentAnchor,
} from "./comment-anchors";

function isParagraphCellContent(
  node: TableCellContentNode
): node is ParagraphNode {
  return node.type === "paragraph";
}

function isTableCellContentTable(
  node: TableCellContentNode
): node is TableNode {
  return node.type === "table";
}

function cloneTableCellContent(
  nodes: TableCellContentNode[]
): TableCellContentNode[] {
  return nodes.map((node) => {
    if (isParagraphCellContent(node)) {
      return cloneParagraph(node);
    }

    if (isTableCellContentTable(node)) {
      return cloneTable(node);
    }

    return node;
  });
}

function cloneParagraphNumbering(
  numbering?: ParagraphNumbering
): ParagraphNumbering | undefined {
  return numbering ? { ...numbering } : undefined;
}

function cloneParagraphSpacing(
  spacing?: ParagraphSpacing
): ParagraphSpacing | undefined {
  return spacing ? { ...spacing } : undefined;
}

function cloneParagraphIndent(
  indent?: ParagraphIndent
): ParagraphIndent | undefined {
  return indent ? { ...indent } : undefined;
}

function cloneParagraphBorderStyle(
  border?: ParagraphBorderStyle
): ParagraphBorderStyle | undefined {
  return border ? { ...border } : undefined;
}

function cloneParagraphBorderSet(
  borders?: ParagraphBorderSet
): ParagraphBorderSet | undefined {
  if (!borders) {
    return undefined;
  }

  return {
    top: cloneParagraphBorderStyle(borders.top),
    right: cloneParagraphBorderStyle(borders.right),
    bottom: cloneParagraphBorderStyle(borders.bottom),
    left: cloneParagraphBorderStyle(borders.left),
    between: cloneParagraphBorderStyle(borders.between),
    bar: cloneParagraphBorderStyle(borders.bar),
  };
}

function cloneParagraphStyle(
  style?: ParagraphStyle
): ParagraphStyle | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    numbering: cloneParagraphNumbering(style.numbering),
    spacing: cloneParagraphSpacing(style.spacing),
    indent: cloneParagraphIndent(style.indent),
    borders: cloneParagraphBorderSet(style.borders),
    tabStops: style.tabStops?.map((tabStop) => ({ ...tabStop })),
    dropCap: style.dropCap
      ? {
          ...style.dropCap,
        }
      : undefined,
  };
}

/** Clone a text run, keeping every run field (including note references). */
export function cloneTextRun(run: TextRunNode): TextRunNode {
  return {
    ...run,
    style: run.style ? { ...run.style } : undefined,
    link: run.link,
    noteReference: run.noteReference ? { ...run.noteReference } : undefined,
  };
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

export function cloneFormFieldRun(run: FormFieldRunNode): FormFieldRunNode {
  return {
    ...run,
    options: run.options?.map((option) => ({
      displayText: option.displayText,
      value: option.value,
    })),
    widget: cloneFormFieldWidget(run.widget),
    style: run.style ? { ...run.style } : undefined,
    link: run.link,
    sourceXml: run.sourceXml,
  };
}

/** Clone an image run, keeping crop/original XML and every other image field. */
export function cloneImageRun(run: ImageRunNode): ImageRunNode {
  return {
    ...run,
    data: run.data ? new Uint8Array(run.data) : undefined,
    sourceXml: run.sourceXml,
    crop: run.crop ? { ...run.crop } : undefined,
    cssFilter: run.cssFilter,
    cssOpacity: run.cssOpacity,
    floating: run.floating ? { ...run.floating } : undefined,
    syntheticTextBox: run.syntheticTextBox,
    textBoxText: run.textBoxText,
  };
}

/** Clone a single paragraph child, dispatching on its run type. */
export function cloneParagraphChild(child: ParagraphChildNode): ParagraphChildNode {
  if (child.type === "text") {
    return cloneTextRun(child);
  }

  if (child.type === "form-field") {
    return cloneFormFieldRun(child);
  }

  return cloneImageRun(child);
}

export function cloneParagraph(paragraph: ParagraphNode): ParagraphNode {
  return {
    type: "paragraph",
    style: cloneParagraphStyle(paragraph.style),
    paragraphMarkDeleted: paragraph.paragraphMarkDeleted,
    sourceXml: paragraph.sourceXml,
    children: paragraph.children.map(cloneParagraphChild),
  };
}

function cloneTableBoxSpacing(
  spacing?: TableBoxSpacing
): TableBoxSpacing | undefined {
  if (!spacing) {
    return undefined;
  }

  return {
    topTwips: spacing.topTwips,
    rightTwips: spacing.rightTwips,
    bottomTwips: spacing.bottomTwips,
    leftTwips: spacing.leftTwips,
  };
}

function cloneTableBorderStyle(
  border?: TableBorderStyle
): TableBorderStyle | undefined {
  if (!border) {
    return undefined;
  }

  return {
    type: border.type,
    color: border.color,
    sizeEighthPt: border.sizeEighthPt,
  };
}

function cloneTableBorderSet(
  borders?: TableBorderSet
): TableBorderSet | undefined {
  if (!borders) {
    return undefined;
  }

  return {
    top: cloneTableBorderStyle(borders.top),
    right: cloneTableBorderStyle(borders.right),
    bottom: cloneTableBorderStyle(borders.bottom),
    left: cloneTableBorderStyle(borders.left),
    insideH: cloneTableBorderStyle(borders.insideH),
    insideV: cloneTableBorderStyle(borders.insideV),
    tl2br: cloneTableBorderStyle(borders.tl2br),
    tr2bl: cloneTableBorderStyle(borders.tr2bl),
  };
}

function cloneTableFloatingStyle(
  floating?: NonNullable<TableStyle["floating"]>
): NonNullable<TableStyle["floating"]> | undefined {
  if (!floating) {
    return undefined;
  }

  return {
    xTwips: floating.xTwips,
    yTwips: floating.yTwips,
    leftFromTextTwips: floating.leftFromTextTwips,
    rightFromTextTwips: floating.rightFromTextTwips,
    topFromTextTwips: floating.topFromTextTwips,
    bottomFromTextTwips: floating.bottomFromTextTwips,
    horizontalAnchor: floating.horizontalAnchor,
    verticalAnchor: floating.verticalAnchor,
    horizontalAlign: floating.horizontalAlign,
    verticalAlign: floating.verticalAlign,
  };
}

function cloneTable(table: TableNode): TableNode {
  return {
    type: "table",
    sourceXml: table.sourceXml,
    style: table.style
      ? {
          widthTwips: table.style.widthTwips,
          indentTwips: table.style.indentTwips,
          layout: table.style.layout,
          cellSpacingTwips: table.style.cellSpacingTwips,
          floating: cloneTableFloatingStyle(table.style.floating),
          cellMarginTwips: cloneTableBoxSpacing(table.style.cellMarginTwips),
          columnWidthsTwips: table.style.columnWidthsTwips
            ? [...table.style.columnWidthsTwips]
            : undefined,
          borders: cloneTableBorderSet(table.style.borders),
        }
      : undefined,
    rows: table.rows.map((row) => ({
      type: "table-row",
      style: row.style ? { ...row.style } : undefined,
      cells: row.cells.map((cell) => ({
        type: "table-cell",
        style: cell.style
          ? {
              ...cell.style,
              marginTwips: cloneTableBoxSpacing(cell.style.marginTwips),
              borders: cloneTableBorderSet(cell.style.borders),
            }
          : undefined,
        nodes: cloneTableCellContent(cell.nodes),
      })),
    })),
  };
}

export function cloneDocNode(node: DocNode): DocNode {
  return node.type === "paragraph" ? cloneParagraph(node) : cloneTable(node);
}

function cloneNumberingDefinitions(
  numberingDefinitions?: NumberingDefinitionSet
): NumberingDefinitionSet | undefined {
  if (!numberingDefinitions) {
    return undefined;
  }

  return {
    abstracts: numberingDefinitions.abstracts.map((abstractDefinition) => ({
      abstractNumId: abstractDefinition.abstractNumId,
      levels: abstractDefinition.levels.map((level) => ({
        ...level,
        runStyle: level.runStyle ? { ...level.runStyle } : undefined,
        pictureBullet: level.pictureBullet
          ? { ...level.pictureBullet }
          : undefined,
      })),
    })),
    instances: numberingDefinitions.instances.map((instanceDefinition) => ({
      numId: instanceDefinition.numId,
      abstractNumId: instanceDefinition.abstractNumId,
      levelStartOverrides: instanceDefinition.levelStartOverrides
        ? { ...instanceDefinition.levelStartOverrides }
        : undefined,
      levelOverrides: instanceDefinition.levelOverrides
        ? instanceDefinition.levelOverrides.map((level) => ({
            ...level,
            runStyle: level.runStyle ? { ...level.runStyle } : undefined,
            pictureBullet: level.pictureBullet
              ? { ...level.pictureBullet }
              : undefined,
          }))
        : undefined,
    })),
  };
}

export function cloneDocModel(model: DocModel): DocModel {
  return {
    nodes: model.nodes.map(cloneDocNode),
    metadata: {
      sourceParts: model.metadata.sourceParts,
      warnings: [...model.metadata.warnings],
      documentPageCount: model.metadata.documentPageCount,
      documentOpenTag: model.metadata.documentOpenTag,
      documentBackgroundColor: model.metadata.documentBackgroundColor,
      sectionPropertiesXml: model.metadata.sectionPropertiesXml,
      sections: model.metadata.sections?.map((section) => ({
        startNodeIndex: section.startNodeIndex,
        sectionPropertiesXml: section.sectionPropertiesXml,
        headerSections: (section.headerSections ?? []).map((headerSection) => ({
          partName: headerSection.partName,
          referenceType: headerSection.referenceType,
          nodes: headerSection.nodes.map(cloneDocNode),
        })),
        footerSections: (section.footerSections ?? []).map((footerSection) => ({
          partName: footerSection.partName,
          referenceType: footerSection.referenceType,
          nodes: footerSection.nodes.map(cloneDocNode),
        })),
      })),
      headerSections: (model.metadata.headerSections ?? []).map((section) => ({
        partName: section.partName,
        referenceType: section.referenceType,
        nodes: section.nodes.map(cloneDocNode),
      })),
      footerSections: (model.metadata.footerSections ?? []).map((section) => ({
        partName: section.partName,
        referenceType: section.referenceType,
        nodes: section.nodes.map(cloneDocNode),
      })),
      paragraphStyles: (model.metadata.paragraphStyles ?? []).map((style) => ({
        ...style,
        runStyle: style.runStyle ? { ...style.runStyle } : undefined,
        numbering: cloneParagraphNumbering(style.numbering),
        spacing: cloneParagraphSpacing(style.spacing),
        indent: cloneParagraphIndent(style.indent),
        borders: cloneParagraphBorderSet(style.borders),
      })),
      defaultParagraphStyleId: model.metadata.defaultParagraphStyleId,
      numberingDefinitions: cloneNumberingDefinitions(
        model.metadata.numberingDefinitions
      ),
      compatibility: model.metadata.compatibility
        ? { ...model.metadata.compatibility }
        : undefined,
      footnotes: model.metadata.footnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(cloneDocNode),
      })),
      endnotes: model.metadata.endnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(cloneDocNode),
      })),
      comments: model.metadata.comments?.map((comment) => {
        const cloned = { ...comment };
        const anchor = getCachedDocumentCommentAnchor(comment);
        if (anchor) {
          setCachedDocumentCommentAnchor(
            cloned,
            cloneDocumentCommentAnchor(anchor)
          );
        }
        return cloned;
      }),
    },
  };
}
