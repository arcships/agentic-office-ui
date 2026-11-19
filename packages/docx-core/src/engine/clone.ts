import type {
  DocModel,
  DocNode,
  DrawingRunNode,
  FooterSection,
  FormFieldRunNode,
  HeaderSection,
  ImageRunNode,
  MathRunNode,
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
  TextBoxContent,
  TextRunNode,
  TextStyle,
} from "./types";
import {
  cloneDocumentCommentAnchor,
  getCachedDocumentCommentAnchor,
  setCachedDocumentCommentAnchor,
} from "./comment-anchors";

/**
 * Returns the encoded image bytes without copying them.
 *
 * Package bytes are large (a long document can carry tens of megabytes of
 * images) and the editor treats them as immutable: decoders only read from them
 * and no edit command writes into a run buffer. Edits therefore alias the
 * original buffer instead of duplicating it on every keystroke.
 * Array-shaped payloads (legacy or JSON input) are still normalised into a
 * `Uint8Array`.
 */
export function shareImageData(
  data: ImageRunNode["data"]
): ImageRunNode["data"] {
  if (!data) {
    return undefined;
  }

  return data instanceof Uint8Array
    ? data
    : new Uint8Array(data as unknown as ArrayLike<number>);
}

/**
 * Returns encoded image bytes as an independent copy.
 *
 * `cloneDocModel` is the isolation boundary (import, normalise, checkpoints), so
 * it keeps the historical "no shared mutable state" guarantee. Edit-path clones
 * use {@link shareImageData} instead.
 */
export function copyImageData(
  data: ImageRunNode["data"]
): ImageRunNode["data"] {
  if (!data) {
    return undefined;
  }

  return data instanceof Uint8Array
    ? new Uint8Array(data)
    : Uint8Array.from(data as unknown as ArrayLike<number>);
}

/**
 * Controls how a node clone treats encoded image payloads.
 *
 * Edit-path clones (`cloneDocNode`, `cloneParagraph`, ...) share bytes by
 * default; `cloneDocModel` opts into copying so the full clone stays isolated.
 */
export interface DocCloneOptions {
  shareImageBytes?: boolean;
}

const SHARE_IMAGE_BYTES: DocCloneOptions = { shareImageBytes: true };
const COPY_IMAGE_BYTES: DocCloneOptions = { shareImageBytes: false };

function cloneImageBytes(
  data: ImageRunNode["data"],
  options: DocCloneOptions
): ImageRunNode["data"] {
  return options.shareImageBytes === false
    ? copyImageData(data)
    : shareImageData(data);
}

function cloneTextStyle(style?: TextStyle): TextStyle | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    runBorder: style.runBorder ? { ...style.runBorder } : undefined,
    textShadow: style.textShadow ? { ...style.textShadow } : undefined,
    textOutline: style.textOutline ? { ...style.textOutline } : undefined,
  };
}

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
  nodes: TableCellContentNode[],
  options: DocCloneOptions
): TableCellContentNode[] {
  return nodes.map((node) => {
    if (isParagraphCellContent(node)) {
      return cloneParagraph(node, options);
    }

    if (isTableCellContentTable(node)) {
      return cloneTable(node, options);
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
    style: cloneTextStyle(run.style),
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
    style: cloneTextStyle(run.style),
    link: run.link,
    sourceXml: run.sourceXml,
  };
}

/** Clone a text box content, keeping paragraph and run styles. */
export function cloneTextBoxContent(
  content?: TextBoxContent,
): TextBoxContent | undefined {
  if (!content) return undefined;
  return {
    paragraphs: content.paragraphs.map((paragraph) => ({
      ...paragraph,
      segments: paragraph.segments.map((segment) => ({
        ...segment,
        style: cloneTextStyle(segment.style),
      })),
    })),
  };
}

/** Clone an image run, keeping crop/original XML and every other image field. */
export function cloneImageRun(
  run: ImageRunNode,
  options: DocCloneOptions = SHARE_IMAGE_BYTES
): ImageRunNode {
  return {
    ...run,
    data: cloneImageBytes(run.data, options),
    sourceXml: run.sourceXml,
    crop: run.crop ? { ...run.crop } : undefined,
    cssFilter: run.cssFilter,
    cssOpacity: run.cssOpacity,
    floating: run.floating ? { ...run.floating } : undefined,
    syntheticTextBox: run.syntheticTextBox,
    textBoxText: run.textBoxText,
    textBox: cloneTextBoxContent(run.textBox),
    textBoxSourceXml: run.textBoxSourceXml,
  };
}

/** Clone a drawing object run (shapes, charts, embedded objects, 3D models). */
export function cloneDrawingRun(run: DrawingRunNode): DrawingRunNode {
  return {
    ...run,
    floating: run.floating ? { ...run.floating } : undefined,
    textBox: cloneTextBoxContent(run.textBox),
    chart: run.chart
      ? {
          ...run.chart,
          categories: run.chart.categories ? [...run.chart.categories] : undefined,
          series: run.chart.series?.map((series) => ({
            ...series,
            categories: series.categories ? [...series.categories] : undefined,
            values: series.values ? [...series.values] : undefined,
          })),
        }
      : undefined,
    shape: run.shape ? { ...run.shape } : undefined,
    ole: run.ole ? { ...run.ole } : undefined,
  };
}

/** Clone an equation run. */
export function cloneMathRun(run: MathRunNode): MathRunNode {
  return {
    ...run,
    style: cloneTextStyle(run.style),
  };
}

/** Clone a single paragraph child, dispatching on its run type. */
export function cloneParagraphChild(
  child: ParagraphChildNode,
  options: DocCloneOptions = SHARE_IMAGE_BYTES
): ParagraphChildNode {
  if (child.type === "text") {
    return cloneTextRun(child);
  }

  if (child.type === "form-field") {
    return cloneFormFieldRun(child);
  }

  if (child.type === "drawing") {
    return cloneDrawingRun(child);
  }

  if (child.type === "math") {
    return cloneMathRun(child);
  }

  return cloneImageRun(child, options);
}

export function cloneParagraph(
  paragraph: ParagraphNode,
  options: DocCloneOptions = SHARE_IMAGE_BYTES
): ParagraphNode {
  return {
    type: "paragraph",
    style: cloneParagraphStyle(paragraph.style),
    paragraphMarkDeleted: paragraph.paragraphMarkDeleted,
    sourceXml: paragraph.sourceXml,
    sectionBreak: paragraph.sectionBreak
      ? { ...paragraph.sectionBreak }
      : undefined,
    commentRanges: paragraph.commentRanges?.map((range) => ({ ...range })),
    children: paragraph.children.map((child) =>
      cloneParagraphChild(child, options)
    ),
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

function cloneTable(
  table: TableNode,
  options: DocCloneOptions = SHARE_IMAGE_BYTES
): TableNode {
  return {
    type: "table",
    sourceXml: table.sourceXml,
    style: table.style
      ? {
          widthTwips: table.style.widthTwips,
          indentTwips: table.style.indentTwips,
          layout: table.style.layout,
          cellSpacingTwips: table.style.cellSpacingTwips,
          styleId: table.style.styleId,
          sourceStyleId: table.style.sourceStyleId,
          collapsedGridColumns: table.style.collapsedGridColumns
            ? [...table.style.collapsedGridColumns]
            : undefined,
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
        nodes: cloneTableCellContent(cell.nodes, options),
      })),
    })),
  };
}

/**
 * Deep-clones one top-level document node. Image bytes are shared by default
 * (see {@link shareImageData}); every other object in the node is copied.
 * Pass `{ shareImageBytes: false }` for a fully isolated copy.
 */
export function cloneDocNode<T extends DocNode>(
  node: T,
  options: DocCloneOptions = SHARE_IMAGE_BYTES
): T {
  return (
    node.type === "paragraph" ? cloneParagraph(node, options) : cloneTable(node, options)
  ) as T;
}

/**
 * Copy-on-write shell for the document model.
 *
 * Only the model object, its `nodes` array and the `metadata` object are
 * copied. Nodes, header/footer sections, metadata arrays and image bytes are
 * shared by reference with `model`, which keeps a single-node edit O(1) in the
 * size of the document instead of O(document).
 *
 * Callers must never mutate a shared node: replace the touched top-level node
 * with `cloneDocNode(model.nodes[index])` first, then mutate the clone. Use
 * `cloneDocModel` when the whole graph is being rewritten (import, normalise,
 * checkpoints that outlive the working model).
 */
export function cloneDocModelShallow(model: DocModel): DocModel {
  return {
    ...model,
    nodes: [...model.nodes],
    metadata: { ...model.metadata }
  };
}

/**
 * Copy-on-write model whose listed top-level nodes are fresh clones, ready to be
 * mutated in place.
 *
 * Commands that touch several nodes (a range format change, a table-cell edit, an
 * image moved between paragraphs) pass every touched node index here instead of
 * cloning the whole document. Indexes may repeat and may be out of range; only
 * existing nodes are copied.
 */
export function cloneDocModelNodes(
  model: DocModel,
  nodeIndexes: Iterable<number>
): DocModel {
  const next = cloneDocModelShallow(model);
  const copied = new Set<number>();

  for (const nodeIndex of nodeIndexes) {
    if (copied.has(nodeIndex)) continue;
    copied.add(nodeIndex);

    const source = model.nodes[nodeIndex];
    if (!source) continue;
    next.nodes[nodeIndex] = cloneDocNode(source);
  }

  return next;
}

/**
 * Copy-on-write model whose header or footer sections are fresh copies, plus
 * those copies.
 *
 * Header/footer edits live in `metadata.headerSections`, in
 * `metadata.sections[].headerSections` and in `metadata.footerSections` /
 * `metadata.sections[].footerSections`. The returned model owns new arrays for
 * every one of those lists, every section object is copied, and every section
 * `nodes` array is copied so a caller can replace a node in it. Nodes that are
 * not edited stay shared.
 *
 * Callers mutate the returned `sections` (whose `nodes` arrays belong to the
 * returned model) instead of `model.metadata`.
 */
export function cloneDocModelSectionLists(
  model: DocModel,
  region: "header" | "footer"
): { model: DocModel; sections: HeaderSection[] | FooterSection[] } {
  const next = cloneDocModelShallow(model);
  const key = region === "header" ? "headerSections" : "footerSections";
  const collected: HeaderSection[] | FooterSection[] = [];

  const copyList = <T extends HeaderSection | FooterSection>(
    sections: T[] | undefined
  ): T[] | undefined =>
    sections?.map((section) => {
      const copy = { ...section, nodes: [...section.nodes] } as T;
      collected.push(copy);
      return copy;
    });

  const topLevel = next.metadata[key];
  if (topLevel) {
    next.metadata[key] = copyList(topLevel) as typeof topLevel;
  }

  if (next.metadata.sections) {
    next.metadata.sections = next.metadata.sections.map((section) => {
      const list = section[key];
      return list ? { ...section, [key]: copyList(list) } : section;
    });
  }

  return { model: next, sections: collected };
}

export function cloneNumberingDefinitions(
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
        runStyle: cloneTextStyle(level.runStyle),
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
            runStyle: cloneTextStyle(level.runStyle),
            pictureBullet: level.pictureBullet
              ? { ...level.pictureBullet }
              : undefined,
          }))
        : undefined,
    })),
  };
}

/**
 * Fully isolated copy of the document model: every node, paragraph style,
 * numbering definition and metadata array is copied, and encoded image bytes are
 * copied too. Use this for import, normalisation and checkpoints; edit commands
 * use the cheaper `cloneDocModelShallow` + `cloneDocNode`. Undo history keeps
 * model references and never calls this.
 */
export function cloneDocModel(model: DocModel): DocModel {
  return {
    nodes: model.nodes.map((node) => cloneDocNode(node, COPY_IMAGE_BYTES)),
    metadata: {
      sourceParts: model.metadata.sourceParts,
      sourcePartNames: model.metadata.sourcePartNames
        ? { ...model.metadata.sourcePartNames }
        : undefined,
      warnings: [...model.metadata.warnings],
      documentPageCount: model.metadata.documentPageCount,
      documentOpenTag: model.metadata.documentOpenTag,
      documentBackgroundColor: model.metadata.documentBackgroundColor,
      documentBackgroundPattern: model.metadata.documentBackgroundPattern
        ? { ...model.metadata.documentBackgroundPattern }
        : undefined,
      sectionPropertiesXml: model.metadata.sectionPropertiesXml,
      sections: model.metadata.sections?.map((section) => ({
        startNodeIndex: section.startNodeIndex,
        sectionPropertiesXml: section.sectionPropertiesXml,
        headerSections: (section.headerSections ?? []).map((headerSection) => ({
          partName: headerSection.partName,
          referenceType: headerSection.referenceType,
          nodes: headerSection.nodes.map((node) =>
            cloneDocNode(node, COPY_IMAGE_BYTES)
          ),
        })),
        footerSections: (section.footerSections ?? []).map((footerSection) => ({
          partName: footerSection.partName,
          referenceType: footerSection.referenceType,
          nodes: footerSection.nodes.map((node) =>
            cloneDocNode(node, COPY_IMAGE_BYTES)
          ),
        })),
      })),
      headerSections: (model.metadata.headerSections ?? []).map((section) => ({
        partName: section.partName,
        referenceType: section.referenceType,
        nodes: section.nodes.map((node) => cloneDocNode(node, COPY_IMAGE_BYTES)),
      })),
      footerSections: (model.metadata.footerSections ?? []).map((section) => ({
        partName: section.partName,
        referenceType: section.referenceType,
        nodes: section.nodes.map((node) => cloneDocNode(node, COPY_IMAGE_BYTES)),
      })),
      paragraphStyles: (model.metadata.paragraphStyles ?? []).map((style) => ({
        ...style,
        runStyle: cloneTextStyle(style.runStyle),
        numbering: cloneParagraphNumbering(style.numbering),
        spacing: cloneParagraphSpacing(style.spacing),
        indent: cloneParagraphIndent(style.indent),
        borders: cloneParagraphBorderSet(style.borders),
      })),
      defaultParagraphStyleId: model.metadata.defaultParagraphStyleId,
      defaultTabStopTwips: model.metadata.defaultTabStopTwips,
      themeFonts: model.metadata.themeFonts
        ? {
            major: model.metadata.themeFonts.major
              ? { ...model.metadata.themeFonts.major }
              : undefined,
            minor: model.metadata.themeFonts.minor
              ? { ...model.metadata.themeFonts.minor }
              : undefined,
          }
        : undefined,
      numberingDefinitions: cloneNumberingDefinitions(
        model.metadata.numberingDefinitions
      ),
      compatibility: model.metadata.compatibility
        ? { ...model.metadata.compatibility }
        : undefined,
      footnotes: model.metadata.footnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map((node) => cloneDocNode(node, COPY_IMAGE_BYTES)),
      })),
      endnotes: model.metadata.endnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map((node) => cloneDocNode(node, COPY_IMAGE_BYTES)),
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
