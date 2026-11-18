export type ParagraphAlignment =
  | "left"
  | "center"
  | "right"
  | "justify"
  | "both"
  | "distribute"
  | "thaiDistribute"
  | "low"
  | "start"
  | "end";
export type ParagraphDirection = "ltr" | "rtl";
export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** OOXML `w:u` underline variants (beyond a plain single line). */
export type TextUnderlineStyle =
  | "single"
  | "words"
  | "double"
  | "thick"
  | "dotted"
  | "dottedHeavy"
  | "dash"
  | "dashedHeavy"
  | "dashLong"
  | "dashLongHeavy"
  | "dotDash"
  | "dashDotHeavy"
  | "dotDotDash"
  | "dashDotDotHeavy"
  | "wave"
  | "wavyHeavy"
  | "wavyDouble";

/** OOXML `w:strike` / `w:dstrike` strike-through variants. */
export type TextStrikeStyle = "single" | "double";

export interface TextRunBorderStyle {
  type: string;
  color?: string;
  sizeEighthPt?: number;
  spacePt?: number;
  frame?: boolean;
  shadow?: boolean;
}

export interface TextShadowStyle {
  color?: string;
  blurPt?: number;
  offsetXPt?: number;
  offsetYPt?: number;
}

export interface TextOutlineStyle {
  color?: string;
  widthPt?: number;
}

export interface DocxThemeFontSet {
  ascii?: string;
  eastAsia?: string;
  complexScript?: string;
}

export interface DocxThemeFonts {
  major?: DocxThemeFontSet;
  minor?: DocxThemeFontSet;
}

export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** OOXML underline variant; `underline` stays the boolean compatibility flag. */
  underlineStyle?: TextUnderlineStyle;
  strike?: boolean;
  /** OOXML strike variant; `strike` stays the boolean compatibility flag. */
  strikeStyle?: TextStrikeStyle;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlight?: string;
  backgroundColor?: string;
  fontSizePt?: number;
  /** Primary (ASCII / high-ANSI) font family. */
  fontFamily?: string;
  /** East-Asian font family (`w:rFonts/@w:eastAsia`). */
  fontFamilyEastAsia?: string;
  /** Complex-script font family (`w:rFonts/@w:cs`). */
  fontFamilyComplexScript?: string;
  characterSpacingTwips?: number;
  /** OOXML `w:w` character scale, 1-600 (100 = normal). */
  characterScalePercent?: number;
  direction?: ParagraphDirection;
  verticalAlign?: "superscript" | "subscript";
  runBorder?: TextRunBorderStyle;
  textShadow?: TextShadowStyle;
  textOutline?: TextOutlineStyle;
}

export type NoteReferenceKind = "footnote" | "endnote";

export interface NoteReference {
  kind: NoteReferenceKind;
  /** Raw OOXML w:id of the referenced note definition. */
  id: number;
}

export interface TextRunNode {
  type: "text";
  text: string;
  style?: TextStyle;
  link?: string;
  noteReference?: NoteReference;
}

export interface ImageRunNode {
  type: "image";
  src?: string;
  alt?: string;
  widthPx?: number;
  heightPx?: number;
  partName?: string;
  contentType?: string;
  data?: Uint8Array;
  sourceXml?: string;
  crop?: {
    leftFraction?: number;
    topFraction?: number;
    rightFraction?: number;
    bottomFraction?: number;
  };
  cssFilter?: string;
  cssOpacity?: number;
  floating?: {
    xPx?: number;
    yPx?: number;
    horizontalAlign?: "left" | "center" | "right" | "inside" | "outside";
    verticalAlign?: "top" | "center" | "bottom" | "inside" | "outside";
    horizontalRelativeTo?: string;
    verticalRelativeTo?: string;
    distLPx?: number;
    distRPx?: number;
    distTPx?: number;
    distBPx?: number;
    wrapType?: "none" | "square" | "tight" | "through" | "topAndBottom";
    wrapText?: "bothSides" | "left" | "right" | "largest";
    behindDocument?: boolean;
    zIndex?: number;
    /** Whether the anchor is laid out inside its table cell (wp:anchor/@layoutInCell). */
    layoutInCell?: boolean;
    /** Whether the anchored object may overlap other floating objects (wp:anchor/@allowOverlap). */
    allowOverlap?: boolean;
    /** Whether the editor must not move or resize the anchor (wp:anchor/@locked). */
    lockedAnchor?: boolean;
    /** Whether x/y are page-relative simple positions (wp:anchor/@simplePos). */
    simplePos?: boolean;
    /** Visual overflow beyond the extent, in CSS pixels (wp:effectExtent). */
    effectExtent?: {
      leftPx?: number;
      topPx?: number;
      rightPx?: number;
      bottomPx?: number;
    };
  };
  syntheticTextBox?: boolean;
  textBoxText?: string;
}

export type FormFieldType = "checkbox" | "text" | "date" | "dropdown";

export type FormFieldSourceKind = "sdt" | "legacy";

export interface FormFieldOption {
  displayText: string;
  value?: string;
}

export interface FormFieldTextWidgetSettings {
  inputType?:
    | "regular"
    | "number"
    | "date"
    | "currentDate"
    | "currentTime"
    | "calculated"
    | (string & {});
  defaultText?: string;
  maxLength?: number;
  textFormat?: string;
}

export interface FormFieldCheckboxWidgetSettings {
  defaultChecked?: boolean;
  sizeMode?: "auto" | "exact";
  sizePt?: number;
}

export interface FormFieldDropdownWidgetSettings {
  defaultValue?: string;
}

export interface FormFieldWidgetSettings {
  name?: string;
  enabled?: boolean;
  calcOnExit?: boolean;
  text?: FormFieldTextWidgetSettings;
  checkbox?: FormFieldCheckboxWidgetSettings;
  dropdown?: FormFieldDropdownWidgetSettings;
}

export interface FormFieldRunNode {
  type: "form-field";
  fieldType: FormFieldType;
  sourceKind?: FormFieldSourceKind;
  id?: number;
  tag?: string;
  title?: string;
  placeholder?: string;
  /** Legacy form-field `w:ffData/w:required` flag. */
  required?: boolean;
  checked?: boolean;
  value?: string;
  options?: FormFieldOption[];
  widget?: FormFieldWidgetSettings;
  checkedSymbol?: string;
  uncheckedSymbol?: string;
  style?: TextStyle;
  link?: string;
  sourceXml?: string;
}

export type ParagraphChildNode = TextRunNode | ImageRunNode | FormFieldRunNode;

export interface ParagraphNumbering {
  numId: number;
  ilvl: number;
}

export interface ParagraphSpacing {
  beforeTwips?: number;
  afterTwips?: number;
  lineTwips?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  /**
   * OOXML `w:spacing/@w:beforeAutospacing`. The amount is consumer-defined, so
   * the model keeps the flag and layout resolves it (see
   * `resolveParagraphAutoSpacingPx`).
   */
  beforeAutoSpacing?: boolean;
  /** OOXML `w:spacing/@w:afterAutospacing`; resolved like `beforeAutoSpacing`. */
  afterAutoSpacing?: boolean;
}

export interface ParagraphIndent {
  leftTwips?: number;
  rightTwips?: number;
  firstLineTwips?: number;
  hangingTwips?: number;
}

export interface ParagraphBorderStyle {
  type: string;
  color?: string;
  sizeEighthPt?: number;
  spacePt?: number;
  frame?: boolean;
  shadow?: boolean;
}

export interface ParagraphBorderSet {
  top?: ParagraphBorderStyle;
  right?: ParagraphBorderStyle;
  bottom?: ParagraphBorderStyle;
  left?: ParagraphBorderStyle;
  between?: ParagraphBorderStyle;
  bar?: ParagraphBorderStyle;
}

export interface ParagraphTabStop {
  alignment?: "left" | "center" | "right" | "decimal" | "bar";
  leader?: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
  positionTwips?: number;
}

export interface ParagraphStyle {
  align?: ParagraphAlignment;
  direction?: ParagraphDirection;
  headingLevel?: HeadingLevel;
  styleId?: string;
  styleName?: string;
  numbering?: ParagraphNumbering;
  spacing?: ParagraphSpacing;
  indent?: ParagraphIndent;
  backgroundColor?: string;
  borders?: ParagraphBorderSet;
  tabStops?: ParagraphTabStop[];
  contextualSpacing?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  dropCap?: {
    type: "drop" | "margin";
    lines?: number;
    wrap?: string;
    horizontalAnchor?: string;
    verticalAnchor?: string;
    xTwips?: number;
    yTwips?: number;
    horizontalSpaceTwips?: number;
    verticalSpaceTwips?: number;
  };
}

export type SectionBreakType =
  | "nextPage"
  | "continuous"
  | "evenPage"
  | "oddPage"
  | "nextColumn";

/**
 * Section break carried by a paragraph (the paragraph that ends a section).
 * The raw `<w:sectPr>` fragment keeps the complete section properties — page
 * size, orientation, margins, header/footer references and `w:type` — so a
 * model mutation that invalidates `sourceXml` cannot silently drop a section.
 */
export interface ParagraphSectionBreak {
  sectionPropertiesXml: string;
}

export interface ParagraphCommentRange {
  /** `w:id` of the comment in `word/comments.xml`. */
  commentId: number;
  /** Inclusive UTF-16 offset into `paragraphText(paragraph)`. */
  start: number;
  /** Exclusive UTF-16 offset into `paragraphText(paragraph)`. */
  end: number;
  /**
   * The commented text when the range was captured. Edits remap the offsets;
   * when an edit cannot be mapped (a wholesale text replacement, for example)
   * the renderer re-anchors the range by searching for this text.
   */
  anchorText?: string;
}

export interface ParagraphNode {
  type: "paragraph";
  children: ParagraphChildNode[];
  style?: ParagraphStyle;
  paragraphMarkDeleted?: boolean;
  sourceXml?: string;
  sectionBreak?: ParagraphSectionBreak;
  /**
   * Comment ranges (`w:commentRangeStart` / `w:commentRangeEnd`) as character
   * offsets into this paragraph's plain text, captured at parse time and kept
   * up to date by the editing commands. Rendering and the comment gutter read
   * these instead of re-parsing `sourceXml`, which every edit clears.
   */
  commentRanges?: ParagraphCommentRange[];
}

export interface TableBoxSpacing {
  topTwips?: number;
  rightTwips?: number;
  bottomTwips?: number;
  leftTwips?: number;
}

export interface TableBorderStyle {
  type: string;
  color?: string;
  sizeEighthPt?: number;
}

export interface TableBorderSet {
  top?: TableBorderStyle;
  right?: TableBorderStyle;
  bottom?: TableBorderStyle;
  left?: TableBorderStyle;
  insideH?: TableBorderStyle;
  insideV?: TableBorderStyle;
  tl2br?: TableBorderStyle;
  tr2bl?: TableBorderStyle;
}

export interface TableCellStyle {
  backgroundColor?: string;
  gridSpan?: number;
  rowSpan?: number;
  vMergeContinuation?: boolean;
  widthTwips?: number;
  marginTwips?: TableBoxSpacing;
  verticalAlign?: "top" | "center" | "bottom";
  borders?: TableBorderSet;
}

export type TableCellContentNode = ParagraphNode | TableNode;

export interface TableCellNode {
  type: "table-cell";
  nodes: TableCellContentNode[];
  style?: TableCellStyle;
}

export interface TableRowStyle {
  backgroundColor?: string;
  heightTwips?: number;
  heightRule?: "auto" | "atLeast" | "exact";
  cantSplit?: boolean;
  isHeader?: boolean;
}

export interface TableRowNode {
  type: "table-row";
  cells: TableCellNode[];
  style?: TableRowStyle;
}

export interface TableStyle {
  widthTwips?: number;
  indentTwips?: number;
  layout?: "fixed" | "autofit";
  cellSpacingTwips?: number;
  cellMarginTwips?: TableBoxSpacing;
  columnWidthsTwips?: number[];
  borders?: TableBorderSet;
  floating?: {
    xTwips?: number;
    yTwips?: number;
    leftFromTextTwips?: number;
    rightFromTextTwips?: number;
    topFromTextTwips?: number;
    bottomFromTextTwips?: number;
    horizontalAnchor?: string;
    verticalAnchor?: string;
    horizontalAlign?: "left" | "center" | "right" | "inside" | "outside";
    verticalAlign?: "top" | "center" | "bottom" | "inside" | "outside";
  };
}

export interface TableNode {
  type: "table";
  rows: TableRowNode[];
  style?: TableStyle;
  sourceXml?: string;
}

export type DocNode = ParagraphNode | TableNode;

export interface HeaderSection {
  partName: string;
  referenceType?: string;
  nodes: DocNode[];
}

export interface FooterSection {
  partName: string;
  referenceType?: string;
  nodes: DocNode[];
}

export interface DocumentSection {
  startNodeIndex: number;
  sectionPropertiesXml?: string;
  headerSections: HeaderSection[];
  footerSections: FooterSection[];
}

export interface ParagraphStyleDefinition {
  id: string;
  name: string;
  basedOnId?: string;
  nextStyleId?: string;
  align?: ParagraphAlignment;
  direction?: ParagraphDirection;
  headingLevel?: HeadingLevel;
  numbering?: ParagraphNumbering;
  spacing?: ParagraphSpacing;
  indent?: ParagraphIndent;
  backgroundColor?: string;
  borders?: ParagraphBorderSet;
  tabStops?: ParagraphTabStop[];
  contextualSpacing?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  runStyle?: TextStyle;
  uiPriority?: number;
  isDefault?: boolean;
  isPrimary?: boolean;
}

export interface NumberingLevelDefinition {
  ilvl: number;
  start?: number;
  format?: string;
  text?: string;
  suffix?: "tab" | "space" | "nothing";
  /** OOXML `w:lvlJc` label alignment. */
  justification?: ParagraphAlignment;
  indent?: ParagraphIndent;
  runStyle?: TextStyle;
  bulletFontFamily?: string;
  bulletColor?: string;
  pictureBulletId?: number;
  pictureBullet?: NumberingPictureBulletDefinition;
}

export interface NumberingAbstractDefinition {
  abstractNumId: number;
  levels: NumberingLevelDefinition[];
}

export interface NumberingInstanceDefinition {
  numId: number;
  abstractNumId: number;
  levelStartOverrides?: Record<string, number>;
  levelOverrides?: NumberingLevelDefinition[];
}

export interface NumberingDefinitionSet {
  abstracts: NumberingAbstractDefinition[];
  instances: NumberingInstanceDefinition[];
}

export interface NumberingPictureBulletDefinition {
  numPicBulletId: number;
  src?: string;
  widthPx?: number;
  heightPx?: number;
  partName?: string;
  contentType?: string;
}

export interface DocumentNoteDefinition {
  id: number;
  text: string;
  nodes?: DocNode[];
}

export interface DocumentCommentAnchor {
  /** Region that owns the anchored paragraph. */
  region: "body" | "header" | "footer";
  /** OOXML part name for header/footer anchors (e.g. `word/header1.xml`). */
  partName?: string;
  /** Index into `model.nodes` (body) or the section's `nodes` array. */
  nodeIndex: number;
  /** Row index when the anchor is inside a table cell. */
  rowIndex?: number;
  /** Cell index when the anchor is inside a table cell. */
  cellIndex?: number;
  /** Paragraph index within the cell, counting nested-table paragraphs. */
  paragraphIndex?: number;
  /** UTF-16 offset of the anchor within the paragraph's plain text. */
  textOffset?: number;
  /** Index among the paragraph's visible text runs. */
  runIndex?: number;
  /** Offset within the anchored text run. */
  offset?: number;
  /** Plain-text excerpt covered by the comment range, when available. */
  anchorText?: string;
}

export interface DocumentCommentDefinition {
  id: number;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
  /** Comment id this comment replies to (from commentsExtended threading). */
  parentId?: number;
  /** True when the comment thread is marked done in commentsExtended. */
  resolved?: boolean;
}

export interface DocumentCompatibilitySettings {
  suppressSpacingBeforeAfterPageBreak?: boolean;
  usePrinterMetrics?: boolean;
  useFixedHtmlParagraphSpacing?: boolean;
  doNotBreakWrappedTables?: boolean;
  doNotBreakConstrainedForcedTable?: boolean;
  evenAndOddHeaders?: boolean;
  /**
   * OOXML `w:compat/w:compatSetting[@w:name='compatibilityMode']/@w:val`:
   * 11 = Word 2003, 12 = Word 2007, 13 = Word 2010, 14 = Word 2010 (current
   * default), 15 = Word 2013+. Word honours HTML auto spacing only in 12/15.
   */
  compatibilityMode?: number;
}

export type NoteNumberRestart = "continuous" | "eachSect" | "eachPage";
export type NoteNumberPosition = "docEnd" | "sectEnd";

/** Subset of w:footnotePr / w:endnotePr that changes note numbering/placement. */
export interface DocumentNoteNumberingProperties {
  /** OOXML w:numFmt/@w:val, e.g. "decimal", "lowerRoman", "upperLetter". */
  numFmt?: string;
  /** OOXML w:numRestart/@w:val. */
  numRestart?: NoteNumberRestart;
  /** OOXML w:numStart/@w:val — first display number of the numbering run. */
  numStart?: number;
  /** Endnote placement (w:endnotePr/w:pos): "docEnd" or "sectEnd". */
  position?: NoteNumberPosition;
}

export interface DocumentNoteProperties {
  footnote?: DocumentNoteNumberingProperties;
  endnote?: DocumentNoteNumberingProperties;
}

/** Page background pattern/shading resolved from w:background. */
export interface DocumentBackgroundPattern {
  /** OOXML preset pattern name, e.g. "pct20", "ltHorz" or "diagCross". */
  preset: string;
  /** Pattern foreground colour as a CSS hex colour. */
  foregroundColor?: string;
  /** Pattern background fill as a CSS hex colour. */
  backgroundColor?: string;
}

export interface DocModel {
  nodes: DocNode[];
  metadata: {
    sourceParts: number;
    warnings: string[];
    documentPageCount?: number;
    documentOpenTag?: string;
    documentBackgroundColor?: string;
    documentBackgroundPattern?: DocumentBackgroundPattern;
    sectionPropertiesXml?: string;
    sections?: DocumentSection[];
    headerSections: HeaderSection[];
    footerSections: FooterSection[];
    paragraphStyles: ParagraphStyleDefinition[];
    defaultParagraphStyleId?: string;
    /**
     * OOXML `w:defaultTabStop/@w:val` from word/settings.xml, in twips. The
     * default is 720 twips (0.5"), which resolves to 48 CSS pixels.
     */
    defaultTabStopTwips?: number;
    themeFonts?: DocxThemeFonts;
    numberingDefinitions?: NumberingDefinitionSet;
    compatibility?: DocumentCompatibilitySettings;
    footnotes?: DocumentNoteDefinition[];
    endnotes?: DocumentNoteDefinition[];
    comments?: DocumentCommentDefinition[];
    /** Document-wide w:footnotePr / w:endnotePr from word/settings.xml. */
    noteProperties?: DocumentNoteProperties;
  };
}
