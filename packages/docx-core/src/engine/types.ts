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

/** OOXML `w14:ligatures` / Word's "Use ligatures" list. */
export type TextLigatures =
  | "none"
  | "standard"
  | "contextual"
  | "historical"
  | "discretional"
  | "standardContextual"
  | "standardHistorical"
  | "contextualHistorical"
  | "standardDiscretional"
  | "contextualDiscretional"
  | "historicalDiscretional"
  | "standardContextualHistorical"
  | "standardContextualDiscretional"
  | "standardHistoricalDiscretional"
  | "contextualHistoricalDiscretional"
  | "all";

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
  /**
   * Font family as imported before any user edit. Script-font resolution uses
   * it to tell a structured `w:rFonts` style from a legacy single-family one.
   */
  sourceFontFamily?: string;
  /** OOXML `w:rFonts/@w:ascii` slot. */
  fontFamilyAscii?: string;
  /** OOXML `w:rFonts/@w:hAnsi` slot. */
  fontFamilyHAnsi?: string;
  /** East-Asian font family (`w:rFonts/@w:eastAsia`). */
  fontFamilyEastAsia?: string;
  /** Complex-script font family (`w:rFonts/@w:cs`). */
  fontFamilyComplexScript?: string;
  /** Complex-script font size (`w:szCs`), in points. */
  fontSizeCsPt?: number;
  /** Minimum size (`w:kern/@w:val`, points) at which kerning is enabled. */
  kerningMinFontSizePt?: number;
  /** OOXML `w14:ligatures`; `none` unless the document asks otherwise. */
  ligatures?: TextLigatures;
  /** OOXML `w14:cntxtAlts`; enables contextual alternates when true. */
  contextualAlternates?: boolean;
  /** OOXML `w:rFonts/@w:hint` (`eastAsia`, `cs`, `default`). */
  fontHint?: string;
  /** OOXML `w:lang/@w:val`, used to classify neutral characters. */
  language?: string;
  /** OOXML `w:lang/@w:eastAsia`. */
  languageEastAsia?: string;
  /** OOXML `w:lang/@w:bidi`. */
  languageBidi?: string;
  /** OOXML `w:rtl`; keeps `direction` as the rendering compatibility flag. */
  rightToLeft?: boolean;
  /** OOXML `w:cs`; forces complex-script shaping for the run. */
  complexScript?: boolean;
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

/**
 * Anchor, wrap and paint-order of a drawing object (`wp:anchor`), shared by
 * floating images and the drawing objects this model carries (shapes, charts,
 * embedded objects).
 */
export interface DrawingAnchor {
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
}

/** One styled text segment of a text box paragraph. */
export interface TextBoxRunSegment {
  text: string;
  style?: TextStyle;
}

/**
 * One paragraph of a text box (`w:txbxContent`). The text box content is part of
 * the model so it stays editable and renderable after an edit clears the run's
 * `sourceXml`.
 */
export interface TextBoxParagraph {
  align?: ParagraphAlignment;
  lineHeightPx: number;
  segments: TextBoxRunSegment[];
}

/** Text content of a shape/text box drawing (`w:txbxContent`). */
export interface TextBoxContent {
  paragraphs: TextBoxParagraph[];
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
  floating?: DrawingAnchor;
  /**
   * True when the image is a synthetic render of a Word text box / shape.
   */
  syntheticTextBox?: boolean;
  textBoxText?: string;
  /**
   * The text box content parsed at import. The renderer prefers it over
   * re-parsing `sourceXml`, which an edit clears.
   */
  textBox?: TextBoxContent;
  /**
   * The original `w:drawing`/`w:pict` markup of a synthetic text box. Editing a
   * text box clears `sourceXml` (the engine must rebuild the run) but the
   * drawing itself still has to reach the export; this copy is what the
   * serializer writes back.
   */
  textBoxSourceXml?: string;
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

/**
 * The drawing objects a Word document can carry that are not pictures: text
 * boxes, autoshapes, WordArt, shape groups, SmartArt, charts, embedded OLE
 * objects and 3D models. Everything is preserved verbatim in `sourceXml` so an
 * export writes the object back exactly as it was imported.
 */
export type DrawingObjectKind =
  | "textBox"
  | "shape"
  | "wordArt"
  | "group"
  | "smartArt"
  | "chart"
  | "oleObject"
  | "model3d";

/** One cached chart series (`c:ser`): name, categories and plotted values. */
export interface ChartSeriesData {
  name?: string;
  categories?: string[];
  values?: Array<number | null>;
}

/** Chart part data (`word/charts/chartN.xml`) carried by a `chart` drawing. */
export interface ChartData {
  /** OOXML plot element, e.g. "barChart", "lineChart" or "pieChart". */
  chartType?: string;
  title?: string;
  categories?: string[];
  series?: ChartSeriesData[];
}

/** Fill/stroke a shape placeholder paints with. */
export interface DrawingObjectShapeStyle {
  preset?: string;
  fillColor?: string;
  strokeColor?: string;
  strokeWidthPx?: number;
}

/**
 * A non-picture drawing object inside a paragraph. Layout uses `widthPx` /
 * `heightPx` and the optional `floating` anchor; rendering falls back to a
 * placeholder built from the available content (`textBox`, `chart`, `shape`,
 * `ole`) because the document model cannot render the object itself.
 */
export interface DrawingRunNode {
  type: "drawing";
  kind: DrawingObjectKind;
  name?: string;
  alt?: string;
  widthPx?: number;
  heightPx?: number;
  /** Anchor and z-order (`wp:anchor`); absent for inline (`wp:inline`) objects. */
  floating?: DrawingAnchor;
  /** Paint order derived from Word's `relativeHeight`, when floating. */
  zIndex?: number;
  /** Local preview (data URI) when the importer could produce one. */
  src?: string;
  textBox?: TextBoxContent;
  chart?: ChartData;
  shape?: DrawingObjectShapeStyle;
  ole?: {
    progId?: string;
    drawAspect?: string;
  };
  /** Original `w:drawing` / `w:object` / `w:pict` markup, written back on export. */
  sourceXml: string;
}

/**
 * An equation (`m:oMath` / `m:oMathPara`). The linearized `text` keeps the
 * equation visible when the browser cannot render OMML; `sourceXml` is written
 * back unchanged on export.
 */
export interface MathRunNode {
  type: "math";
  /** Linearized equation text, e.g. `x^2 + a/b`. */
  text: string;
  /** True for a display equation (`m:oMathPara`). */
  display?: boolean;
  style?: TextStyle;
  sourceXml: string;
}

export type ParagraphChildNode =
  | TextRunNode
  | ImageRunNode
  | FormFieldRunNode
  | DrawingRunNode
  | MathRunNode;

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
  /**
   * Logical grid widths. When a merge collapsed zero-width columns the array
   * keeps the zeroes and `collapsedGridColumns` names them; serialization
   * projects the table onto the physical grid before the width list is written.
   */
  columnWidthsTwips?: number[];
  /**
   * OOXML `w:tblStyle/@w:val` of the table, when it carries an explicit style
   * reference.
   */
  styleId?: string;
  /**
   * The explicit style id the table was imported with. A collapsed grid is only
   * safe while the two agree: a different style can supply its own grid widths.
   */
  sourceStyleId?: string;
  /**
   * @internal Zero-width logical columns a merge collapsed. The export
   * preparation projects them away before the physical grid is written.
   */
  collapsedGridColumns?: number[];
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
