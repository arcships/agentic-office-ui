/**
 * Platform-neutral DOCX model entry.
 *
 * Keep this file explicit: adding a browser, Worker, WASM, Vue, or mutable
 * process-wide dependency here must fail the core-boundary gate.
 */

export type {
  DocModel,
  DocNode,
  DocumentCommentDefinition,
  DocumentCompatibilitySettings,
  DocumentNoteDefinition,
  DocumentSection,
  FooterSection,
  FormFieldCheckboxWidgetSettings,
  FormFieldDropdownWidgetSettings,
  FormFieldOption,
  FormFieldRunNode,
  FormFieldSourceKind,
  FormFieldTextWidgetSettings,
  FormFieldType,
  FormFieldWidgetSettings,
  HeaderSection,
  HeadingLevel,
  ImageRunNode,
  NumberingAbstractDefinition,
  NumberingDefinitionSet,
  NumberingInstanceDefinition,
  NumberingLevelDefinition,
  NumberingPictureBulletDefinition,
  ParagraphAlignment,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphChildNode,
  ParagraphIndent,
  ParagraphNode,
  ParagraphNumbering,
  ParagraphSpacing,
  ParagraphStyle,
  ParagraphStyleDefinition,
  ParagraphTabStop,
  TableBorderSet,
  TableBorderStyle,
  TableBoxSpacing,
  TableCellContentNode,
  TableCellNode,
  TableCellStyle,
  TableNode,
  TableRowNode,
  TableRowStyle,
  TableStyle,
  TextRunBorderStyle,
  TextRunNode,
  TextStyle,
} from "./engine/types";

export { cloneDocModel } from "./engine/clone";
export { normalizeDocModel } from "./engine/normalize";

export type {
  LayoutBlock,
  LayoutImageRun,
  LayoutOptions,
  LayoutPage,
  LayoutParagraphBlock,
  LayoutRun,
  LayoutTableBlock,
  LayoutTableCell,
  LayoutTableRow,
  LayoutTextRun,
} from "./layout/layout-engine";
export { layoutDocument } from "./layout/layout-engine";

export type {
  InsertParagraphOptions,
  UpdateTextOptions,
} from "./editor/paragraph-ops";
export {
  duplicateParagraph,
  insertParagraph,
  paragraphFromText,
  removeParagraph,
  updateParagraphText,
} from "./editor/paragraph-ops";
export {
  applyRunStyle,
  replaceText,
  setParagraphAlignment,
  toggleRunStyleFlag,
} from "./editor/run-style-ops";
export {
  updateTableCellParagraphText,
  updateTableCellText,
} from "./editor/table-ops";
