export {
  addWorkbookConditionalFormatRule,
  addWorkbookTable,
  appendWorkbookSheetDrawingAttachment,
  createConditionalFormatRuleFromDraft,
  normalizeConditionalFormatColor,
  deleteWorkbookConditionalFormatRule,
  deleteWorkbookTable,
  readWorkbookSheetAuthoringState,
  replaceWorkbookConditionalFormatRules,
  replaceWorkbookDataValidations,
  setWorkbookCellComment,
  setWorkbookCellHyperlink,
  setWorkbookDataValidation,
  setWorkbookFreezePanes,
  updateWorkbookConditionalFormatRule,
  updateWorkbookTableRange,
} from "./sheet-authoring";
export type {
  WorkbookAuthoringAssets,
  WorkbookSheetAuthoringState,
  XlsxCellCommentInput,
  XlsxCellHyperlinkInput,
  XlsxTableInsertInput,
} from "./sheet-authoring";
export {
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
  applyWorkbookDefinedNames,
  normalizeWorkbookDefinedNames,
  readWorkbookDefinedNames,
  remapA1References,
  remapAxisIndexForStructure,
  remapCellForStructure,
  remapConditionalFormatRuleForStructure,
  remapDataValidationForStructure,
  remapDefinedNameForStructure,
  remapFreezePanesForStructure,
  remapImageAnchorForStructure,
  remapRangeForStructure,
  remapRangesForStructure,
  remapSparklineForStructure,
  remapTableForStructure,
  shiftWorkbookSheetStructure,
} from "./sheet-structure";
export type {
  XlsxStructureAxis,
  XlsxStructureChange,
} from "./sheet-structure";
export type { XlsxConditionalFormatDraft } from "../types";
export { mergeWorkbookSheetAuthoringAssets } from "./sheet-merge";
