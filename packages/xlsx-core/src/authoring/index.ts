export {
  addWorkbookConditionalFormatRule,
  addWorkbookTable,
  appendWorkbookSheetDrawingAttachment,
  createConditionalFormatRuleFromDraft,
  normalizeConditionalFormatColor,
  deleteWorkbookConditionalFormatRule,
  deleteWorkbookTable,
  readWorkbookSheetAuthoringState,
  setWorkbookCellComment,
  setWorkbookCellHyperlink,
  setWorkbookDataValidation,
  setWorkbookFreezePanes,
  updateWorkbookConditionalFormatRule,
} from "./sheet-authoring";
export type {
  WorkbookAuthoringAssets,
  WorkbookSheetAuthoringState,
  XlsxCellCommentInput,
  XlsxCellHyperlinkInput,
  XlsxTableInsertInput,
} from "./sheet-authoring";
export type { XlsxConditionalFormatDraft } from "../types";
export { mergeWorkbookSheetAuthoringAssets } from "./sheet-merge";
