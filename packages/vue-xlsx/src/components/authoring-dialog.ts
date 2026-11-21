export type XlsxAuthoringDialogKind =
  | "chart"
  | "comment"
  | "hyperlink"
  | "table"
  | "conditional"
  | "validation"
  | "pasteSpecial"
  | "numberFormat"
  | "names"
  /** Worksheet page set-up; renders `XlsxPageSetupDialog` instead of a section. */
  | "pageSetup";
