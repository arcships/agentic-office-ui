// Public API type definitions for the DOCX editor (part 1).
// Upstream editor.tsx: lines 3055-3481 — foundational selection/range/history/
// transaction/location/tracked-change/comment/spacing/border types.
//
// Framework-agnostic type contracts shared by editor helpers, composables, and
// Vue components. React-specific types are replaced with plain equivalents:
//   React.CSSProperties -> Record<string, string | number | undefined>
//   React.Dispatch<React.SetStateAction<T>> -> (value: T | ((prev: T) => T)) => void

import type {
  DocModel,
  FormFieldRunNode,
  HeadingLevel,
  ParagraphStyleDefinition,
  TextRunNode
} from "../../engine/types";

export type DocxDocumentTheme = "light" | "dark";
export type DocxHeadingStyleMap = Partial<
  Record<HeadingLevel, Record<string, string | number | undefined>>
>;

export type DocxListType = "unordered" | "ordered";

export type DocxEditorSelection =
  | {
      kind: "paragraph";
      nodeIndex: number;
    }
  | {
      kind: "table-cell";
      tableIndex: number;
      rowIndex: number;
      cellIndex: number;
    };

export interface DocxTableCellLocation {
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
}

export interface DocxTableCellSelectionRange {
  tableIndex: number;
  anchorRowIndex: number;
  anchorCellIndex: number;
  focusRowIndex: number;
  focusCellIndex: number;
}

export interface ParagraphLocationInBody {
  kind: "paragraph";
  nodeIndex: number;
}

export interface ParagraphLocationInCell {
  kind: "table-cell";
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
  paragraphIndex: number;
}

export type DocxTextRangeLocation =
  | ParagraphLocationInBody
  | ParagraphLocationInCell;

export interface DocxTextRangeBoundary {
  location: DocxTextRangeLocation;
  offset: number;
}

export interface DocxTextRange {
  start: DocxTextRangeBoundary;
  end: DocxTextRangeBoundary;
}

export interface DocxHistorySnapshot {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
}

export interface DocxHistoryRestoreRequest {
  nonce: number;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
}

export interface DocxEditorTransactionContext {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingRunStyle?: TextRunNode["style"];
}

export interface DocxEditorTransactionPatch {
  model?: DocModel;
  selection?: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingRunStyle?: TextRunNode["style"];
  status?: string;
  clearSelectedFormField?: boolean;
  pushHistory?: boolean;
}

export type ParagraphLocation = DocxTextRangeLocation;

export type DocxImageLocation = ParagraphLocation & {
  childIndex: number;
};

export type DocxSectionRegion = "header" | "footer";

export interface DocxSectionParagraphLocation {
  region: DocxSectionRegion;
  partName: string;
  nodeIndex: number;
  rowIndex?: number;
  cellIndex?: number;
  paragraphIndex?: number;
}

export type DocxSectionImageLocation = DocxSectionParagraphLocation & {
  childIndex: number;
};

export type DocxFormFieldLocation = ParagraphLocation & {
  childIndex: number;
};

export interface DocxSelectedFormField {
  location: DocxFormFieldLocation;
  field: FormFieldRunNode;
}

export type DocxImageDropTarget = ParagraphLocation & {
  childIndex: number;
};

export type DocxTrackedChangeKind =
  | "insertion"
  | "deletion"
  | "move-from"
  | "move-to"
  | "format-change"
  | "paragraph-format-change";

export interface DocxTrackedChange {
  id: string;
  inlineAnchorId?: string;
  kind: DocxTrackedChangeKind;
  author?: string;
  date?: string;
  text?: string;
  nodeIndex: number;
  location: DocxTextRangeLocation;
}

export interface DocxComment {
  /** Stable identifier (unique per rendered anchor). */
  id: string;
  /** Comment id from `word/comments.xml` (`w:id`). */
  commentId: number;
  author?: string;
  initials?: string;
  date?: string;
  /** Plain-text comment body. */
  text: string;
  /** Comment id this comment replies to, when part of a thread. */
  parentId?: number;
  /** True when the comment thread is marked done. */
  resolved?: boolean;
  /** Plain-text excerpt of the commented document range. */
  anchorText?: string;
  nodeIndex: number;
  location: DocxTextRangeLocation;
}

export type DocxLineSpacingRule = "auto" | "exact" | "atLeast";

export type DocxSelectionSessionKind =
  | "idle"
  | "pointer"
  | "keyboard"
  | "composition"
  | "command"
  | "history-restore";

export interface DocxLineSpacingInfo {
  lineRule: DocxLineSpacingRule;
  lineTwips?: number;
  multiple: number;
}

export type DocxBorderContext = "paragraph" | "table";

export type DocxBorderPreset =
  | "bottom"
  | "top"
  | "left"
  | "right"
  | "none"
  | "all"
  | "outside"
  | "inside"
  | "inside-horizontal"
  | "inside-vertical"
  | "diagonal-down"
  | "diagonal-up"
  | "horizontal-line";

export type DocxBorderPresetState = Record<DocxBorderPreset, boolean>;

export interface UseDocxEditorOptions {
  /**
   * Model used for `editor.newDocument()` and for the initial empty editor.
   *
   * @defaultValue `defaultStarterModel`
   */
  starterModel?: DocModel;
  /**
   * Initial display name before a user imports a file.
   *
   * @defaultValue `"(new document)"`
   */
  initialFileName?: string;
  /**
   * Initial status text exposed on `editor.status`.
   *
   * @defaultValue `"Ready"`
   */
  initialStatus?: string;
  /**
   * Initial document surface theme.
   *
   * @defaultValue `"light"`
   */
  initialDocumentTheme?: DocxDocumentTheme;
  /**
   * Whether tracked changes are visible when the editor first mounts.
   *
   * @defaultValue `false`
   */
  initialShowTrackedChanges?: boolean;
  /**
   * Whether document comments are visible when the editor first mounts.
   *
   * @defaultValue `false`
   */
  initialShowComments?: boolean;
}
