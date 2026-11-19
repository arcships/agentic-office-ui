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
/**
 * How revision markup is drawn on the page.
 *
 * - `inline` draws insertions and deletions in the text flow and hides the
 *   tracked-change gutter.
 * - `balloons` keeps insertions inline but moves deleted content out of the text
 *   flow onto gutter cards.
 * - `none` shows the final text only: no inline markup and no revision cards.
 *
 * The editor model is never rewritten by this choice; export still writes every
 * revision.
 */
export type DocxRevisionDisplay = "inline" | "balloons" | "none";
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
      /** Runtime key for a cell in a nested table; absent for body tables. */
      tableRuntimeKey?: string;
      /**
       * Multi-cell rectangular selection inside this table. `rowIndex` and
       * `cellIndex` point at the focus (active) cell; `cellRange` carries the
       * anchor so merge/split and table commands can work on whole ranges.
       */
      cellRange?: DocxTableCellSelectionRange;
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
  /**
   * Physical table-grid columns covered by the rectangle, after merged cells
   * and their grid spans are taken into account. 0-based and inclusive.
   *
   * Anchor/focus address model cells; this range addresses the drawn grid, so
   * a cell that spans two columns stays inside a selection that only touches
   * one of them. Commands, highlighting and clipboard export all use it.
   */
  gridRange?: DocxTableCellGridRange;
}

export interface DocxTableCellGridRange {
  startColumnIndex: number;
  endColumnIndex: number;
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

/**
 * Clipboard flavours a paste event or the async clipboard API can provide.
 *
 * Paste prefers the internal structured payload, then `text/html`, then
 * `text/rtf`, then `text/plain`.
 */
export interface DocxClipboardSources {
  /** Internal structured DOCX fragment (`DOCX_CLIPBOARD_MIME`). */
  internal?: string;
  /** `text/html` flavour, e.g. copied from Word. */
  html?: string;
  /** `text/rtf` flavour, for editors that only take RTF. */
  rtf?: string;
  /** `text/plain` flavour; always usable as the last resort. */
  text?: string;
}

/**
 * Minimal `DataTransfer` surface the copy path writes into.
 *
 * Declared structurally so the editor does not depend on DOM lib types.
 */
export interface DocxClipboardDataSink {
  setData(type: string, value: string): void;
}

// Range-comparison helpers (upstream editor.tsx:3413-3480). Pure functions
// colocated with the range types; consumed by selection-helpers and the
// pagination/normalize paths.
export function compareParagraphLocations(
  a: ParagraphLocation,
  b: ParagraphLocation
): number {
  if (a.kind === "paragraph" && b.kind === "paragraph") {
    return Math.sign(a.nodeIndex - b.nodeIndex);
  }

  if (a.kind === "paragraph" && b.kind === "table-cell") {
    if (a.nodeIndex !== b.tableIndex) {
      return Math.sign(a.nodeIndex - b.tableIndex);
    }

    return -1;
  }

  if (a.kind === "table-cell" && b.kind === "paragraph") {
    if (a.tableIndex !== b.nodeIndex) {
      return Math.sign(a.tableIndex - b.nodeIndex);
    }

    return 1;
  }

  if (a.kind === "table-cell" && b.kind === "table-cell") {
    if (a.tableIndex !== b.tableIndex) {
      return Math.sign(a.tableIndex - b.tableIndex);
    }

    if (a.rowIndex !== b.rowIndex) {
      return Math.sign(a.rowIndex - b.rowIndex);
    }

    if (a.cellIndex !== b.cellIndex) {
      return Math.sign(a.cellIndex - b.cellIndex);
    }

    return Math.sign(a.paragraphIndex - b.paragraphIndex);
  }

  return 0;
}

export function compareTextRangeBoundaries(
  a: DocxTextRangeBoundary,
  b: DocxTextRangeBoundary
): number {
  const locationCompare = compareParagraphLocations(a.location, b.location);
  if (locationCompare !== 0) {
    return locationCompare;
  }

  return Math.sign(a.offset - b.offset);
}

export function normalizeTextRange(range: DocxTextRange): DocxTextRange {
  if (compareTextRangeBoundaries(range.start, range.end) <= 0) {
    return {
      start: range.start,
      end: range.end,
    };
  }

  return {
    start: range.end,
    end: range.start,
  };
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
  /**
   * Coalesces consecutive text commits into one undo entry.
   *
   * A patch that carries the same key as the previous patch, with no
   * intervening transaction, folds into the undo entry opened by the first
   * commit of the burst. Selection changes, formatting/structural commands,
   * undo/redo, a different key or `endTextCommitBurst()` close the burst.
   */
  textCommitKey?: string;
}

export type ParagraphLocation = DocxTextRangeLocation;

/** Story that owns a search result. */
export type DocxReplaceRegion =
  | "body"
  | "header"
  | "footer"
  | "footnote"
  | "endnote";

/**
 * One search result handed to `replaceMatch()`.
 *
 * `location` points at the paragraph, `start`/`end` are UTF-16 offsets into
 * that paragraph's plain text. `region` is advisory: only `body` results can be
 * replaced, so a header/footer/note result is rejected instead of silently
 * rewriting a read-only story.
 */
export interface DocxReplaceTarget {
  region: DocxReplaceRegion;
  location: DocxTextRangeLocation | {
    kind: "nested-table-cell";
    tableIndex: number;
    path: readonly { rowIndex: number; cellIndex: number; contentIndex: number }[];
    rowIndex: number;
    cellIndex: number;
    paragraphIndex: number;
  };
  start: number;
  end: number;
}

/** Query options shared by replace-all and the search session. */
export interface DocxReplaceAllOptions {
  caseSensitive?: boolean;
  wholeWord?: boolean;
  regex?: boolean;
}

/** What one replace command actually changed. */
export interface DocxReplaceSummary {
  /** Number of occurrences replaced. */
  replaced: number;
  /** Matches skipped because their story is not editable. */
  skipped: number;
}

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
  | "paragraph-format-change"
  | "table-format-change"
  | "row-format-change"
  | "cell-format-change";

export type DocxInlineTrackedChangeKind = Exclude<
  DocxTrackedChangeKind,
  | "format-change"
  | "paragraph-format-change"
  | "table-format-change"
  | "row-format-change"
  | "cell-format-change"
>;

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
  /** Region that owns the anchored paragraph (defaults to the body). */
  region?: "body" | "header" | "footer";
  /** OOXML part name for header/footer comments. */
  partName?: string;
}

export type DocxLineSpacingRule = "auto" | "exact" | "atLeast";

/**
 * State of the inline comment editor: a new comment on a selection, or a reply
 * to an existing comment. The composer owns no model state until it is
 * submitted, so cancelling it leaves the document and its history untouched.
 */
export interface DocxCommentComposer {
  mode: "new" | "reply";
  /** Reply target, set when `mode` is `"reply"`. */
  commentId?: number;
  /** Selection to anchor the new comment to, set when `mode` is `"new"`. */
  range?: DocxTextRange;
  /** Start location of `range`, used to place the editor beside its page. */
  anchor?: DocxTextRangeLocation;
}

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

/** Status messages emitted by editor commands, keyed for host localization. */
export interface DocxEditorStatusMessages {
  ready: string;
  copied: string;
  pasted: string;
  listParagraphNotFound: string;
  listAppliedBullet: string;
  listAppliedNumbered: string;
  /** Template with a `{level}` placeholder. */
  maximumListLevel: string;
  discardChanges: string;
  /** Template with a `{name}` placeholder. */
  loadingFile: string;
  /** Template with a `{message}` placeholder. */
  importFailed: string;
  unsupportedFile: string;
  exportCancelledChanged: string;
  exportNotReady: string;
  /** Template with a `{message}` placeholder. */
  exportFailed: string;
  /** Template with a `{label}` placeholder. */
  exportRequiredField: string;
  exportRequiredFieldEmpty: string;
  exporting: string;
  /** Template with a `{phase}` placeholder. */
  exportingPhase: string;
  exported: string;
  unknownImportError: string;
  unknownExportError: string;
  /** Template with a `{count}` placeholder. */
  acceptedChanges: string;
  acceptedChange: string;
  commentAdded: string;
  commentReplied: string;
  commentResolved: string;
  commentReopened: string;
  commentDeleted: string;
  /** The selection or comment the command targets is gone. */
  commentTargetMissing: string;
  commentTextRequired: string;
  /** Template with a `{count}` placeholder. */
  rejectedChanges: string;
  rejectedChange: string;
  /** Template with a `{count}` placeholder. */
  formattingCleared: string;
  formattingClearedOne: string;
  /** Template with a `{feature}` placeholder. */
  unsupportedFeature: string;
  imageTargetUnavailable: string;
  imageSizePositive: string;
  resizeTargetUnavailable: string;
  lockedImageAnchorResize: string;
  imageOverlapDisabled: string;
  syntheticTextBoxUnavailable: string;
  imageMoveTargetUnavailable: string;
  lockedImageAnchorMove: string;
  sectionImageTargetUnavailable: string;
  dropCapTargetUnavailable: string;
  dropCapFontSizeTargetUnavailable: string;
  dropCapTextTargetUnavailable: string;
  imageMoveSourceUnavailable: string;
  selectImageFile: string;
  imageInsertionCancelled: string;
  /** Template with a `{name}` placeholder. */
  insertedImage: string;
  /** Template with `{width}` and `{height}` placeholders. */
  resizedImage: string;
  updatedTextBox: string;
  /** Template with a `{mode}` placeholder. */
  imageWrapSet: string;
  /** The image is not floating, so it has no layer order. */
  imageLayerOrderRequiresFloating: string;
  imageLayerOrderChanged: string;
  movedImage: string;
  movedSectionImage: string;
  movedDropCap: string;
  /** Template with a `{size}` placeholder. */
  dropCapFontSize: string;
  updatedDropCapText: string;
  tableColumnWidthPositive: string;
  tableColumnWidthTargetUnavailable: string;
  cannotInsertRowTableNotFound: string;
  cannotInsertTableRow: string;
  cannotInsertColumnTableNotFound: string;
  cannotInsertTableColumn: string;
  cannotDeleteRowTableNotFound: string;
  cannotDeleteOnlyTableRow: string;
  cannotDeleteColumnTableNotFound: string;
  cannotDeleteOnlyTableColumn: string;
  cannotDeleteTableColumn: string;
  cannotDeleteTableNotFound: string;
  cannotClearCellsTableNotFound: string;
  cannotMergeCellsTableNotFound: string;
  cannotMergeSingleCell: string;
  cannotSplitCellTableNotFound: string;
  cannotSplitCellNotMerged: string;
  cannotMoveTableNotFound: string;
  cannotMoveEmbeddedHostTableNotFound: string;
  cannotMoveEmbeddedHostCellNotFound: string;
  cannotMoveEmbeddedTableNotFound: string;
  insertedTable: string;
  /** Template with a `{direction}` placeholder. */
  insertedTableRow: string;
  /** Template with a `{direction}` placeholder. */
  insertedTableColumn: string;
  deletedTableRow: string;
  deletedTableColumn: string;
  deletedTable: string;
  clearedTableCells: string;
  resizedTableColumns: string;
  mergedTableCells: string;
  splitTableCell: string;
  movedTable: string;
  movedEmbeddedTable: string;
  resizedTableRow: string;
  tableRowHeightPositive: string;
  tableRowHeightTargetUnavailable: string;
  floatedTable: string;
  inlinedTable: string;
  cannotFloatTableNotFound: string;
  movedBodyNode: string;
  cannotMoveBodyNodeNotFound: string;
  updatedFormField: string;
  formFieldTargetUnavailable: string;
  formFieldOptionsRequired: string;
  /** Template with a `{preset}` placeholder. */
  unsupportedDiagonalTableCellBorders: string;
  /** Template with a `{preset}` placeholder. */
  unsupportedDiagonalParagraphBorders: string;
  revisionLossOne: string;
  /** Template with a `{count}` placeholder. */
  revisionLossMany: string;
  pageMarginsUpdated: string;
  pageSizeUpdated: string;
  sectionColumnsUpdated: string;
  headerAdded: string;
  footerAdded: string;
  /** Template with a `{count}` placeholder. */
  replacedMatches: string;
  replaceTargetUnavailable: string;
}

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
   * Localized status messages for commands. Can be a getter so a host can
   * switch locale while the editor stays mounted.
   */
  statusMessages?:
    | Partial<DocxEditorStatusMessages>
    | (() => Partial<DocxEditorStatusMessages>);
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
   * How revisions render while they are visible.
   *
   * `inline` keeps insertions and deletions in the text flow, `balloons` moves
   * deleted content onto gutter cards, and `none` shows the final text only.
   *
   * @defaultValue `"balloons"`
   */
  initialRevisionDisplay?: DocxRevisionDisplay;
  /**
   * Whether document comments are visible when the editor first mounts.
   *
   * @defaultValue `false`
   */
  initialShowComments?: boolean;
  /**
   * Author name stored on comments created through the editor commands.
   *
   * @defaultValue `"Author"`
   */
  commentAuthor?: string;
  /** Comment initials; derived from `commentAuthor` when omitted. */
  commentInitials?: string;
  /**
   * Clock used for a new comment's `w:date`. Hosts can pin it for
   * reproducible exports and deterministic tests.
   */
  commentClock?: () => Date;
  /**
   * Maximum number of undo and redo snapshots retained by this editor
   * instance.
   *
   * @defaultValue `100`
   */
  historyMaxEntries?: number;
  /**
   * Approximate byte budget for undo and redo snapshots retained by this
   * editor instance. The newest oversized snapshot is kept by itself so the
   * latest edit remains undoable.
   *
   * @defaultValue `32 * 1024 * 1024`
   */
  historyMaxBytes?: number;
}
