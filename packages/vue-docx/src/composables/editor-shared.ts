// Shared context type for DOCX editor sub-modules.
// Each sub-module (editor-transaction, editor-history, etc.) receives an
// EditorCore instance that holds all shared reactive state, mutable
// snapshots, and pre-computed derived values.

import type { Ref, ComputedRef } from "vue"
import type {
  DocModel,
  ParagraphNode,
  TextRunNode,
  OoxmlPackage,
  ParagraphStyleDefinition,
} from "@arcships/docx-core"
import type {
  DocxDocumentTheme,
  DocxEditorSelection,
  DocxTextRange,
  DocxHistorySnapshot,
  DocxHistoryRestoreRequest,
  DocxFormFieldLocation,
  DocxSelectedFormField,
  DocxSelectionSessionKind,
  DocxLineSpacingInfo,
  DocxBorderContext,
  DocxBorderPresetState,
  DocxListType,
  DocxTrackedChange,
  DocxComment,
  DocxPaginationInfo,
  ParagraphLocation,
  DocxEditorStatusMessages,
} from "@arcships/docx-core"

export interface ApplyModelChangeOptions {
  selection?: DocxEditorSelection
  activeTextRange?: DocxTextRange | undefined
  /**
   * Form-field edits keep the field selected so its focus ring survives the
   * model update; regular edits clear the selection as before.
   */
  clearSelectedFormField?: boolean
  /** Set to false to fold a change into the previous undo entry. */
  pushHistory?: boolean
  /**
   * Coalesces consecutive text commits into one undo entry.
   *
   * Commits that carry the same key with no intervening transaction share the
   * undo entry opened by the first commit of the burst; any other transaction
   * (selection change, formatting command, structural edit, undo/redo, a
   * different key) or `endTextCommitBurst()` closes the burst, so the next text
   * commit starts a new entry.
   */
  textCommitKey?: string
}

export interface EditorCore {
  // ── core reactive refs ──────────────────────────────────────────
  model: Ref<DocModel>
  /** Model snapshot as of the last load/new/export, used for dirty tracking. */
  cleanModel: Ref<DocModel>
  isDirty: ComputedRef<boolean>
  selection: Ref<DocxEditorSelection>
  activeTextRange: Ref<DocxTextRange | undefined>
  pendingRunStyle: Ref<TextRunNode["style"] | undefined>
  history: Ref<{ past: DocxHistorySnapshot[]; future: DocxHistorySnapshot[] }>
  historyBudget: { maxEntries: number; maxBytes: number }
  estimateHistorySnapshotBytes: (snapshot: DocxHistorySnapshot) => number
  status: Ref<string>
  /** Resolve a localized editor status message with placeholder values. */
  statusMessage: (
    key: keyof DocxEditorStatusMessages,
    values?: Record<string, string | number>,
  ) => string
  selectedFormFieldLocation: Ref<DocxFormFieldLocation | undefined>
  documentLoadNonce: Ref<number>
  fileName: Ref<string>
  /** File name restored by newDocument(). */
  initialFileName: string
  importError: Ref<Error | undefined>
  isImporting: Ref<boolean>
  documentTheme: Ref<DocxDocumentTheme>
  showTrackedChanges: Ref<boolean>
  showComments: Ref<boolean>
  paginationInfo: Ref<DocxPaginationInfo>
  historyRestoreRequest: Ref<DocxHistoryRestoreRequest | undefined>
  basePackage: Ref<OoxmlPackage | undefined>
  selectionSessionKind: Ref<DocxSelectionSessionKind>
  selectionSessionKindInternal: Ref<DocxSelectionSessionKind>
  embeddedFontLoadNonce: Ref<number>

  // ── mutable snapshots (synchronous access for transaction resolvers) ─
  modelSnapshot: { value: DocModel }
  selectionSnapshot: { value: DocxEditorSelection }
  activeTextRangeSnapshot: { value: DocxTextRange | undefined }
  pendingRunStyleSnapshot: { value: TextRunNode["style"] | undefined }

  // ── mutable flags & bookkeeping ─────────────────────────────────
  historyRestoreNonce: { value: number }
  suppressNextDomSelectionRestore: { value: boolean }
  domSelectionRestoreModel: { value: DocModel }
  suppressSelectionReset: { value: boolean }
  selectionSessionTimeout: { value: ReturnType<typeof setTimeout> | null }
  loadedEmbeddedFontFaces: { value: FontFace[] }
  activeImportAbortController: { value: AbortController | undefined }
  pendingExportModelTransformer: { value: ((model: DocModel) => DocModel) | undefined }
  flushTextDrafts: () => void
  /** Ends the current text-commit burst so the next keystroke opens a new undo entry. */
  endTextCommitBurst: () => void
  activeFormFieldEditKey: { value: string | undefined }
  starterTemplate: DocModel

  // ── computed refs (pre-computed in useDocxEditor, read by sub-modules) ─
  selectedParagraph: ComputedRef<ParagraphNode | undefined>
  selectedParagraphLocation: ComputedRef<ParagraphLocation>
  selectedRunStyle: ComputedRef<TextRunNode["style"] | undefined>
  selectedLink: ComputedRef<string | undefined>
  selectedParagraphStyleId: ComputedRef<string | undefined>
  selectedLineSpacing: ComputedRef<DocxLineSpacingInfo>
  selectedBorderContext: ComputedRef<DocxBorderContext>
  activeBorderPresets: ComputedRef<DocxBorderPresetState>
  availableParagraphStyles: ComputedRef<ParagraphStyleDefinition[]>
  trackedChanges: ComputedRef<DocxTrackedChange[]>
  comments: ComputedRef<DocxComment[]>
  selectedListType: ComputedRef<DocxListType | undefined>
  selectedFormField: ComputedRef<DocxSelectedFormField | undefined>
  hasUnorderedList: ComputedRef<boolean>
  hasOrderedList: ComputedRef<boolean>
  canUndo: ComputedRef<boolean>
  canRedo: ComputedRef<boolean>
}
