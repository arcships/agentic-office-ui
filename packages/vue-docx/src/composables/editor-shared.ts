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
} from "@extend-ai/docx-core"
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
} from "@extend-ai/docx-core"

export interface EditorCore {
  // ── core reactive refs ──────────────────────────────────────────
  model: Ref<DocModel>
  selection: Ref<DocxEditorSelection>
  activeTextRange: Ref<DocxTextRange | undefined>
  pendingRunStyle: Ref<TextRunNode["style"] | undefined>
  history: Ref<{ past: DocxHistorySnapshot[]; future: DocxHistorySnapshot[] }>
  status: Ref<string>
  selectedFormFieldLocation: Ref<DocxFormFieldLocation | undefined>
  documentLoadNonce: Ref<number>
  fileName: Ref<string>
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
