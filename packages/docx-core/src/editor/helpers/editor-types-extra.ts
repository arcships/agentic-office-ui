// Public API type definitions for the DOCX editor (part 2).
// Upstream editor.tsx: lines 3482-4545 — DocxEditorController, context-menu,
// virtualization, viewer props, and thumbnail/observer result types.
//
// React-specific types are replaced with plain equivalents (see editor-types.ts).

import type {
  DocModel,
  FormFieldRunNode,
  HeadingLevel,
  ImageRunNode,
  ParagraphAlignment,
  ParagraphNode,
  ParagraphStyleDefinition,
  TextRunNode
} from "../../engine/types";
import type {
  DocxBorderContext,
  DocxBorderPreset,
  DocxBorderPresetState,
  DocxComment,
  DocxDocumentTheme,
  DocxEditorSelection,
  DocxFormFieldLocation,
  DocxHeadingStyleMap,
  DocxHistoryRestoreRequest,
  DocxImageDropTarget,
  DocxImageLocation,
  DocxLineSpacingInfo,
  DocxListType,
  DocxSectionImageLocation,
  DocxSectionParagraphLocation,
  DocxSelectedFormField,
  DocxSelectionSessionKind,
  DocxTextRange,
  DocxTextRangeLocation,
  DocxTrackedChange,
  ParagraphLocation
} from "./editor-types";
// Re-export the foundational types so consumers can import from one barrel.
export type {
  DocxDocumentTheme,
  DocxHeadingStyleMap,
  DocxListType,
  DocxEditorSelection,
  DocxTextRangeLocation,
  DocxTextRangeBoundary,
  DocxTextRange,
  ParagraphLocation,
  DocxImageLocation,
  DocxSectionRegion,
  DocxSectionParagraphLocation,
  DocxSectionImageLocation,
  DocxFormFieldLocation,
  DocxSelectedFormField,
  DocxImageDropTarget,
  DocxTrackedChangeKind,
  DocxTrackedChange,
  DocxComment,
  DocxLineSpacingRule,
  DocxSelectionSessionKind,
  DocxLineSpacingInfo,
  DocxBorderContext,
  DocxBorderPreset,
  DocxBorderPresetState,
  UseDocxEditorOptions
} from "./editor-types";

// Image wrap-mode types (upstream lines 1350-1381).
export type DocxImageWrapMode =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

export interface DocxImageWrapState {
  mode: DocxImageWrapMode;
  moveWithText: boolean;
  fixedPositionOnPage: boolean;
}

export interface DocxImageWrapMenuOption {
  actionId: DocxContextMenuActionId | (string & {});
  label: string;
  checked: boolean;
  disabled?: boolean;
  onSelect: () => void;
}

export interface UseDocxImageWrapMenuResult {
  state: DocxImageWrapState;
  wrapOptions: DocxImageWrapMenuOption[];
  positioningOptions: DocxImageWrapMenuOption[];
  editWrapBoundaryOption: DocxImageWrapMenuOption;
  moreLayoutOptionsOption: DocxImageWrapMenuOption;
  setMode: (mode: DocxImageWrapMode) => void;
  setMoveWithText: (moveWithText: boolean) => void;
}

export interface DocxEditorController {
  model: DocModel;
  documentLoadNonce: number;
  fileName: string;
  status: string;
  importError?: Error;
  isImporting: boolean;
  documentTheme: DocxDocumentTheme;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  historyRestoreRequest?: DocxHistoryRestoreRequest;
  selectedFormField?: DocxSelectedFormField;
  selectedParagraph?: ParagraphNode;
  selectedRunStyle?: TextRunNode["style"];
  selectedLink?: string;
  pendingRunStyle?: TextRunNode["style"];
  selectionSessionKind: DocxSelectionSessionKind;
  suppressNextDomSelectionRestore: () => void;
  beginSelectionSession: (
    kind: Exclude<DocxSelectionSessionKind, "idle">,
    options?: {
      settleAfterMs?: number;
    }
  ) => void;
  clearSelectionSession: (expectedKind?: DocxSelectionSessionKind) => void;
  selectedParagraphStyleId?: string;
  selectedLineSpacing: DocxLineSpacingInfo;
  selectedBorderContext: DocxBorderContext;
  activeBorderPresets: DocxBorderPresetState;
  availableParagraphStyles: ParagraphStyleDefinition[];
  trackedChanges: DocxTrackedChange[];
  showTrackedChanges: boolean;
  comments: DocxComment[];
  showComments: boolean;
  currentPage: number;
  totalPages: number;
  hasUnorderedList: boolean;
  hasOrderedList: boolean;
  canUndo: boolean;
  canRedo: boolean;
  registerPendingExportModelTransformer: (
    transformer?: (model: DocModel) => DocModel
  ) => void;
  setStatus: (value: string | ((prev: string) => string)) => void;
  setDocumentTheme: (theme: DocxDocumentTheme) => void;
  setShowTrackedChanges: (showTrackedChanges: boolean) => void;
  setShowComments: (showComments: boolean) => void;
  syncPaginationInfo: (pagination: DocxPaginationInfo) => void;
  toggleShowTrackedChanges: () => void;
  toggleShowComments: () => void;
  importDocxFile: (file: File) => Promise<void>;
  newDocument: () => void;
  exportDocx: () => void;
  undo: () => void;
  redo: () => void;
  setHeading: (heading?: HeadingLevel) => void;
  setParagraphStyle: (styleId?: string) => void;
  setLineSpacing: (lineMultiple: number) => void;
  setFontFamily: (fontFamily: string) => void;
  setFontSize: (fontSizePt: number) => void;
  toggleBold: () => void;
  toggleItalic: () => void;
  toggleUnderline: () => void;
  toggleStrike: () => void;
  toggleSuperscript: () => void;
  toggleSubscript: () => void;
  setHighlight: (highlight?: string) => void;
  setTextColor: (color?: string) => void;
  setLink: (link?: string) => void;
  selectFormField: (location?: DocxFormFieldLocation) => void;
  toggleFormCheckbox: (location: DocxFormFieldLocation) => void;
  setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void;
  updateFormFieldWidget: (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void;
  applyBorderPreset: (preset: DocxBorderPreset) => void;
  setAlignment: (align?: ParagraphAlignment) => void;
  toggleList: (listType: DocxListType) => void;
  adjustSelectedListDepth: (levelDelta: number, draftText?: string) => boolean;
  insertListItemAfterSelection: (
    draftText: string,
    startOffset: number,
    endOffset?: number,
    targetLocation?: ParagraphLocation
  ) =>
    | {
        paragraphIndex: number;
        caretOffset: number;
      }
    | undefined;
  splitParagraphAtSelection: (
    draftText: string,
    startOffset: number,
    endOffset?: number,
    targetLocation?: ParagraphLocation
  ) =>
    | {
        paragraphIndex: number;
        caretOffset: number;
      }
    | undefined;
  insertTable: () => void;
  insertImageFile: (file: File) => Promise<void>;
  appendParagraph: (text?: string) => number;
  resizeImage: (
    location: DocxImageLocation,
    widthPx: number,
    heightPx: number
  ) => void;
  setSyntheticTextBoxText: (location: DocxImageLocation, text: string) => void;
  setImageWrapMode: (
    location: DocxImageLocation,
    mode: DocxImageWrapMode,
    seedFloating?: Partial<NonNullable<ImageRunNode["floating"]>>
  ) => void;
  moveFloatingImage: (
    location: DocxImageLocation,
    patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ) => void;
  moveSectionFloatingImage: (
    location: DocxSectionImageLocation,
    patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ) => void;
  moveParagraphDropCap: (
    nodeIndex: number,
    patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>
  ) => void;
  setParagraphDropCapFontSizePt: (
    nodeIndex: number,
    fontSizePt: number
  ) => void;
  setParagraphDropCapText: (nodeIndex: number, text: string) => void;
  moveImage: (source: DocxImageLocation, target: DocxImageDropTarget) => void;
  setActiveTextRange: (range?: DocxTextRange) => void;
  selectParagraph: (nodeIndex: number) => void;
  selectTableCell: (
    tableIndex: number,
    rowIndex: number,
    cellIndex: number
  ) => void;
  clearTableCellContents: (
    tableIndex: number,
    cells: Array<{ rowIndex: number; cellIndex: number }>
  ) => void;
  insertTableRow: (
    tableIndex: number,
    rowIndex: number,
    direction: "above" | "below"
  ) => void;
  insertTableColumn: (
    tableIndex: number,
    cellIndex: number,
    direction: "left" | "right",
    rowIndex?: number
  ) => void;
  deleteTableRow: (tableIndex: number, rowIndex: number) => void;
  deleteTableColumn: (
    tableIndex: number,
    cellIndex: number,
    rowIndex?: number
  ) => void;
  deleteTable: (tableIndex: number) => void;
  moveTable: (tableIndex: number, targetNodeIndex: number) => void;
  moveEmbeddedTableToBody: (
    tableRuntimeKey: string,
    targetNodeIndex: number
  ) => void;
  replaceExpandedSelection: (
    text: string,
    range?: DocxTextRange
  ) => DocxTextRange | undefined;
  deleteExpandedSelection: (range?: DocxTextRange) => DocxTextRange | undefined;
  commitParagraphText: (nodeIndex: number, text: string) => void;
  commitTableCellText: (
    tableIndex: number,
    rowIndex: number,
    cellIndex: number,
    text: string
  ) => void;
  commitTableCellParagraphTextRecursive: (
    tableIndex: number,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ) => void;
  commitSectionParagraphText: (
    location: DocxSectionParagraphLocation,
    text: string
  ) => void;
  copy: () => Promise<void>;
  paste: () => Promise<void>;
}

export type DocxTableContextMenuActionId =
  | "insert-row-above"
  | "insert-row-below"
  | "insert-column-left"
  | "insert-column-right"
  | "delete-row"
  | "delete-column"
  | "delete-table";

export type DocxContextMenuActionId =
  | DocxTableContextMenuActionId
  | "cut"
  | "copy"
  | "paste"
  | "image-wrap-inline"
  | "image-wrap-square"
  | "image-wrap-tight"
  | "image-wrap-through"
  | "image-wrap-top-and-bottom"
  | "image-wrap-behind-text"
  | "image-wrap-in-front-of-text"
  | "image-edit-wrap-boundary"
  | "image-move-with-text"
  | "image-fix-position-on-page"
  | "image-more-layout-options"
  | "image-bring-to-front"
  | "image-bring-forward"
  | "image-in-front-of-text"
  | "image-send-to-back"
  | "image-send-backward"
  | "image-behind-text";

export interface DocxTableContextMenuContext {
  /** Zero-based table index in the document body. */
  tableIndex: number;
  /** Zero-based row index under the pointer. */
  rowIndex: number;
  /** Zero-based cell index under the pointer. */
  cellIndex: number;
}

export interface DocxTableContextMenuAction {
  /** Built-in table command id. */
  id: DocxTableContextMenuActionId;
  /** Human-readable menu label. */
  label: string;
  /** True when the action removes content or structure. */
  destructive?: boolean;
}

export interface DocxContextMenuAction {
  /** Built-in command id, or a custom id from your menu renderer. */
  id: DocxContextMenuActionId | (string & {});
  /** Human-readable menu label. */
  label: string;
  /** Optional keyboard shortcut text to display next to the label. */
  shortcut?: string;
  /** True when the action removes content or structure. */
  destructive?: boolean;
  /** Prevents the action from being selected. */
  disabled?: boolean;
  /** Marks toggle-style menu items as active. */
  checked?: boolean;
  /** Adds a visual separator before this item in default-style menus. */
  separatorBefore?: boolean;
  /** Nested submenu actions. */
  children?: DocxContextMenuAction[];
}

export interface DocxTableContextMenuRenderProps {
  /** Table cell that opened the menu. */
  context: DocxTableContextMenuContext;
  /** Built-in table actions available for this context. */
  actions: DocxTableContextMenuAction[];
  /** Runs a built-in table action. */
  runAction: (actionId: DocxTableContextMenuActionId) => void;
  /** Closes the menu without running an action. */
  closeMenu: () => void;
  /** Viewport coordinates where the menu should be placed. */
  position: {
    x: number;
    y: number;
  };
  /** Current document theme, useful for custom menu styling. */
  documentTheme: DocxDocumentTheme;
}

export interface DocxContextMenuContext {
  /** Type of target that opened the menu. */
  kind: "text" | "table" | "image";
  /** Active text range when the menu was opened from text. */
  activeTextRange?: DocxTextRange;
  /** Text location under the pointer when available. */
  location?: DocxTextRangeLocation;
  /** Table context when `kind` is `"table"`. */
  tableContext?: DocxTableContextMenuContext;
  /** Image context when `kind` is `"image"`. */
  image?:
    | {
        location: DocxImageLocation;
        floating?: NonNullable<ImageRunNode["floating"]>;
        wrap?: DocxImageWrapState;
      }
    | undefined;
}

export interface DocxContextMenuRenderProps {
  /** Target context that opened the menu. */
  context: DocxContextMenuContext;
  /** Built-in actions available for this context. */
  actions: DocxContextMenuAction[];
  /** Runs a built-in or custom action id. */
  runAction: (actionId: DocxContextMenuActionId | (string & {})) => void;
  /** Closes the menu without running an action. */
  closeMenu: () => void;
  /** Viewport coordinates where the menu should be placed. */
  position: {
    x: number;
    y: number;
  };
  /** Current document theme, useful for custom menu styling. */
  documentTheme: DocxDocumentTheme;
}

export interface DocxPageVirtualizationOptions {
  /**
   * Enables or disables internal page virtualization.
   *
   * Disable only when you need every page mounted at once, for example when
   * generating thumbnails for all pages from the live DOM.
   *
   * @defaultValue `true`
   */
  enabled?: boolean;
  /**
   * Number of pages to keep mounted before and after the visible viewport.
   *
   * Higher values reduce mount/unmount churn while scrolling but increase
   * initial DOM work.
   *
   * @defaultValue `2`
   */
  overscan?: number;
  /**
   * Compatibility option from earlier releases.
   *
   * Page virtualization now starts immediately, so this option is accepted but
   * does not delay the initial page window.
   *
   * @defaultValue `0`
   */
  settleDelayMs?: number;
  /**
   * Explicit scroll container for internal page virtualization.
   *
   * Omit this to let the viewer discover the nearest scrollable ancestor.
   * Provide it when the host app owns scrolling, such as when the viewer is
   * mounted inside a custom scroll-area viewport.
   */
  scrollElement?: HTMLElement | null;
  /**
   * Effective visual scale applied around the viewer.
   *
   * Omit this to let the viewer infer CSS `zoom` from its ancestor chain.
   * Provide it when toolbar zoom is controlled outside the viewer so virtual
   * page offsets update synchronously with the selected zoom.
   */
  zoomScale?: number;
}

export interface DocxVisiblePageRange {
  /**
   * Zero-based first page index to render.
   *
   * @example
   * ```tsx
   * <DocxEditorViewer
   *   editor={editor}
   *   visiblePageRange={{ startPageIndex: 0, endPageIndex: 2 }}
   * />
   * ```
   */
  startPageIndex: number;
  /** Zero-based last page index to render. */
  endPageIndex: number;
}

export interface DocxEditorViewerProps {
  /**
   * Editor controller returned by `useDocxEditor`.
   */
  editor: DocxEditorController;
  /**
   * CSS class applied to the outer viewer root.
   */
  className?: string;
  /**
   * Inline styles applied to the outer viewer root.
   */
  style?: Record<string, string | number | undefined>;
  /**
   * Background color of each rendered page surface.
   *
   * @defaultValue Viewer theme page color.
   */
  pageBackgroundColor?: string;
  /**
   * Background color shown between pages.
   *
   * @defaultValue `"transparent"`
   */
  pageGapBackgroundColor?: string;
  /**
   * Hides the document behind `loadingState` until initial pagination settles.
   *
   * Enable only when you prefer stable first paint over immediate approximate
   * layout.
   *
   * @defaultValue `false`
   */
  deferInitialPaginationPaint?: boolean;
  /**
   * Custom content shown while initial pagination is settling.
   *
   * Used only when `deferInitialPaginationPaint` is true.
   *
   * @defaultValue A compact `"Loading..."` pill.
   */
  loadingState?: unknown;
  /**
   * Configures internal page virtualization.
   *
   * @defaultValue `{ enabled: true, overscan: 2 }`
   */
  pageVirtualization?: DocxPageVirtualizationOptions;
  /**
   * Controlled page range to mount.
   *
   * When provided, internal page virtualization is bypassed and only this
   * range is rendered.
   */
  visiblePageRange?: DocxVisiblePageRange;
  /**
   * Called whenever the viewer's resolved page count changes.
   */
  onPageCountChange?: (pageCount: number) => void;
  /**
   * Called when the viewer needs a page to become visible, such as after a
   * bookmark or cross-reference navigation.
   *
   * If omitted, the internal virtualizer scrolls the page into view.
   */
  onRequestPageReveal?: (pageIndex: number) => void;
  /**
   * Overrides visual styles for heading levels.
   */
  headingStyles?: DocxHeadingStyleMap;
  /**
   * Interaction mode for the viewer.
   *
   * @defaultValue `"edit"`
   */
  mode?: DocxEditorViewerMode;
  /**
   * Overrides whether tracked changes are shown.
   *
   * If omitted, the value from `useDocxTrackChanges(editor)` or
   * `editor.showTrackedChanges` is used.
   */
  showTrackedChanges?: boolean;
  /**
   * Custom renderer for tracked-change cards in the page gutter.
   */
  renderTrackedChangeCard?: (
    props: DocxTrackedChangeCardRenderProps
  ) => unknown;
  /**
   * Overrides whether document comments are shown.
   *
   * If omitted, the value from `useDocxComments(editor)` or
   * `editor.showComments` is used.
   */
  showComments?: boolean;
  /**
   * Custom renderer for comment cards in the page gutter.
   */
  renderCommentCard?: (props: DocxCommentCardRenderProps) => unknown;
  /**
   * Custom renderer for table context menus.
   *
   * Call `props.runAction(action.id)` to execute built-in actions.
   */
  renderTableContextMenu?: (
    props: DocxTableContextMenuRenderProps
  ) => unknown;
  /**
   * Custom renderer for text, image, and table context menus.
   *
   * Call `props.closeMenu()` after handling a custom action.
   */
  renderContextMenu?: (props: DocxContextMenuRenderProps) => unknown;
  /**
   * Called when a form field is double-clicked.
   *
   * Use this to open a custom field settings panel.
   */
  onFormFieldDoubleClick?: (location: DocxFormFieldLocation) => void;
}

export interface DocxTrackedChangeCardRenderProps {
  /** Tracked-change data represented by the card. */
  change: DocxTrackedChange;
  /** Short display label for the change kind. */
  kindLabel: string;
  /** Plain-text excerpt for the changed content. */
  snippet: string;
  /** Formatted change date, if the source document provided one. */
  formattedDate?: string;
  /** Accent color chosen for this change kind. */
  accentColor: string;
  /** Current document theme. */
  documentTheme: DocxDocumentTheme;
  /** Zero-based page index that owns the card. */
  pageIndex: number;
  /** Positioning style computed by the viewer. Apply this to the card root. */
  style: Record<string, string | number | undefined>;
}

export interface DocxCommentCardRenderProps {
  /** Comment data represented by the card. */
  comment: DocxComment;
  /** Plain-text comment body (already normalized for display). */
  snippet: string;
  /** Formatted comment date, if the source document provided one. */
  formattedDate?: string;
  /** Accent color chosen for comments. */
  accentColor: string;
  /** Current document theme. */
  documentTheme: DocxDocumentTheme;
  /** Zero-based page index that owns the card. */
  pageIndex: number;
  /** Positioning style computed by the viewer. Apply this to the card root. */
  style: Record<string, string | number | undefined>;
}

export type DocxEditorViewerMode = "edit" | "read-only";

export interface UseDocxDocumentThemeResult {
  documentTheme: DocxDocumentTheme;
  isDarkDocument: boolean;
  setDocumentTheme: (theme: DocxDocumentTheme) => void;
  toggleDocumentTheme: () => void;
}

export interface UseDocxParagraphStylesResult {
  paragraphStyles: ParagraphStyleDefinition[];
  selectedParagraphStyleId?: string;
  setParagraphStyle: (styleId?: string) => void;
}

export interface UseDocxLineSpacingResult {
  lineSpacing: DocxLineSpacingInfo;
  setLineSpacing: (lineMultiple: number) => void;
}

export interface UseDocxBordersResult {
  borderContext: DocxBorderContext;
  activeBorderPresets: DocxBorderPresetState;
  applyBorderPreset: (preset: DocxBorderPreset) => void;
}

export interface UseDocxFormFieldsResult {
  formFields: DocxSelectedFormField[];
  selectedFormField?: DocxSelectedFormField;
  selectFormField: (location?: DocxFormFieldLocation) => void;
  setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void;
  toggleFormCheckbox: (location: DocxFormFieldLocation) => void;
  updateFormFieldWidget: (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void;
  updateSelectedFormFieldWidget: (
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void;
}

export interface UseDocxTrackChangesResult {
  trackedChanges: DocxTrackedChange[];
  showTrackedChanges: boolean;
  setShowTrackedChanges: (showTrackedChanges: boolean) => void;
  toggleShowTrackedChanges: () => void;
  changesByLocation: Map<string, DocxTrackedChange[]>;
  getChangesForLocation: (
    location: DocxTextRangeLocation
  ) => DocxTrackedChange[];
}

export interface UseDocxCommentsResult {
  comments: DocxComment[];
  showComments: boolean;
  setShowComments: (showComments: boolean) => void;
  toggleShowComments: () => void;
  commentsByLocation: Map<string, DocxComment[]>;
  getCommentsForLocation: (location: DocxTextRangeLocation) => DocxComment[];
}

export interface DocxSectionColumnLayout {
  count: number;
  gapPx: number;
}

export interface DocxPageLayoutInfo {
  pageWidthPx: number;
  pageHeightPx: number;
  contentWidthPx: number;
  contentHeightPx: number;
  marginsPx: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  headerDistancePx: number;
  footerDistancePx: number;
  pageNumberStart: number;
  columns?: DocxSectionColumnLayout;
  viewportDefaults: {
    zoomPercent: number;
    pageGapPx: number;
  };
}

export interface UseDocxPageLayoutResult {
  layout: DocxPageLayoutInfo;
}

export interface DocxPaginationInfo {
  currentPage: number;
  totalPages: number;
}

export interface UseDocxPaginationResult {
  pagination: DocxPaginationInfo;
}

export interface DocxPageThumbnailResolutionOptions {
  /** Source page width in CSS pixels. */
  sourceWidthPx: number;
  /** Source page height in CSS pixels. */
  sourceHeightPx: number;
  /**
   * Desired thumbnail bounds.
   *
   * A number means "fit within this max width and max height". An object lets
   * you constrain width and height independently.
   *
   * @defaultValue `160`
   */
  resolution?: DocxPageThumbnailBounds;
  /**
   * Maximum CSS pixel width for the thumbnail.
   *
   * Overrides `resolution.maxWidth` when provided.
   */
  maxWidthPx?: number;
  /**
   * Maximum CSS pixel height for the thumbnail.
   *
   * Overrides `resolution.maxHeight` when provided.
   */
  maxHeightPx?: number;
  /**
   * Canvas backing-store pixel ratio.
   *
   * Increase for sharper thumbnails on high-density displays.
   *
   * @defaultValue `window.devicePixelRatio`, capped internally.
   */
  pixelRatio?: number;
}

export type DocxPageThumbnailBounds =
  | number
  | {
      /** Maximum thumbnail CSS height. */
      maxHeight?: number;
      /** Maximum thumbnail CSS width. */
      maxWidth?: number;
    };

export interface DocxPageThumbnailResolution {
  /** Thumbnail CSS width. */
  widthPx: number;
  /** Thumbnail CSS height. */
  heightPx: number;
  /** Canvas backing-store width. */
  pixelWidthPx: number;
  /** Canvas backing-store height. */
  pixelHeightPx: number;
  /** Scale from source page pixels to thumbnail CSS pixels. */
  scale: number;
}

export interface DocxPageThumbnailRenderWindow {
  /**
   * Page indexes whose attached thumbnail canvases should render first.
   *
   * Use this for the thumbnails currently visible in a virtualized sidebar.
   */
  visiblePageIndexes?: readonly number[];
  /**
   * Page indexes to rasterize into the thumbnail surface cache after visible
   * thumbnails. Prefetched pages paint quickly once their canvases mount.
   */
  prefetchPageIndexes?: readonly number[];
}

export interface UseDocxPageThumbnailsOptions {
  /**
   * Desired thumbnail bounds.
   *
   * @defaultValue `160`
   */
  resolution?: DocxPageThumbnailBounds;
  /** Maximum thumbnail CSS width. */
  maxWidthPx?: number;
  /** Maximum thumbnail CSS height. */
  maxHeightPx?: number;
  /**
   * Canvas backing-store pixel ratio.
   *
   * @defaultValue `window.devicePixelRatio`, capped internally.
   */
  pixelRatio?: number;
  /**
   * Minimum interval between repeat raster jobs for the same thumbnail canvas.
   *
   * Lower this when the consumer already limits work to a small visible
   * thumbnail window.
   *
   * @defaultValue `200`
   */
  minRasterIntervalMs?: number;
  /**
   * Prioritizes thumbnails for consumer-owned virtualized thumbnail rails.
   */
  renderWindow?: DocxPageThumbnailRenderWindow;
  /**
   * Prevents thumbnail rendering while keeping stable item metadata.
   *
   * @defaultValue `false`
   */
  disabled?: boolean;
}

export type DocxPageThumbnailStatus =
  | "idle"
  | "rendering"
  | "ready"
  | "unavailable"
  | "error";

export interface DocxPageThumbnailItem extends DocxPageThumbnailResolution {
  /** Source page aspect ratio. */
  aspectRatio: number;
  /** Alias for `sourceHeightPx` for compatibility with other preview APIs. */
  contentHeight: number;
  /** Alias for `sourceWidthPx` for compatibility with other preview APIs. */
  contentWidth: number;
  /** Alias for `heightPx`. */
  height: number;
  /** Zero-based page index. */
  pageIndex: number;
  /** One-based page number for display. */
  pageNumber: number;
  /** Source page width in CSS pixels. */
  sourceWidthPx: number;
  /** Source page height in CSS pixels. */
  sourceHeightPx: number;
  /** True when the source page DOM is currently mounted. */
  isMounted: boolean;
  /** Current thumbnail render status. */
  status: DocxPageThumbnailStatus;
  /** Last thumbnail rendering error, if any. */
  error?: Error;
  /** Paints this thumbnail into a canvas. Returns false when unavailable. */
  paint: (canvas: HTMLCanvasElement | null) => boolean;
  /** Ref callback that keeps an attached canvas rendered as the page changes. */
  canvasRef: (canvas: HTMLCanvasElement | null) => void;
  /** Asynchronously renders this thumbnail into a canvas. */
  renderToCanvas: (canvas: HTMLCanvasElement) => Promise<void>;
  /** Alias for `widthPx`. */
  width: number;
}

export interface UseDocxPageThumbnailsResult {
  /** Paints the requested page thumbnail into a canvas. */
  paintThumbnail: (
    pageIndex: number,
    canvas: HTMLCanvasElement | null
  ) => boolean;
  /** Thumbnail metadata and paint helpers for each known page. */
  thumbnails: DocxPageThumbnailItem[];
  /** Re-renders every thumbnail canvas currently attached through `canvasRef`. */
  rerenderAttachedThumbnails: () => Promise<void>;
}

export type UseDocxViewerThumbnailsOptions = UseDocxPageThumbnailsOptions;

export type DocxViewerThumbnails = UseDocxPageThumbnailsResult;

