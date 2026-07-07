import {
  ref,
  computed,
  watch,
  type Ref,
  type ComputedRef,
} from "vue";
import {
  cloneDocModel,
  type DocModel,
  type ParagraphNode,
  type TableNode,
  type TextRunNode,
  type HeadingLevel,
  type ParagraphAlignment,
  type ParagraphStyleDefinition,
  type TableCellContentNode,
} from "@extend-ai/docx-core";
import {
  updateParagraphText,
  toggleRunStyleFlag,
} from "@extend-ai/docx-core";
import { serializeDocx } from "@extend-ai/docx-core";
import { importDocxBuffer } from "@extend-ai/docx-core";
import type { DocxImportResult, DocxImportOptions } from "@extend-ai/docx-core";
import type { OoxmlPackage } from "@extend-ai/docx-core";
import {
  buildDocumentPageNodeSegments,
  type PageSegmentationCallbacks,
} from "@extend-ai/docx-core";
import {
  defaultStarterModel,
  type DocxEditorSelection,
  type DocxTextRange,
  type DocxTextRangeLocation,
  type DocxTextRangeBoundary,
  type DocxDocumentTheme,
  type DocxSelectionSessionKind,
  type DocxFormFieldLocation,
  type DocxSelectedFormField,
  type DocxPaginationInfo,
  type DocxLineSpacingInfo,
  type DocxBorderContext,
  type DocxBorderPreset,
  type DocxBorderPresetState,
  type DocxTrackedChange,
  type DocxComment,
  type UseDocxEditorOptions,
  type DocxEditorController,
} from "@extend-ai/docx-core";

// ---------------------------------------------------------------------------
// Internal types (not exported by docx-core)
// ---------------------------------------------------------------------------

interface DocxHistoryRestoreRequest {
  nonce: number;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
}

// ---------------------------------------------------------------------------
// Helpers (ported from upstream editor.tsx, zero React dependency)
// ---------------------------------------------------------------------------

function textRuns(paragraph: ParagraphNode): TextRunNode[] {
  return paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text",
  );
}

function paragraphText(paragraph: ParagraphNode): string {
  return textRuns(paragraph)
    .map((run) => run.text)
    .join("");
}

function cloneTextStyle(
  style?: TextRunNode["style"],
): TextRunNode["style"] | undefined {
  return style ? { ...style } : undefined;
}

function cloneTextRangeLocation(
  loc: DocxTextRangeLocation,
): DocxTextRangeLocation {
  if (loc.kind === "paragraph") {
    return { kind: "paragraph", nodeIndex: loc.nodeIndex };
  }
  return {
    kind: "table-cell",
    tableIndex: loc.tableIndex,
    rowIndex: loc.rowIndex,
    cellIndex: loc.cellIndex,
    paragraphIndex: loc.paragraphIndex,
  };
}

function cloneTextRange(
  range?: DocxTextRange,
): DocxTextRange | undefined {
  if (!range) return undefined;
  return {
    start: {
      location: cloneTextRangeLocation(range.start.location),
      offset: range.start.offset,
    },
    end: {
      location: cloneTextRangeLocation(range.end.location),
      offset: range.end.offset,
    },
  };
}

function cloneEditorSelection(
  selection: DocxEditorSelection,
): DocxEditorSelection {
  if (selection.kind === "paragraph") {
    return { kind: "paragraph", nodeIndex: selection.nodeIndex };
  }
  return {
    kind: "table-cell",
    tableIndex: selection.tableIndex,
    rowIndex: selection.rowIndex,
    cellIndex: selection.cellIndex,
  };
}

function sameEditorSelection(
  a: DocxEditorSelection,
  b: DocxEditorSelection,
): boolean {
  if (a.kind !== b.kind) return false;
  if (a.kind === "paragraph" && b.kind === "paragraph") {
    return a.nodeIndex === b.nodeIndex;
  }
  if (a.kind === "table-cell" && b.kind === "table-cell") {
    return (
      a.tableIndex === b.tableIndex &&
      a.rowIndex === b.rowIndex &&
      a.cellIndex === b.cellIndex
    );
  }
  return false;
}

function locationKey(loc: DocxTextRangeLocation): string {
  if (loc.kind === "paragraph") {
    return `p:${loc.nodeIndex}`;
  }
  return `tc:${loc.tableIndex}:${loc.rowIndex}:${loc.cellIndex}:${loc.paragraphIndex}`;
}

function compareTextRangeBoundaries(
  a: DocxTextRangeBoundary,
  b: DocxTextRangeBoundary,
): number {
  const keyCmp = locationKey(a.location).localeCompare(locationKey(b.location));
  if (keyCmp !== 0) return keyCmp;
  return a.offset - b.offset;
}

function sameTextRange(
  a?: DocxTextRange,
  b?: DocxTextRange,
): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return (
    compareTextRangeBoundaries(a.start, b.start) === 0 &&
    compareTextRangeBoundaries(a.end, b.end) === 0
  );
}

function normalizeTextRange(range: DocxTextRange): DocxTextRange {
  const cmp = compareTextRangeBoundaries(range.start, range.end);
  if (cmp <= 0) return range;
  return {
    start: { location: cloneTextRangeLocation(range.end.location), offset: range.end.offset },
    end: { location: cloneTextRangeLocation(range.start.location), offset: range.start.offset },
  };
}

function tableCellParagraphs(nodes: TableCellContentNode[]): ParagraphNode[] {
  return nodes.filter(
    (n): n is ParagraphNode => n.type === "paragraph",
  );
}

function getParagraphAtLocation(
  model: DocModel,
  location: DocxTextRangeLocation,
): { paragraph?: ParagraphNode; tableNode?: TableNode } {
  if (location.kind === "paragraph") {
    const node = model.nodes[location.nodeIndex];
    if (!node || node.type !== "paragraph") return {};
    return { paragraph: node as ParagraphNode };
  }

  const tableNode = model.nodes[location.tableIndex];
  if (!tableNode || tableNode.type !== "table") return {};

  const cell = (tableNode as TableNode).rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) return {};

  const paragraph = tableCellParagraphs(cell.nodes)[location.paragraphIndex];
  if (!paragraph) return {};

  return { paragraph, tableNode: tableNode as TableNode };
}

// ---------------------------------------------------------------------------
// Editor transaction types (internal)
// ---------------------------------------------------------------------------

interface DocxHistorySnapshot {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
}

interface DocxEditorTransactionContext {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingRunStyle?: TextRunNode["style"];
}

interface DocxEditorTransactionPatch {
  model?: DocModel;
  selection?: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingRunStyle?: TextRunNode["style"];
  status?: string;
  clearSelectedFormField?: boolean;
  pushHistory?: boolean;
}

// ---------------------------------------------------------------------------
// Default pagination callbacks (minimal)
// ---------------------------------------------------------------------------

function makeDefaultPaginationCallbacks(): PageSegmentationCallbacks {
  return {
    estimateDocNodeHeightPx: (_node, _width) => 22,
    paragraphHasVisibleText: (p) => paragraphText(p).length > 0,
    paragraphIsStructuralSectionBreakSpacer: () => false,
    estimateParagraphHeightPx: (_p, _width) => 22,
    estimateParagraphLineHeightPx: () => 22,
    paragraphLineCountWithinWidth: (_p, _width) =>
      Math.max(1, Math.ceil(paragraphText(_p).length / 60)),
    paragraphWidowControlEnabled: () => true,
    paragraphCanSplitAcrossPages: () => true,
    estimateTableRowHeightsPx: (_tableNode, _availableWidthPx) => [22],
  };
}

// ---------------------------------------------------------------------------
// Style toggle helper: applies a style flag to all runs within a text range
// ---------------------------------------------------------------------------

function toggleRunStyleInRange(
  model: DocModel,
  nodeIndex: number,
  startOffset: number,
  endOffset: number,
  key: "bold" | "italic" | "underline" | "strike",
): DocModel {
  const paragraph = model.nodes[nodeIndex];
  if (!paragraph || paragraph.type !== "paragraph") return model;

  // Count text runs and their offsets to find which ones overlap
  const runs = textRuns(paragraph);
  let cursor = 0;
  let runIdx = 0;
  let next = model;

  for (const run of runs) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;

    // Skip runs that don't overlap with the selection range
    if (runEnd <= startOffset || runStart >= endOffset) {
      runIdx++;
      continue;
    }

    // Split the run if selection is partial within this run
    const localStart = Math.max(0, startOffset - runStart);
    const localEnd = Math.min(run.text.length, endOffset - runStart);

    if (localStart > 0 || localEnd < run.text.length) {
      // Need to split this run to isolate the selected portion
      next = cloneDocModel(next);
      const nextParagraph = next.nodes[nodeIndex] as ParagraphNode;
      const nextRuns = textRuns(nextParagraph);

      // Find the updated run index after previous splits
      const currentRun = nextRuns[runIdx];
      if (!currentRun) { runIdx++; continue; }

      const originalText = currentRun.text;
      const beforeText = originalText.slice(0, localStart);
      const selectedText = originalText.slice(localStart, localEnd);
      const afterText = originalText.slice(localEnd);

      // Replace the single run with up to 3 runs
      const newChildren: ParagraphNode["children"] = [];
      for (const child of nextParagraph.children) {
        if (child.type === "text" && child === currentRun) {
          if (beforeText.length > 0) {
            newChildren.push({
              type: "text" as const,
              text: beforeText,
              style: cloneTextStyle(child.style),
              link: child.link,
            });
          }
          newChildren.push({
            type: "text" as const,
            text: selectedText,
            style: cloneTextStyle(child.style),
            link: child.link,
          });
          if (afterText.length > 0) {
            newChildren.push({
              type: "text" as const,
              text: afterText,
              style: cloneTextStyle(child.style),
              link: child.link,
            });
          }
        } else {
          newChildren.push(child);
        }
      }
      nextParagraph.children = newChildren;

      // Now toggle the isolated run
      const isolatedRunIdx = beforeText.length > 0 ? runIdx + 1 : runIdx;
      next = toggleRunStyleFlag(next, nodeIndex, isolatedRunIdx, key);
      nextParagraph.sourceXml = undefined;

      // Skip the rest of this run's contribution
      runIdx += (beforeText.length > 0 ? 2 : 1) + (afterText.length > 0 ? 1 : 0);
    } else {
      // Run is fully within the selection range
      next = toggleRunStyleFlag(next, nodeIndex, runIdx, key);
      (next.nodes[nodeIndex] as ParagraphNode).sourceXml = undefined;
      runIdx++;
    }
  }

  return next;
}

// ---------------------------------------------------------------------------
// useDocxEditor
// ---------------------------------------------------------------------------

export function useDocxEditor(
  options: UseDocxEditorOptions = {},
): DocxEditorController {
  const starterTemplateRef = cloneDocModel(
    options.starterModel ?? defaultStarterModel,
  );

  // ---- Reactive state ----
  const model = ref<DocModel>(cloneDocModel(starterTemplateRef));
  const basePackage = ref<OoxmlPackage | undefined>();
  const documentLoadNonce = ref(0);
  const fileName = ref(options.initialFileName ?? "(new document)");
  const selection = ref<DocxEditorSelection>({
    kind: "paragraph",
    nodeIndex: 0,
  });
  const selectedFormFieldLocation = ref<DocxFormFieldLocation | undefined>();
  const status = ref(options.initialStatus ?? "Ready");
  const importError = ref<Error | undefined>();
  const isImporting = ref(false);
  const documentTheme = ref<DocxDocumentTheme>(
    options.initialDocumentTheme ?? "light",
  );
  const showTrackedChanges = ref(options.initialShowTrackedChanges ?? false);
  const showComments = ref(options.initialShowComments ?? false);
  const paginationInfo = ref<DocxPaginationInfo>({
    currentPage: 1,
    totalPages: 1,
  });
  const activeImportAbortController = ref<AbortController | undefined>();
  const history = ref<{ past: DocxHistorySnapshot[]; future: DocxHistorySnapshot[] }>({
    past: [],
    future: [],
  });
  const activeTextRange = ref<DocxTextRange | undefined>();
  const pendingRunStyle = ref<TextRunNode["style"] | undefined>();
  const historyRestoreRequest = ref<DocxHistoryRestoreRequest | undefined>();
  const selectionSessionKind = ref<DocxSelectionSessionKind>("idle");

  // Mirrors for synchronous access inside closures
  const modelRef = { current: model.value };
  const selectionRef = { current: selection.value };
  const activeTextRangeRef = { current: activeTextRange.value };
  const pendingRunStyleRef = { current: pendingRunStyle.value };

  watch(model, (v) => { modelRef.current = v; });
  watch(selection, (v) => { selectionRef.current = v; });
  watch(activeTextRange, (v) => { activeTextRangeRef.current = v; });
  watch(pendingRunStyle, (v) => { pendingRunStyleRef.current = v; });

  // ---- Computed properties ----
  const selectedParagraph = computed<ParagraphNode | undefined>(() => {
    const sel = selection.value;
    if (sel.kind !== "paragraph") return undefined;
    const node = model.value.nodes[sel.nodeIndex];
    return node?.type === "paragraph" ? (node as ParagraphNode) : undefined;
  });

  const selectedFormField = computed<DocxSelectedFormField | undefined>(() => {
    const loc = selectedFormFieldLocation.value;
    if (!loc) return undefined;
    const { paragraph, tableNode } = getParagraphAtLocation(model.value, loc);
    if (!paragraph) return undefined;
    const child = paragraph.children[loc.childIndex];
    if (!child || child.type !== "form-field") return undefined;
    return {
      location: {
        ...loc,
        childIndex: loc.childIndex,
      } as DocxFormFieldLocation,
      field: child,
    };
  });

  const selectedRunStyle = computed<TextRunNode["style"] | undefined>(() => {
    const range = activeTextRange.value ?? undefined;
    if (!range) return undefined;
    const { paragraph } = getParagraphAtLocation(model.value, range.start.location);
    if (!paragraph) return undefined;
    const runs = textRuns(paragraph);
    for (const run of runs) {
      if (run.style) return run.style;
    }
    return undefined;
  });

  const selectedLink = computed<string | undefined>(() => {
    const range = activeTextRange.value ?? undefined;
    if (!range) return undefined;
    const { paragraph } = getParagraphAtLocation(model.value, range.start.location);
    if (!paragraph) return undefined;
    const runs = textRuns(paragraph);
    for (const run of runs) {
      if (run.link) return run.link;
    }
    return undefined;
  });

  const canUndo = computed(() => history.value.past.length > 0);
  const canRedo = computed(() => history.value.future.length > 0);

  // Stub computed properties for remaining DocxEditorController fields
  const selectedParagraphStyleId = computed<string | undefined>(() => undefined);
  const selectedLineSpacing = computed<DocxLineSpacingInfo>(() => ({
    lineRule: "auto" as const,
    multiple: 1,
  }));
  const selectedBorderContext = computed<DocxBorderContext>(() => "paragraph");
  const activeBorderPresets = computed<DocxBorderPresetState>(() => {
    const presets: DocxBorderPreset[] = [
      "none", "all", "outside", "inside", "top", "bottom", "left", "right",
      "inside-horizontal", "inside-vertical", "diagonal-down", "diagonal-up", "horizontal-line",
    ];
    const state = {} as DocxBorderPresetState;
    for (const p of presets) {
      state[p] = false;
    }
    return state;
  });
  const availableParagraphStyles = computed<ParagraphStyleDefinition[]>(
    () => model.value.metadata.paragraphStyles ?? [],
  );
  const trackedChanges = computed<DocxTrackedChange[]>(() => []);
  const comments = computed<DocxComment[]>(() => []);
  const hasUnorderedList = computed(() => false);
  const hasOrderedList = computed(() => false);
  const currentPage = computed(() => paginationInfo.value.currentPage);
  const totalPages = computed(() => paginationInfo.value.totalPages);

  // ---- History helpers ----
  function pushHistorySnapshot(): void {
    history.value = {
      past: [
        ...history.value.past,
        {
          model: cloneDocModel(model.value),
          selection: cloneEditorSelection(selection.value),
          activeTextRange: cloneTextRange(activeTextRange.value),
        },
      ],
      future: [],
    };
  }

  // ---- Core transaction dispatch (D4: cloneDocModel + editor-ops) ----
  function dispatchEditorTransaction(
    resolver: (
      current: DocxEditorTransactionContext,
    ) => DocxEditorTransactionPatch | undefined,
  ): boolean {
    const currentModel = modelRef.current;
    const currentSelection = cloneEditorSelection(selectionRef.current);
    const currentRange = cloneTextRange(activeTextRangeRef.current);
    const currentPendingRunStyle = cloneTextStyle(pendingRunStyleRef.current);

    const patch = resolver({
      model: currentModel,
      selection: currentSelection,
      activeTextRange: currentRange,
      pendingRunStyle: currentPendingRunStyle,
    });
    if (!patch) return false;

    const nextModel = patch.model ?? currentModel;
    const nextSelection = patch.selection ?? currentSelection;
    const nextRange =
      "activeTextRange" in patch ? patch.activeTextRange : currentRange;
    const nextPendingRunStyle =
      "pendingRunStyle" in patch
        ? cloneTextStyle(patch.pendingRunStyle)
        : currentPendingRunStyle;

    const modelChanged = nextModel !== currentModel;

    if (modelChanged && patch.pushHistory !== false) {
      pushHistorySnapshot();
    }

    model.value = nextModel;
    selection.value = nextSelection;
    if ("activeTextRange" in patch) {
      activeTextRange.value = nextRange;
    }
    if ("pendingRunStyle" in patch) {
      pendingRunStyle.value = nextPendingRunStyle;
    }
    if (patch.status) {
      status.value = patch.status;
    }
    if (patch.clearSelectedFormField) {
      selectedFormFieldLocation.value = undefined;
    }

    return true;
  }

  // ---- applyModelChange (D4: calls cloneDocModel) ----
  function applyModelChange(
    mutator: (current: DocModel) => DocModel | undefined,
  ): void {
    dispatchEditorTransaction((ctx) => {
      const next = mutator(ctx.model);
      if (!next || next === ctx.model) return undefined;
      return { model: next };
    });
  }

  // ---- D1 + D2: importDocxFile (wasm + worker) ----
  async function importDocxFile(file: File): Promise<void> {
    activeImportAbortController.value?.abort();
    activeImportAbortController.value = undefined;

    if (!/\.docx?$/i.test(file.name)) {
      importError.value = new Error("Only .docx and .doc files are supported");
      return;
    }

    isImporting.value = true;
    importError.value = undefined;
    status.value = `Loading ${file.name}...`;
    const importAbortController = new AbortController();
    activeImportAbortController.value = importAbortController;

    try {
      const buffer = await file.arrayBuffer();
      const importResult: DocxImportResult = await importDocxBuffer(buffer, {
        signal: importAbortController.signal,
        transferBuffer: true,
      });

      const pkg = importResult.package;
      const nextModel = importResult.model;

      model.value = nextModel;
      documentLoadNonce.value = documentLoadNonce.value + 1;
      history.value = { past: [], future: [] };
      historyRestoreRequest.value = undefined;
      basePackage.value = pkg;
      fileName.value = file.name;
      selection.value = { kind: "paragraph", nodeIndex: 0 };
      activeTextRange.value = undefined;
      pendingRunStyle.value = undefined;
      selectedFormFieldLocation.value = undefined;
      importError.value = undefined;
      status.value = `Loaded ${file.name}`;
    } catch (error) {
      if (
        error instanceof Error &&
        error.name === "AbortError" &&
        activeImportAbortController.value !== importAbortController
      ) {
        return;
      }
      importError.value =
        error instanceof Error ? error : new Error("Unknown error");
    } finally {
      if (activeImportAbortController.value === importAbortController) {
        activeImportAbortController.value = undefined;
        isImporting.value = false;
      }
    }
  }

  // ---- D1 + D5: exportDocx (wasmSerializeDocx + serializeDocx + basePackage) ----
  async function exportDocx(): Promise<void> {
    const sourceModel = modelRef.current;
    const output = await serializeDocx(sourceModel, basePackage.value);
    const blob = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = /\.docx?$/i.test(fileName.value)
      ? fileName.value.replace(/\.docx?$/i, "") + "-edited.docx"
      : "edited.docx";
    anchor.style.display = "none";
    document.body.append(anchor);
    anchor.click();
    window.setTimeout(() => {
      anchor.remove();
      URL.revokeObjectURL(url);
    }, 1000);
    status.value = "Exported DOCX";
  }

  // ---- newDocument ----
  function newDocument(): void {
    activeImportAbortController.value?.abort();
    activeImportAbortController.value = undefined;
    model.value = cloneDocModel(starterTemplateRef);
    documentLoadNonce.value = documentLoadNonce.value + 1;
    basePackage.value = undefined;
    fileName.value = "(new document)";
    selection.value = { kind: "paragraph", nodeIndex: 0 };
    activeTextRange.value = undefined;
    pendingRunStyle.value = undefined;
    selectedFormFieldLocation.value = undefined;
    history.value = { past: [], future: [] };
    historyRestoreRequest.value = undefined;
    importError.value = undefined;
    status.value = "Created new document";
  }

  // ---- undo / redo ----
  function undo(): void {
    const currentHistory = history.value;
    const snapshot = currentHistory.past[currentHistory.past.length - 1];
    if (!snapshot) return;

    history.value = {
      past: currentHistory.past.slice(0, -1),
      future: [
        {
          model: cloneDocModel(model.value),
          selection: cloneEditorSelection(selection.value),
          activeTextRange: cloneTextRange(activeTextRange.value),
        },
        ...currentHistory.future,
      ],
    };

    const nonce = (historyRestoreRequest.value?.nonce ?? 0) + 1;
    historyRestoreRequest.value = {
      nonce,
      selection: cloneEditorSelection(snapshot.selection),
      activeTextRange: cloneTextRange(snapshot.activeTextRange),
    };
    model.value = cloneDocModel(snapshot.model);
    selection.value = cloneEditorSelection(snapshot.selection);
    activeTextRange.value = cloneTextRange(snapshot.activeTextRange);
  }

  function redo(): void {
    const currentHistory = history.value;
    const snapshot = currentHistory.future[0];
    if (!snapshot) return;

    history.value = {
      past: [
        ...currentHistory.past,
        {
          model: cloneDocModel(model.value),
          selection: cloneEditorSelection(selection.value),
          activeTextRange: cloneTextRange(activeTextRange.value),
        },
      ],
      future: currentHistory.future.slice(1),
    };

    const nonce = (historyRestoreRequest.value?.nonce ?? 0) + 1;
    historyRestoreRequest.value = {
      nonce,
      selection: cloneEditorSelection(snapshot.selection),
      activeTextRange: cloneTextRange(snapshot.activeTextRange),
    };
    model.value = cloneDocModel(snapshot.model);
    selection.value = cloneEditorSelection(snapshot.selection);
    activeTextRange.value = cloneTextRange(snapshot.activeTextRange);
  }

  // ---- D4: commitParagraphText (updateParagraphText) ----
  function commitParagraphText(nodeIndex: number, text: string): void {
    applyModelChange((current) => {
      return updateParagraphText(current, nodeIndex, text, {
        insertedStyle: cloneTextStyle(pendingRunStyleRef.current),
      });
    });
  }

  // ---- D4: toggleBold (cloneDocModel + toggleRunStyleFlag) ----
  function toggleBold(): void {
    const range = activeTextRange.value;
    if (range && range.start.location.kind === "paragraph") {
      const nodeIndex = range.start.location.nodeIndex;
      const startOffset = Math.min(range.start.offset, range.end.offset);
      const endOffset = Math.max(range.start.offset, range.end.offset);
      if (startOffset === endOffset) {
        pendingRunStyle.value = {
          ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
          bold: !Boolean(selectedRunStyle.value?.bold),
        };
        return;
      }
      applyModelChange((current) =>
        toggleRunStyleInRange(current, nodeIndex, startOffset, endOffset, "bold"),
      );
      return;
    }
    pendingRunStyle.value = {
      ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
      bold: !Boolean(selectedRunStyle.value?.bold),
    };
  }

  function toggleItalic(): void {
    const range = activeTextRange.value;
    if (range && range.start.location.kind === "paragraph") {
      const nodeIndex = range.start.location.nodeIndex;
      const startOffset = Math.min(range.start.offset, range.end.offset);
      const endOffset = Math.max(range.start.offset, range.end.offset);
      if (startOffset === endOffset) {
        pendingRunStyle.value = {
          ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
          italic: !Boolean(selectedRunStyle.value?.italic),
        };
        return;
      }
      applyModelChange((current) =>
        toggleRunStyleInRange(current, nodeIndex, startOffset, endOffset, "italic"),
      );
      return;
    }
    pendingRunStyle.value = {
      ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
      italic: !Boolean(selectedRunStyle.value?.italic),
    };
  }

  function toggleUnderline(): void {
    const range = activeTextRange.value;
    if (range && range.start.location.kind === "paragraph") {
      const nodeIndex = range.start.location.nodeIndex;
      const startOffset = Math.min(range.start.offset, range.end.offset);
      const endOffset = Math.max(range.start.offset, range.end.offset);
      if (startOffset === endOffset) {
        pendingRunStyle.value = {
          ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
          underline: !Boolean(selectedRunStyle.value?.underline),
        };
        return;
      }
      applyModelChange((current) =>
        toggleRunStyleInRange(current, nodeIndex, startOffset, endOffset, "underline"),
      );
      return;
    }
    pendingRunStyle.value = {
      ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
      underline: !Boolean(selectedRunStyle.value?.underline),
    };
  }

  function toggleStrike(): void {
    const range = activeTextRange.value;
    if (range && range.start.location.kind === "paragraph") {
      const nodeIndex = range.start.location.nodeIndex;
      const startOffset = Math.min(range.start.offset, range.end.offset);
      const endOffset = Math.max(range.start.offset, range.end.offset);
      if (startOffset === endOffset) {
        pendingRunStyle.value = {
          ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
          strike: !Boolean(selectedRunStyle.value?.strike),
        };
        return;
      }
      applyModelChange((current) =>
        toggleRunStyleInRange(current, nodeIndex, startOffset, endOffset, "strike"),
      );
      return;
    }
    pendingRunStyle.value = {
      ...(cloneTextStyle(pendingRunStyle.value) ?? {}),
      strike: !Boolean(selectedRunStyle.value?.strike),
    };
  }

  function setHeading(heading?: HeadingLevel): void {
    status.value = heading ? `Set heading ${heading}` : "Clear heading";
  }

  function setParagraphStyle(styleId?: string): void {
    status.value = styleId ? `Style: ${styleId}` : "Clear style";
  }

  function setLineSpacing(lineMultiple: number): void {
    status.value = `Line spacing: ${lineMultiple}`;
  }

  function setFontFamily(fontFamily: string): void {
    status.value = `Font: ${fontFamily}`;
  }

  function setFontSize(fontSizePt: number): void {
    status.value = `Font size: ${fontSizePt}pt`;
  }

  function setActiveTextRange(range?: DocxTextRange): void {
    if (!range) {
      activeTextRange.value = undefined;
      pendingRunStyle.value = undefined;
      return;
    }
    const normalized = normalizeTextRange(range);
    activeTextRange.value = cloneTextRange(normalized);
    pendingRunStyle.value = undefined;
  }

  // ---- D3: Pagination (buildDocumentPageNodeSegments) ----
  function computePageSegments(
    pageContentHeightPx: number,
    pageContentWidthPx: number,
  ) {
    const callbacks = makeDefaultPaginationCallbacks();
    return buildDocumentPageNodeSegments(
      model.value,
      pageContentHeightPx,
      pageContentWidthPx,
      callbacks,
      model.value.metadata.numberingDefinitions,
      [],
      {
        allowParagraphLineSplitting: true,
      },
    );
  }

  // ---- Controller object ----
  const controller: DocxEditorController = {
    // State
    get model() { return model.value; },
    get documentLoadNonce() { return documentLoadNonce.value; },
    get fileName() { return fileName.value; },
    get status() { return status.value; },
    get importError() { return importError.value; },
    get isImporting() { return isImporting.value; },
    get documentTheme() { return documentTheme.value; },
    get selection() { return selection.value; },
    get activeTextRange() { return activeTextRange.value; },
    get historyRestoreRequest() { return historyRestoreRequest.value; },
    get selectedFormField() { return selectedFormField.value; },
    get selectedParagraph() { return selectedParagraph.value; },
    get selectedRunStyle() { return selectedRunStyle.value; },
    get selectedLink() { return selectedLink.value; },
    get pendingRunStyle() { return pendingRunStyle.value; },
    get selectionSessionKind() { return selectionSessionKind.value; },
    get selectedParagraphStyleId() { return selectedParagraphStyleId.value; },
    get selectedLineSpacing() { return selectedLineSpacing.value; },
    get selectedBorderContext() { return selectedBorderContext.value; },
    get activeBorderPresets() { return activeBorderPresets.value; },
    get availableParagraphStyles() { return availableParagraphStyles.value; },
    get trackedChanges() { return trackedChanges.value; },
    get showTrackedChanges() { return showTrackedChanges.value; },
    get comments() { return comments.value; },
    get showComments() { return showComments.value; },
    get currentPage() { return currentPage.value; },
    get totalPages() { return totalPages.value; },
    get hasUnorderedList() { return hasUnorderedList.value; },
    get hasOrderedList() { return hasOrderedList.value; },
    get canUndo() { return canUndo.value; },
    get canRedo() { return canRedo.value; },

    // Setters
    setStatus: (value: string | ((prev: string) => string)) => {
      if (typeof value === "function") {
        status.value = value(status.value);
      } else {
        status.value = value;
      }
    },
    syncPaginationInfo: (p: DocxPaginationInfo) => { paginationInfo.value = p; },

    // Core operations
    importDocxFile,
    newDocument,
    exportDocx,
    undo,
    redo,

    // Formatting
    toggleBold,
    toggleItalic,
    toggleUnderline,
    toggleStrike,
    toggleSuperscript: () => { status.value = "superscript toggle (stub)"; },
    toggleSubscript: () => { status.value = "subscript toggle (stub)"; },
    setHighlight: (_highlight?: string) => { status.value = "highlight (stub)"; },
    setTextColor: (_color?: string) => { status.value = "text color (stub)"; },
    setLink: (_link?: string) => { status.value = "link (stub)"; },
    setHeading,
    setParagraphStyle,
    setLineSpacing,
    setFontFamily,
    setFontSize,
    setAlignment: (_align?: ParagraphAlignment) => { status.value = "alignment (stub)"; },
    setActiveTextRange,

    // Selection
    suppressNextDomSelectionRestore: () => {},
    selectParagraph: (nodeIndex: number) => {
      selection.value = { kind: "paragraph", nodeIndex };
    },
    selectTableCell: (_tableIndex: number, _rowIndex: number, _cellIndex: number) => {
      status.value = "select table cell (stub)";
    },

    // Text editing
    commitParagraphText,
    commitTableCellText: (_ti: number, _ri: number, _ci: number, _text: string) => {
      status.value = "commit table cell text (stub)";
    },
    commitTableCellParagraphTextRecursive: () => {
      status.value = "commit table cell text recursive (stub)";
    },
    commitSectionParagraphText: () => {
      status.value = "commit section paragraph text (stub)";
    },
    replaceExpandedSelection: (_text: string, _range?: DocxTextRange) => undefined,
    deleteExpandedSelection: (_range?: DocxTextRange) => undefined,

    // Form fields
    selectFormField: (_location?: DocxFormFieldLocation) => {},
    toggleFormCheckbox: (_location: DocxFormFieldLocation) => {},
    setFormFieldValue: (_location: DocxFormFieldLocation, _value: string) => {},
    updateFormFieldWidget: () => {},

    // Tables
    insertTable: () => { status.value = "insert table (stub)"; },
    clearTableCellContents: () => {},
    insertTableRow: () => {},
    insertTableColumn: () => {},
    deleteTableRow: () => {},
    deleteTableColumn: () => {},
    deleteTable: () => {},
    moveTable: () => {},
    moveEmbeddedTableToBody: () => {},

    // Images
    insertImageFile: async (_file: File) => {},
    resizeImage: () => {},
    setSyntheticTextBoxText: () => {},
    setImageWrapMode: () => {},
    moveFloatingImage: () => {},
    moveSectionFloatingImage: () => {},
    moveParagraphDropCap: () => {},
    setParagraphDropCapFontSizePt: () => {},
    setParagraphDropCapText: () => {},
    moveImage: () => {},

    // Lists
    toggleList: () => { status.value = "list toggle (stub)"; },
    adjustSelectedListDepth: () => false,
    insertListItemAfterSelection: () => undefined,
    splitParagraphAtSelection: () => undefined,
    appendParagraph: (_text?: string) => {
      const nodeIndex = model.value.nodes.length;
      const next = cloneDocModel(model.value);
      next.nodes.push({
        type: "paragraph",
        children: [{ type: "text", text: _text ?? "" }],
      });
      model.value = next;
      return nodeIndex;
    },

    // Borders
    applyBorderPreset: () => {},

    // Themes
    setDocumentTheme: (theme: DocxDocumentTheme) => { documentTheme.value = theme; },
    toggleShowTrackedChanges: () => { showTrackedChanges.value = !showTrackedChanges.value; },
    toggleShowComments: () => { showComments.value = !showComments.value; },
    setShowTrackedChanges: (v: boolean) => { showTrackedChanges.value = v; },
    setShowComments: (v: boolean) => { showComments.value = v; },

    // Export transformer
    registerPendingExportModelTransformer: () => {},

    // Selection session
    beginSelectionSession: () => {},
    clearSelectionSession: () => {},
  };

  return controller;
}
