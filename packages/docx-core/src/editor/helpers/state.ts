// Editor state machine: history (undo/redo), composition, and layout-cache
// state plus the transaction reducer.
// Upstream: react-viewer/src/core/state.ts (227 lines).
//
// This is the framework-agnostic state core consumed by the editor composable.

import type {
  DocModel,
  ParagraphNode,
  TextRunNode
} from "../../engine/types";
import type {
  DocxEditorSelection,
  DocxTextRange
} from "./editor-types";
import {
  DEFAULT_DOCX_HISTORY_MAX_BYTES,
  trimHistoryEntriesToBudget
} from "./history-budget";

export interface EditorHistoryEntry {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingTypingStyle?: TextRunNode["style"];
}

export interface EditorHistoryState {
  past: EditorHistoryEntry[];
  future: EditorHistoryEntry[];
  maxEntries: number;
  /** Added compatibly; missing values use the default byte budget. */
  maxBytes?: number;
  /** Approximate retained bytes, populated by `createEditorStateV2`. */
  pastBytes?: number;
  /** Approximate retained bytes, populated by `createEditorStateV2`. */
  futureBytes?: number;
}

export interface EditorCompositionState {
  isComposing: boolean;
  buffer: string;
}

export interface EditorLayoutCacheState {
  version: number;
  lastMeasuredAt: number;
}

export interface EditorStateV2 {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingTypingStyle?: TextRunNode["style"];
  history: EditorHistoryState;
  composition: EditorCompositionState;
  layoutCache: EditorLayoutCacheState;
}

export interface EditorTransactionV2 {
  type: string;
  model?: DocModel;
  selection?: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingTypingStyle?: TextRunNode["style"];
  mergeWithPrevious?: boolean;
  pushHistory?: boolean;
  clearFuture?: boolean;
  updateLayoutVersion?: boolean;
}

export function createEditorStateV2(options: {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingTypingStyle?: TextRunNode["style"];
  historyMaxEntries?: number;
  historyMaxBytes?: number;
}): EditorStateV2 {
  const requestedMaxEntries = options.historyMaxEntries ?? 200;
  const requestedMaxBytes = options.historyMaxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES;
  const maxEntries = Number.isFinite(requestedMaxEntries) && requestedMaxEntries > 0
    ? Math.max(1, Math.floor(requestedMaxEntries))
    : 200;
  const maxBytes = Number.isFinite(requestedMaxBytes) && requestedMaxBytes > 0
    ? Math.max(1, Math.floor(requestedMaxBytes))
    : DEFAULT_DOCX_HISTORY_MAX_BYTES;
  return {
    model: options.model,
    selection: options.selection,
    activeTextRange: options.activeTextRange,
    pendingTypingStyle: options.pendingTypingStyle,
    history: {
      past: [],
      future: [],
      maxEntries,
      maxBytes,
      pastBytes: 0,
      futureBytes: 0
    },
    composition: {
      isComposing: false,
      buffer: ""
    },
    layoutCache: {
      version: 1,
      lastMeasuredAt: Date.now()
    }
  };
}

function snapshotState(state: EditorStateV2): EditorHistoryEntry {
  return {
    model: state.model,
    selection: state.selection,
    activeTextRange: state.activeTextRange,
    pendingTypingStyle: state.pendingTypingStyle
  };
}

export function applyEditorTransactionV2(
  state: EditorStateV2,
  transaction: EditorTransactionV2
): EditorStateV2 {
  const nextModel = transaction.model ?? state.model;
  const nextSelection = transaction.selection ?? state.selection;
  const nextActiveTextRange = transaction.activeTextRange ?? state.activeTextRange;
  const nextPendingTypingStyle =
    transaction.pendingTypingStyle ?? state.pendingTypingStyle;

  const modelChanged = nextModel !== state.model;
  const selectionChanged = nextSelection !== state.selection;
  const rangeChanged = nextActiveTextRange !== state.activeTextRange;
  const styleChanged = nextPendingTypingStyle !== state.pendingTypingStyle;

  const shouldPushHistory =
    transaction.pushHistory !== false &&
    (modelChanged || selectionChanged || rangeChanged || styleChanged);

  let past = state.history.past;
  let future = state.history.future;
  let pastBytes = state.history.pastBytes ?? 0;
  let futureBytes = state.history.futureBytes ?? 0;
  const historyMaxBytes = state.history.maxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES;

  if (shouldPushHistory) {
    if (transaction.mergeWithPrevious && past.length > 0) {
      past = [...past.slice(0, -1), snapshotState(state)];
    } else {
      past = [...past, snapshotState(state)];
    }
    const trimmedPast = trimHistoryEntriesToBudget(past, {
      maxEntries: state.history.maxEntries,
      maxBytes: historyMaxBytes
    });
    past = trimmedPast.entries;
    pastBytes = trimmedPast.estimatedBytes;
  }

  if (transaction.clearFuture !== false && shouldPushHistory) {
    future = [];
    futureBytes = 0;
  }

  const nextLayoutVersion =
    transaction.updateLayoutVersion === false
      ? state.layoutCache.version
      : modelChanged
        ? state.layoutCache.version + 1
        : state.layoutCache.version;

  return {
    ...state,
    model: nextModel,
    selection: nextSelection,
    activeTextRange: nextActiveTextRange,
    pendingTypingStyle: nextPendingTypingStyle,
    history: {
      ...state.history,
      past,
      future,
      pastBytes,
      futureBytes
    },
    layoutCache: {
      version: nextLayoutVersion,
      lastMeasuredAt: Date.now()
    }
  };
}

export function undoEditorStateV2(state: EditorStateV2): EditorStateV2 {
  const previous = state.history.past[state.history.past.length - 1];
  if (!previous) {
    return state;
  }

  const updatedPast = trimHistoryEntriesToBudget(
    state.history.past.slice(0, -1),
    {
      maxEntries: state.history.maxEntries,
      maxBytes: state.history.maxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES
    }
  );
  const updatedFuture = trimHistoryEntriesToBudget(
    [snapshotState(state), ...state.history.future],
    {
      maxEntries: state.history.maxEntries,
      maxBytes: state.history.maxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES,
      newestAt: "start"
    }
  );

  return {
    ...state,
    model: previous.model,
    selection: previous.selection,
    activeTextRange: previous.activeTextRange,
    pendingTypingStyle: previous.pendingTypingStyle,
    history: {
      ...state.history,
      past: updatedPast.entries,
      future: updatedFuture.entries,
      pastBytes: updatedPast.estimatedBytes,
      futureBytes: updatedFuture.estimatedBytes
    },
    layoutCache: {
      version: state.layoutCache.version + 1,
      lastMeasuredAt: Date.now()
    }
  };
}

export function redoEditorStateV2(state: EditorStateV2): EditorStateV2 {
  const [next, ...remainingFuture] = state.history.future;
  if (!next) {
    return state;
  }

  const updatedPast = trimHistoryEntriesToBudget(
    [...state.history.past, snapshotState(state)],
    {
      maxEntries: state.history.maxEntries,
      maxBytes: state.history.maxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES
    }
  );
  const updatedFuture = trimHistoryEntriesToBudget(
    remainingFuture,
    {
      maxEntries: state.history.maxEntries,
      maxBytes: state.history.maxBytes ?? DEFAULT_DOCX_HISTORY_MAX_BYTES,
      newestAt: "start"
    }
  );

  return {
    ...state,
    model: next.model,
    selection: next.selection,
    activeTextRange: next.activeTextRange,
    pendingTypingStyle: next.pendingTypingStyle,
    history: {
      ...state.history,
      past: updatedPast.entries,
      future: updatedFuture.entries,
      pastBytes: updatedPast.estimatedBytes,
      futureBytes: updatedFuture.estimatedBytes
    },
    layoutCache: {
      version: state.layoutCache.version + 1,
      lastMeasuredAt: Date.now()
    }
  };
}

export function paragraphAtSelection(
  model: DocModel,
  selection: DocxEditorSelection
): ParagraphNode | undefined {
  if (selection.kind !== "paragraph") {
    return undefined;
  }

  const node = model.nodes[selection.nodeIndex];
  return node && node.type === "paragraph" ? node : undefined;
}
