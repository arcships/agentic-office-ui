import { ref, type Ref } from "vue";
import type { XlsxNamedRange, XlsxViewerController } from "@arcships/xlsx-core";
import {
  resolveNamedRangeTarget,
  type FormulaReferenceNameTarget,
} from "./formula-reference-highlight";

/** The grid's inline editor owns the shared formula-editing session. */
export const FORMULA_REFERENCE_OWNER_GRID = "grid";
/** The formula bar owns the shared formula-editing session. */
export const FORMULA_REFERENCE_OWNER_FORMULA_BAR = "formula-bar";

export interface FormulaReferenceEditingSession {
  active: Ref<boolean>;
  owner: Ref<string | null>;
  text: Ref<string>;
  caret: Ref<number>;
}

export interface FormulaReferenceEditing {
  readonly session: FormulaReferenceEditingSession;
  begin(owner: string, text: string, caret?: number): void;
  update(owner: string, text: string, caret?: number): void;
  end(owner: string): void;
}

// One session per controller, so the formula bar and the grid (in the viewer or
// in a host-composed `XlsxSheetSurface`) agree without a new controller API or
// event plumbing between the two components. A different controller is a
// different workbook and therefore a different session.
const sessions = new WeakMap<object, FormulaReferenceEditingSession>();

function sessionFor(controller: object): FormulaReferenceEditingSession {
  let session = sessions.get(controller);
  if (!session) {
    session = {
      active: ref(false),
      caret: ref(0),
      owner: ref<string | null>(null),
      text: ref(""),
    };
    sessions.set(controller, session);
  }
  return session;
}

/**
 * Shared, transient "what is being typed" state for formula reference
 * highlighting. It never touches the workbook: committing or cancelling the
 * edit is what changes the document, `end()` only clears the highlight.
 */
export function useFormulaReferenceEditing(controller: object): FormulaReferenceEditing {
  const session = sessionFor(controller);

  function begin(owner: string, text: string, caret: number = text.length) {
    session.active.value = true;
    session.owner.value = owner;
    session.text.value = text;
    session.caret.value = Number.isFinite(caret) ? Math.max(0, caret) : text.length;
  }

  function update(owner: string, text: string, caret: number = text.length) {
    if (!session.active.value || session.owner.value !== owner) return;
    session.text.value = text;
    session.caret.value = Number.isFinite(caret) ? Math.max(0, caret) : text.length;
  }

  function end(owner: string) {
    // A stale blur from the surface that just lost focus must not clear the
    // session the newly focused surface already started.
    if (session.owner.value !== owner) return;
    session.active.value = false;
    session.owner.value = null;
    session.text.value = "";
    session.caret.value = 0;
  }

  return { begin, end, session, update };
}

function namedRangeMatchesSheet(
  range: XlsxNamedRange,
  requestedSheetIndex: number | undefined,
): boolean {
  return range.sheetIndex === undefined || range.sheetIndex === requestedSheetIndex;
}

/**
 * Builds the `resolveName` callback for the controller's Name Manager entries.
 * Worksheet-scoped names only resolve on their own sheet; workbook-scoped names
 * resolve everywhere. A name that points at another sheet keeps that sheet name
 * so the formula bar can colour it while the grid leaves it un-outlined.
 */
export function createNamedRangeResolver(
  controller: XlsxViewerController,
): (name: string, sheetName: string | null) => FormulaReferenceNameTarget | null {
  return (name, sheetName) => {
    const ranges = controller.getNamedRanges?.() ?? [];
    const activeSheetIndex = controller.activeSheet?.workbookSheetIndex;
    const requestedSheetIndex = sheetName
      ? controller.sheets?.find((sheet) => sheet.name.toLowerCase() === sheetName.toLowerCase())?.workbookSheetIndex
      : activeSheetIndex;
    const lowered = name.toLowerCase();
    const candidates = ranges.filter(
      (entry) => entry.name.toLowerCase() === lowered
        && namedRangeMatchesSheet(entry, requestedSheetIndex),
    );
    if (candidates.length === 0) return null;

    const scoped = candidates.find((entry) => entry.sheetIndex === requestedSheetIndex);
    const workbookScoped = candidates.find((entry) => entry.sheetIndex === undefined);
    const match = scoped ?? workbookScoped ?? candidates[0];
    if (!match) return null;

    const scopeSheetName = match.sheetIndex !== undefined
      ? controller.sheets?.find((sheet) => sheet.workbookSheetIndex === match.sheetIndex)?.name ?? null
      : sheetName;
    return resolveNamedRangeTarget(match.refersTo, scopeSheetName ?? null);
  };
}

/** True when a reference's target lives on the active sheet and can be outlined there. */
export function isReferenceOnActiveSheet(
  reference: { sheetName: string | null },
  activeSheetName: string | null,
): boolean {
  if (!reference.sheetName) return true;
  if (!activeSheetName) return false;
  return reference.sheetName.toLowerCase() === activeSheetName.toLowerCase();
}
