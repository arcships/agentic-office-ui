import type {
  ManualRegionLocator,
  OfficeCandidateNavigationState,
  OfficeHitCandidate,
  OfficeObjectReference,
  OfficeObjectReferenceDraft,
  OfficeReferenceCandidatePreview,
  OfficeReferenceConfirmEvent,
  OfficeReferenceSnapshot,
  OfficeSelectionKeyboardContext,
  OfficeSelectionKeyboardInput,
  OfficeSelectionKeyboardResult,
  OfficeSelectionMode,
  OfficeSelectionSessionState,
  OfficeSelectionTrigger,
} from "./types"
import { parseOfficeObjectReference, parseOfficeReferenceConfirmEvent } from "./validation"

export type OfficeCandidateNavigationAction =
  | {
      type: "set-candidates"
      candidates: readonly OfficeReferenceCandidatePreview[]
      activeCandidateId?: string
    }
  | { type: "activate-candidate"; candidateId: string }
  | { type: "move-candidate"; direction: 1 | -1 }
  | { type: "clear-candidates" }

export type OfficeSelectionSessionAction =
  | { type: "set-mode"; mode: OfficeSelectionMode }
  | { type: "candidates"; action: OfficeCandidateNavigationAction }
  | { type: "begin-region"; region: ManualRegionLocator }
  | { type: "update-region"; region: ManualRegionLocator }
  | { type: "cancel" }
  | { type: "reset"; mode?: OfficeSelectionMode }

export interface OfficeSelectionSessionOptions {
  mode?: OfficeSelectionMode
  candidates?: readonly OfficeReferenceCandidatePreview[]
  activeCandidateId?: string
  regionDraft?: ManualRegionLocator
}

export interface OfficeReferenceConfirmationInput {
  referenceId: string
  trigger: OfficeSelectionTrigger
  additiveRequested?: boolean
  snapshot?: OfficeReferenceSnapshot
}

export interface OfficeCandidateConfirmationInput extends OfficeReferenceConfirmationInput {
  candidateId?: string
}

export interface OfficeCandidateConfirmationResult {
  state: OfficeSelectionSessionState
  event: OfficeReferenceConfirmEvent
}

export const MAX_OFFICE_VISIBLE_CANDIDATES = 20

function candidateIds(candidates: readonly OfficeReferenceCandidatePreview[]): Set<string> {
  if (candidates.length > MAX_OFFICE_VISIBLE_CANDIDATES) {
    throw new RangeError(`visible candidate list must contain at most ${MAX_OFFICE_VISIBLE_CANDIDATES} items`)
  }
  const ids = new Set<string>()
  for (const candidate of candidates) {
    if (!candidate.candidateId) throw new TypeError("candidateId must not be empty")
    if (ids.has(candidate.candidateId)) throw new TypeError(`duplicate candidateId: ${candidate.candidateId}`)
    ids.add(candidate.candidateId)
  }
  return ids
}

export function createOfficeCandidateNavigationState(
  candidates: readonly OfficeReferenceCandidatePreview[] = [],
  activeCandidateId?: string,
): OfficeCandidateNavigationState {
  const ids = candidateIds(candidates)
  if (activeCandidateId !== undefined && !ids.has(activeCandidateId)) {
    throw new RangeError(`unknown candidateId: ${activeCandidateId}`)
  }
  const active = activeCandidateId ?? candidates[0]?.candidateId
  return {
    candidates: [...candidates],
    ...(active === undefined ? {} : { activeCandidateId: active }),
  }
}

export function reduceOfficeCandidateNavigation(
  state: OfficeCandidateNavigationState,
  action: OfficeCandidateNavigationAction,
): OfficeCandidateNavigationState {
  const ids = candidateIds(state.candidates)
  if (state.activeCandidateId !== undefined && !ids.has(state.activeCandidateId)) {
    throw new RangeError(`unknown active candidateId: ${state.activeCandidateId}`)
  }
  switch (action.type) {
    case "set-candidates": {
      const nextIds = candidateIds(action.candidates)
      const requested = action.activeCandidateId ?? state.activeCandidateId
      const active = requested !== undefined && nextIds.has(requested)
        ? requested
        : action.candidates[0]?.candidateId
      return {
        candidates: [...action.candidates],
        ...(active === undefined ? {} : { activeCandidateId: active }),
      }
    }
    case "activate-candidate":
      if (!ids.has(action.candidateId)) throw new RangeError(`unknown candidateId: ${action.candidateId}`)
      return { ...state, activeCandidateId: action.candidateId }
    case "move-candidate": {
      if (state.candidates.length === 0) return state
      const current = state.activeCandidateId === undefined
        ? -1
        : state.candidates.findIndex((candidate) => candidate.candidateId === state.activeCandidateId)
      const index = current < 0
        ? action.direction === 1 ? 0 : state.candidates.length - 1
        : (current + action.direction + state.candidates.length) % state.candidates.length
      return { ...state, activeCandidateId: state.candidates[index]?.candidateId }
    }
    case "clear-candidates":
      return { candidates: [] }
  }
}

function restingPhase(mode: OfficeSelectionMode): OfficeSelectionSessionState["phase"] {
  return mode === "object" ? "pointing" : "idle"
}

function assertSelectionMode(mode: string): asserts mode is OfficeSelectionMode {
  if (mode !== "content" && mode !== "object" && mode !== "region") {
    throw new RangeError(`unsupported Office selection mode: ${mode}`)
  }
}

function sessionFromCandidateNavigation(
  state: OfficeSelectionSessionState,
  candidates: OfficeCandidateNavigationState,
): OfficeSelectionSessionState {
  return {
    mode: state.mode,
    phase: candidates.candidates.length > 0 ? "choosing" : restingPhase(state.mode),
    candidates: candidates.candidates,
    ...(candidates.activeCandidateId === undefined ? {} : { activeCandidateId: candidates.activeCandidateId }),
  }
}

export function createOfficeSelectionSessionState(
  options: OfficeSelectionSessionOptions = {},
): OfficeSelectionSessionState {
  const mode = options.mode ?? "content"
  assertSelectionMode(mode)
  const navigation = createOfficeCandidateNavigationState(options.candidates, options.activeCandidateId)
  if (options.regionDraft !== undefined && mode !== "region") {
    throw new TypeError("regionDraft requires mode=region")
  }
  return {
    mode,
    phase: options.regionDraft !== undefined
      ? "drawing"
      : navigation.candidates.length > 0
        ? "choosing"
        : restingPhase(mode),
    candidates: navigation.candidates,
    ...(navigation.activeCandidateId === undefined ? {} : { activeCandidateId: navigation.activeCandidateId }),
    ...(options.regionDraft === undefined ? {} : { regionDraft: options.regionDraft }),
  }
}

export function reduceOfficeSelectionSession(
  state: OfficeSelectionSessionState,
  action: OfficeSelectionSessionAction,
): OfficeSelectionSessionState {
  switch (action.type) {
    case "set-mode":
      return createOfficeSelectionSessionState({ mode: action.mode })
    case "candidates": {
      if (state.mode === "region") throw new TypeError("candidate navigation is unavailable in region mode")
      const navigation = reduceOfficeCandidateNavigation(
        createOfficeCandidateNavigationState(state.candidates, state.activeCandidateId),
        action.action,
      )
      return sessionFromCandidateNavigation(state, navigation)
    }
    case "begin-region":
      if (state.mode !== "region") throw new TypeError("begin-region requires mode=region")
      return { mode: "region", phase: "drawing", candidates: [], regionDraft: action.region }
    case "update-region":
      if (state.mode !== "region" || state.phase !== "drawing") {
        throw new TypeError("update-region requires an active region draft")
      }
      return { ...state, regionDraft: action.region }
    case "cancel":
      return createOfficeSelectionSessionState({ mode: state.mode })
    case "reset":
      return createOfficeSelectionSessionState({ mode: action.mode ?? "content" })
  }
}

export function activeOfficeReferenceCandidate(
  state: Pick<OfficeSelectionSessionState, "candidates" | "activeCandidateId">,
): OfficeReferenceCandidatePreview | undefined {
  return state.candidates.find((candidate) => candidate.candidateId === state.activeCandidateId)
}

export function applyOfficeSelectionKeyboard(
  state: OfficeSelectionSessionState,
  input: OfficeSelectionKeyboardInput,
  context: OfficeSelectionKeyboardContext = {},
): OfficeSelectionKeyboardResult {
  let command: OfficeSelectionKeyboardResult["command"]
  let next = state
  if (input.key === "Escape") {
    if (state.candidates.length > 0) {
      command = "dismiss-candidates"
      next = reduceOfficeSelectionSession(state, {
        type: "candidates",
        action: { type: "clear-candidates" },
      })
    } else if (state.phase === "drawing") {
      command = "cancel-selection"
      next = reduceOfficeSelectionSession(state, { type: "cancel" })
    }
  } else if (input.key === "Enter" && activeOfficeReferenceCandidate(state)) {
    command = "confirm-candidate"
  } else if (
    state.candidates.length > 0 &&
    (input.key === "Tab" || input.key === "ArrowDown" || input.key === "ArrowUp")
  ) {
    const previous = input.key === "ArrowUp" || (input.key === "Tab" && input.shiftKey === true)
    command = previous ? "previous-candidate" : "next-candidate"
    next = reduceOfficeSelectionSession(state, {
      type: "candidates",
      action: { type: "move-candidate", direction: previous ? -1 : 1 },
    })
  } else if (input.key === "ArrowRight" && context.canEnterChild) {
    command = "enter-child"
  } else if (input.key === "ArrowLeft" && context.canReturnParent) {
    command = "return-parent"
  }
  return {
    state: next,
    handled: command !== undefined,
    ...(command === undefined ? {} : { command }),
    ...(command === "confirm-candidate"
      ? { activeCandidate: activeOfficeReferenceCandidate(state) }
      : {}),
  }
}

export function createOfficeReferenceId(): string {
  if (typeof globalThis.crypto?.randomUUID !== "function") {
    throw new Error("crypto.randomUUID is required to create an Office reference id")
  }
  return `ref_${globalThis.crypto.randomUUID()}`
}

export function createOfficeObjectReference(
  draft: OfficeObjectReferenceDraft,
  referenceId: string,
): OfficeObjectReference {
  return parseOfficeObjectReference({ ...draft, referenceId })
}

export function confirmOfficeReferenceDraft(
  draft: OfficeObjectReferenceDraft,
  input: OfficeReferenceConfirmationInput,
): OfficeReferenceConfirmEvent {
  return parseOfficeReferenceConfirmEvent({
    reference: createOfficeObjectReference(draft, input.referenceId),
    ...(input.snapshot === undefined ? {} : { snapshot: input.snapshot }),
    trigger: input.trigger,
    additiveRequested: input.additiveRequested ?? false,
  })
}

export function confirmOfficeHitCandidate(
  candidate: OfficeHitCandidate,
  input: OfficeReferenceConfirmationInput,
): OfficeReferenceConfirmEvent {
  return confirmOfficeReferenceDraft(candidate.draft, input)
}

export function confirmOfficeCandidate(
  state: OfficeSelectionSessionState,
  input: OfficeCandidateConfirmationInput,
): OfficeCandidateConfirmationResult {
  const candidate = input.candidateId === undefined
    ? activeOfficeReferenceCandidate(state)
    : state.candidates.find((entry) => entry.candidateId === input.candidateId)
  if (!candidate) {
    throw new RangeError(input.candidateId === undefined
      ? "no active candidate to confirm"
      : `unknown candidateId: ${input.candidateId}`)
  }
  return {
    state: reduceOfficeSelectionSession(state, {
      type: "candidates",
      action: { type: "clear-candidates" },
    }),
    event: confirmOfficeReferenceDraft(candidate.draft, input),
  }
}
