// useDocxFindReplace — the state behind the editor's find & replace bar.
//
// The surface owns the search session (query, results, active result); this
// composable owns the bar's own inputs (replace text, toggle state, open/closed)
// and turns "Replace" / "Replace All" into the controller's replace commands.
// One undo entry per command is guaranteed by the controller.

import { computed, nextTick, shallowReadonly, shallowRef, type ComputedRef, type ShallowRef } from "vue"
import type { DocxEditorController } from "@arcships/docx-core"
import {
  canReplaceDocxSearchMatch,
  docxReplaceTargetFromMatch,
  type DocxFindReplaceSurface,
} from "./editor-search-replace"
import type { DocxSearchMatch, DocxSearchOptions, DocxSearchState, DocxSearchStatus } from "./useDocxSearch"

/** Optional message the bar renders after a replace command. */
export type DocxReplaceFeedback =
  | { kind: "replaced"; count: number }
  | { kind: "unavailable" }

export interface DocxFindReplaceUiState {
  open: boolean
  query: string
  replacement: string
  caseSensitive: boolean
  wholeWord: boolean
  regex: boolean
  status: DocxSearchStatus
  matches: readonly DocxSearchMatch[]
  activeIndex: number
  feedback?: DocxReplaceFeedback
}

export interface DocxFindReplaceApi {
  state: Readonly<ShallowRef<DocxFindReplaceUiState>>
  /** False until a result the editor can rewrite is active. */
  canReplace: ComputedRef<boolean>
  /** False until at least one result of the active query can be rewritten. */
  canReplaceAll: ComputedRef<boolean>
  openFind(): void
  closeFind(): void
  toggleFind(): void
  setQuery(query: string): void
  setReplacement(text: string): void
  setOption(key: "caseSensitive" | "wholeWord" | "regex", value: boolean): void
  searchNext(): void
  searchPrevious(): void
  replaceActive(): Promise<void>
  replaceAll(): Promise<void>
  /** Feed surface `searchStateChange` events into the bar. */
  syncSearchState(state: DocxSearchState): void
  /** Re-run the active query once the surface is mounted. */
  refresh(): Promise<void>
}

const IDLE_SEARCH_STATE: DocxSearchState = {
  status: "idle",
  query: "",
  matches: [],
  activeIndex: -1,
}

function isAbortError(error: unknown): boolean {
  return error instanceof Error && error.name === "AbortError"
}

export function useDocxFindReplace(options: {
  controller: () => DocxEditorController
  surface: () => DocxFindReplaceSurface | undefined
}): DocxFindReplaceApi {
  const state = shallowRef<DocxFindReplaceUiState>({
    open: false,
    query: "",
    replacement: "",
    caseSensitive: false,
    wholeWord: false,
    regex: false,
    status: "idle",
    matches: [],
    activeIndex: -1,
  })

  let requestToken = 0

  function patch(next: Partial<DocxFindReplaceUiState>): void {
    state.value = { ...state.value, ...next }
  }

  function searchOptions(): DocxSearchOptions {
    const current = state.value
    return {
      caseSensitive: current.caseSensitive,
      wholeWord: current.wholeWord,
      regex: current.regex,
    }
  }

  function syncSearchState(next: DocxSearchState): void {
    patch({
      status: next.status,
      matches: next.matches,
      activeIndex: next.activeIndex,
    })
  }

  async function runSearch(): Promise<void> {
    const token = ++requestToken
    const surface = options.surface()
    const query = state.value.query
    if (!surface) return
    if (!query.trim()) {
      surface.clearSearch()
      syncSearchState(IDLE_SEARCH_STATE)
      return
    }
    try {
      const next = await surface.search(query, searchOptions())
      if (token !== requestToken) return
      syncSearchState(next)
    } catch (error) {
      if (isAbortError(error) || token !== requestToken) return
      patch({ status: "error" })
    }
  }

  async function refresh(): Promise<void> {
    const surface = options.surface()
    if (!surface || !state.value.query.trim()) return
    const token = ++requestToken
    try {
      const next = await surface.refreshSearch()
      if (token === requestToken) syncSearchState(next)
    } catch {
      // A superseded refresh is expected; genuine failures are published by the
      // surface through `syncSearchState` (status/error), so a replace command
      // must not fail because the follow-up search raced an edit.
    }
  }

  function openFind(): void {
    if (state.value.open) return
    patch({ open: true })
    void runSearch()
  }

  function closeFind(): void {
    requestToken += 1
    options.surface()?.clearSearch()
    state.value = {
      ...state.value,
      open: false,
      query: "",
      status: "idle",
      matches: [],
      activeIndex: -1,
      feedback: undefined,
    }
  }

  function toggleFind(): void {
    if (state.value.open) closeFind()
    else openFind()
  }

  function setQuery(query: string): void {
    patch({ query, feedback: undefined })
    void runSearch()
  }

  function setReplacement(text: string): void {
    patch({ replacement: text })
  }

  function setOption(key: "caseSensitive" | "wholeWord" | "regex", value: boolean): void {
    patch({ [key]: value, feedback: undefined } as Partial<DocxFindReplaceUiState>)
    void runSearch()
  }

  function searchNext(): void {
    void options.surface()?.searchNext().catch((error: unknown) => {
      if (!isAbortError(error)) patch({ status: "error" })
    })
  }

  function searchPrevious(): void {
    void options.surface()?.searchPrevious().catch((error: unknown) => {
      if (!isAbortError(error)) patch({ status: "error" })
    })
  }

  async function afterReplace(feedback: DocxReplaceFeedback): Promise<void> {
    patch({ feedback })
    // The surface refreshes itself when the model changes; awaiting the tick and
    // refreshing explicitly keeps the counts and active index in step with the
    // edit that just happened.
    await nextTick()
    await refresh()
  }

  async function replaceActive(): Promise<void> {
    const current = state.value
    const match = current.matches[current.activeIndex]
    if (!match) return
    const target = docxReplaceTargetFromMatch(match)
    if (!target) {
      patch({ feedback: { kind: "unavailable" } })
      return
    }
    const replaced = options.controller().replaceMatch(target, current.replacement)
    if (!replaced) {
      // The result the bar held was already gone (undo, another edit) — the
      // refresh below replaces the stale list with the current one.
      await afterReplace({ kind: "replaced", count: 0 })
      return
    }
    await afterReplace({ kind: "replaced", count: 1 })
  }

  async function replaceAll(): Promise<void> {
    const current = state.value
    if (!current.query.trim()) return
    let summary
    try {
      summary = options.controller().replaceAll(
        current.query,
        current.replacement,
        searchOptions(),
      )
    } catch {
      // An invalid regular expression is reported by the search field itself;
      // a replace command must not throw into an event handler.
      patch({ status: "error" })
      return
    }
    if (summary.replaced > 0) {
      await afterReplace({ kind: "replaced", count: summary.replaced })
      return
    }
    await afterReplace(summary.skipped > 0 ? { kind: "unavailable" } : { kind: "replaced", count: 0 })
  }

  const canReplace = computed(() => {
    const current = state.value
    const match = current.matches[current.activeIndex]
    return Boolean(match) && canReplaceDocxSearchMatch(match)
  })

  const canReplaceAll = computed(() =>
    state.value.matches.some((match) => canReplaceDocxSearchMatch(match)),
  )

  return {
    state: shallowReadonly(state) as Readonly<ShallowRef<DocxFindReplaceUiState>>,
    canReplace,
    canReplaceAll,
    openFind,
    closeFind,
    toggleFind,
    setQuery,
    setReplacement,
    setOption,
    searchNext,
    searchPrevious,
    replaceActive,
    replaceAll,
    syncSearchState,
    refresh,
  }
}
