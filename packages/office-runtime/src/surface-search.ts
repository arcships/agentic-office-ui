export type SurfaceSearchStatus = "idle" | "searching" | "ready" | "error"

/**
 * Query options shared by every office surface search (DOCX, XLSX, …).
 *
 * Defaults mirror Word/Excel: literal matching, case-insensitive, matching
 * inside words. `regex` treats the query as a JavaScript regular expression
 * and composes with `wholeWord`/`caseSensitive`.
 */
export interface SurfaceSearchOptions {
  caseSensitive?: boolean
  wholeWord?: boolean
  regex?: boolean
}

export interface SurfaceSearchError {
  code: "SEARCH_FAILED" | "ACTIVATION_FAILED"
  message: string
}

export interface SurfaceSearchState<M> {
  status: SurfaceSearchStatus
  query: string
  matches: readonly M[]
  activeIndex: number
  error?: SurfaceSearchError
}

export interface SurfaceSearchApi<M> {
  search(query: string, options?: SurfaceSearchOptions): Promise<SurfaceSearchState<M>>
  /**
   * Re-run the active query against the current surface content and keep the
   * active match index where possible. Hosts call this after an edit replaces
   * the model so the result list and highlights survive the edit.
   */
  refreshSearch(): Promise<SurfaceSearchState<M>>
  activateSearchMatch(index: number): Promise<void>
  searchNext(): Promise<void>
  searchPrevious(): Promise<void>
  clearSearch(): void
  getSearchState(): SurfaceSearchState<M>
}

export interface SurfaceSearchTaskContext {
  signal: AbortSignal
}

export interface SurfaceSearchAdapter<M> {
  search(
    query: string,
    context: SurfaceSearchTaskContext,
    options: Readonly<SurfaceSearchOptions>,
  ): Promise<readonly M[]> | readonly M[]
  activate(
    match: M,
    index: number,
    context: SurfaceSearchTaskContext,
  ): Promise<void> | void
  clear?(): void
}

export interface SurfaceSearchSession<M> extends SurfaceSearchApi<M> {
  subscribe(listener: (state: SurfaceSearchState<M>) => void): () => void
  dispose(): void
}

function createAbortError(): Error {
  if (typeof DOMException !== "undefined") {
    return new DOMException("Surface search task was replaced.", "AbortError")
  }
  const error = new Error("Surface search task was replaced.")
  error.name = "AbortError"
  return error
}

function isAbortError(error: unknown): boolean {
  return error instanceof Error && error.name === "AbortError"
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function escapeSurfaceSearchPattern(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

/**
 * Build the regular expression used by a surface search.
 *
 * Offsets stay UTF-16 code-unit offsets in the source text, matching DOM
 * Range's offset model. `wholeWord` uses Unicode letter/number/underscore
 * boundaries instead of `\b`, which misclassifies CJK text and combining
 * marks.
 */
export function createSurfaceSearchExpression(
  query: string,
  options: Readonly<SurfaceSearchOptions> = {},
): RegExp {
  const flags = options.caseSensitive ? "gu" : "giu"
  const pattern = options.regex ? query : escapeSurfaceSearchPattern(query)
  const bounded = options.wholeWord
    ? `(?<![\\p{L}\\p{N}_])(?:${pattern})(?![\\p{L}\\p{N}_])`
    : pattern
  return new RegExp(bounded, flags)
}

export function createSurfaceSearchSession<M>(
  adapter: SurfaceSearchAdapter<M>,
): SurfaceSearchSession<M> {
  let state: SurfaceSearchState<M> = {
    status: "idle",
    query: "",
    matches: [],
    activeIndex: -1,
  }
  let requestVersion = 0
  let activationVersion = 0
  let requestedActiveIndex = -1
  let searchController: AbortController | undefined
  let activationController: AbortController | undefined
  let activeQuery: { query: string; options: SurfaceSearchOptions } | undefined
  let disposed = false
  const listeners = new Set<(value: SurfaceSearchState<M>) => void>()

  function assertUsable(): void {
    if (disposed) throw new Error("Surface search session has been disposed.")
  }

  function commit(next: SurfaceSearchState<M>): void {
    state = next
    for (const listener of listeners) listener(state)
  }

  function cancelActivation(): void {
    activationVersion += 1
    activationController?.abort()
    activationController = undefined
  }

  async function activate(
    index: number,
    expectedRequest = requestVersion,
    rethrowFailure = true,
  ): Promise<void> {
    assertUsable()
    if (!Number.isInteger(index) || index < 0 || index >= state.matches.length) {
      throw new RangeError(`Search result index ${index} is out of range.`)
    }

    cancelActivation()
    const currentActivation = activationVersion
    const controller = new AbortController()
    activationController = controller
    requestedActiveIndex = index
    const match = state.matches[index]

    try {
      await adapter.activate(match, index, { signal: controller.signal })
      if (
        disposed
        || controller.signal.aborted
        || expectedRequest !== requestVersion
        || currentActivation !== activationVersion
      ) {
        // Superseded by a newer query or activation. This is normal
        // interactive behaviour, not a host-visible failure.
        return
      }
      commit({
        status: "ready",
        query: state.query,
        matches: state.matches,
        activeIndex: index,
      })
    } catch (error) {
      if (
        isAbortError(error)
        || controller.signal.aborted
        || expectedRequest !== requestVersion
        || currentActivation !== activationVersion
      ) {
        return
      }
      requestedActiveIndex = state.activeIndex
      commit({
        status: "error",
        query: state.query,
        matches: state.matches,
        activeIndex: state.activeIndex,
        error: { code: "ACTIVATION_FAILED", message: errorMessage(error) },
      })
      if (rethrowFailure) throw error
    } finally {
      if (activationController === controller) activationController = undefined
    }
  }

  async function runSearch(
    rawQuery: string,
    options: SurfaceSearchOptions,
    preserveActiveIndex: boolean,
  ): Promise<SurfaceSearchState<M>> {
    assertUsable()
    const query = rawQuery.trim()
    if (!query) {
      session.clearSearch()
      return state
    }

    const previousMatches = preserveActiveIndex ? state.matches : []
    const previousActiveIndex = preserveActiveIndex
      ? Math.max(0, state.activeIndex)
      : -1
    activeQuery = { query, options: { ...options } }

    const currentRequest = ++requestVersion
    searchController?.abort()
    cancelActivation()
    const controller = new AbortController()
    searchController = controller
    requestedActiveIndex = previousActiveIndex
    commit({
      status: "searching",
      query,
      matches: previousMatches,
      activeIndex: previousActiveIndex,
    })

    let matches: readonly M[]
    try {
      matches = await adapter.search(query, { signal: controller.signal }, { ...options })
    } catch (error) {
      if (
        isAbortError(error)
        || controller.signal.aborted
        || currentRequest !== requestVersion
      ) {
        throw createAbortError()
      }
      commit({
        status: "error",
        query,
        matches: previousMatches,
        activeIndex: previousActiveIndex,
        error: { code: "SEARCH_FAILED", message: errorMessage(error) },
      })
      throw error
    } finally {
      if (searchController === controller) searchController = undefined
    }

    if (disposed || controller.signal.aborted || currentRequest !== requestVersion) {
      throw createAbortError()
    }
    commit({ status: "ready", query, matches: [...matches], activeIndex: -1 })
    if (matches.length) {
      // The implicit activation is a convenience for hosts that type into a
      // search box. A refresh keeps the previously active result (clamped to
      // the new list); a new query activates the first result. A superseded or
      // failed activation is already reflected in `state`, so resolve with the
      // latest state instead of leaking an AbortError or activation failure.
      const preferredIndex = preserveActiveIndex
        ? Math.min(previousActiveIndex, matches.length - 1)
        : 0
      try {
        await activate(preferredIndex, currentRequest, false)
      } catch {
        // `state` already carries ACTIVATION_FAILED.
      }
    }
    return state
  }

  const session: SurfaceSearchSession<M> = {
    search(rawQuery, options = {}) {
      return runSearch(rawQuery, options, false)
    },

    refreshSearch() {
      assertUsable()
      if (!activeQuery) return Promise.resolve(state)
      return runSearch(activeQuery.query, activeQuery.options, true)
    },

    activateSearchMatch(index) {
      return activate(index)
    },

    searchNext() {
      assertUsable()
      if (!state.matches.length || state.status === "searching") return Promise.resolve()
      const base = requestedActiveIndex >= 0 ? requestedActiveIndex : state.activeIndex
      return activate((base + 1 + state.matches.length) % state.matches.length)
    },

    searchPrevious() {
      assertUsable()
      if (!state.matches.length || state.status === "searching") return Promise.resolve()
      const base = requestedActiveIndex >= 0 ? requestedActiveIndex : state.activeIndex
      return activate((base - 1 + state.matches.length) % state.matches.length)
    },

    clearSearch() {
      if (disposed) return
      requestVersion += 1
      searchController?.abort()
      searchController = undefined
      cancelActivation()
      requestedActiveIndex = -1
      activeQuery = undefined
      adapter.clear?.()
      commit({ status: "idle", query: "", matches: [], activeIndex: -1 })
    },

    getSearchState() {
      return state
    },

    subscribe(listener) {
      assertUsable()
      listeners.add(listener)
      listener(state)
      return () => listeners.delete(listener)
    },

    dispose() {
      if (disposed) return
      session.clearSearch()
      disposed = true
      listeners.clear()
    },
  }

  return session
}
