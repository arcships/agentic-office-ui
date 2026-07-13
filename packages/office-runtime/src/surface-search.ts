export type SurfaceSearchStatus = "idle" | "searching" | "ready" | "error"

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
  search(query: string): Promise<SurfaceSearchState<M>>
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

  async function activate(index: number, expectedRequest = requestVersion): Promise<void> {
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
        throw createAbortError()
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
        throw createAbortError()
      }
      requestedActiveIndex = state.activeIndex
      commit({
        status: "error",
        query: state.query,
        matches: state.matches,
        activeIndex: state.activeIndex,
        error: { code: "ACTIVATION_FAILED", message: errorMessage(error) },
      })
      throw error
    } finally {
      if (activationController === controller) activationController = undefined
    }
  }

  const session: SurfaceSearchSession<M> = {
    async search(rawQuery) {
      assertUsable()
      const query = rawQuery.trim()
      if (!query) {
        session.clearSearch()
        return state
      }

      const currentRequest = ++requestVersion
      searchController?.abort()
      cancelActivation()
      const controller = new AbortController()
      searchController = controller
      requestedActiveIndex = -1
      commit({ status: "searching", query, matches: [], activeIndex: -1 })

      let matches: readonly M[]
      try {
        matches = await adapter.search(query, { signal: controller.signal })
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
          matches: [],
          activeIndex: -1,
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
      if (matches.length) await activate(0, currentRequest)
      return state
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
