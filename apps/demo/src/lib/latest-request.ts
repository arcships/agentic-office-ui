// Latest-wins request gate for demo pages that load a sample asynchronously.
//
// A page that awaits `fetch` and then assigns the result lets an older, slower
// response overwrite a newer document (the surface page had no guard at all).
// The gate owns the sequence number and the AbortController together so a new
// load always aborts the previous fetch and stale results are ignored.

export type LatestRequest = {
  readonly token: number
  readonly signal: AbortSignal
  isCurrent(): boolean
}

export type LatestRequestGate = {
  begin(): LatestRequest
  isCurrent(token: number): boolean
  abort(): void
}

export function createLatestRequestGate(): LatestRequestGate {
  let sequence = 0
  let controller: AbortController | undefined

  return {
    begin(): LatestRequest {
      sequence += 1
      controller?.abort()
      const active = new AbortController()
      controller = active
      const token = sequence
      return {
        token,
        signal: active.signal,
        isCurrent: () => token === sequence && !active.signal.aborted,
      }
    },
    isCurrent(token: number): boolean {
      return token === sequence
    },
    abort(): void {
      sequence += 1
      controller?.abort()
      controller = undefined
    },
  }
}
