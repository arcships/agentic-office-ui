// ---------------------------------------------------------------------------
// XLSX Web Worker Client — postMessage RPC wrapper around @dukelib/sheets-wasm
// Zero React/Vue dependency. Works in any browser or Web Worker–compatible
// runtime.
// ---------------------------------------------------------------------------

/**
 * Messages exchanged between the main thread and the worker.
 */
export interface XlsxWorkerRequest {
  /** Unique correlation id echoed back in the response. */
  id: number
  /** RPC method name. */
  method: string
  /** Method arguments (structured-clone-safe). */
  args: unknown[]
}

export interface XlsxWorkerResponse {
  /** Correlation id matching the request. */
  id: number
  /** Result when the call succeeded. */
  result?: unknown
  /** Serialized error when the call failed. */
  error?: { message: string; name: string }
}

export interface XlsxWorkerCallbacks {
  /**
   * Called after the workbook has been fully loaded and parsed.
   * Receives the parsed workbook metadata.
   */
  onLoaded?: (data: unknown) => void
  /**
   * Called when an unrecoverable error occurs.
   */
  onError?: (error: Error) => void
  /**
   * Called when progress information is available (e.g. parse progress).
   */
  onProgress?: (progress: { loaded: number; total: number }) => void
}

export type XlsxWorkerMessageHandler = (event: MessageEvent<XlsxWorkerResponse>) => void

// ── Client ─────────────────────────────────────────────────────────────────

export class XlsxWorkerClient {
  private worker: Worker | null = null
  private requestId = 0
  private pending = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (reason: Error) => void }
  >()
  private callbacks: XlsxWorkerCallbacks
  private messageHandler: XlsxWorkerMessageHandler
  private disposed = false

  /**
   * @param workerScript URL of the worker script that imports and initialises
   *   `@dukelib/sheets-wasm`. The script must implement the server-side of
   *   this RPC protocol.
   * @param callbacks Optional lifecycle callbacks for load / error / progress.
   */
  constructor(
    workerScript: string | URL,
    callbacks: XlsxWorkerCallbacks = {}
  ) {
    this.callbacks = callbacks

    this.messageHandler = (event: MessageEvent<XlsxWorkerResponse>) => {
      const { id, result, error } = event.data
      const pending = this.pending.get(id)
      if (!pending) return

      this.pending.delete(id)
      if (error) {
        pending.reject(new Error(error.message))
      } else {
        pending.resolve(result)
      }
    }

    this.worker = new Worker(workerScript, { type: "module" })
    this.worker.addEventListener("message", this.messageHandler)
    this.worker.addEventListener("error", (evt) => {
      const err = evt.error ?? new Error("Worker error")
      this.callbacks.onError?.(err)
    })
  }

  // ── RPC core ──────────────────────────────────────────────────────────

  /**
   * Send a request to the worker and return a promise that resolves with the
   * response.
   */
  async call<T = unknown>(method: string, ...args: unknown[]): Promise<T> {
    if (this.disposed || !this.worker) {
      throw new Error("XlsxWorkerClient is disposed")
    }

    const id = ++this.requestId

    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (v: unknown) => void, reject })

      const request: XlsxWorkerRequest = { id, method, args }
      this.worker!.postMessage(request)
    })
  }

  // ── Public API ────────────────────────────────────────────────────────

  /**
   * Load a workbook from raw bytes.
   */
  async load(buffer: ArrayBuffer): Promise<unknown> {
    return this.call("load", buffer, [buffer.byteLength])
  }

  /**
   * Fetch and load a remote workbook by URL.
   */
  async loadFromUrl(url: string): Promise<unknown> {
    return this.call("loadFromUrl", url)
  }

  /**
   * Retrieve metadata for a specific sheet.
   */
  async getSheetData(sheetIndex: number): Promise<unknown> {
    return this.call("getSheetData", sheetIndex)
  }

  /**
   * Retrieve all workbook sheet metadata.
   */
  async getAllSheetData(): Promise<unknown[]> {
    return this.call("getAllSheetData")
  }

  /**
   * Read a batch of rows from a sheet (optimised for large grids).
   */
  async getRowsBatch(
    workbookSheetIndex: number,
    startRow: number,
    rowCount: number
  ): Promise<unknown[] | null> {
    return this.call("getRowsBatch", workbookSheetIndex, startRow, rowCount)
  }

  /**
   * Get the stylesheet metadata for the loaded workbook.
   */
  async getStyles(): Promise<unknown> {
    return this.call("getStyles")
  }

  /**
   * Write workbook bytes (e.g. for download / export).
   */
  async getWorkbookBytes(): Promise<ArrayBuffer> {
    return this.call("getWorkbookBytes")
  }

  /**
   * Write workbook bytes to CSV.
   */
  async getCsvBytes(sheetIndex?: number): Promise<ArrayBuffer> {
    return this.call("getCsvBytes", sheetIndex)
  }

  /**
   * Terminate the worker and release all resources.
   */
  dispose(): void {
    if (this.disposed) return
    this.disposed = true

    if (this.worker) {
      this.worker.removeEventListener("message", this.messageHandler)
      this.worker.terminate()
      this.worker = null
    }

    // Reject all pending requests
    for (const [, pending] of this.pending) {
      pending.reject(new Error("XlsxWorkerClient disposed"))
    }
    this.pending.clear()
  }

  /**
   * True when the client has been disposed.
   */
  get isDisposed(): boolean {
    return this.disposed
  }
}

// ── Worker-side helper ─────────────────────────────────────────────────────
//
// The code below is a reference skeleton for the worker script. It is NOT
// bundled into this module — worker scripts should import what they need from
// this package or copy the relevant pieces.
//
// ```ts
// // worker-entry.ts (bundled separately with @dukelib/sheets-wasm)
// import { initWasm } from "@extend-ai/xlsx-core"
//
// interface XlsxWorkerRequest { id: number; method: string; args: unknown[] }
//
// const handlers: Record<string, (...args: unknown[]) => Promise<unknown>> = {
//   async load(buffer: ArrayBuffer) {
//     await initWasm()
//     // ... parse and store workbook
//   },
// }
//
// self.onmessage = async (e: MessageEvent<XlsxWorkerRequest>) => {
//   const { id, method, args } = e.data
//   const handler = handlers[method]
//   if (!handler) {
//     self.postMessage({ id, error: { message: `Unknown method: ${method}`, name: "Error" } })
//     return
//   }
//   try {
//     const result = await handler(...args)
//     self.postMessage({ id, result })
//   } catch (err) {
//     self.postMessage({ id, error: { message: (err as Error).message, name: (err as Error).name } })
//   }
// }
// ```
