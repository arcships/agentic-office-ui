import type { XlsxChart, XlsxChartsheet, XlsxSheetData, XlsxTable, XlsxWorkbookTab } from "./types";
import { bundledXlsxWasmUrl } from "./wasm-asset";
import { getConfiguredWorkerWasmSource, type WorkerWasmSource } from "./wasm";

type WorkerMessage =
  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
      };
    }
  | {
      id: number;
      type: "parseCharts";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
      };
    }
  | {
      id: number;
      type: "getCellSnapshot";
      payload: {
        workbookSheetIndex: number;
        row: number;
        col: number;
      };
    }
  | {
      id: number;
      type: "getRowsBatch";
      payload: {
        workbookSheetIndex: number;
        startRow: number;
        rowCount: number;
      };
    };

type WorkerSuccessMessage =
  | {
      id: number;
      success: true;
      result: {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        tabs: XlsxWorkbookTab[];
      };
    }
  | {
      id: number;
      success: true;
      result: {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        sheets: XlsxSheetData[];
        tablesByWorkbookSheetIndex: XlsxTable[][];
        tabs: XlsxWorkbookTab[];
      };
    }
  | {
      id: number;
      success: true;
      result: {
        displayValue: string;
        formula: string;
      };
    }
  | {
      id: number;
      success: true;
      result: unknown[] | null;
    };

type WorkerErrorMessage = {
  id: number;
  success: false;
  error: string;
};

type WorkerResponse = WorkerSuccessMessage | WorkerErrorMessage;

type PendingRequest = {
  cleanup: () => void;
  reject: (error: Error) => void;
  resolve: (value: unknown) => void;
};

export interface XlsxWorkerClientOptions {
  /** Uses a host-provided Worker factory for an isolated client instance. */
  createWorker?: () => Worker;
  /** Uses a host-served module Worker URL for an isolated client instance. */
  workerUrl?: string;
  /** Overrides the WASM source sent to this client Worker only. */
  wasmSource?: WorkerWasmSource;
  /** Keep the legacy module default only for callers that construct clients directly. */
  useLegacyDefaultWasmSource?: boolean;
  /** Instance runtime lifecycle hook. */
  onDispose?: () => void;
}

export function createBundledXlsxWorker(): Worker {
  return new Worker(new URL("./xlsx-worker.js", import.meta.url), { type: "module" });
}

function createAbortError() {
  if (typeof DOMException !== "undefined") {
    return new DOMException("Aborted", "AbortError");
  }

  const error = new Error("Aborted");
  error.name = "AbortError";
  return error;
}

function workerFailureError(message: string) {
  const error = new Error(message);
  error.name = "XlsxWorkerError";
  return error;
}

export class XlsxWorkerClient {
  private readonly worker: Worker;

  private readonly wasmSource: WorkerWasmSource | undefined;

  private readonly useLegacyDefaultWasmSource: boolean;

  private readonly onDispose: (() => void) | undefined;

  private nextRequestId = 1;

  private wasmSourceSent = false;

  private readonly pendingRequests = new Map<number, PendingRequest>();

  private disposed = false;

  private terminalError: Error | null = null;

  constructor(options: XlsxWorkerClientOptions = {}) {
    this.wasmSource = options.wasmSource;
    this.useLegacyDefaultWasmSource = options.useLegacyDefaultWasmSource ?? true;
    this.onDispose = options.onDispose;
    this.worker = options.createWorker
      ? options.createWorker()
      : options.workerUrl
        ? new Worker(options.workerUrl, { type: "module" })
        : createBundledXlsxWorker();
    this.worker.addEventListener("message", this.handleMessage);
    this.worker.addEventListener("error", this.handleError);
    this.worker.addEventListener("messageerror", this.handleMessageError);
  }

  dispose() {
    this.close(createAbortError());
  }

  loadWorkbook(
    buffer: ArrayBuffer,
    skipXmlParsing = false,
    showHiddenSheets = false,
    signal?: AbortSignal
  ) {
    if (signal?.aborted) {
      return Promise.reject(createAbortError());
    }
    const workerBuffer = cloneArrayBufferForTransfer(buffer);
    const wasmSource = this.resolveNextWasmSource();
    return this.request<{
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      sheets: XlsxSheetData[];
      tablesByWorkbookSheetIndex: XlsxTable[][];
      tabs: XlsxWorkbookTab[];
    }>({
      id: 0,
      payload: {
        buffer: workerBuffer,
        showHiddenSheets,
        skipXmlParsing,
        wasmSource
      },
      type: "load"
    }, [workerBuffer], signal, wasmSource !== undefined);
  }

  getCellSnapshot(
    workbookSheetIndex: number,
    row: number,
    col: number,
    signal?: AbortSignal
  ) {
    if (signal?.aborted) {
      return Promise.reject(createAbortError());
    }
    return this.request<{
      displayValue: string;
      formula: string;
    }>({
      id: 0,
      payload: { col, row, workbookSheetIndex },
      type: "getCellSnapshot"
    }, [], signal);
  }

  parseCharts(
    buffer: ArrayBuffer,
    skipXmlParsing = false,
    showHiddenSheets = false,
    signal?: AbortSignal
  ) {
    if (signal?.aborted) {
      return Promise.reject(createAbortError());
    }
    const workerBuffer = cloneArrayBufferForTransfer(buffer);
    const wasmSource = this.resolveNextWasmSource();
    return this.request<{
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      tabs: XlsxWorkbookTab[];
    }>({
      id: 0,
      payload: {
        buffer: workerBuffer,
        showHiddenSheets,
        skipXmlParsing,
        wasmSource
      },
      type: "parseCharts"
    }, [workerBuffer], signal, wasmSource !== undefined);
  }

  getRowsBatch(
    workbookSheetIndex: number,
    startRow: number,
    rowCount: number,
    signal?: AbortSignal
  ) {
    if (signal?.aborted) {
      return Promise.reject(createAbortError());
    }
    return this.request<unknown[] | null>({
      id: 0,
      payload: { rowCount, startRow, workbookSheetIndex },
      type: "getRowsBatch"
    }, [], signal);
  }

  private request<TResult>(
    message: WorkerMessage,
    transfer: Transferable[] = [],
    signal?: AbortSignal,
    marksWasmSourceSent = false
  ) {
    return new Promise<TResult>((resolve, reject) => {
      if (this.disposed) {
        reject(this.terminalError ?? createAbortError());
        return;
      }
      if (signal?.aborted) {
        reject(createAbortError());
        return;
      }

      const id = this.nextRequestId;
      this.nextRequestId += 1;
      let cleaned = false;
      const handleAbort = () => {
        const request = this.pendingRequests.get(id);
        if (!request) return;
        this.pendingRequests.delete(id);
        request.cleanup();
        request.reject(createAbortError());
      };
      const cleanup = () => {
        if (cleaned) return;
        cleaned = true;
        signal?.removeEventListener("abort", handleAbort);
      };
      this.pendingRequests.set(id, {
        cleanup,
        reject,
        resolve: resolve as (value: unknown) => void
      });
      signal?.addEventListener("abort", handleAbort, { once: true });
      try {
        this.worker.postMessage({ ...message, id }, transfer);
        if (marksWasmSourceSent) this.wasmSourceSent = true;
      } catch (error) {
        const request = this.pendingRequests.get(id);
        if (!request) return;
        this.pendingRequests.delete(id);
        request.cleanup();
        request.reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  }

  private resolveNextWasmSource(): WorkerWasmSource | undefined {
    if (this.wasmSourceSent) return undefined;
    return this.wasmSource ?? (
      this.useLegacyDefaultWasmSource ? getConfiguredWorkerWasmSource() : undefined
    ) ?? bundledXlsxWasmUrl;
  }

  private close(error: Error) {
    if (this.disposed) return;
    this.disposed = true;
    this.terminalError = error;
    this.worker.removeEventListener("message", this.handleMessage);
    this.worker.removeEventListener("error", this.handleError);
    this.worker.removeEventListener("messageerror", this.handleMessageError);
    this.worker.terminate();
    const pending = [...this.pendingRequests.values()];
    this.pendingRequests.clear();
    for (const request of pending) {
      request.cleanup();
      request.reject(error);
    }
    try { this.onDispose?.(); } catch { /* lifecycle callbacks are isolated */ }
  }

  private readonly handleError = (event: ErrorEvent) => {
    this.close(workerFailureError(event.message || "Worker request failed."));
  };

  private readonly handleMessageError = () => {
    this.close(workerFailureError("Worker returned an unreadable response."));
  };

  private readonly handleMessage = (event: MessageEvent<WorkerResponse>) => {
    const message = event.data;
    if (!message || typeof message.id !== "number") {
      return;
    }
    const request = this.pendingRequests.get(message.id);
    if (!request) {
      return;
    }

    this.pendingRequests.delete(message.id);
    request.cleanup();
    if (!message.success) {
      request.reject(new Error(message.error));
      return;
    }

    request.resolve(message.result);
  };
}

function cloneArrayBufferForTransfer(buffer: ArrayBuffer) {
  return buffer.slice(0);
}
