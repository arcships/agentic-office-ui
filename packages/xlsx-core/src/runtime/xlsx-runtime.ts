import { bundledXlsxWasmUrl } from "../wasm-asset";
import {
  XlsxWorkerClient,
  type XlsxWorkerClientOptions,
} from "../worker-client";
import type { WorkerWasmSource } from "../wasm";
import {
  resolveXlsxRuntimeLimits,
  type XlsxRuntimeLimits,
} from "../resource-limits";

export interface XlsxRuntimeParseOptions {
  showHiddenSheets?: boolean;
  skipXmlParsing?: boolean;
}

export interface XlsxRuntimeDiagnostic {
  type: "worker-client-created" | "worker-client-disposed" | "runtime-disposed";
  runtimeId: string;
  workerUrl?: string;
  wasmUrl?: string;
}

/** Stable failures raised by an XLSX runtime instance. */
export type XlsxRuntimeErrorCode = "RUNTIME_DISPOSED";

/**
 * Public error returned when an XLSX runtime operation cannot be started.
 * The shape deliberately stays independent from the private shared runtime.
 */
export class XlsxRuntimeError extends Error {
  readonly code: XlsxRuntimeErrorCode;
  readonly format = "xlsx" as const;
  readonly runtimeId: string;

  constructor(code: XlsxRuntimeErrorCode, message: string, runtimeId: string) {
    super(message);
    this.name = "XlsxRuntimeError";
    this.code = code;
    this.runtimeId = runtimeId;
  }

  toJSON(): Record<string, string> {
    return {
      name: this.name,
      code: this.code,
      message: this.message,
      format: this.format,
      runtimeId: this.runtimeId,
    };
  }
}

export interface XlsxRuntimeConfig {
  wasmSource?: WorkerWasmSource;
  workerUrl?: string;
  createWorker?: () => Worker;
  parseOptions?: XlsxRuntimeParseOptions;
  /** Per-instance input, archive, XML and parsing budgets. */
  limits?: XlsxRuntimeLimits;
  onDiagnostic?: (event: XlsxRuntimeDiagnostic) => void;
}

export interface XlsxRuntime {
  readonly id: string;
  readonly wasmSource: WorkerWasmSource;
  readonly workerUrl?: string;
  readonly parseOptions: Readonly<XlsxRuntimeParseOptions>;
  readonly limits: Readonly<XlsxRuntimeLimits>;
  createWorkerClient(): XlsxWorkerClient;
  dispose(): void;
}

function createRuntimeId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `xlsx-runtime-${crypto.randomUUID()}`;
  }
  return `xlsx-runtime-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

export function createXlsxRuntime(config: XlsxRuntimeConfig = {}): XlsxRuntime {
  const id = createRuntimeId();
  const wasmSource = typeof config.wasmSource === "string"
    ? String(config.wasmSource)
    : config.wasmSource instanceof ArrayBuffer
      ? config.wasmSource.slice(0)
      : config.wasmSource ?? bundledXlsxWasmUrl;
  const workerUrl = config.workerUrl ? String(config.workerUrl) : undefined;
  const createWorker = config.createWorker;
  const parseOptions = Object.freeze({ ...(config.parseOptions ?? {}) });
  const limits = resolveXlsxRuntimeLimits(config.limits);
  const onDiagnostic = config.onDiagnostic;
  const clients = new Set<XlsxWorkerClient>();
  let disposed = false;

  const report = (event: XlsxRuntimeDiagnostic): void => {
    try { onDiagnostic?.(Object.freeze(event)); } catch { /* diagnostics are isolated */ }
  };

  return Object.freeze({
    id,
    wasmSource,
    workerUrl,
    parseOptions,
    limits,
    createWorkerClient(): XlsxWorkerClient {
      if (disposed) {
        throw new XlsxRuntimeError(
          "RUNTIME_DISPOSED",
          "XLSX 运行实例已经销毁。",
          id,
        );
      }
      let client: XlsxWorkerClient;
      const options: XlsxWorkerClientOptions = {
        createWorker,
        workerUrl,
        wasmSource,
        limits,
        useLegacyDefaultWasmSource: false,
        onDispose: () => {
          clients.delete(client);
          report({
            type: "worker-client-disposed",
            runtimeId: id,
            workerUrl,
            wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
          });
        },
      };
      client = new XlsxWorkerClient(options);
      clients.add(client);
      report({
        type: "worker-client-created",
        runtimeId: id,
        workerUrl,
        wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
      });
      return client;
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      for (const client of [...clients]) client.dispose();
      clients.clear();
      report({
        type: "runtime-disposed",
        runtimeId: id,
        workerUrl,
        wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
      });
    },
  });
}
