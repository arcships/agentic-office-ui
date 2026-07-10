import { bundledXlsxWasmUrl } from "./wasm-asset";

export type XlsxWasmSource =
  | string
  | URL
  | Request
  | Response
  | BufferSource
  | WebAssembly.Module;

export type WorkerWasmSource = string | ArrayBuffer | WebAssembly.Module;


function bufferSourceToArrayBuffer(source: ArrayBuffer | ArrayBufferView<ArrayBufferLike>): ArrayBuffer {
  if (source instanceof ArrayBuffer) {
    return source.slice(0);
  }

  const bytes = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  const copy = new Uint8Array(bytes);
  return copy.buffer;
}

function sourceToWorkerSource(source: XlsxWasmSource): WorkerWasmSource | undefined {
  if (typeof source === "string") {
    return source;
  }
  if (typeof URL !== "undefined" && source instanceof URL) {
    return source.href;
  }
  if (typeof Request !== "undefined" && source instanceof Request) {
    return source.url;
  }
  if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) {
    return bufferSourceToArrayBuffer(source);
  }
  if (typeof WebAssembly !== "undefined" && source instanceof WebAssembly.Module) {
    return source;
  }

  return undefined;
}

function createLegacyDefaultWasmRuntime() {
  let modulePromise: Promise<typeof import("@dukelib/sheets-wasm")> | undefined;
  let hasConfiguredSource = false;
  let source: XlsxWasmSource | undefined;
  let workerSource: WorkerWasmSource | undefined;
  return Object.freeze({
    setSource(nextSource: XlsxWasmSource): void {
      if (modulePromise) {
        throw new Error(
          "@extend-ai/xlsx-core: setWasmSource must be called before the default main-thread WASM module initializes",
        );
      }
      hasConfiguredSource = true;
      source = nextSource;
      workerSource = sourceToWorkerSource(nextSource);
    },
    canUseInWorker(): boolean {
      return !hasConfiguredSource || workerSource !== undefined;
    },
    getWorkerSource(): WorkerWasmSource | undefined {
      return workerSource;
    },
    getModule(): Promise<typeof import("@dukelib/sheets-wasm")> {
      if (!modulePromise) {
        modulePromise = import("@dukelib/sheets-wasm").then(async (mod) => {
          await mod.default({ module_or_path: source ?? bundledXlsxWasmUrl });
          return mod;
        });
      }
      return modulePromise;
    },
  });
}

// Compatibility-only default instance. createXlsxRuntime always supplies an
// explicit Worker source and therefore never reads this state.
const legacyDefaultWasmRuntime = createLegacyDefaultWasmRuntime();

/**
 * Configure the compatibility default XLSX runtime.
 *
 * @deprecated Since 0.2.0. Create an isolated runtime with
 * `createXlsxRuntime({ wasmSource })`. This compatibility entry remains
 * available throughout 0.x and will not be removed before 1.0.0.
 */
export function setWasmSource(source: XlsxWasmSource): void {
  legacyDefaultWasmRuntime.setSource(source);
}

export function initWasm(source?: XlsxWasmSource) {
  if (source !== undefined) setWasmSource(source);
  return legacyDefaultWasmRuntime.getModule();
}

/**
 * Inspect the compatibility default XLSX runtime.
 *
 * @deprecated Since 0.2.0. Supply a Worker-transferable `wasmSource` to
 * `createXlsxRuntime()` instead. Kept throughout 0.x; earliest removal 1.0.0.
 */
export function canUseConfiguredWasmSourceInWorker(): boolean {
  return legacyDefaultWasmRuntime.canUseInWorker();
}

/**
 * Read the compatibility default XLSX runtime's Worker source.
 *
 * @deprecated Since 0.2.0. Read configuration from the owning
 * `createXlsxRuntime()` call instead. Kept throughout 0.x; earliest removal
 * 1.0.0.
 */
export function getConfiguredWorkerWasmSource(): WorkerWasmSource | undefined {
  return legacyDefaultWasmRuntime.getWorkerSource();
}

/**
 * Access the module owned by the compatibility default XLSX runtime.
 *
 * @deprecated Since 0.2.0. Use an instance created by `createXlsxRuntime()`.
 * Kept throughout 0.x; earliest removal 1.0.0.
 */
export function getSheetsWasmModule() {
  return legacyDefaultWasmRuntime.getModule();
}
