import {
  initWasm as initRuntimeWasm,
  setWasmSource as setRuntimeWasmSource,
  type WasmSource
} from "../engine/wasm";

export type WorkerWasmSource = string | ArrayBuffer | WebAssembly.Module;

function bufferSourceToArrayBuffer(
  source: ArrayBuffer | ArrayBufferView<ArrayBufferLike>
): ArrayBuffer {
  if (source instanceof ArrayBuffer) {
    return source.slice(0);
  }

  const bytes = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  const copy = new Uint8Array(bytes);
  return copy.buffer;
}

function sourceToWorkerSource(source: WasmSource): WorkerWasmSource | undefined {
  if (typeof source === "string") {
    return source;
  }
  if (typeof URL !== "undefined" && source instanceof URL) {
    return source.href;
  }
  if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) {
    return bufferSourceToArrayBuffer(source);
  }
  if (typeof WebAssembly !== "undefined" && source instanceof WebAssembly.Module) {
    return source;
  }
  if (typeof Request !== "undefined" && source instanceof Request) {
    return source.url;
  }

  return undefined;
}

function createLegacyWorkerWasmBridge() {
  let hasConfiguredSource = false;
  let workerSource: WorkerWasmSource | undefined;
  return Object.freeze({
    remember(source: WasmSource): void {
      hasConfiguredSource = true;
      workerSource = sourceToWorkerSource(source);
    },
    canUseInWorker(): boolean {
      return !hasConfiguredSource || workerSource !== undefined;
    },
    getWorkerSource(): WorkerWasmSource | undefined {
      return workerSource;
    },
  });
}

const legacyWorkerWasmBridge = createLegacyWorkerWasmBridge();

/**
 * Configure the compatibility default DOCX runtime.
 *
 * @deprecated Since 0.2.0. Create an isolated runtime with
 * `createDocxRuntime({ wasmSource })` or `createDocxRuntime({ wasmUrl })`.
 * This compatibility entry remains available throughout 0.x and will not be
 * removed before 1.0.0.
 */
export function setWasmSource(source: WasmSource): void {
  setRuntimeWasmSource(source);
  legacyWorkerWasmBridge.remember(source);
}

export function initWasm(source?: WasmSource): ReturnType<typeof initRuntimeWasm> {
  if (source !== undefined) {
    legacyWorkerWasmBridge.remember(source);
  }
  return initRuntimeWasm(source);
}

export function canUseConfiguredWasmSourceInWorker(): boolean {
  return legacyWorkerWasmBridge.canUseInWorker();
}

export function getConfiguredWorkerWasmSource(): WorkerWasmSource | undefined {
  return legacyWorkerWasmBridge.getWorkerSource();
}

export type { WasmSource };
