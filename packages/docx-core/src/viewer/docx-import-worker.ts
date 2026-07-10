import { buildDocModel } from "../engine/doc-model";
import { parseDocx } from "../engine/ooxml-core";
import { setWasmSource } from "../engine/wasm";

import type {
  DocxImportErrorCode,
  DocxImportWorkerRequest,
  DocxImportWorkerResponse,
  DocxImportWorkerTimings,
} from "./docx-import";

function performanceNow(): number {
  return typeof performance !== "undefined" &&
    typeof performance.now === "function"
    ? performance.now()
    : Date.now();
}

function serializeError(error: unknown): {
  code: DocxImportErrorCode;
  name?: string;
  message: string;
  stack?: string;
} {
  const message = error instanceof Error ? error.message : String(error);
  const code: DocxImportErrorCode =
    error instanceof WebAssembly.CompileError ||
    /webassembly|wasm|instantiate|magic word/i.test(message)
      ? "WASM_LOAD_FAILED"
      : "PARSE_FAILED";

  if (error instanceof Error) {
    return {
      code,
      name: error.name,
      message,
      stack: error.stack,
    };
  }

  return {
    code,
    message,
  };
}

self.addEventListener(
  "message",
  async (event: MessageEvent<DocxImportWorkerRequest>) => {
    const request = event.data;
    if (!request || request.type !== "import-docx") {
      return;
    }

    try {
      if (request.wasmSource !== undefined) {
        setWasmSource(request.wasmSource);
      }
      const startedAt = performanceNow();
      const pkg = await parseDocx(request.buffer);
      const parsedAt = performanceNow();
      const model = await buildDocModel(pkg);
      const finishedAt = performanceNow();
      const timings: DocxImportWorkerTimings = {
        totalMs: finishedAt - startedAt,
        parseMs: parsedAt - startedAt,
        buildModelMs: finishedAt - parsedAt,
      };
      const response: DocxImportWorkerResponse = {
        id: request.id,
        type: "success",
        package: pkg,
        model,
        timings,
      };
      self.postMessage(response);
    } catch (error) {
      const response: DocxImportWorkerResponse = {
        id: request.id,
        type: "error",
        error: serializeError(error),
      };
      self.postMessage(response);
    }
  }
);
