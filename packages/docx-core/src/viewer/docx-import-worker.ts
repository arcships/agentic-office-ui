import { buildDocModel } from "../engine/doc-model";
import { parseDocx } from "../engine/ooxml-core";
import { setWasmSource } from "../engine/wasm";
import {
  OfficeLoadError,
  validateOfficeArchive,
} from "@arcships/office-runtime";
import {
  assertDocxModelBudget,
  assertDocxParseTime,
} from "../resource-limits";
import { validateDocxImageAssets } from "../image-budget";

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
  phase?: string;
  limit?: string;
  actual?: number;
  allowed?: number;
} {
  const message = error instanceof Error ? error.message : String(error);
  const code: DocxImportErrorCode =
    error instanceof OfficeLoadError && error.code === "TIMEOUT"
      ? "TIMEOUT"
      : error instanceof OfficeLoadError && error.code === "LIMIT_EXCEEDED"
      ? "LIMIT_EXCEEDED"
      : error instanceof OfficeLoadError && error.code === "IMAGE_LIMIT_EXCEEDED"
      ? "IMAGE_LIMIT_EXCEEDED"
      : error instanceof OfficeLoadError && error.code === "INVALID_IMAGE"
      ? "INVALID_IMAGE"
      : error instanceof OfficeLoadError && error.code === "IMAGE_DECODE_FAILED"
      ? "IMAGE_DECODE_FAILED"
      : error instanceof OfficeLoadError && error.code === "ABORTED"
        ? "ABORTED"
        : error instanceof WebAssembly.CompileError ||
    /webassembly|wasm|instantiate|magic word/i.test(message)
      ? "WASM_LOAD_FAILED"
      : "PARSE_FAILED";

  if (error instanceof Error) {
    return {
      code,
      name: error.name,
      message,
      stack: error.stack,
      ...(error instanceof OfficeLoadError && error.phase ? { phase: error.phase } : {}),
      ...(error instanceof OfficeLoadError && error.limit ? { limit: error.limit } : {}),
      ...(error instanceof OfficeLoadError && error.actual !== undefined ? { actual: error.actual } : {}),
      ...(error instanceof OfficeLoadError && error.allowed !== undefined ? { allowed: error.allowed } : {}),
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
      validateOfficeArchive(request.buffer, request.limits, { format: "docx" });
      const pkg = await parseDocx(request.buffer);
      validateDocxImageAssets(pkg, request.limits);
      const parsedAt = performanceNow();
      const model = await buildDocModel(pkg);
      const finishedAt = performanceNow();
      assertDocxModelBudget(model, request.limits);
      assertDocxParseTime(finishedAt - startedAt, request.limits);
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
