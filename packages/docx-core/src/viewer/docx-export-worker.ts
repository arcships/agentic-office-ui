import type { DocModel } from "../engine/types";
import type { OoxmlPackage } from "../engine/ooxml-core";
import { serializeDocx } from "../engine/serializer";
import { DocxSerializationError } from "../engine/prepare-model";
import { setWasmSource } from "../engine/wasm";
import type {
  DocxExportErrorCode,
  DocxExportWorkerRequest,
  DocxExportWorkerResponse,
} from "./docx-export";

function serializeError(error: unknown): {
  code: DocxExportErrorCode;
  name?: string;
  message: string;
  stack?: string;
} {
  const message = error instanceof Error ? error.message : String(error);
  if (error instanceof Error) {
    return {
      code: error instanceof DocxSerializationError ? "INVALID_MODEL" : "EXPORT_FAILED",
      name: error.name,
      message,
      stack: error.stack,
    };
  }
  return { code: "EXPORT_FAILED", message };
}

self.addEventListener(
  "message",
  async (event: MessageEvent<DocxExportWorkerRequest>) => {
    const workerSelf = self as unknown as {
      postMessage(message: unknown, transfer?: Transferable[]): void;
    };
    const request = event.data;
    if (!request || request.type !== "export-docx") {
      return;
    }

    try {
      if (request.wasmSource !== undefined) {
        setWasmSource(request.wasmSource);
      }
      const progress: DocxExportWorkerResponse = {
        id: request.id,
        type: "progress",
        phase: "serializing",
        percent: 10,
      };
      self.postMessage(progress);

      const bytes = await serializeDocx(
        request.model as DocModel,
        request.basePackage as OoxmlPackage | undefined,
      );

      const finalizing: DocxExportWorkerResponse = {
        id: request.id,
        type: "progress",
        phase: "finalizing",
        percent: 90,
      };
      self.postMessage(finalizing);

      const response: DocxExportWorkerResponse = {
        id: request.id,
        type: "success",
        bytes,
      };
      workerSelf.postMessage(response, [bytes]);
    } catch (error) {
      const response: DocxExportWorkerResponse = {
        id: request.id,
        type: "error",
        error: serializeError(error),
      };
      self.postMessage(response);
    }
  },
);
