import type { DocModel } from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import { mapsToWasmPackage, wasmModelToDocumentXml, wasmSerializeDocx } from "./wasm";

export async function modelToDocumentXml(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<string> {
  return wasmModelToDocumentXml(
    model,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
}

export async function serializeDocModel(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<OoxmlPackage> {
  const bytes = await wasmSerializeDocx(
    model,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  const { parseDocx } = await import("./ooxml-core");
  return parseDocx(bytes);
}

export async function serializeDocx(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<ArrayBuffer> {
  return wasmSerializeDocx(model, basePackage ? mapsToWasmPackage(basePackage) : undefined);
}
