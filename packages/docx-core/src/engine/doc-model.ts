// doc-model aggregation — builds DocModel from OOXML packages/bytes.
// Ported from upstream doc-model/src/index.ts (commit 6f70b92).

import type { OoxmlPackage } from "./ooxml-core";
import {
  mapsToWasmPackage,
  wasmBuildDocModelFromPackage,
  wasmPackageToMaps,
} from "./wasm";
import { parseDocx } from "./ooxml-core";

import { normalizeDocModel } from "./normalize";
import type { DocModel } from "./types";

export * from "./types";
export { cloneDocModel } from "./clone";
export { normalizeDocModel } from "./normalize";

export async function buildDocModel(pkg: OoxmlPackage): Promise<DocModel> {
  const wasmPackage = mapsToWasmPackage({
    parts: pkg.parts,
    binaryAssets: pkg.binaryAssets,
  });
  const model = (await wasmBuildDocModelFromPackage(wasmPackage)) as DocModel;
  return normalizeDocModel(model);
}

export async function buildDocModelFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<{ package: OoxmlPackage; model: DocModel }> {
  const payload = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const buffer = payload.buffer.slice(
    payload.byteOffset,
    payload.byteOffset + payload.byteLength,
  ) as ArrayBuffer;
  const pkg = await parseDocx(buffer);
  const model = await buildDocModel(pkg);
  return { package: pkg, model };
}
