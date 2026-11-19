import {
  mapsToWasmPackage,
  wasmPackageToArrayBuffer,
  wasmPackageToMaps,
  wasmParseDocx,
  type LegacyWasmOoxmlPackage,
  type WasmOoxmlPackage
} from "./wasm";

export interface OoxmlPart {
  name: string;
  content: string;
}

export interface OoxmlPackage {
  parts: Map<string, OoxmlPart>;
  binaryAssets: Map<string, Uint8Array>;
}

/**
 * OpenDocument text is a ZIP (or flat XML) with no Word package. The WASM
 * parser only reads OOXML, so reject it here with the reason the viewer shows
 * instead of reporting a generic ZIP failure.
 */
function isOpenDocumentPackage(input: ArrayBuffer): boolean {
  const bytes = new Uint8Array(input)
  const head = new TextDecoder("utf-8", { fatal: false }).decode(bytes.subarray(0, 4096))
  const isZip = bytes.length >= 4 &&
    bytes[0] === 0x50 && bytes[1] === 0x4b &&
    bytes[2] === 0x03 && bytes[3] === 0x04
  if (isZip) {
    return head.includes("mimetype") &&
      /application\/vnd\.oasis\.opendocument\.text/i.test(head)
  }
  return /<office:document\b/i.test(head) &&
    /application\/vnd\.oasis\.opendocument\.text/i.test(head)
}

export async function parseDocx(input: ArrayBuffer): Promise<OoxmlPackage> {
  if (isOpenDocumentPackage(input)) {
    throw new Error(
      "OpenDocument text has no Word package; use the document runtime to open .odt files in view-only mode.",
    );
  }
  const wasmPackage = await wasmParseDocx(input);
  const { parts, binaryAssets } = wasmPackageToMaps(wasmPackage);
  return { parts, binaryAssets };
}

export async function packageToArrayBuffer(pkg: OoxmlPackage): Promise<ArrayBuffer> {
  return wasmPackageToArrayBuffer(mapsToWasmPackage(pkg));
}

export function createMinimalDocxPackage(
  documentXml = DEFAULT_DOCUMENT_XML
): OoxmlPackage {
  return {
    parts: new Map([
      ["[Content_Types].xml", { name: "[Content_Types].xml", content: DEFAULT_CONTENT_TYPES_XML }],
      ["_rels/.rels", { name: "_rels/.rels", content: DEFAULT_ROOT_RELS_XML }],
      ["word/document.xml", { name: "word/document.xml", content: documentXml }],
      [
        "word/_rels/document.xml.rels",
        { name: "word/_rels/document.xml.rels", content: DEFAULT_DOCUMENT_RELS_XML }
      ]
    ]),
    binaryAssets: new Map()
  };
}

export function getPart(pkg: OoxmlPackage, partName: string): OoxmlPart | undefined {
  return pkg.parts.get(partName);
}

export function withPart(pkg: OoxmlPackage, part: OoxmlPart): OoxmlPackage {
  const parts = new Map(pkg.parts);
  parts.set(part.name, part);
  return {
    parts,
    binaryAssets: new Map(pkg.binaryAssets)
  };
}

const WORD_MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DEFAULT_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_MAIN_NS}">
  <w:body>
    <w:p><w:r><w:t/></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const DEFAULT_CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const DEFAULT_ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DEFAULT_DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

export type { LegacyWasmOoxmlPackage, WasmOoxmlPackage };
