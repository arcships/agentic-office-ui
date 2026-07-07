// OOXML package parsing, embedded font collection, twips/points conversion,
// and XML attribute helpers.
// Upstream editor.tsx: lines 470-476 (xmlAttribute), 1488-1691 (OOXML helpers).

import type { TableBorderSet } from "../../engine/types";
import type { OoxmlPackage } from "../../engine/ooxml-core";
import { TWIPS_PER_PIXEL } from "../../viewer/section-layout";

/** Extract the value of an XML attribute from a tag string. */
export function xmlAttribute(tagXml: string, attribute: string): string | undefined {
  const escaped = attribute.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tagXml.match(
    new RegExp(`${escaped}=(?:"([^"]+)"|'([^']+)')`, "i")
  );
  return match?.[1] ?? match?.[2];
}

export function createDefaultEditorTableBorders(): TableBorderSet {
  return {
    top: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    right: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    bottom: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    left: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    insideH: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    insideV: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
  };
}

export function twipsToSignedPixels(twips?: number): number | undefined {
  if (!Number.isFinite(twips)) {
    return undefined;
  }

  return Math.round((twips as number) / TWIPS_PER_PIXEL);
}

export function pointsToPixels(points?: number): number | undefined {
  if (!Number.isFinite(points)) {
    return undefined;
  }

  return Math.max(0, Number((((points as number) * 96) / 72).toFixed(2)));
}

export type EmbeddedFontFaceDescriptor = {
  family: string;
  style: "normal" | "italic";
  weight: string;
  source: ArrayBuffer;
};

export function relationshipPartNameForOoxmlPart(partName: string): string {
  const segments = partName.split("/");
  const fileName = segments.pop() ?? partName;
  const directory = segments.join("/");
  return directory.length > 0
    ? `${directory}/_rels/${fileName}.rels`
    : `_rels/${fileName}.rels`;
}

export function resolveRelativeOoxmlPartName(
  basePartName: string,
  target: string
): string {
  if (/^[a-z]+:/i.test(target)) {
    return target;
  }

  const normalizedBasePartName = basePartName.replace(/^\/+/, "");
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }

  const baseSegments = normalizedBasePartName.split("/");
  baseSegments.pop();
  target.split("/").forEach((segment) => {
    if (!segment || segment === ".") {
      return;
    }

    if (segment === "..") {
      if (baseSegments.length > 0) {
        baseSegments.pop();
      }
      return;
    }

    baseSegments.push(segment);
  });

  return baseSegments.join("/");
}

export function parseOoxmlRelationships(
  pkg: OoxmlPackage,
  partName: string
): Map<string, string> {
  const relationships = new Map<string, string>();
  const relationshipsXml = pkg.parts.get(
    relationshipPartNameForOoxmlPart(partName)
  )?.content;
  if (!relationshipsXml) {
    return relationships;
  }

  const relationshipPattern =
    /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?>/gi;
  let match: RegExpExecArray | null;
  while ((match = relationshipPattern.exec(relationshipsXml))) {
    const relationshipId = match[1]?.trim();
    const target = match[2]?.trim();
    if (!relationshipId || !target) {
      continue;
    }

    relationships.set(
      relationshipId,
      resolveRelativeOoxmlPartName(partName, target)
    );
  }

  return relationships;
}

export function deobfuscateEmbeddedFontData(
  fontData: Uint8Array,
  fontKey?: string
): ArrayBuffer {
  const output = Uint8Array.from(fontData);
  const normalizedFontKey = (fontKey ?? "").replace(/[{}-]/g, "");
  if (!/^[0-9a-f]{32}$/i.test(normalizedFontKey)) {
    return output.buffer.slice(
      output.byteOffset,
      output.byteOffset + output.byteLength
    );
  }

  const keyBytes = Uint8Array.from(
    normalizedFontKey.match(/../g)?.map((pair) => Number.parseInt(pair, 16)) ??
      []
  ).reverse();
  const xorLength = Math.min(32, output.length);
  for (let index = 0; index < xorLength; index += 1) {
    output[index] ^= keyBytes[index % keyBytes.length] ?? 0;
  }

  return output.buffer.slice(
    output.byteOffset,
    output.byteOffset + output.byteLength
  );
}

export function collectEmbeddedFontFaceDescriptors(
  pkg: OoxmlPackage
): EmbeddedFontFaceDescriptor[] {
  const fontTableXml = pkg.parts.get("word/fontTable.xml")?.content;
  if (!fontTableXml) {
    return [];
  }

  const fontRelationships = parseOoxmlRelationships(pkg, "word/fontTable.xml");
  const descriptors: EmbeddedFontFaceDescriptor[] = [];
  const fontPattern =
    /<w:font\b[^>]*w:name="([^"]+)"[^>]*>([\s\S]*?)<\/w:font>/gi;
  const fontVariants = [
    { tagName: "embedRegular", style: "normal" as const, weight: "400" },
    { tagName: "embedBold", style: "normal" as const, weight: "700" },
    { tagName: "embedItalic", style: "italic" as const, weight: "400" },
    { tagName: "embedBoldItalic", style: "italic" as const, weight: "700" },
  ];

  let fontMatch: RegExpExecArray | null;
  while ((fontMatch = fontPattern.exec(fontTableXml))) {
    const family = fontMatch[1]?.trim();
    const fontXml = fontMatch[2] ?? "";
    if (!family) {
      continue;
    }

    fontVariants.forEach((variant) => {
      const tagXml =
        fontXml.match(
          new RegExp(`<w:${variant.tagName}\\b[^>]*\\/?>`, "i")
        )?.[0] ?? "";
      const relationshipId = xmlAttribute(tagXml, "r:id");
      if (!relationshipId) {
        return;
      }

      const partName = fontRelationships.get(relationshipId);
      const fontData = partName ? pkg.binaryAssets.get(partName) : undefined;
      if (!fontData) {
        return;
      }

      descriptors.push({
        family,
        style: variant.style,
        weight: variant.weight,
        source: deobfuscateEmbeddedFontData(
          fontData,
          xmlAttribute(tagXml, "w:fontKey")
        ),
      });
    });
  }

  return descriptors;
}
