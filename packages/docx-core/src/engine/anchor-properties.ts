// Floating anchor properties that the compact WASM model deliberately does not
// surface. They live in the drawing's source XML, so the TS model reads them
// back from there and exposes them on `ImageRunNode.floating`.

import type { ImageRunNode } from "./types";

export type FloatingAnchorProperties = Partial<
  NonNullable<ImageRunNode["floating"]>
>;

export interface FloatingEffectExtentPx {
  leftPx: number;
  topPx: number;
  rightPx: number;
  bottomPx: number;
}

const ANCHOR_TAG_PATTERN = /<(?:[A-Za-z_][\w.-]*:)?anchor\b([^>]*)>/i;
const EFFECT_EXTENT_TAG_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?effectExtent\b([^>]*)\/?>/i;
const EMU_PER_PIXEL = 9525;

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function readXmlAttribute(
  attributes: string | undefined,
  name: string,
): string | undefined {
  if (!attributes) {
    return undefined;
  }
  const match = attributes.match(
    new RegExp(`(?:^|\\s)${escapeRegExp(name)}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, "i"),
  );
  return match?.[1] ?? match?.[2];
}

function parseOnOffAttribute(value: string | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  const normalized = value.trim().toLowerCase();
  if (normalized === "1" || normalized === "true" || normalized === "on") {
    return true;
  }
  if (normalized === "0" || normalized === "false" || normalized === "off") {
    return false;
  }
  return undefined;
}

function emuAttributeToPixels(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const parsed = Number(value.trim());
  if (!Number.isFinite(parsed)) {
    return undefined;
  }
  const pixels = Number((Math.max(0, parsed) / EMU_PER_PIXEL).toFixed(3));
  return pixels > 0 ? pixels : undefined;
}

/**
 * Read `layoutInCell`, `allowOverlap`, `locked`, `simplePos` and
 * `effectExtent` from a `<wp:anchor>` (or equivalent namespace prefix).
 */
export function parseFloatingAnchorProperties(
  sourceXml: string | undefined,
): FloatingAnchorProperties | undefined {
  if (!sourceXml) {
    return undefined;
  }

  const anchor = sourceXml.match(ANCHOR_TAG_PATTERN);
  if (!anchor) {
    return undefined;
  }

  const properties: FloatingAnchorProperties = {};
  const anchorAttributes = anchor[1];
  const layoutInCell = parseOnOffAttribute(
    readXmlAttribute(anchorAttributes, "layoutInCell"),
  );
  if (layoutInCell !== undefined) {
    properties.layoutInCell = layoutInCell;
  }
  const allowOverlap = parseOnOffAttribute(
    readXmlAttribute(anchorAttributes, "allowOverlap"),
  );
  if (allowOverlap !== undefined) {
    properties.allowOverlap = allowOverlap;
  }
  const lockedAnchor = parseOnOffAttribute(
    readXmlAttribute(anchorAttributes, "locked"),
  );
  if (lockedAnchor !== undefined) {
    properties.lockedAnchor = lockedAnchor;
  }
  const simplePos = parseOnOffAttribute(
    readXmlAttribute(anchorAttributes, "simplePos"),
  );
  if (simplePos !== undefined) {
    properties.simplePos = simplePos;
  }

  const effectExtent = sourceXml.match(EFFECT_EXTENT_TAG_PATTERN);
  if (effectExtent) {
    const effectAttributes = effectExtent[1];
    const leftPx = emuAttributeToPixels(readXmlAttribute(effectAttributes, "l"));
    const topPx = emuAttributeToPixels(readXmlAttribute(effectAttributes, "t"));
    const rightPx = emuAttributeToPixels(readXmlAttribute(effectAttributes, "r"));
    const bottomPx = emuAttributeToPixels(readXmlAttribute(effectAttributes, "b"));
    if (
      leftPx !== undefined ||
      topPx !== undefined ||
      rightPx !== undefined ||
      bottomPx !== undefined
    ) {
      properties.effectExtent = { leftPx, topPx, rightPx, bottomPx };
    }
  }

  return Object.keys(properties).length > 0 ? properties : undefined;
}

/** Effect overflow in CSS pixels, normalized for geometry math. */
export function resolveFloatingEffectExtentPx(
  floating: ImageRunNode["floating"],
): FloatingEffectExtentPx {
  const effect = floating?.effectExtent;
  const toPx = (value: number | undefined): number =>
    Number.isFinite(value) ? Math.max(0, value as number) : 0;
  return {
    leftPx: toPx(effect?.leftPx),
    topPx: toPx(effect?.topPx),
    rightPx: toPx(effect?.rightPx),
    bottomPx: toPx(effect?.bottomPx),
  };
}
