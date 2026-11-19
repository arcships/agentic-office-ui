// OMML → linear text.
//
// The document model has no math layout engine, and an equation that renders as
// nothing is a data-loss bug (issue #33). The importer therefore stores a
// linearized form (`x^2`, `a/b`, `√(x)`) next to the original `m:oMath` markup:
// the renderer shows the linear form and the export writes the untouched OMML.

import { parseXmlElements, type XmlElement } from "./xml-elements";

const FRACTION_NAMES = new Set(["m:f"]);
const SCRIPT_NAMES = new Set(["m:sSup", "m:sSub", "m:sSubSup", "m:sPre"]);

function child(elements: readonly XmlElement[], name: string): XmlElement | undefined {
  return elements.find((element) => element.name === name);
}

/** Members of a group element (`m:e`, `m:num`, ...), concatenated. */
function memberText(element: XmlElement | undefined): string {
  return element ? renderElements(element.children) : "";
}

function wrapGroup(value: string): string {
  const trimmed = value.trim();
  if (trimmed.length <= 1) return trimmed;
  if (/^\(.*\)$/.test(trimmed)) return trimmed;
  return `(${trimmed})`;
}

/**
 * `m:chr` of a n-ary/group character, which Word stores as a hex code point in
 * `m:val` (e.g. `2211` for ∑).
 */
function operatorCharacter(element: XmlElement | undefined, fallback: string): string {
  const value = element?.attributes.match(/m:val="([^"]*)"/)?.[1];
  if (!value) return fallback;
  const codePoint = Number.parseInt(value, 16);
  return Number.isFinite(codePoint) && codePoint > 0
    ? String.fromCodePoint(codePoint)
    : fallback;
}

function renderElement(element: XmlElement): string {
  const children = element.children;

  if (element.name === "m:t") {
    return element.text;
  }

  if (element.name === "m:r") {
    return renderElements(children);
  }

  if (FRACTION_NAMES.has(element.name)) {
    return `${wrapGroup(memberText(child(children, "m:num")))}/${wrapGroup(
      memberText(child(children, "m:den")),
    )}`;
  }

  if (element.name === "m:sSup") {
    return `${memberText(child(children, "m:e"))}^${wrapGroup(
      memberText(child(children, "m:sup")),
    )}`;
  }

  if (element.name === "m:sSub") {
    return `${memberText(child(children, "m:e"))}_${wrapGroup(
      memberText(child(children, "m:sub")),
    )}`;
  }

  if (element.name === "m:sSubSup" || element.name === "m:sPre") {
    const base = memberText(child(children, "m:e"));
    const sub = memberText(child(children, "m:sub"));
    const sup = memberText(child(children, "m:sup"));
    return `${base}_${wrapGroup(sub)}^${wrapGroup(sup)}`;
  }

  if (element.name === "m:rad") {
    const degree = memberText(child(children, "m:deg")).trim();
    const radicand = memberText(child(children, "m:e")).trim();
    return degree
      ? `√[${degree}](${radicand})`
      : `√(${radicand})`;
  }

  if (element.name === "m:d") {
    const properties = child(children, "m:dPr");
    const begin = operatorCharacter(
      properties?.children.find((entry) => entry.name === "m:begChr"),
      "(",
    );
    const end = operatorCharacter(
      properties?.children.find((entry) => entry.name === "m:endChr"),
      ")",
    );
    return `${begin}${memberText(child(children, "m:e"))}${end}`;
  }

  if (element.name === "m:nary") {
    const properties = child(children, "m:naryPr");
    const character = operatorCharacter(
      properties?.children.find((entry) => entry.name === "m:chr"),
      "∑",
    );
    const sub = memberText(child(children, "m:sub")).trim();
    const sup = memberText(child(children, "m:sup")).trim();
    const body = memberText(child(children, "m:e")).trim();
    const limits = `${sub ? `_{${sub}}` : ""}${sup ? `^{${sup}}` : ""}`;
    return `${character}${limits}(${body})`;
  }

  if (element.name === "m:func") {
    return `${memberText(child(children, "m:fName"))}${wrapGroup(
      memberText(child(children, "m:e")),
    )}`;
  }

  if (element.name === "m:limLow") {
    return `${memberText(child(children, "m:e"))}_{${memberText(
      child(children, "m:lim"),
    )}}`;
  }

  if (element.name === "m:limUpp") {
    return `${memberText(child(children, "m:e"))}^{${memberText(
      child(children, "m:lim"),
    )}}`;
  }

  if (element.name === "m:eqArr") {
    return children
      .filter((entry) => entry.name === "m:e")
      .map((entry) => memberText(entry).trim())
      .join(", ");
  }

  if (element.name === "m:bar" || element.name === "m:acc") {
    return memberText(child(children, "m:e"));
  }

  if (SCRIPT_NAMES.has(element.name)) {
    return renderElements(children);
  }

  return renderElements(children);
}

function renderElements(elements: readonly XmlElement[]): string {
  return elements.map(renderElement).join("");
}

/** Strips the markup and keeps the text of an equation as one line. */
export function linearizeOmml(xml: string | undefined): string {
  if (!xml) return "";
  const elements = parseXmlElements(xml);
  return renderElements(elements)
    .replace(/\s+/g, " ")
    .trim();
}
