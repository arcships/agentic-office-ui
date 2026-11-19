// Small DOM helpers shared by the OpenDocument text reader.
//
// The DOCX import Worker has no `DOMParser`; `@xmldom/xmldom` provides one and
// is used whenever the host does not already expose the browser parser. The
// helpers below only use DOM Level 2 methods that both implementations provide.

import { DOMParser as XmldomDOMParser } from "@xmldom/xmldom"

export const ODF_NAMESPACE = Object.freeze({
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
})

export const ELEMENT_NODE = 1
export const TEXT_NODE = 3

type DomParserConstructor = new () => {
  parseFromString(source: string, mimeType: string): unknown
}

/**
 * Parses XML with the host parser when one exists and `@xmldom/xmldom`
 * otherwise. The returned document is cast to the DOM interface; both
 * implementations implement the subset this reader uses.
 */
export function parseXmlDocument(xml: string): Document {
  const parser = typeof DOMParser !== "undefined"
    ? (DOMParser as unknown as DomParserConstructor)
    : (XmldomDOMParser as unknown as DomParserConstructor)
  return new parser().parseFromString(xml, "application/xml") as Document
}

export function localName(node: { localName?: string | null; nodeName?: string }): string {
  const name = node.localName ?? node.nodeName ?? ""
  const separator = name.lastIndexOf(":")
  return separator >= 0 ? name.slice(separator + 1) : name
}

export function isElement(node: Node): node is Element {
  return node.nodeType === ELEMENT_NODE
}

/** Direct element children, optionally filtered by local name. */
export function childElements(parent: Node, name?: string): Element[] {
  const result: Element[] = []
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (!isElement(child)) continue
    if (name !== undefined && localName(child) !== name) continue
    result.push(child)
  }
  return result
}

export function firstChildElement(parent: Node, name: string): Element | undefined {
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (isElement(child) && localName(child) === name) return child
  }
  return undefined
}

/** All descendant elements with the given local name. */
export function descendantElements(parent: Node, name: string): Element[] {
  const result: Element[] = []
  const visit = (node: Node): void => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (!isElement(child)) continue
      if (localName(child) === name) result.push(child)
      visit(child)
    }
  }
  visit(parent)
  return result
}

export function firstDescendantElement(parent: Node, name: string): Element | undefined {
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (!isElement(child)) continue
    if (localName(child) === name) return child
    const nested = firstDescendantElement(child, name)
    if (nested) return nested
  }
  return undefined
}

/**
 * Attribute lookup that accepts both a namespace-aware parser and a plain
 * prefixed attribute (`text:style-name`).
 */
export function attribute(
  element: Element,
  name: string,
  namespace?: string,
): string | undefined {
  if (namespace !== undefined) {
    const namespaced = element.getAttributeNS(namespace, name)
    if (namespaced !== null && namespaced !== "") return namespaced
  }
  const prefixed = element.getAttribute(name)
  if (prefixed !== null && prefixed !== "") return prefixed
  const attributes = element.attributes
  for (let index = 0; index < attributes.length; index += 1) {
    const item = attributes.item(index)
    if (item && localName(item) === name && item.value !== "") return item.value
  }
  return undefined
}

export function textContent(node: Node): string {
  let value = ""
  const visit = (current: Node): void => {
    for (let child = current.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === TEXT_NODE) {
        value += child.nodeValue ?? ""
      } else if (isElement(child)) {
        visit(child)
      }
    }
  }
  visit(node)
  return value
}

/** Parses an ODF length (in, cm, mm, pt, pc, px) into CSS pixels. */
export function lengthToPx(value: string | undefined): number | undefined {
  if (!value) return undefined
  const match = /^\s*(-?\d+(?:\.\d+)?)\s*(in|cm|mm|pt|pc|px)?\s*$/i.exec(value)
  if (!match) return undefined
  const amount = Number(match[1])
  if (!Number.isFinite(amount)) return undefined
  switch ((match[2] ?? "px").toLowerCase()) {
    case "in": return amount * 96
    case "cm": return (amount / 2.54) * 96
    case "mm": return (amount / 25.4) * 96
    case "pt": return (amount / 72) * 96
    case "pc": return amount * 16
    default: return amount
  }
}

/** Parses an ODF length into twips (1/1440 inch). */
export function lengthToTwips(value: string | undefined): number | undefined {
  const pixels = lengthToPx(value)
  return pixels === undefined ? undefined : Math.round(pixels * 15)
}

/** Parses an ODF point size (`12pt`) into points. */
export function lengthToPt(value: string | undefined): number | undefined {
  if (!value) return undefined
  const match = /^\s*(-?\d+(?:\.\d+)?)\s*(pt|px|in|cm|mm|pc)?\s*$/i.exec(value)
  if (!match) return undefined
  const amount = Number(match[1])
  if (!Number.isFinite(amount)) return undefined
  switch ((match[2] ?? "pt").toLowerCase()) {
    case "px": return (amount / 96) * 72
    case "in": return amount * 72
    case "cm": return (amount / 2.54) * 72
    case "mm": return (amount / 25.4) * 72
    case "pc": return amount * 12
    default: return amount
  }
}

export function percentage(value: string | undefined): number | undefined {
  if (!value) return undefined
  const match = /^\s*(-?\d+(?:\.\d+)?)\s*%\s*$/.exec(value)
  if (!match) return undefined
  const amount = Number(match[1])
  return Number.isFinite(amount) ? amount : undefined
}

/** Decodes standard base64 without assuming a browser or Node global. */
export function base64ToBytes(value: string): Uint8Array | undefined {
  const cleaned = value.replace(/\s+/g, "")
  if (!cleaned) return undefined
  if (typeof atob === "function") {
    try {
      const binary = atob(cleaned)
      const bytes = new Uint8Array(binary.length)
      for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
      return bytes
    } catch {
      return undefined
    }
  }
  const buffer = (globalThis as {
    Buffer?: { from(value: string, encoding: string): Uint8Array }
  }).Buffer
  if (!buffer) return undefined
  try {
    return new Uint8Array(buffer.from(cleaned, "base64"))
  } catch {
    return undefined
  }
}

/** Decodes `_20_` style-name escapes used by ODF for spaces and punctuation. */
export function decodeStyleName(value: string): string {
  return value.replace(/_([0-9A-Fa-f]{2})_/g, (_match, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16))
  )
}
