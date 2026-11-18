// Minimal XML element scanner for the drawing-object / equation import passes.
//
// The DOCX parser lives in WASM and the engine only receives already-parsed
// parts, so the few OOXML fragments the model has to understand (OMML math,
// chart parts, drawing properties) are scanned here. This is deliberately not a
// general XML parser: it reads balanced tags and text content, rejects nothing,
// and keeps the raw inner markup so callers can still look at unknown children
// with a regex.

export interface XmlElement {
  /** Qualified tag name, e.g. `m:f` or `c:ser`. */
  name: string;
  /** Raw attribute text of the open tag (may be empty). */
  attributes: string;
  /** `<tag/>` with no children and no text. */
  selfClosing: boolean;
  children: XmlElement[];
  /** Direct text content (concatenated, not trimmed). */
  text: string;
  /** Raw markup between the open and the close tag. */
  innerXml: string;
}

const TAG_PATTERN = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<(\/?)([A-Za-z_][\w.:-]*)\b([^>]*?)(\/?)>/g;

interface OpenElement {
  element: XmlElement;
  /** Offset of the first character of the element's inner markup. */
  innerStart: number;
}

/**
 * Parses one XML fragment into its top-level elements. Unknown or unbalanced
 * markup is tolerated: an element without a close tag keeps its children.
 */
export function parseXmlElements(xml: string): XmlElement[] {
  const roots: XmlElement[] = [];
  const stack: OpenElement[] = [];
  const textByDepth = new Map<OpenElement, string>();
  let lastIndex = 0;

  const appendText = (value: string): void => {
    if (!value) return;
    const current = stack[stack.length - 1];
    if (!current) return;
    textByDepth.set(current, (textByDepth.get(current) ?? "") + value);
  };

  for (const match of xml.matchAll(TAG_PATTERN)) {
    const index = match.index ?? 0;
    appendText(xml.slice(lastIndex, index));
    lastIndex = index + match[0].length;

    if (match[1] !== undefined) {
      // CDATA: text content only.
      appendText(match[1]);
      continue;
    }
    const name = match[3];
    if (!name) continue; // comment or processing instruction

    const closing = match[2] === "/";
    const selfClosing = match[5] === "/";
    const current = stack[stack.length - 1];

    if (closing) {
      const open = stack.pop();
      if (!open) continue;
      open.element.innerXml = xml.slice(open.innerStart, index);
      open.element.text = textByDepth.get(open) ?? "";
      textByDepth.delete(open);
      const parent = stack[stack.length - 1];
      if (parent) parent.element.children.push(open.element);
      else roots.push(open.element);
      continue;
    }

    const element: XmlElement = {
      name,
      attributes: match[4] ?? "",
      selfClosing,
      children: [],
      text: "",
      innerXml: "",
    };
    if (selfClosing) {
      if (current) current.element.children.push(element);
      else roots.push(element);
      continue;
    }
    stack.push({ element, innerStart: lastIndex });
  }

  // Unclosed elements: close them at the end of the input, deepest first.
  const endIndex = xml.length;
  while (stack.length > 0) {
    const open = stack.pop();
    if (!open) break;
    open.element.innerXml = xml.slice(open.innerStart, endIndex);
    open.element.text = textByDepth.get(open) ?? "";
    textByDepth.delete(open);
    const parent = stack[stack.length - 1];
    if (parent) parent.element.children.push(open.element);
    else roots.push(open.element);
  }

  return roots;
}

/** All descendants (depth-first, document order) whose name matches. */
export function findXmlElements(
  elements: readonly XmlElement[],
  name: string,
): XmlElement[] {
  const found: XmlElement[] = [];
  const visit = (items: readonly XmlElement[]): void => {
    for (const element of items) {
      if (element.name === name) found.push(element);
      visit(element.children);
    }
  };
  visit(elements);
  return found;
}

/** First descendant with the given name, if any. */
export function findFirstXmlElement(
  elements: readonly XmlElement[],
  name: string,
): XmlElement | undefined {
  for (const element of elements) {
    if (element.name === name) return element;
    const nested = findFirstXmlElement(element.children, name);
    if (nested) return nested;
  }
  return undefined;
}

/** Attribute value of one element (qualified names are matched verbatim). */
export function xmlElementAttribute(
  element: XmlElement | undefined,
  name: string,
): string | undefined {
  if (!element) return undefined;
  const pattern = new RegExp(`(?:^|\\s)${escapeRegExp(name)}\\s*=\\s*"([^"]*)"`);
  return pattern.exec(element.attributes)?.[1];
}

export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Text content of every matching element, concatenated in document order. */
export function xmlElementsText(
  elements: readonly XmlElement[],
  name: string,
): string {
  return findXmlElements(elements, name)
    .map((element) => element.text)
    .join("");
}
