// Minimal DOM shim for block/word geometry tests.
//
// The geometry resolver only needs element trees, a text tree walker, range
// client rects and computed sizes. This shim implements exactly that surface so
// the unit and component suites can verify page-space boxes without a browser.

function normalizeRect(rect) {
  return {
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
    right: rect.left + rect.width,
    bottom: rect.top + rect.height,
  };
}

function parseSimpleSelector(token) {
  const tag = token.match(/^[a-zA-Z][\w-]*/)?.[0]?.toLowerCase();
  const attributes = [...token.matchAll(/\[([^\]=~^$*|]+)(?:([~^$*|]?=)['"]?([^'"\]]*)['"]?)?\]/g)]
    .map((match) => ({
      name: match[1].trim().toLowerCase(),
      operator: match[2],
      value: match[3],
    }));
  return { tag, attributes };
}

function attributeValue(element, name) {
  if (Object.prototype.hasOwnProperty.call(element.attributes, name)) {
    return element.attributes[name];
  }
  if (name.startsWith("data-")) {
    const key = name
      .slice(5)
      .replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
    if (Object.prototype.hasOwnProperty.call(element.dataset, key)) {
      return String(element.dataset[key]);
    }
  }
  if (name === "role") return element.attributes.role;
  return undefined;
}

function matchesSimple(element, token) {
  const { tag, attributes } = parseSimpleSelector(token);
  if (tag && element.tagName.toLowerCase() !== tag) return false;
  return attributes.every(({ name, operator, value }) => {
    const actual = attributeValue(element, name);
    if (actual === undefined || actual === null) return false;
    if (!operator) return true;
    if (operator === "=") return String(actual) === String(value);
    if (operator === "~=") return String(actual).split(/\s+/).includes(String(value));
    return true;
  });
}

function matchesSelector(element, selector) {
  return selector
    .split(",")
    .some((part) => {
      const tokens = part.trim().split(/\s+/).filter(Boolean);
      let current = element;
      for (let index = tokens.length - 1; index >= 0; index -= 1) {
        if (!current || !matchesSimple(current, tokens[index])) return false;
        if (index > 0) current = current.parentElement;
      }
      return true;
    });
}

function* descendants(element) {
  for (const child of element.children) {
    if (child.nodeType !== 1) continue;
    yield child;
    yield* descendants(child);
  }
}

function* textDescendants(element) {
  for (const child of element.children) {
    if (child.nodeType === 3) yield child;
    else if (child.nodeType === 1) yield* textDescendants(child);
  }
}

function rectUnion(rects) {
  if (rects.length === 0) return { left: 0, top: 0, width: 0, height: 0, right: 0, bottom: 0 };
  const left = Math.min(...rects.map((rect) => rect.left));
  const top = Math.min(...rects.map((rect) => rect.top));
  const right = Math.max(...rects.map((rect) => rect.right));
  const bottom = Math.max(...rects.map((rect) => rect.bottom));
  return { left, top, width: right - left, height: bottom - top, right, bottom };
}

function rectsForText(textNode, start, end) {
  const chars = textNode.charRects ?? [];
  const covered = chars.slice(start, end).filter(Boolean).map(normalizeRect);
  if (covered.length > 0) {
    // One box per visual line, mirroring Range.getClientRects().
    const lines = new Map();
    for (const rect of covered) {
      const line = lines.get(rect.top) ?? [];
      line.push(rect);
      lines.set(rect.top, line);
    }
    return [...lines.values()].map(rectUnion);
  }
  const rect = textNode.rect;
  if (!rect) return [];
  return [normalizeRect(rect)];
}

function createTextNode(doc, data, parent = null) {
  return {
    nodeType: 3,
    data,
    charRects: [],
    rect: undefined,
    parentElement: parent,
    ownerDocument: doc,
    get length() {
      return this.data.length;
    },
  };
}

function createDocument() {
  const doc = {
    fonts: {
      addEventListener() {},
      removeEventListener() {},
      ready: Promise.resolve(),
    },
    body: null,
    createElement(tagName) {
      return createElement(doc, tagName);
    },
    createTextNode(data) {
      return createTextNode(doc, data);
    },
    createTreeWalker(root) {
      const nodes = [...textDescendants(root)];
      let index = -1;
      return {
        get currentNode() {
          return nodes[index];
        },
        nextNode() {
          index += 1;
          return index < nodes.length ? nodes[index] : null;
        },
      };
    },
    createRange() {
      let node;
      let start = 0;
      let end = 0;
      let whole = false;
      const range = {
        setStart(textNode, offset) {
          node = textNode;
          start = offset;
        },
        setEnd(textNode, offset) {
          node = textNode;
          end = offset;
        },
        selectNodeContents(textNode) {
          node = textNode;
          start = 0;
          end = textNode.length;
          whole = true;
        },
        getClientRects() {
          if (!node) return [];
          return rectsForText(node, whole ? 0 : start, whole ? node.length : end);
        },
        getBoundingClientRect() {
          return rectUnion(this.getClientRects());
        },
        _whole: () => whole,
      };
      return range;
    },
  };
  doc.body = createElement(doc, "body");
  return doc;
}

function createElement(doc, tagName) {
  const element = {
    nodeType: 1,
    tagName: String(tagName).toUpperCase(),
    ownerDocument: doc,
    dataset: {},
    attributes: {},
    style: {
      setProperty(name, value) {
        this[name] = value;
      },
      removeProperty(name) {
        delete this[name];
      },
    },
    computedStyle: {},
    children: [],
    parentElement: null,
    rect: undefined,
    scrollLeft: 0,
    offsetWidth: 0,
    offsetHeight: 0,
    getBoundingClientRect() {
      return element.rect
        ? normalizeRect(element.rect)
        : { left: 0, top: 0, width: 0, height: 0, right: 0, bottom: 0 };
    },
    setAttribute(name, value) {
      element.attributes[name] = String(value);
      if (name.startsWith("data-")) {
        const key = name
          .slice(5)
          .replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
        element.dataset[key] = String(value);
      }
    },
    getAttribute(name) {
      return attributeValue(element, name) ?? null;
    },
    appendChild(child) {
      child.parentElement = element;
      child.parentNode = element;
      element.children.push(child);
      return child;
    },
    append(...nodes) {
      for (const node of nodes) element.appendChild(node);
    },
    remove() {
      if (!element.parentElement) return;
      const index = element.parentElement.children.indexOf(element);
      if (index >= 0) element.parentElement.children.splice(index, 1);
      element.parentElement = null;
    },
    querySelectorAll(selector) {
      return [...descendants(element)].filter((child) => matchesSelector(child, selector));
    },
    querySelector(selector) {
      return [...descendants(element)].find((child) => matchesSelector(child, selector)) ?? null;
    },
    matches(selector) {
      return element.nodeType === 1 && matchesSelector(element, selector);
    },
    closest(selector) {
      let current = element;
      while (current) {
        if (current.matches?.(selector)) return current;
        current = current.parentElement;
      }
      return null;
    },
  };
  return element;
}

/** Sets an element's client rect and returns the element. */
export function setRect(element, rect) {
  element.rect = normalizeRect(rect);
  return element;
}

/** Appends a text node with optional per-character client rects. */
export function appendText(element, data, charRects = []) {
  const text = createTextNode(element.ownerDocument, data, element);
  text.charRects = charRects;
  element.children.push(text);
  return text;
}

/**
 * Creates a fake document/root pair and installs the browser globals the
 * geometry resolver reads. Returns `restore()` to put the previous globals back.
 */
export function installGeometryDom(options = {}) {
  const doc = createDocument();
  const root = createElement(doc, "div");
  const previous = {
    document: globalThis.document,
    getComputedStyle: globalThis.getComputedStyle,
    CSS: globalThis.CSS,
    window: globalThis.window,
  };
  const defaultComputedStyle = {
    display: "block",
    visibility: "visible",
    overflowX: "visible",
    overflowY: "visible",
    width: "",
    height: "",
    boxSizing: "content-box",
    direction: "ltr",
    textAlign: "left",
    getPropertyValue(name) {
      return options.fontStyle?.[name] ?? "";
    },
  };
  globalThis.document = doc;
  globalThis.getComputedStyle = (element) => ({
    ...defaultComputedStyle,
    ...(element?.computedStyle ?? {}),
  });
  globalThis.CSS = {
    escape: (value) => String(value).replace(/["\\]/g, "\\$&"),
  };
  if (options.window !== false) {
    globalThis.window = {
      addEventListener() {},
      removeEventListener() {},
      requestAnimationFrame(callback) {
        callback(0);
        return 1;
      },
      cancelAnimationFrame() {},
    };
  }
  return {
    doc,
    root,
    createElement: (tagName) => createElement(doc, tagName),
    appendText,
    setRect,
    restore() {
      if (previous.document === undefined) delete globalThis.document;
      else globalThis.document = previous.document;
      if (previous.getComputedStyle === undefined) delete globalThis.getComputedStyle;
      else globalThis.getComputedStyle = previous.getComputedStyle;
      if (previous.CSS === undefined) delete globalThis.CSS;
      else globalThis.CSS = previous.CSS;
      if (previous.window === undefined) delete globalThis.window;
      else globalThis.window = previous.window;
    },
  };
}
