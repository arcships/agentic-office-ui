// Render a VNode tree to a static HTML string.
//
// This intentionally does not mount the supplied VNodes. Paragraph and table
// hosts reuse those VNodes in Vue's live tree; mounting them in a temporary
// app mutates their DOM references and breaks the next live update.

import { Comment, Fragment, Text, type VNode } from "vue"

const BOOLEAN_ATTRIBUTES = new Set([
  "allowfullscreen",
  "async",
  "autofocus",
  "autoplay",
  "checked",
  "controls",
  "default",
  "defer",
  "disabled",
  "formnovalidate",
  "hidden",
  "ismap",
  "itemscope",
  "loop",
  "multiple",
  "muted",
  "nomodule",
  "novalidate",
  "open",
  "readonly",
  "required",
  "reversed",
  "selected",
])

const UNITLESS_STYLE_PROPERTIES = new Set([
  "animation-iteration-count",
  "aspect-ratio",
  "border-image-outset",
  "border-image-slice",
  "border-image-width",
  "column-count",
  "columns",
  "flex",
  "flex-grow",
  "flex-shrink",
  "font-weight",
  "grid-column",
  "grid-column-end",
  "grid-column-start",
  "grid-row",
  "grid-row-end",
  "grid-row-start",
  "line-height",
  "opacity",
  "order",
  "orphans",
  "scale",
  "tab-size",
  "widows",
  "z-index",
  "zoom",
])

/**
 * Render native VNodes to a static HTML string without attaching a Vue app.
 *
 * The DOCX run renderer only produces native elements. Components are handled
 * defensively by serializing their children, which is safer than trying to
 * evaluate a component during an editor update.
 */
export function renderStaticHtml(node: VNode | VNode[]): string {
  if (!node) return ""

  if (typeof document === "undefined") {
    return renderStaticString(node)
  }

  const container = document.createElement("div")
  appendStaticNode(container, node)
  return container.innerHTML
}

function renderStaticString(value: unknown): string {
  if (value === null || value === undefined || value === false || value === true) {
    return ""
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "bigint") {
    return escapeHtml(String(value))
  }
  if (Array.isArray(value)) {
    return value.map(renderStaticString).join("")
  }
  if (!isVNode(value) || value.type === Comment) return ""
  if (value.type === Text || value.type === Fragment || typeof value.type !== "string") {
    return renderStaticString(value.children)
  }

  const attributes = staticAttributeString(value.props as Record<string, unknown> | null)
  return `<${value.type}${attributes}>${renderStaticString(value.children)}</${value.type}>`
}

function staticAttributeString(props: Record<string, unknown> | null): string {
  if (!props) return ""
  const attributes: string[] = []
  for (const [key, value] of Object.entries(props)) {
    if (
      value === null ||
      value === undefined ||
      value === false ||
      key === "key" ||
      key === "ref" ||
      key === "children" ||
      key === "innerHTML" ||
      key === "textContent" ||
      /^on[A-Z]/.test(key)
    ) {
      continue
    }
    if (key === "style") {
      const style = staticStyleString(value)
      if (style) attributes.push(`style="${escapeHtml(style)}"`)
      continue
    }
    const name = key === "className" ? "class" : normalizeAttributeName(key)
    const rendered = key === "class" ? normalizeClass(value) : value === true ? "" : String(value)
    if (rendered !== "") attributes.push(`${name}="${escapeHtml(rendered)}"`)
    else if (value === true) attributes.push(name)
  }
  return attributes.length > 0 ? ` ${attributes.join(" ")}` : ""
}

function staticStyleString(value: unknown): string {
  if (typeof value === "string") return value
  if (Array.isArray(value)) return value.map(staticStyleString).filter(Boolean).join(";")
  if (typeof value !== "object" || value === null) return ""
  return Object.entries(value)
    .filter(([, propertyValue]) => propertyValue !== null && propertyValue !== undefined)
    .map(([name, propertyValue]) =>
      `${toKebabCase(name)}:${staticStyleValue(name, propertyValue)}`
    )
    .join(";")
}

function staticStyleValue(name: string, value: unknown): string {
  const propertyName = toKebabCase(name)
  if (
    typeof value === "number" &&
    value !== 0 &&
    !propertyName.startsWith("--") &&
    !UNITLESS_STYLE_PROPERTIES.has(propertyName)
  ) {
    return `${value}px`
  }
  return String(value)
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
}

function appendStaticNode(parent: Node, value: unknown): void {
  if (value === null || value === undefined || value === false || value === true) {
    return
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "bigint") {
    parent.appendChild(document.createTextNode(String(value)))
    return
  }

  if (Array.isArray(value)) {
    value.forEach((child) => appendStaticNode(parent, child))
    return
  }

  if (!isVNode(value)) return

  if (value.type === Comment) return

  if (value.type === Text || value.type === Fragment || typeof value.type !== "string") {
    appendStaticNode(parent, value.children)
    return
  }

  const element = document.createElement(value.type)
  applyStaticProps(element, value.props as Record<string, unknown> | null)
  appendStaticNode(element, value.children)
  parent.appendChild(element)
}

function isVNode(value: unknown): value is VNode {
  return typeof value === "object" && value !== null && "type" in value
}

function applyStaticProps(
  element: HTMLElement,
  props: Record<string, unknown> | null
): void {
  if (!props) return

  for (const [key, value] of Object.entries(props)) {
    if (
      key === "key" ||
      key === "ref" ||
      key === "children" ||
      key === "innerHTML" ||
      key === "textContent" ||
      /^on[A-Z]/.test(key)
    ) {
      continue
    }

    if (key === "class" || key === "className") {
      const className = normalizeClass(value)
      if (className) element.setAttribute("class", className)
      continue
    }

    if (key === "style") {
      applyStaticStyle(element.style, value)
      continue
    }

    if (value === null || value === undefined) continue

    const attributeName = normalizeAttributeName(key)
    if (value === true) {
      element.setAttribute(
        attributeName,
        BOOLEAN_ATTRIBUTES.has(attributeName) ? "" : "true"
      )
      continue
    }

    if (value === false) {
      if (attributeName === "draggable" || attributeName === "contenteditable") {
        element.setAttribute(attributeName, "false")
      }
      continue
    }

    element.setAttribute(attributeName, String(value))
  }
}

function normalizeClass(value: unknown): string {
  if (typeof value === "string") return value
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeClass(entry)).filter(Boolean).join(" ")
  }
  if (typeof value === "object" && value !== null) {
    return Object.entries(value)
      .filter(([, enabled]) => Boolean(enabled))
      .map(([name]) => name)
      .join(" ")
  }
  return ""
}

function applyStaticStyle(style: CSSStyleDeclaration, value: unknown): void {
  if (typeof value === "string") {
    style.cssText = value
    return
  }

  if (Array.isArray(value)) {
    value.forEach((entry) => applyStaticStyle(style, entry))
    return
  }

  if (typeof value !== "object" || value === null) return

  for (const [name, propertyValue] of Object.entries(value)) {
    if (Array.isArray(propertyValue)) {
      propertyValue.forEach((entry) => setStaticStyle(style, name, entry))
    } else {
      setStaticStyle(style, name, propertyValue)
    }
  }
}

function setStaticStyle(
  style: CSSStyleDeclaration,
  name: string,
  value: unknown
): void {
  if (value === null || value === undefined) return

  const text = staticStyleValue(name, value)
  const important = text.match(/\s*!important$/)
  if (important) {
    style.setProperty(
      toKebabCase(name),
      text.slice(0, important.index).trimEnd(),
      "important"
    )
    return
  }

  if (name.startsWith("--") || name.includes("-")) {
    style.setProperty(name, text)
    return
  }

  ;(style as unknown as Record<string, string>)[name] = text
}

function normalizeAttributeName(name: string): string {
  if (name === "htmlFor") return "for"
  if (name === "httpEquiv") return "http-equiv"
  return name
}

function toKebabCase(value: string): string {
  return value.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`)
}
