// clipboard-html-parse.ts — `text/html` clipboard parsing for DOCX.
//
// Word and browsers put the same tags on the clipboard but disagree on the
// surrounding markup (Word uses `mso-*` styles, `<font>` and paragraph-level
// numbering; browsers use inline CSS), so the parser walks the markup with a
// small stack machine instead of relying on a DOM. That keeps it usable in a
// Worker and testable without a browser.
//
// External input rules: element content that can execute is dropped, only
// whitelisted inline declarations are read, links must pass the caller's
// resolver, and paragraph/run caps bound the work.

import type { HeadingLevel, ParagraphAlignment, ParagraphStyle, TextStyle } from "../engine/types"
import type {
  DocxClipboardListInfo,
  DocxClipboardParagraph,
  DocxClipboardParseOptions,
  DocxClipboardRun,
  ParsedInlineStyle
} from "./clipboard-html"
import {
  HEADING_FONT_SIZES_PT,
  MAX_DOCX_CLIPBOARD_HTML_CHARS,
  MAX_DOCX_CLIPBOARD_PARAGRAPHS,
  MAX_DOCX_CLIPBOARD_RUNS,
  collapseClipboardWhitespace,
  decodeClipboardHtmlEntities,
  mergeAdjacentRuns,
  mergeStyles,
  parseClipboardInlineStyle,
  sanitizeClipboardColor,
  sanitizeClipboardFontFamily,
  sanitizeClipboardParagraphs,
  trimRunEdges
} from "./clipboard-html"

const VOID_TAGS = new Set([
  "area",
  "base",
  "br",
  "col",
  "embed",
  "hr",
  "img",
  "input",
  "link",
  "meta",
  "param",
  "source",
  "track",
  "wbr"
])

const SKIP_CONTENT_TAGS = new Set([
  "applet",
  "audio",
  "base",
  "canvas",
  "embed",
  "form",
  "frame",
  "frameset",
  "head",
  "iframe",
  "link",
  "math",
  "meta",
  "noscript",
  "object",
  "script",
  "style",
  "svg",
  "template",
  "title",
  "video",
  "xml"
])

const BLOCK_TAGS = new Set([
  "address",
  "article",
  "aside",
  "blockquote",
  "body",
  "caption",
  "center",
  "dd",
  "div",
  "dl",
  "dt",
  "figcaption",
  "figure",
  "footer",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "header",
  "html",
  "li",
  "main",
  "nav",
  "ol",
  "p",
  "pre",
  "section",
  "table",
  "tbody",
  "td",
  "tfoot",
  "th",
  "thead",
  "tr",
  "ul"
])


const LEGACY_FONT_SIZES_PT: Record<string, number> = {
  "1": 8,
  "2": 10,
  "3": 12,
  "4": 14,
  "5": 18,
  "6": 24,
  "7": 36
}


// ── HTML parsing ───────────────────────────────────────────────────

interface HtmlFrame {
  style: TextStyle
  link?: string
  list?: DocxClipboardListInfo
  headingLevel?: HeadingLevel
  headingStyle?: TextStyle
  align?: ParagraphAlignment
  msoList?: DocxClipboardListInfo
  /** Word's bullet/number marker run (`mso-list:Ignore`); its text is a duplicate. */
  ignoreText?: boolean
}

interface HtmlAttribute {
  name: string
  value: string
}

function parseAttributes(source: string): HtmlAttribute[] {
  const attributes: HtmlAttribute[] = []
  const pattern = /([A-Za-z_:][\w:.-]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>/]+)))?/g
  let match: RegExpExecArray | null
  while ((match = pattern.exec(source))) {
    const name = match[1]?.toLowerCase()
    if (!name) continue
    const raw = match[2] ?? match[3] ?? match[4] ?? ""
    attributes.push({ name, value: decodeClipboardHtmlEntities(raw) })
  }
  return attributes
}

function attributeOf(attributes: readonly HtmlAttribute[], name: string): string | undefined {
  return attributes.find((attribute) => attribute.name === name)?.value
}

function headingLevelFromTag(tag: string): HeadingLevel | undefined {
  const match = /^h([1-6])$/.exec(tag)
  if (!match) return undefined
  return Number(match[1]) as HeadingLevel
}

function listFrame(
  tag: string,
  frame: HtmlFrame
): HtmlFrame | undefined {
  if (tag !== "ul" && tag !== "ol" && tag !== "menu") return undefined
  const level = Math.max(0, Math.min(8, (frame.list?.level ?? -1) + 1))
  return {
    ...frame,
    list: { type: tag === "ol" ? "ordered" : "unordered", level }
  }
}

function inlineTagStyle(tag: string): TextStyle | undefined {
  switch (tag) {
    case "b":
    case "strong":
      return { bold: true }
    case "i":
    case "em":
      return { italic: true }
    case "u":
    case "ins":
      return { underline: true }
    case "s":
    case "strike":
    case "del":
      return { strike: true }
    case "sup":
      return { verticalAlign: "superscript" }
    case "sub":
      return { verticalAlign: "subscript" }
    case "code":
    case "kbd":
    case "samp":
    case "tt":
      return { fontFamily: "monospace" }
    case "mark":
      return { backgroundColor: "#ffff00" }
    default:
      return undefined
  }
}

function legacyFontStyle(attributes: readonly HtmlAttribute[]): TextStyle | undefined {
  const style: TextStyle = {}
  const color = sanitizeClipboardColor(attributeOf(attributes, "color"))
  if (color) style.color = color
  const face = sanitizeClipboardFontFamily(attributeOf(attributes, "face"))
  if (face) style.fontFamily = face
  const size = attributeOf(attributes, "size")?.trim()
  if (size) {
    const pt = LEGACY_FONT_SIZES_PT[size.replace(/^\+/, "")]
    if (pt !== undefined) style.fontSizePt = pt
  }
  return Object.keys(style).length > 0 ? style : undefined
}

function blockAlign(attributes: readonly HtmlAttribute[], inline: ParsedInlineStyle): ParagraphAlignment | undefined {
  if (inline.align) return inline.align
  const align = attributeOf(attributes, "align")?.trim().toLowerCase()
  if (align === "left" || align === "right" || align === "center" || align === "justify") {
    return align
  }
  return undefined
}

/**
 * Parses the `text/html` clipboard flavour into clipboard paragraphs.
 *
 * Element content that can execute is dropped, only whitelisted inline styles
 * are read, and links survive only when `sanitizeLink` accepts them.
 */
export function parseDocxClipboardHtml(
  html: string,
  options?: DocxClipboardParseOptions
): DocxClipboardParagraph[] {
  if (!html) return []
  const maxChars = Math.max(1, options?.maxHtmlChars ?? MAX_DOCX_CLIPBOARD_HTML_CHARS)
  const source = html.length > maxChars ? html.slice(0, maxChars) : html
  const maxParagraphs = Math.max(1, options?.maxParagraphs ?? MAX_DOCX_CLIPBOARD_PARAGRAPHS)
  const maxRuns = Math.max(1, options?.maxRuns ?? MAX_DOCX_CLIPBOARD_RUNS)

  const paragraphs: DocxClipboardParagraph[] = []
  const stack: Array<{ name: string; frame: HtmlFrame }> = [
    { name: "#document", frame: { style: {} } }
  ]
  let runs: DocxClipboardRun[] = []
  let opened = false
  let blockStyle: ParagraphStyle | undefined
  let blockList: DocxClipboardListInfo | undefined
  let runCount = 0
  let skipDepth = 0

  const topFrame = (): HtmlFrame => stack[stack.length - 1]?.frame ?? { style: {} }

  function flushBlock(): void {
    if (!opened && runs.length === 0) return
    if (paragraphs.length >= maxParagraphs) {
      opened = false
      runs = []
      blockStyle = undefined
      blockList = undefined
      return
    }
    const paragraph: DocxClipboardParagraph = { runs: trimRunEdges(mergeAdjacentRuns(runs)) }
    if (blockStyle) paragraph.style = blockStyle
    if (blockList) paragraph.list = blockList
    paragraphs.push(paragraph)
    opened = false
    runs = []
    blockStyle = undefined
    blockList = undefined
  }

  function openBlock(frame: HtmlFrame): void {
    if (opened) return
    opened = true
    blockStyle = undefined
    blockList = undefined
    if (frame.headingLevel) blockStyle = { ...(blockStyle ?? {}), headingLevel: frame.headingLevel }
    if (frame.align) blockStyle = { ...(blockStyle ?? {}), align: frame.align }
    if (frame.list) blockList = { ...frame.list }
    else if (frame.msoList) blockList = { ...frame.msoList }
  }

  function appendText(text: string, frame: HtmlFrame): void {
    if (!text || frame.ignoreText) return
    if (!opened) openBlock(frame)
    const style = mergeStyles(frame.headingStyle, frame.style)
    if (!style && !frame.link) {
      runs.push({ text })
      runCount += 1
      return
    }
    runs.push({
      text,
      style: style ? { ...style } : undefined,
      link: frame.link
    })
    runCount += 1
  }

  const pattern = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<![^>]*>|<\/([A-Za-z][\w:-]*)\s*>|<([A-Za-z][\w:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*)(\/?)>/g
  let cursor = 0
  let match: RegExpExecArray | null

  while ((match = pattern.exec(source))) {
    if (runCount >= maxRuns) break
    if (match.index > cursor) {
      const text = collapseClipboardWhitespace(
        decodeClipboardHtmlEntities(source.slice(cursor, match.index))
      )
      if (skipDepth === 0 && text) {
        if (text.trim() || opened) appendText(text, topFrame())
      }
    }
    cursor = pattern.lastIndex

    const endTag = match[1]?.toLowerCase()
    const startTag = match[2]?.toLowerCase()

    if (endTag) {
      if (skipDepth > 0) {
        const skipIndex = findLastStackIndex(stack, endTag)
        if (skipIndex > 0 && isSkipTag(stack[skipIndex].name)) {
          stack.length = skipIndex
          skipDepth = countSkipFrames(stack)
        }
        continue
      }
      if (endTag === "br") continue
      const index = findLastStackIndex(stack, endTag)
      if (index > 0) {
        if (BLOCK_TAGS.has(endTag)) flushBlock()
        stack.length = index
      }
      continue
    }

    if (!startTag) continue
    const attributeSource = match[3] ?? ""
    const selfClosing = match[4] === "/" || attributeSource.trimEnd().endsWith("/")
    const attributes = parseAttributes(attributeSource)
    const parent = topFrame()
    const inline = parseClipboardInlineStyle(attributeOf(attributes, "style"))

    if (skipDepth > 0) {
      if (!selfClosing && isSkipTag(startTag) && !VOID_TAGS.has(startTag)) {
        stack.push({ name: startTag, frame: { ...parent, style: { ...parent.style } } })
        skipDepth = countSkipFrames(stack)
      }
      continue
    }

    if (startTag === "br") {
      flushBlock()
      openBlock(parent)
      continue
    }

    if (startTag === "hr") {
      flushBlock()
      continue
    }

    if (isSkipTag(startTag)) {
      if (!selfClosing && !VOID_TAGS.has(startTag)) {
        stack.push({ name: startTag, frame: { ...parent, style: { ...parent.style } } })
        skipDepth = countSkipFrames(stack)
      }
      continue
    }

    // Media and controls carry no text worth pasting; `<img alt>` is not text.
    if (startTag === "img" || startTag === "input") continue

    if (BLOCK_TAGS.has(startTag)) {
      flushBlock()
    }

    const frame: HtmlFrame = {
      style: { ...parent.style },
      link: parent.link,
      list: parent.list,
      headingLevel: parent.headingLevel,
      headingStyle: parent.headingStyle,
      align: parent.align,
      msoList: parent.msoList,
      ignoreText: parent.ignoreText
    }
    if (inline.listIgnore) frame.ignoreText = true

    const inlineTag = inlineTagStyle(startTag)
    if (inlineTag) frame.style = mergeStyles(frame.style, inlineTag) ?? frame.style
    if (startTag === "font") {
      const legacy = legacyFontStyle(attributes)
      if (legacy) frame.style = mergeStyles(frame.style, legacy) ?? frame.style
    }
    if (inline.style && Object.keys(inline.style).length > 0) {
      frame.style = mergeStyles(frame.style, inline.style) ?? frame.style
    }
    if (startTag === "a") {
      const href = attributeOf(attributes, "href")?.trim()
      frame.link = href && options?.sanitizeLink ? options.sanitizeLink(href) : undefined
    }

    const listContext = listFrame(startTag, parent)
    if (listContext) {
      frame.list = listContext.list
    }
    if (startTag === "li") {
      frame.list = parent.list ?? { type: "unordered", level: 0 }
    }

    const heading = headingLevelFromTag(startTag)
    if (heading) {
      frame.headingLevel = heading
      frame.headingStyle = { bold: true, fontSizePt: HEADING_FONT_SIZES_PT[heading - 1] }
    }

    const align = BLOCK_TAGS.has(startTag) ? blockAlign(attributes, inline) : undefined
    if (align) frame.align = align

    if (inline.list) {
      frame.msoList = inline.list
      // `mso-list` paragraphs are the paragraph itself: open eagerly so list
      // info is attached even when the marker run is dropped.
      if (!opened) openBlock(frame)
    }

    if (!selfClosing && !VOID_TAGS.has(startTag)) {
      stack.push({ name: startTag, frame })
    }
  }

  if (cursor < source.length && skipDepth === 0) {
    const tail = collapseClipboardWhitespace(decodeClipboardHtmlEntities(source.slice(cursor)))
    if (tail) appendText(tail, topFrame())
  }
  flushBlock()

  return sanitizeClipboardParagraphs(paragraphs, options)
}


function isSkipTag(name: string): boolean {
  return SKIP_CONTENT_TAGS.has(name)
}

function findLastStackIndex(stack: Array<{ name: string }>, name: string): number {
  for (let index = stack.length - 1; index > 0; index -= 1) {
    if (stack[index].name === name) return index
  }
  return -1
}

function countSkipFrames(stack: Array<{ name: string }>): number {
  return stack.filter((entry) => isSkipTag(entry.name)).length
}
