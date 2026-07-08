// Render VNode tree to static HTML string.
// Upstream editor.tsx: 4902 — renderStaticHtml using renderToStaticMarkup.
//
// Vue equivalent: uses a temporary DOM container and renders into it
// via the Vue renderer. Returns outerHTML of the container's first child.

import { type VNode, createApp, h } from "vue"

/**
 * Render a VNode tree to a static HTML string.
 *
 * Uses a temporary DOM mount to serialize VNodes. Suitable for
 * copy-to-clipboard and export scenarios where a static HTML
 * representation is needed.
 */
export function renderStaticHtml(node: VNode | VNode[]): string {
  if (!node) return ""

  const container = document.createElement("div")
  const nodes = Array.isArray(node) ? node : [node]

  // Create a wrapper component that renders the nodes
  const app = createApp({
    render() {
      return h("div", nodes)
    },
  })
  app.mount(container)
  const html = container.innerHTML
  app.unmount()

  return html
}
