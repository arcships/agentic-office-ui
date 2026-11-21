// Copied from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn.
const canonicalUrl = "https://www.extend.ai/ui"

export const siteConfig = {
  name: "Extend UI",
  url: canonicalUrl,
  ogImage: `${canonicalUrl}/opengraph-image.png`,
  description:
    "Open source UI primitives for building document processing products with viewers, review surfaces, and validation workflows.",
  links: {
    github: "https://github.com/extend-hq/ui",
  },
  navItems: [
    {
      href: "/docs",
      label: "Docs",
    },
    {
      href: "/docs/components",
      label: "Components",
    },
    {
      href: "/blocks",
      label: "Blocks",
    },
  ],
}

export const META_THEME_COLORS = {
  light: "#ffffff",
  dark: "#09090b",
}
