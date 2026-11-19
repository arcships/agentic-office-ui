# Upstream snapshot

| Field | Value |
|---|---|
| Repository | https://github.com/extend-hq/ui |
| Commit | `70fe70c6f1e1eee53954c6923f3d9f583557a57e` |
| Commit date | 2026-11-17 |
| License | MIT (Copyright (c) 2026 CrowdView Inc, dba Extend; Portions Copyright (c) 2023 shadcn) |
| License copy | [`docs/upstream/licenses/extend-ui-MIT.txt`](../../../docs/upstream/licenses/extend-ui-MIT.txt) |
| Copied on | 2026-11-20 |
| Local tool versions | react 19.2.7, vite 6.4.3, tailwindcss 4.2.4, @tailwindcss/vite 4.2.4, @hugeicons/react 1.1.10, @hugeicons/core-free-icons 4.4.0 |

## Copied files

| Local path | Upstream path | Mode |
|---|---|---|
| `src/vendor/ui/button.tsx` | `apps/v4/components/ui/button.tsx` | adapted |
| `src/vendor/ui/spinner.tsx` | `apps/v4/registry/new-york-v4/ui/spinner.tsx` | adapted |
| `src/vendor/ui/card.tsx` | `apps/v4/registry/new-york-v4/ui/card.tsx` | adapted |
| `src/vendor/ui/separator.tsx` | `apps/v4/registry/new-york-v4/ui/separator.tsx` | adapted |
| `src/vendor/ui/select.tsx` | `apps/v4/registry/new-york-v4/ui/select.tsx` | adapted |
| `src/vendor/extend/file-thumbnail.tsx` | `apps/v4/components/extend/file-thumbnail.tsx` | copied |
| `src/vendor/extend/file-upload.tsx` | `apps/v4/components/extend/file-upload.tsx` | adapted |
| `src/vendor/extend/document-viewer-sidebar.tsx` | `apps/v4/components/extend/document-viewer-sidebar.tsx` | copied |
| `src/vendor/extend/pdf-toolbar.tsx` | `apps/v4/components/extend/pdf-viewer.tsx` | adapted |
| `src/styles/extend-globals.css` | `apps/v4/app/globals.css` | adapted (import graph only) |
| `src/styles/extend-shadcn-tailwind.css` | `apps/v4/app/shadcn-tailwind.css` | copied (provenance header added) |
| `src/lib/cn.ts` | `apps/v4/lib/utils.ts` | adapted |

Integrity: `scripts/ci/extract-extend-tokens.mjs` recomputes the sha256 of
`extend-globals.css` from `@custom-variant style-vega` to EOF and compares it
with the pinned upstream body hash
(`44bf6997bde067b701806833e88db08cfa75ae62b1805b63ee9b382107b44eca`). Passing
`--extend-dir <checkout>` additionally verifies the whole file against
`f929d2c1de31312f6fc53bd9a7fff9f6a3feff46d4d5fd65805fea2deddbc759`.

## Not vendored

`package.json`, lockfiles, font binaries, the Next.js app shell, the docs
registry, and every component not rendered by the parity matrix. Hugeicons and
Geist are consumed as npm dependencies instead of being copied.
