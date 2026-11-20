# Extend reference renderer (parity harness input)

A minimal Vite + React app that renders the pinned
[`extend-hq/ui`](https://github.com/extend-hq/ui) components that
`@arcships/vue-ui` is compared against. It is **not** the Extend Next.js site:
only the components the harness screenshots are vendored, with the upstream
commit pinned in every file header.

- Upstream commit: `70fe70c6f1e1eee53954c6923f3d9f583557a57e` (2026-11-17)
- Vendored files and their adaptations: [`src/vendor/NOTICE.md`](src/vendor/NOTICE.md)
- Attribution index: [`docs/upstream/extend-ui.md`](../../docs/upstream/extend-ui.md)

## Run

```bash
# from the repository root
pnpm --filter @aou/extend-reference build
pnpm --filter @aou/extend-reference dev      # http://127.0.0.1:4179/?surface=button&theme=dark
```

Query parameters:

| Parameter | Values | Meaning |
|---|---|---|
| `surface` | `button`, `file-upload`, `file-thumbnail`, `pdf-toolbar`, `pptx-toolbar`, `pptx-sorter`, `document-sidebar` | render one surface; omit for all |
| `theme` | `light` (default), `dark` | toggles the `.dark` scope, same as the reference site |

The comparison page is [`tools/vue-reference`](../vue-reference). The matrix and
thresholds live in [`tools/parity-fixtures/manifest.json`](../parity-fixtures/manifest.json)
and are executed by `pnpm ui:parity`.

## Shared harness DOM contract

```html
<div data-parity-surface="file-upload" data-parity-theme="light">
  <div data-parity-cell="default" data-parity-matrix-state="empty">
    <div data-parity-host><!-- measured target is the first element child --></div>
  </div>
</div>
```

`scripts/ci/pixel-parity.py` waits for fonts, disables animations, applies the
cell's interaction (hover / programmatic focus / mouse-down / dragenter /
file-input) and captures the target box plus a 12px ring margin from both apps.

## What is deliberately missing

- No custom colour theme (`style-vega` …) overlays, no legacy themes, no docs
  chrome: `src/styles/extend-globals.css` narrows only the `@import` graph.
- No `border-beam` dependency. The FileUpload dropzone is rendered unchanged;
  only the drag-active beam is absent, and the report records that.
- No PDF engine, no search popover, no dropdown menus: the PDF toolbar cells
  render the toolbar chrome and the closed trigger states.
- No PPTX engine and no PPTX presenter: `pptx-toolbar` vendors the toolbar row
  (closed tooltips/select/menu), `pptx-sorter` renders the vendored thumbnail
  card in the production sorter grid metrics. `pptx-presenter` has no upstream
  reference at the pinned commit and is reported as `unreachable` by the
  harness instead of being compared against a harness-authored composition.
