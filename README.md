# Agentic Office UI

Vue 3 workspace for office/document UI components and browser verification demos.

## Upstream attribution

This project is a Vue-oriented implementation inspired by the public Extend UI / Extend AI React packages. The original projects, package names, API ideas, and design direction belong to Extend and their maintainers.

Upstream references checked from the npm registry:

- `@extend-ai/react-docx`
  - npm version observed: `0.7.6`
  - repository: `https://github.com/extend-hq/react-docx`
  - homepage: `https://github.com/extend-hq/react-docx#readme`
  - license reported by npm: MIT
- `@extend-ai/react-xlsx`
  - npm version observed: `0.13.4`
  - repository: `https://github.com/extend-hq/react-xlsx`
  - homepage: `https://github.com/extend-hq/react-xlsx#readme`
  - license reported by npm: MIT

This repository currently does not vendor or commit the original Extend UI source code. It records upstream provenance so the original authors are credited and future API compatibility work can be traced back to the correct projects.

See [`docs/upstream-extend-ui.md`](docs/upstream-extend-ui.md) for the detailed provenance and sync policy.

## Workspace

```text
apps/demo/              Browser verification demo
packages/docx-core/     Framework-agnostic DOCX model/layout helpers
packages/vue-docx/      Vue DOCX components and composables
packages/xlsx-core/     XLSX utility layer
packages/vue-xlsx/      Vue XLSX components and composables
packages/vue-extend/    Shared Vue components: PDF viewer, upload, thumbnails, citations, layout blocks, spinner, tooltip
scripts/                Test material generation and verification helpers
docs/                   Verification, acceptance, and upstream documentation
```

## Commands

```bash
pnpm install
pnpm dev
pnpm typecheck
pnpm build
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts\verify_test_materials.py
```

## Browser acceptance

The primary acceptance workflow uses the shared browser and real interactions. The current UI acceptance method is documented in [`docs/shadcn-components-browser-acceptance.md`](docs/shadcn-components-browser-acceptance.md).
