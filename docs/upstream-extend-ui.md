# Upstream Extend UI Attribution

## Purpose

This repository is a Vue 3 workspace for office/document UI components. It is inspired by the public Extend UI / Extend AI React packages and aims to provide Vue components, composables, demos, and verification workflows with comparable API coverage.

The original Extend UI / Extend AI projects, package names, API concepts, and product direction belong to Extend and their maintainers. This repository should preserve attribution and avoid obscuring the upstream origin of the ideas it is implementing.

## Upstream references

Information below was checked via the npm registry during repository setup.

### `@extend-ai/react-docx`

- npm package: `@extend-ai/react-docx`
- npm version observed: `0.7.6`
- repository: `https://github.com/extend-hq/react-docx`
- homepage: `https://github.com/extend-hq/react-docx#readme`
- repository metadata: `git+https://github.com/extend-hq/react-docx.git`, directory `packages/react-viewer`
- npm license field: MIT

### `@extend-ai/react-xlsx`

- npm package: `@extend-ai/react-xlsx`
- npm version observed: `0.13.4`
- repository: `https://github.com/extend-hq/react-xlsx`
- homepage: `https://github.com/extend-hq/react-xlsx#readme`
- repository metadata: `git+https://github.com/extend-hq/react-xlsx.git`
- npm license field: MIT

## Current repository policy

- The original Extend UI source code is not vendored in this repository.
- This repository keeps provenance documentation so upstream authors are credited.
- Vue package names in this workspace are internal implementation names for the porting/verification work.
- API compatibility claims should be backed by tests, browser acceptance evidence, and source references.
- Any future vendored snapshot or submodule should include upstream license files and the exact commit/tag.

## Recommended sync strategy

When aligning behavior with upstream:

1. Record the upstream package, version, commit/tag, and checked date.
2. Map the upstream React API to the Vue API in a docs table.
3. Implement Vue behavior in local source files without copying untracked upstream source into the repository.
4. Add real browser verification using the demo app and sample materials.
5. Update [`docs/component-browser-verification-evidence.md`](component-browser-verification-evidence.md) and [`docs/shadcn-components-browser-acceptance.md`](shadcn-components-browser-acceptance.md).
6. Keep intentional differences explicit.

## API mapping notes

| Upstream reference | Local Vue workspace area | Notes |
|---|---|---|
| `@extend-ai/react-docx` | `packages/docx-core`, `packages/vue-docx` | DOCX model, viewer/editor components, composables |
| `@extend-ai/react-xlsx` | `packages/xlsx-core`, `packages/vue-xlsx` | XLSX viewer, provider, controller/composables, utility helpers |
| Extend UI shared components | `packages/vue-extend` | PDF viewer, signature pad, upload, thumbnails, citations, layout blocks, spinner, tooltip |

## Respect statement

This project should clearly credit Extend and its maintainers whenever it references upstream APIs or behavior. Public documentation should describe this codebase as a Vue implementation inspired by and compatible with the relevant public Extend AI React packages, while keeping the original upstream ownership clear.
