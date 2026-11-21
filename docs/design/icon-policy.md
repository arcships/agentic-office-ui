# 图标策略：Lucide 默认 + Hugeicons 像素例外（#180）

> 决策：**A（受控双轨）**。`OfficeIcon` 继续作为默认图标入口，数据源是 `@lucide/vue`；
> `@hugeicons/core-free-icons` 只在为与固定上游字形逐像素对齐而存在的 viewer chrome /
> Extend 组件文件里直连，并且逐个文件登记在 `tests/icons/allowlist.json`。
> 第三个图标库不允许出现，新增图标不允许绕过 `OFFICE_ICONS`。
>
> 门禁：`pnpm icons:check`（`scripts/ci/verify-icons.mjs`），接入 `pnpm test:docs` 与
> [`ui-parity.yml`](../../.github/workflows/ui-parity.yml) 的 `ui-contracts` job。

## 1. 为什么是受控双轨

#166 的技术决策是“图标统一 `@lucide/vue`，提供 `OfficeIcon` 适配层”。0.8 为了让
viewer chrome 与上游 Extend 参考页做 0px 像素对照，`PdfViewerToolbar` / `DocxViewerToolbar` /
`XlsxToolbar` / `PptxViewerToolbar` / `XlsxRibbon` 直接使用了 `@hugeicons/core-free-icons`
的数据；#163/#164 的 `FileUpload`、`BoundingBoxCitations`、`LayoutBlocks`、`DocumentSplits`、
`SignaturePad`、`ESignatureFields` 出于同一原因也直连。于是“统一 Lucide”名不副实。

三个选项的取舍：

| 选项 | 内容 | 结论 |
|---|---|---|
| A. 受控双轨（**采用**） | `OfficeIcon` 默认 Lucide；pixel-parity chrome 走 `HugeIcon` + 逐文件 allowlist | 保留 0.8 的像素收益，同时把第二套图标收敛成一份可审计清单 |
| B. 单一 Hugeicons | 把 `OfficeIcon` 的数据源换成 Hugeicons，Lucide 只留兼容层 | 影响全部组件与 bundle、所有宿主自定义图标，需要单独评估；不是 0.8.1 的范围 |
| C. 单一 Lucide | viewer chrome 放弃厂商字形，接受图标 delta | 与 0.8 的像素目标冲突，会让已归档的 parity 基线全部失效 |

关键约束：**例外是文件级的，不是目录级的，也不是“第二个默认”**。新增文件默认拿不到
Hugeicons；要拿，必须像现有 11 个文件一样，在 ledger 里写清 parity surface、上游快照、
owner issue 和复核日期。

## 2. 两条轨道

| 轨道 | 入口 | 数据包 | 谁可以直连 | 用途 |
|---|---|---|---|---|
| 默认 | `OfficeIcon` / `OfficeIconButton` / `OFFICE_ICONS` | `@lucide/vue` | 仅 `packages/vue-ui/src/icons.ts`（+ 登记的 `FileThumbnail.vue` 历史例外） | 所有普通 UI：按钮、菜单、状态、表单、工具栏通用图标 |
| 像素例外 | `HugeIcon`（`@arcships/vue-ui` 导出的 Vue 渲染器） | `@hugeicons/core-free-icons` | 仅 `tests/icons/allowlist.json` 列出的文件 | 与固定上游 Extend 字形逐像素对照的 viewer chrome / Extend 组件 |

`HugeIcon` 只负责把 icon data 渲染成内联 SVG（24×24 viewBox、`currentColor`、尺寸由 CSS 决定），
不提供名称注册表：调用方显式传入 `@hugeicons/core-free-icons` 的图标对象。
`OfficeIcon` 则通过 `OFFICE_ICONS` 名称注册表解析，宿主还可以用 `provideOfficeIcons()` 扩展。

## 3. 例外清单

机器可读版本是 [`tests/icons/allowlist.json`](../../tests/icons/allowlist.json)；
下表与它一一对应，`verify-icons.mjs` 会拒绝任何“ledger 有条目、本文档没有行”的漂移。

### 3.1 Viewer chrome（office-chrome，5 个文件）

| 文件 | parity surface | 上游快照 | issue |
|---|---|---|---|
| `packages/vue-pdf/src/components/PdfViewerToolbar.vue` | `pdf-toolbar` | `tools/extend-reference/src/vendor/extend/pdf-toolbar.tsx` | #162 |
| `packages/vue-docx/src/components/DocxViewerToolbar.vue` | `docx-toolbar` | `tools/extend-reference/src/vendor/extend/docx-toolbar.tsx` | #159 |
| `packages/vue-pptx/src/PptxViewerToolbar.vue` | `pptx-toolbar` | `tools/extend-reference/src/vendor/extend/pptx-toolbar.tsx` | #161 |
| `packages/vue-xlsx/src/components/XlsxToolbar.vue` | `xlsx-toolbar` | `tools/extend-reference/src/vendor/extend/xlsx-toolbar.tsx` | #160 |
| `packages/vue-xlsx/src/components/XlsxRibbon.vue` | `xlsx-ribbon` | `tools/extend-reference/src/vendor/extend/xlsx-ribbon.tsx` | #170 |

理由：这些 toolbar/ribbon 的缩放、旋转、侧栏、插入/删除行列、B/I/U/S、文字/填充色字形
就是上游的 Hugeicons 数据；换成 Lucide 等价图标会改变笔画长度、端帽和点阵几何，
`pdf-toolbar` 40/40、`docx-toolbar` 32/32、`pptx-toolbar` 40/40、`xlsx-toolbar` 64/64、
`xlsx-ribbon` 32/32 的 0px 结果会立刻失效（`tools/parity-fixtures/manifest.json` 的 blocking surface）。
`XlsxRibbon` 另有 `xlsx-office-chrome` 测试断言 29 个 Hugeicons 字形。

### 3.2 Extend 对齐组件（extend-parity，6 个文件）

| 文件 | parity surface | 上游快照 | issue |
|---|---|---|---|
| `packages/vue-ui/src/components/FileUpload.vue` | `file-upload` | `tools/extend-reference/src/vendor/extend/file-upload.tsx` | #163 |
| `packages/vue-ui/src/components/BoundingBoxCitations.vue` | `bounding-box-citations` | `tools/extend-reference/src/vendor/extend/bounding-box-citations.tsx` | #164 |
| `packages/vue-ui/src/components/LayoutBlocks.vue` | `layout-blocks` | `tools/extend-reference/src/vendor/extend/layout-blocks-panel.tsx` | #164 |
| `packages/vue-ui/src/components/DocumentSplits.vue` | `document-splits` | `tools/extend-reference/src/vendor/extend/document-splits.tsx` | #164 |
| `packages/vue-ui/src/components/SignaturePad.vue` | `e-signature` | `tools/extend-reference/src/vendor/extend/signature-fields.tsx` | #164 |
| `packages/vue-ui/src/components/ESignatureFields.vue` | `e-signature` | `tools/extend-reference/src/vendor/extend/signature-fields.tsx` | #164 |

理由同上：这 6 个组件是 Extend 组件的 Vue 重写，#163/#164 的像素 case（file 组件 19 个
`pass-with-AA`、document blocks 232/232）都建立在同一份 icon data 上；本批不重写它们。

### 3.3 默认轨道的直连（default-adapter，2 个文件）

| 文件 | 原因 |
|---|---|
| `packages/vue-ui/src/icons.ts` | 唯一默认适配层：Lucide 命名导入 + `OFFICE_ICONS` 注册表都在这里，其他文件只写 `<OfficeIcon name="...">` |
| `packages/vue-ui/src/components/FileThumbnail.vue` | 历史例外：文件类型 glyph 用了 7 个未进注册表的 Lucide 图标（archive/audio/image/spreadsheet/video 等）。迁进注册表需要新增 5 个注册名，留作 0.8.1 follow-up；现在先登记，避免“又一个直连”变成惯例 |

## 4. 新增图标：只有一条路

1. 在 `packages/vue-ui/src/icons.ts` 里从 `@lucide/vue` 命名导入图标，并加进 `OFFICE_ICONS`
   （键名用 kebab-case；`normalizeOfficeIconName()` 也接受 `ChevronDown` / `chevron_down`）。
   **不要**在业务组件里直接 `import ... from "@lucide/vue"`。
2. 组件用 `<OfficeIcon name="..." />`。需要可访问名称时传 `label`；需要宿主自定义图标时用
   `:icon="MyIcon"` 或 `provideOfficeIcons()`。
3. 只有“必须和上游 Extend 字形逐像素一致且有 parity surface / 基线”时才考虑 `HugeIcon`，
   并同时完成 §5 的登记。没有 parity surface 的组件一律走默认轨道。
4. 不允许引入第三个图标库（`@heroicons/*`、`@tabler/icons-*`、`@fortawesome/*`、`@iconify/*`、
   `lucide-react`、`bootstrap-icons`……）。`verify-icons.mjs` 对源码import 和
   `package.json` 依赖都会拦。
5. 不允许同一个文件同时直连 Lucide 和 Hugeicons。确实需要（例如一个 toolbar 里既有注册表图标
   又有上游字形）时，在 ledger 里写 `"mixed": true` + `mixedReason`，并说明为什么不拆文件。

## 5. 新增例外：流程

1. 确认该文件有 `tools/parity-fixtures/manifest.json` 里的 parity surface，且基线已提交。
2. 在 `tests/icons/allowlist.json` 加一条**精确文件**条目（不允许 glob、目录、`..`）：

   ```json
   {
     "file": "packages/vue-pdf/src/components/PdfViewerToolbar.vue",
     "iconPackages": ["@hugeicons/core-free-icons"],
     "track": "office-chrome",
     "surface": "pdf-toolbar",
     "upstreamSource": "tools/extend-reference/src/vendor/extend/pdf-toolbar.tsx",
     "issue": "#162",
     "reason": "为什么 token 化 / 换图标会改变像素",
     "reviewBy": "2027-05-21"
   }
   ```

   字段规则：`track` ∈ `default-adapter | office-chrome | extend-parity`；`default-adapter`
   只能声明 Lucide 且只能位于 `packages/vue-ui/src/`；其余轨道必须声明 Hugeicons 数据；
   `surface` 必须在 parity manifest 里；`upstreamSource` 必须是仓库里真实存在的
   `tools/extend-reference/src/vendor/**` 快照；`issue`、`reason`、`reviewBy` 必填。
   条目按 `file` 排序，过期或失效的条目一律 fail closed（不会静默豁免）。
3. 在本文档 §3 对应表格加同一行（`verify-icons.mjs` 会校验两边的文件路径一致）。
4. `pnpm icons:check` 必须 PASS；如果改了图标，同时按
   [UI 门禁](../testing/ui-parity-ci.md) 重跑对应 parity surface 并更新基线审批。

## 6. 门禁行为

```bash
pnpm icons:check        # = node scripts/ci/verify-icons.mjs
node scripts/ci/verify-icons.mjs --json   # 机器可读 summary
```

扫描范围：`packages/*/src/**` 与 `apps/*/src/**` 的 `.ts/.tsx/.mts/.cts/.js/.jsx/.mjs/.cjs/.vue`。
`tools/**`（固定上游参考 build）和 `tests/**`（测试代码）不在范围内：前者原样 import
上游 wrapper，后者不进入发布物；两者由 attribution/parity 门禁覆盖。

| 代码 | 触发条件 |
|---|---|
| `ICON_IMPORT_NOT_ALLOWLISTED` | 源码直连 `@hugeicons/*`（或未登记的 Lucide）但文件不在 ledger |
| `ICON_IMPORT_PACKAGE_NOT_LISTED` | 文件在 ledger，但条目没有声明它实际 import 的图标包 |
| `ICON_MIXED_ICON_SYSTEMS` | 同一文件同时 import Lucide 和 Hugeicons，条目没有 `mixed: true` |
| `ICON_THIRD_PARTY_LIBRARY` / `ICON_THIRD_PARTY_DEPENDENCY` | 源码或 `package.json` 出现第三个图标库 |
| `ICON_ALLOWLIST_STALE` / `_UNUSED` / `_UNUSED_PACKAGE` / `_DUPLICATE` / `_ORDER` | 条目指向不存在的文件、文件已不再 import、条目过期未用、重复或未排序 |
| `ICON_ALLOWLIST_SCHEMA` / `_FILE` / `_PACKAGES` / `_TRACK` / `_SURFACE` / `_UPSTREAM` / `_OWNER` / `_EXPIRED` | 条目字段非法、surface 不在 manifest、上游快照不存在、`reviewBy` 过期 |
| `ICON_ORPHAN_DEPENDENCY` / `ICON_UNDECLARED_DEPENDENCY` | 声明了图标依赖但没有登记的直连文件；或登记文件 import 了没声明的包 |
| `ICON_POLICY_DOC_MISSING` / `ICON_POLICY_DOC_DRIFT` | 本文档缺失，或 ledger 条目在本文件里没有对应行 |

PASS 时脚本会打印允许文件清单（文件、轨道、图标包、surface），并用
`ICON_GUARD_EVIDENCE_DIR` 写 `summary.json`。单元测试
[`tests/unit/icon-policy.test.mjs`](../../tests/unit/icon-policy.test.mjs) 覆盖允许/拒绝两侧：
新文件直连失败、混用失败、第三个库失败、条目失效失败、真实仓库通过。

## 7. 依赖与包边界

| 包 | 声明 | 说明 |
|---|---|---|
| `@arcships/vue-ui` | `@lucide/vue`、`@hugeicons/core-free-icons` | 默认适配层 + `HugeIcon` 渲染器都在这里；构建保持 external |
| `@arcships/vue-pdf` / `vue-docx` / `vue-xlsx` / `vue-pptx` | `@hugeicons/core-free-icons` | 只为各自登记的 toolbar/ribbon 直连数据；`HugeIcon` 仍从 `@arcships/vue-ui` 取 |

两个包都以 runtime dependency 消费，不内联 SVG path：见
[`THIRD_PARTY_NOTICES.md`](../../THIRD_PARTY_NOTICES.md) 与
[`docs/upstream/attribution.json`](../upstream/attribution.json)。新增依赖必须同时更新这两处，
否则 `verify-attribution` 失败。

## 8. 已知遗留

- `FileThumbnail.vue` 的 7 个 Lucide 直连（§3.3）尚未并进 `OFFICE_ICONS`；这是唯一一处
  “默认轨道但不在适配层”的直连，迁移后应删除该 ledger 条目。
- 长期单轨（选项 B）需要一次性评估 `OfficeIcon` 数据源切换、宿主自定义图标和 bundle 体积，
  不在 0.8.1 范围内；如果要做，本 ledger 就是迁移的输入清单。
