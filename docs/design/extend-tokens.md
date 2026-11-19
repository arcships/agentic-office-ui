# Extend token 报告（`docs/design/extend-tokens.json`）

> 由 `scripts/ci/extract-extend-tokens.mjs` 生成，固定基准
> [`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`](https://github.com/extend-hq/ui/blob/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/app/globals.css)。

```bash
pnpm ui:tokens        # 重新生成 docs/design/extend-tokens.json + extend-token-sources.json
pnpm ui:tokens:check  # 报告是否与 vendored CSS / Tailwind 默认值同步（CI 用）
```

## 数据来源

| 来源 | 内容 | 位置 |
|---|---|---|
| pinned `apps/v4/app/globals.css` | 语义色、`--radius`、暗色覆盖、ring、`@theme inline` 映射 | `tools/extend-reference/src/styles/extend-globals.css` |
| pinned `apps/v4/app/shadcn-tailwind.css` | `data-*` 自定义变体、keyframes | `tools/extend-reference/src/styles/extend-shadcn-tailwind.css` |
| pinned `tailwindcss@4.2.4` 默认主题 | `--text-*`、`--text-*--line-height`、`--spacing`、`--shadow-*`、`--default-transition-*` | `node_modules/tailwindcss/theme.css`（版本在 `docs/design/extend-token-sources.json` 固定） |
| 组件类名 | 焦点 ring 宽度、控件高度、动效 duration/easing | gutted vendored 组件（`src/vendor/**`） |

完整性：

- `extend-token-sources.json` 固定上游 `globals.css` 的 sha256（`f929d2…`）与
  `@custom-variant style-vega` 之后正文的 sha256（`44bf6997…`）。生成器每次都重算正文哈希，
  只允许改动 `@import` 图。
- `node scripts/ci/extract-extend-tokens.mjs --extend-dir /tmp/extend-ui` 会与真实上游 checkout 交叉校验整份文件。

## 报告结构

| 字段 | 内容 |
|---|---|
| `provenance` | commit、上游/本地文件哈希、Tailwind theme 文件、字体来源、解析规则说明 |
| `colors.light` / `colors.dark` | 46 个语义色：`raw`（上游原文）、`resolved`（`rgb()/alpha`）、`hex`、`kind`、近似标记 |
| `radius` | `--radius: 0.625rem` → 10px 及 sm/md/lg/xl/2xl/3xl/4xl 换算 |
| `typography` | xs/sm/base/lg/xl… 的 font-size、line-height（含 `calc(1/0.75)` 求值）、letter-spacing、weight，以及字体栈 |
| `spacing` | `--spacing: 0.25rem` → 4px，常用步进表 |
| `shadows` | Tailwind `--shadow-*` / `--inset-shadow-*` 原文 |
| `ring` | ring 颜色（light/dark）、button 的 `ring-2` + `ring-offset-1`、select trigger 的 `ring-[3px]` |
| `motion` | 默认 150ms + `cubic-bezier(0.4, 0, 0.2, 1)`、组件里出现的 duration/easing、skeleton/spin 动画 |
| `tokenLayer` | 与 #156 `--office-*` 层逐项比较（27 项：radius/ring/border/shadow/type/control），`matched/mismatched/missing` |
| `mapping` | 46 个语义色 → `--office-*`：上游值、office 值、`lightDelta`/`darkDelta`（单通道 ≤2/255、alpha ≤0.005 视为一致）、当前仓库 `var(--x, fallback)` 用法 |

### 颜色解析规则

- `oklch()` 使用 OKLab→sRGB 转换；输出 `rgb(r g b / a)` 与 hex。
- `--alpha(color / n%)` 解析为 `color-mix(in oklab, color n%, transparent)`。上游只对 black/white 使用，
  两者在 sRGB 与 oklab 中结果一致，因此本文件里的值是精确的。
- `color-mix()` 按**预乘 alpha** 插值；`in oklab` 会被标记 `approximate`。
- 未知值不会静默跳过：生成器直接抛错，避免报告悄悄失真。

## 与 #156 的结论（2026-11-20）

- 27 个标量 token：26 个一致，1 个语义差异
  `control-padding-inline`（Extend `calc(4px*3 - 1px)` = 11px，Office token = 12px；Extend 用 1px border 补足）。
- 46 个语义色中 42 个一致；4 个 `--code*` token 在 #156 未定义，报告标为 `missing-in-token-layer`。
- 唯一“当前仓库仍读旧变量名”的地方是尚未迁移的组件（`FileUpload`/`FileThumbnail`/`PdfViewer` 读
  `var(--border, …)` 等）。harness 的 `tools/vue-reference/src/parity.css` 只做名称桥接，不含数值。

## 使用方式

- **改 token**：先改 `packages/vue-ui/src/styles/tokens.css` / `theme.css`，再 `pnpm ui:tokens`，
  最后 `pnpm ui:parity` 看截图是否收敛。`mapping[].lightDelta.matches=false` 即说明与上游不一致。
- **升级 Extend commit**：同时更新 `docs/design/extend-token-sources.json` 的 `globalsSha256` / `bodySha256`、
  `scripts/ci/extract-extend-tokens.mjs` 里的 `PINNED`、`tools/parity-fixtures/{fixtures,manifest}.json` 的
  `referenceCommit`、vendored 文件头与 `docs/upstream/attribution.json`（#167 的 gate 会检查）。
