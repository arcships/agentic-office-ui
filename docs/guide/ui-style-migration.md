# 从旧样式迁移到 Office token 层

> 适用批次：0.8（#165）。目标：组件 CSS 只消费 `--office-*` token，视觉回归由
> `pnpm ui:parity` + `tests/visual/extend/**` 基线兜住，新增裸值由 `pnpm lint:style` 拦住。
>
> 相关文档：[token 表](../design/office-theme.md) · [token 报告](../design/extend-tokens.md) ·
> [component map](../components/ui-component-map.md) · [CI 门禁](ui-parity-ci.md)

## 0. 为什么迁移

0.8 之前的组件 CSS 里散落着 `#fff`、`12px`、`0 1px 2px rgb(0 0 0 / 0.05)`、`150ms ease`。
结果是：

- 改一个 Extend token（例如 `--radius`）不会传导到所有组件；
- light/dark 各自硬编码，暗色覆盖靠再写一遍十六进制；
- 视觉回归无法定位是“token 变了”还是“组件写死了值”。

迁移后：**token 文件是唯一的裸值来源**，组件 CSS 只写 `var(--office-*)`。

## 1. 先认领 token

| 需要的东西 | 用哪个 token | 定义位置 |
|---|---|---|
| 品牌/语义色（背景、前景、边框、accent、destructive、chart…） | `--office-background`、`--office-foreground`、`--office-border`、`--office-accent`、`--office-destructive`、`--office-muted`、`--office-sidebar-*` … | `packages/vue-ui/src/styles/tokens.css` + `theme.css` |
| 圆角 | `--office-radius`（基准 10px）、`--office-radius-sm/md/lg/xl/2xl/3xl/4xl` | 同上 |
| 字号 / 行高 / 字族 | `--office-text-xs/sm/base/lg/xl`、`--office-leading-xs/sm/base/lg/xl`、`--office-font-sans/mono` | 同上 |
| 阴影 | `--office-shadow-xs/sm/md`；控件层还有 `--office-ui-ring-shadow`、`--office-ui-shadow-lg` | token 层 + `ui.css` |
| 控件高度、内边距、gap | `--office-control-height`、`--office-control-height-sm/xs`、`--office-control-padding-inline`、`--office-control-gap` | token 层 |
| 动效时长 | `--office-ui-duration-fast`（120ms）、`--office-ui-duration`（160ms） | `packages/vue-ui/src/styles/ui.css`（组件层变量） |
| 焦点环 | `--office-ring`、`--office-ring-width`、`--office-ring-offset-width` | token 层 |
| 其余 | `--office-border-width`、`--office-icon-size`、`--office-selection(-foreground)`、`--office-surface(-foreground)`、`--office-code(-foreground/-highlight/-number)` | token 层 |

完整清单（`tokens.css` 中 85 条 `--office-*` 声明，深色/系统主题另有覆盖）见 [token 表](../design/office-theme.md)。
缓动暂时没有 token：`ease` / `linear` / `ease-in-out` 这类 CSS 关键字允许直接写，
`cubic-bezier()` / `steps()` 这类无法用关键字表达的定时函数才会被 guard 拦住（应先加 token）。

找不到合适的 token 时：先在 `tokens.css`/`theme.css` 里加（带 Extend 对应值），再 `pnpm ui:tokens`，
最后才改组件——不要先把值写进组件。

## 2. 四条替换规则

| 旧写法 | 新写法 | 被哪条 guard 拦住 |
|---|---|---|
| `color: #262626;` / `background: #fff;` | `color: var(--office-foreground);` | `color-no-hex` |
| `border-radius: 8px;` | `border-radius: var(--office-radius-md);` | `declaration-property-unit-disallowed-list`（radius） |
| `font-size: 14px;` | `font-size: var(--office-text-sm);` | 同上（font-size） |
| `box-shadow: 0 1px 2px rgb(0 0 0 / 0.05);` | `box-shadow: var(--office-shadow-xs);` | 同上（box-shadow 几何 + 颜色） |
| `transition: color 150ms ease;` | `transition: color var(--office-ui-duration-fast) ease;` | 同上（transition 时长 + `cubic-bezier()`/`steps()`） |

`transparent`、`currentColor`、`inherit`、纯 `0` 数值、以及 `ease`/`linear` 这类关键字不受限制；
`var(--office-*)` 内部可以带 fallback，但 fallback 里不要再写裸值（fallback 只在缺少 token 时生效，
属于隐藏的第二套数值）。

## 3. 迁移一个组件的最小步骤

以 `OfficeButton` 的 hover 阴影为例：

1. 在 `packages/vue-ui/src/styles/tokens.css` 找到对应语义：`--office-shadow-xs`。
2. 在组件的 `<style scoped>` 里把 `box-shadow: 0 1px 2px rgb(0 0 0 / 0.05)` 换成
   `box-shadow: var(--office-shadow-xs)`。
3. 跑 `pnpm lint:style`：应为 PASS（新违规会带 `file:line:column` 输出）。
4. 跑该组件的像素对照与基线：

   ```bash
   pnpm ui:parity -- --only button --viewports 1440x900 --themes light --baseline --gate-mode ready
   ```

   - 参考侧（Extend）与对照侧仍然 0px / 0 像素 → 迁移无视觉变化。
   - 出现像素差 → 先看 `output/ui-parity/SUMMARY.md` 的 top-N 差异区域，再决定是 token 值不对还是组件写法不对。
5. 如果视觉**故意**变化（例如对齐新的 Extend 版本），按 [基线审批流程](ui-parity-ci.md#3-视觉基线审批不允随手-update)
   重新采集 `tests/visual/extend/**`，并在 PR 里附上 report 链接。

## 4. 主题与作用域

- 组件不再自己判断 light/dark：所有值来自 `[data-office-theme]` 作用域内的 token。
- 宿主用 `OfficeThemeProvider` 或祖先元素上的 `data-office-theme="light|dark|system"`。
- 只引入 `@arcships/vue-ui/style.css` 而不引入 `theme.css` 时，语义色为空——这是刻意的，避免库去改宿主的全局变量。
- `color-scheme` 由 `theme.css` 设置；自建主题时同步设置，否则原生控件（滚动条、焦点环）会和 token 不一致。

## 5. 存量债务

样式 gate 有两本账（#175）：`tests/style/component-style-baseline.json` 只冻结**历史债务**，
当前是迁移后剩下的 971 条违规（527 条裸色、402 条裸尺寸/时长、42 条裸阴影色/定时函数），
分布在 61 个文件；`tests/style/parity-literal-allowlist.json` 只登记**为像素 parity 故意保留的上游字面值**，
逐条精确到 file/rule/property/value/count，并带 upstreamPath/commit、issue、reason、reviewBy。
两类白名单的边界和条目格式见 [ui-parity-ci.md §2.2](../testing/ui-parity-ci.md#22-两套样式白名单175)。
规则永远不会按目录关闭：

- 新文件、或某文件某规则超出现有计数，且不在 parity allowlist → `pnpm lint:style` 失败；
- 修好一条 → 基线计数自动“可收缩”，打印提示，不会失败；
- parity literal 匹配数少于条目 `count` → 报告 stale；条目过期或字段非法 → 直接失败；
- 历史 baseline 只能收缩；`--approve` 只用于记录收缩/刷新，新增裸值只能走 parity allowlist：

  ```bash
  pnpm lint:style:baseline --approve --reviewer <handle> --issue "#nnn" \
    --reason "XlsxGrid 的 71 处裸色由 #160 迁移"
  ```

审批记录写在基线文件的 `approved` 字段里，随 PR 一起评审。

## 6. 迁移完成的判定

- `pnpm lint:style` 的 `baselined` 计数下降到 0；
- `tests/style/component-style-baseline.json` 只剩空 `counts`；
- 每个 gate-blocking surface 的 `--baseline` 比对仍是 PASS；
- axe/keyboard 门禁（`pnpm test:a11y`）没有新增 allowlist 条目。
