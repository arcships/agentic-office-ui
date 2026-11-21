# Office 设计 token、主题与图标（0.8）

> 状态：已实现（#156）。参考基准：Extend UI / shadcn `apps/v4/app/globals.css`，commit
> `70fe70c6f1e1eee53954c6923f3d9f583557a57e`。
>
> 本文是该批次的唯一视觉来源说明。组件迁移由后续 issue 逐步完成，本 issue 只提供
> token、主题、字体与图标入口，不迁移旧组件。

## 1. 入口与文件

| 文件 / 入口 | 作用 |
|---|---|
| `packages/vue-ui/src/styles/tokens.css` | token 契约、字体、圆角、阴影、密度和浅色基线 |
| `packages/vue-ui/src/styles/theme.css` | 引入 token 并追加 `light` / `dark` / `system`，公开入口为 `@arcships/vue-ui/theme.css` |
| `@arcships/vue-ui/tokens.css` | 只要 token 契约和浅色基线时使用 |
| `@arcships/vue-ui/style.css` | 既有组件样式，保持不变 |
| `OfficeThemeProvider` | 写入主题属性、密度属性和 `lang`，并通过 provide 暴露主题上下文 |
| `OfficeIcon` | 以 `@lucide/vue` 为底座的统一图标入口 |

`theme.css` 在发布产物中是自包含的：构建会把 `@import "./tokens.css"` 内联进
`dist/theme.css`，因此用 `<link>` 直接加载也只需要一个请求。

```ts
import "@arcships/vue-ui/style.css"
import "@arcships/vue-ui/theme.css"
```

## 2. 主题与 Provider

```vue
<script setup lang="ts">
import { ref } from "vue"
import { OfficeThemeProvider, OfficeIcon } from "@arcships/vue-ui"

const theme = ref<"light" | "dark" | "system">("system")
const density = ref<"comfortable" | "compact">("comfortable")
</script>

<template>
  <OfficeThemeProvider :theme="theme" :density="density" locale="zh-CN">
    <button class="office-button">
      <OfficeIcon name="check" />
      已保存
    </button>
  </OfficeThemeProvider>
</template>
```

| Prop | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `theme` | `"light" \| "dark" \| "system"` | `"system"` | `system` 跟随 `prefers-color-scheme`，运行时监听变化 |
| `locale` | `string` | — | 写入 `lang`，并驱动共享消息目录；未提供时继承上层 Office locale |
| `density` | `"comfortable" \| "compact"` | `"comfortable"` | 只切换控件尺寸 token，不改变颜色 |
| `messages` | `OfficeMessagesOverride` | — | 在 locale 目录上合并宿主文案 |
| `as` | `string` | `"div"` | Provider 根元素标签；根元素使用 `display: contents`，不参与布局 |

Provider 写入的属性：

| 属性 | 示例 | 用途 |
|---|---|---|
| `data-office-theme` | `dark` | 唯一的主题选择器；CSS 只匹配这个属性 |
| `data-office-theme-resolved` | `dark` | 解析后的实际明暗，供宿主脚本读取；SSR 初值为 `light`，挂载后更新 |
| `data-office-density` | `compact` | 密度选择器 |
| `lang` | `zh-CN` | 仅在显式传入 `locale` 时写入 |

`useOfficeTheme()` 暴露 `theme`、`resolvedTheme`、`density`、`locale` 和 `isDark` 五个只读 ref。
其中 `locale` 是生效值（显式 prop 或继承的上层 Office locale），而 `lang` 属性只在显式传入时
写入，避免意外覆盖宿主页面的语言声明。

没有 Provider 时，宿主可以直接在任意祖先元素上写 `data-office-theme="dark"`；token 会正常继承。
嵌套 Provider 可以覆盖祖先：深色外层中放一个 `theme="light"` 的 Provider，内部会重新声明浅色
token，不需要全局 `.dark` 类。未使用 Provider 时，`useOfficeTheme()` 返回
`system` / `light` / `comfortable` 的只读基线，不会抛错。

SSR：主题属性由 props 直接决定，`system` 的首屏由 CSS 媒体查询负责；`matchMedia` 只在
`onMounted` 中访问，因此服务端没有 `window` 也不会崩溃。

## 3. Token 契约

所有声明都在 `[data-office-theme]` 下，不修改 `:root`、`html`、`body`，也不引入全局 reset。

### 3.1 语义颜色

| Token | Light | Dark | Extend 对应 |
|---|---|---|---|
| `--office-background` | `oklch(1 0 0)` | `oklch(14.5% 0 0)` | `--background` |
| `--office-foreground` | `oklch(26.9% 0 0)` | `oklch(97% 0 0)` | `--foreground`（neutral-800 / neutral-100） |
| `--office-card` | `white` | background 98% + white | `--card`（`var(--color-white)`） |
| `--office-card-foreground` | `oklch(26.9% 0 0)` | `oklch(97% 0 0)` | `--card-foreground` |
| `--office-popover` | `white` | background 98% + white | `--popover`（`var(--color-white)`） |
| `--office-popover-foreground` | `oklch(26.9% 0 0)` | `oklch(97% 0 0)` | `--popover-foreground` |
| `--office-primary` | `oklch(26.9% 0 0)` | `oklch(97% 0 0)` | `--primary`（中性色，不是蓝色） |
| `--office-primary-foreground` | `oklch(98.5% 0 0)` | `oklch(26.9% 0 0)` | `--primary-foreground` |
| `--office-secondary` | black 4% alpha | white 4% alpha | `--secondary` |
| `--office-secondary-foreground` | `oklch(26.9% 0 0)` | `oklch(97% 0 0)` | `--secondary-foreground` |
| `--office-muted` | black 4% alpha | white 4% alpha | `--muted` |
| `--office-muted-foreground` | neutral-500 90% + black | neutral-500 90% + white | `--muted-foreground` |
| `--office-accent` | black 4% alpha | white 4% alpha | `--accent` |
| `--office-accent-foreground` | `oklch(26.9% 0 0)` | `oklch(97% 0 0)` | `--accent-foreground` |
| `--office-destructive` | red-500 | red-500 90% + white | `--destructive` |
| `--office-destructive-foreground` | red-700 | red-400 | `--destructive-foreground` |
| `--office-destructive-contrast` | red-600（#169） | red-600（#169） | 无；WCAG AA 派生 token |
| `--office-destructive-contrast-foreground` | white | white | 无；WCAG AA 派生 token |
| `--office-border` | black 8% alpha | white 6% alpha | `--border`（参考值 8% alpha） |
| `--office-input` | black 10% alpha | white 8% alpha | `--input` |
| `--office-ring` | neutral-400 | neutral-500 | `--ring` |
| `--office-sidebar` | neutral-50 | neutral-950 97% + white | `--sidebar` |
| `--office-sidebar-foreground` | neutral-800 64% + sidebar | neutral-100 64% + sidebar | `--sidebar-foreground` |
| `--office-sidebar-primary` | neutral-800 | neutral-100 | `--sidebar-primary` |
| `--office-sidebar-primary-foreground` | neutral-50 | neutral-800 | `--sidebar-primary-foreground` |
| `--office-sidebar-accent` | black 4% alpha | white 4% alpha | `--sidebar-accent` |
| `--office-sidebar-accent-foreground` | neutral-800 | neutral-100 | `--sidebar-accent-foreground` |
| `--office-sidebar-border` | black 6% alpha | white 5% alpha | `--sidebar-border` |
| `--office-sidebar-ring` | neutral-400 | neutral-400 | `--sidebar-ring` |
| `--office-success` | emerald-500 | emerald-500 | `--success` |
| `--office-success-foreground` | emerald-700 | emerald-400 | `--success-foreground` |
| `--office-warning` | amber-500 | amber-500 | `--warning` |
| `--office-warning-foreground` | amber-700 | amber-400 | `--warning-foreground` |
| `--office-info` | blue-500 | blue-500 | `--info` |
| `--office-info-foreground` | blue-700 | blue-400 | `--info-foreground` |
| `--office-surface` | `oklch(98% 0 0)` | `oklch(20% 0 0)` | `--surface` |
| `--office-surface-foreground` | `var(--office-foreground)` | `oklch(70.8% 0 0)` | `--surface-foreground` |
| `--office-code` | `var(--office-surface)` | `var(--office-surface)` | `--code` |
| `--office-code-foreground` | `var(--office-surface-foreground)` | `var(--office-surface-foreground)` | `--code-foreground` |
| `--office-code-highlight` | `oklch(96% 0 0)` | `oklch(27% 0 0)` | `--code-highlight` |
| `--office-code-number` | `oklch(56% 0 0)` | `oklch(72% 0 0)` | `--code-number` |
| `--office-selection` | `oklch(14.5% 0 0)` | `oklch(92.2% 0 0)` | `--selection` |
| `--office-selection-foreground` | `oklch(1 0 0)` | `oklch(20.5% 0 0)` | `--selection-foreground` |

附加 token：`--office-chart-1`、`--office-chart-2`、`--office-chart-3`、`--office-chart-4`、
`--office-chart-5`（blue-300/500/600/700/800）。参考色值使用 Tailwind v4 的 OKLCH 表示；
`--alpha(x / n%)` 统一展开为 `color-mix(in oklab, x n%, transparent)`，不依赖 Tailwind 运行时。
Extend 用 `var(--color-white)` 声明 card / popover；这两项必须写成 `white`，不能用
`oklch(1 0 0)`（Chromium 会序列化为 `color(srgb 0.999871 1.00005 1.00007)`，在描边、焦点环
或阴影合成后偏 1/255；#157 的 input 焦点对照即由此产生）。`--office-background` 仍与上游一样
使用 `oklch(1 0 0)`，保持与参考页同一表示。

`--office-destructive-contrast` / `--office-destructive-contrast-foreground` 是 #169 新增的
WCAG AA 派生对：红-600 + 白字在 light/dark 下均为 **4.83:1**，而上游 `--office-destructive`
（红-500）的实心白字只有 3.81:1。实心 destructive Button/Badge 使用这对 token，
`--office-destructive` 仍用于描边、文本和低强度强调。上游参考页保持 red-500；parity harness 在
`tools/parity-fixtures/manifest.json` 的 `expectedDiffs` 中为 button destructive 声明了有界像素豁免
（报告状态 `waived`，含 issue/reason/evidence 和实际 diff ratio），基线见
`tests/visual/extend/button/destructive/`。

### 3.2 圆角、描边与阴影

| Token | 值 | 像素（16px root） |
|---|---|---|
| `--office-radius` | `0.625rem` | 10px |
| `--office-radius-sm` | `calc(var(--office-radius) * 0.6)` | 6px |
| `--office-radius-md` | `calc(var(--office-radius) * 0.8)` | 8px |
| `--office-radius-lg` | `var(--office-radius)` | 10px |
| `--office-radius-xl` | `calc(var(--office-radius) * 1.4)` | 14px |
| `--office-radius-2xl` | `calc(var(--office-radius) * 1.8)` | 18px |
| `--office-radius-3xl` | `calc(var(--office-radius) * 2.2)` | 22px |
| `--office-radius-4xl` | `calc(var(--office-radius) * 2.6)` | 26px |
| `--office-radius-full` | `9999px` | `rounded-full`（empty state glyph、skeleton caption） |
| `--office-ring-width` | `2px` | 焦点环宽度 |
| `--office-ring-offset-width` | `1px` | 焦点环偏移 |
| `--office-border-width` | `1px` | 默认描边 |
| `--office-shadow-xs` | `0 1px 2px 0 rgb(0 0 0 / 0.05)` | 与 Extend `shadow-xs` 一致 |
| `--office-shadow-sm` | `0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1)` | 与 Extend `shadow-sm` 一致 |
| `--office-shadow-md` | `0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1)` | 与 Extend `shadow-md` 一致 |
| `--office-shadow-lg` | `0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1)` | 与 Extend `shadow-lg` 一致（overlay sidebar 边缘） |

### 3.3 字体与字号

| Token | 默认值 |
|---|---|
| `--office-font-sans` | `Geist, "Geist Sans", ui-sans-serif, system-ui, ... , sans-serif` |
| `--office-font-mono` | `"Geist Mono", ui-monospace, "SF Mono", SFMono-Regular, Menlo, ... , monospace` |
| `--office-text-xs` / `--office-leading-xs` | `0.75rem` / `1rem` |
| `--office-text-sm` / `--office-leading-sm` | `0.875rem` / `1.25rem` |
| `--office-text-base` / `--office-leading-base` | `1rem` / `1.5rem` |
| `--office-text-lg` / `--office-leading-lg` | `1.125rem` / `1.75rem` |
| `--office-text-xl` / `--office-leading-xl` | `1.25rem` / `1.75rem` |

### 3.4 密度与图标

| Token | `comfortable` | `compact` |
|---|---|---|
| `--office-control-height` | `2.25rem` | `2rem` |
| `--office-control-height-sm` | `2rem` | `1.75rem` |
| `--office-control-height-xs` | `1.75rem` | `1.5rem` |
| `--office-control-padding-inline` | `0.75rem` | `0.625rem` |
| `--office-control-padding-inline-sm` | `0.625rem` | `0.5rem` |
| `--office-control-gap` | `0.5rem` | `0.375rem` |

`--office-icon-size` 默认 `1rem`（16px），`OfficeIcon` 的默认线宽为 `2`。`--office-font-mono`
用于代码、公式和需要等宽对齐的 Office 数值区域。

`--office-control-padding-inline` 的语义是**完整内缩**（Extend 的 `px-[calc(--spacing(3)-1px)]`
= 11px padding + 1px border = 12px）：需要自己画 1px 描边的控件使用
`calc(var(--office-control-padding-inline) - var(--office-border-width))`，渲染值才与参考的
11px padding 一致（`ui.css` 中的 Button / Select / Input 均如此）。`-sm` 后缀同理。
不要把 token 直接当作 1px 描边控件的 `padding-inline`，否则会比参考宽 1px。

## 4. 字体：默认 Geist，宿主覆盖

包内**不打包字体**：没有 `@font-face`，也没有字体二进制或 `url()`，避免把 Geist 的
woff2 体积带进所有消费者。token 只声明字体栈，宿主按需加载一次 Geist：

```html
<!-- 方式一：Google Fonts -->
<link rel="preconnect" href="https://fonts.googleapis.com" />
<link
  href="https://fonts.googleapis.com/css2?family=Geist:wght@400..700&family=Geist+Mono:wght@400&display=swap"
  rel="stylesheet"
/>
```

```css
/* 方式二：自托管 woff2 */
@font-face {
  font-family: "Geist";
  src: url("/fonts/geist.woff2") format("woff2");
  font-weight: 100 900;
  font-display: swap;
}
```

覆盖 token 时，选择器必须命中 `[data-office-theme]`（或更深层元素），写在 `theme.css`
之后：

```css
/* app.css */
[data-office-theme] {
  --office-font-sans: "Inter", ui-sans-serif, system-ui, sans-serif;
  --office-font-mono: "JetBrains Mono", ui-monospace, monospace;
}
```

也可以在 Provider 上使用行内样式（行内声明的优先级最高）：

```vue
<OfficeThemeProvider style="--office-font-sans: 'Inter', sans-serif">
```

文档正文（DOCX 页面、PDF 文本层）应继续使用文档自身字体；`--office-font-*` 面向 Office
chrome、工具栏、面板和表单控件。

## 5. 图标：OfficeIcon + Lucide

`@lucide/vue` 是 `@arcships/vue-ui` 的运行时依赖，构建时保持 external，不复制进每个包。

> 图标是受控双轨：本节描述的 `OfficeIcon` 是默认轨道；viewer chrome / Extend 组件的
> Hugeicons 像素例外、逐文件清单和新增图标规则见
> [icon-policy.md](icon-policy.md)（#180，门禁 `pnpm icons:check`）。

| Prop | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `name` | `string` | — | 注册名，支持 `chevron-down`、`ChevronDown`、`chevron_down` |
| `icon` | `Component` | — | 直接传入组件，优先级高于 `name`；用于宿主自定义图标 |
| `size` | `number \| string` | `1rem` token | 数字按 px，字符串按任意 CSS 长度；写入 `--office-icon-size` |
| `strokeWidth` | `number` | `2` | Lucide 线宽 |
| `label` | `string` | — | 可访问名称；省略时输出 `aria-hidden="true"` |

```vue
<OfficeIcon name="search" />
<OfficeIcon name="trash-2" label="删除" />
<OfficeIcon :icon="MyBrandIcon" size="1.25rem" />
```

内置注册表（`OFFICE_ICONS`）只包含 50 个常用图标，按需命名导入；未使用的 Lucide 图标
不会进入消费者 bundle。扩展注册表：

```ts
import { h } from "vue"
import { provideOfficeIcons } from "@arcships/vue-ui"

const Star = () => h("svg", { viewBox: "0 0 24 24" }, [/* path */])

// 在 setup() 中调用，作用域限于该子树
provideOfficeIcons({ "brand-star": Star, check: MyCheckIcon })
```

- `resolveOfficeIcon(name, registry?)`：纯函数解析，未命中返回 `undefined`。
- `OFFICE_ICON_NAMES`：内置名称清单，可用于文档或图标选择器。
- `provideOfficeIcons()` / `useOfficeIcons()`：子树级注册与读取；后注册项覆盖同名内置图标。
- 未知名称不会抛错：开发环境打印一次 `[OfficeIcon] Unknown icon ...` 警告，渲染为空。

## 6. 暗色示例与运行时切换

```vue
<script setup lang="ts">
import { computed, ref } from "vue"
import { OfficeThemeProvider, OfficeIcon } from "@arcships/vue-ui"

const preference = ref<"light" | "dark" | "system">("system")
const next = computed(() => (preference.value === "dark" ? "light" : "dark"))
</script>

<template>
  <OfficeThemeProvider :theme="preference" density="compact">
    <div class="office-toolbar">
      <OfficeIcon name="file-text" label="文档" />
      <button @click="preference = next">切换到 {{ next }}</button>
      <button @click="preference = 'system'">跟随系统</button>
    </div>
  </OfficeThemeProvider>
</template>
```

`theme="system"` 时，CSS 通过 `@media (prefers-color-scheme: dark)` 切换 token；Provider 同时
监听 `matchMedia` 更新 `data-office-theme-resolved` 和 `useOfficeTheme()` 的 `isDark`。

## 7. Demo 接入方式

`apps/demo/src/main.ts` 目前已经导入 `@arcships/vue-ui/style.css`。接入主题只需要追加：

```ts
import "@arcships/vue-ui/style.css"
import "@arcships/vue-ui/theme.css"
```

然后在需要演示主题的页面外层包一层：

```vue
<OfficeThemeProvider theme="system" density="comfortable">
  <RouterView />
</OfficeThemeProvider>
```

Route 级演示可以在 `apps/demo/src/pages/ComponentsPage.vue` 或后续的组件画廊页面中加入
`light` / `dark` / `system` 切换按钮；本 issue 不修改 demo 代码，避免与并行分支冲突。

## 8. 新代码约束

- 新组件只使用 `--office-*` token：颜色、圆角、字体、字号、阴影、焦点环。
- 不写裸 hex / `rgb()` 颜色；不写裸 `px` 圆角；不写裸 `font-size`。
- 不新增全局选择器、不引入 Tailwind/preflight、不依赖 `.dark` 类。
- 旧组件在各自 issue 中迁移；迁移前保持现状，避免本批次出现大面积视觉回归。
- 暗色值在 `theme.css` 中显式写了两份（`dark` 与媒体查询内的 `system`），单测会比较两块
  声明，防止漂移。

## 9. 验证

```bash
pnpm --filter @arcships/vue-ui typecheck
pnpm --filter @arcships/vue-ui build

node --test --test-concurrency=1 \
  tests/unit/vue-ui-design-tokens.test.mjs \
  tests/component/vue-ui-theme-icon.test.mjs
```

覆盖内容：token 名称与数值、8% 边框 / 4% muted / neutral primary / `--radius` / ring 2px /
shadow-xs-sm、`light`/`dark`/`system` 作用域、无 `:root`/`html`/`body` 污染、字体默认 Geist
且不打包字体、密度切换、Provider 属性与运行时切换、`matchMedia` 清理、无 `window` 的 SSR、
`OfficeIcon` 的尺寸/线宽/可访问属性、自定义图标与注册表扩展、构建产物 `dist/theme.css`。
