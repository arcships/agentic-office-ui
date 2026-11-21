# PPTX 演讲者视图参考规格（#183）

> 状态：研究 + 规格（Issue A/B/C 未实现；#202 已把全屏目标收窄到 audience surface，
> 备注/presenter rail 不再进入投影）。本文回答"演讲者视图在办公软件里怎么组成和渲染"，
> 给出我们的目标区域模型、尺寸、交互与验收方式，供实现 issue 直接拆分。
>
> 相关：[pptx-chrome.md](../components/pptx-chrome.md)（现状与截图证据）、
> [viewer-editor-architecture.md](viewer-editor-architecture.md) §7.4（`office-reference` 分类）、
> [ui-parity-ci.md](../testing/ui-parity-ci.md)（门禁）、
> [extend-parity-harness.md](extend-parity-harness.md)（surface 清单）。

## 0. 结论

1. **五个产品的区域模型收敛**：当前页、下一页、备注、计时、导航控件是必备五项；
   缩略图/跳页、当前时间、标注（笔/激光/黑屏）是可选三项。摆放主流有两种：
   PowerPoint/Google 式"左当前页 + 右侧列（下一页在上、备注在下）"，和
   LibreOffice 式"模式切换"（Normal 只有当前页+下一页，Notes 模式才加备注）；
   ONLYOFFICE 是第三种变体（备注在预览区下方、幻灯片列表常驻左侧）。
2. **演讲者视图完全不来自文件**。OOXML 里与备注相关的只有 `notesSlide`（文本）和
   `notesMaster`/`notesSz`（备注页这一"纸面"的尺寸与默认文本样式，供打印/导出）。
   ISO/IEC 29500-1 对 `notesSz` 的定义就是"备注页和讲义页的纸张表面尺寸"。区域、
   分栏、计时器、缩略图全部是应用 UI，因此**没有可 vendored 的上游布局**这一点是
   文件格式层面的事实，不是本仓库的借口。
3. **我们的现状是"备注栏"，不是演讲者视图**：`PptxPresenterPanel` 有下一页缩略图、
   只读计时、备注滚动区；缺当前页区域（由 stage 承担，面板独立使用时没有）、
   控制条、备注字号、可拖拽分栏、跳页/缩略图、黑屏。**观众面与演讲者面分离的
   全屏漏洞已在 #202 修复**：`enterFullscreen()` 只请求 audience surface
   （`div[data-testid=pptx-audience-surface]`），备注 rail、工具栏与播放控制条
   都在其外，不再进入投影区域；真正的双窗口演讲者面仍属 P2（§7.3）。
4. **建议**：保留 `PptxPresenterPanel` 作为 rail，新增 `PptxPresenterView` 布局组件
   由它组合当前页 + 侧列；**不重写** `PptxViewer` 的播放状态机。分三个 issue 落地
   （面板增强 → 演讲者布局 → 验收接线），多窗口/多显示器单列，0.8 内只做宿主
   组合 seam 与文档化限制。
5. **验收定为 `office-reference`**：不以像素对照 Office 为目标，而是
   "书面规格 checklist（阻断）+ 自基线截图（回归）+ 组件测试（行为）"。
   `unreachable` 在本 surface 上退役：它原本表达"我们说不清什么是对的"，
   而本文已经说清了。

## 1. 证据等级与版权边界

| 等级 | 含义 | 用法 |
|---|---|---|
| **A** | 厂商官方文档 / 公开源码 / 标准文本 | 可直接引用为规格依据 |
| **B** | 第三方教程、社区贴 | 只用于补足官方未写明的细节，本文显式标注 |
| **C** | 本文给定的目标值（我们的设计决定） | 不是 Office 实测像素，不得表述为 parity |

版权边界：

- **不**把 Microsoft/Apple/Google/ONLYOFFICE 的界面截图或帮助图片 vendor 进仓库，
  只保留链接（§10）。仓库内的截图一律是**我们自己组件的自基线**。
- LibreOffice 帮助文本为 MPL-2.0，可引用文字并注明出处；本文引用其帮助页与源码路径。
- 不复制任何私有包（`@extend-ai/react-pptx`）的代码或样式；§8 只记录观察结论。

## 2. 产品区域矩阵

`✓` = 官方文档写明；`◐` = 存在但由文档间接说明；`—` = 官方文档未写（不代表产品没有）。

| 区域 / 能力 | PowerPoint | Keynote | Google Slides | LibreOffice Impress | ONLYOFFICE | 我们（现状） |
|---|---|---|---|---|---|---|
| 当前页 | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ stage（不在面板内） |
| 下一页 | ✓ 右侧列上方 | ✓ 可开关的独立项 | ✓ | ✓ Normal 右 / Notes 当前页下方 | ◐ 左侧幻灯片列表 | ✓ 缩略图 |
| 备注 | ✓ 右侧，自动换行 + 需要时滚动条 | ✓ Presenter Notes 字段 | ✓ | ✓ Notes 模式 | ✓ 预览区下方 | ✓ ScrollArea |
| 计时（elapsed） | ✓ 左上，暂停/重置 | ✓ 顶部，可换时钟 | ✓ 左侧，暂停/重置（B 级） | ✓ Restart 重开 | ✓ `hh.mm.ss` + 暂停/重置 | ◐ 只读 |
| 当前时间（时钟） | ✓ | ✓ 与计时二选一 | — | — | — | — |
| 上一页/下一页控件 | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ 底部播放条 |
| 跳页（数字 + Enter / 缩略图） | ✓ See all slides | ✓ 导航器 + 数字 | ✓ | ✓ 幻灯片浏览模式 | ✓ 页码指示器 | —（仅浏览模式） |
| 备注字号 | ✓ 备注区左下两个按钮 | ✓ 备注字段左上字体按钮 | ✓ +/− | ✓ `G`/`S` | —（社区请求中） | — |
| 备注滚动 | ✓ | ◐ 字段可调大小 | ◐ | ✓ `A`/`Z` | ◐ | ✓ |
| 分栏可拖拽 | ✓ 垂直分隔线，可拖到最左隐藏当前页 | ✓ 自由拖动/缩放各项 | ✓ 拖分隔线 | — 模式切换 | — | — |
| 布局自定义 | — | ✓ 显示项勾选 + 拖放 + Auto Layout | — | — | — | — |
| 黑屏 | ✓ `B` | ✓ | ✓ `b` | ✓ `B` / `.` | — | — |
| 笔 / 激光 | ✓ | ✓ `C` 指针（黑/白屏 `B`/`W`） | ✓ | ✓ `P` | ✓ 指针 + 笔 | — |
| 独立窗口 / 多显示器 | ✓ 每屏一窗口 + Swap Displays | ✓ 每屏可配置（最多 6 屏） | ✓ Chrome 多显示器权限 + display options | ✓ Full screen / Windowed | ✓ 独立窗口（单屏 Alt+Tab） | —（同窗口） |
| 暗色 / 反色 | —（截图可见，无文字来源） | ✓ Invert Colors | — | — | — | ✓ present 模式强制 dark |

依据见 §10。三个结论：

- **"当前页 + 下一页 + 备注 + 计时 + 控件"是跨产品的最小公倍数**；我们缺演讲者控件
  （跳页/缩略图/黑屏）、计时写权限、独立窗口，其余具备。
- **备注字号和分栏拖拽出现频率高于预期**（PPT/Keynote/Google/LO 都有字号；
  PPT/Keynote/Google 都有分栏），是"长备注可读性"的直接手段，不能算锦上添花。
- **观众面必须不含备注**是五个产品的共同硬约束。#202 已把同窗口全屏目标
  收窄到 audience surface（舞台），备注 rail 不进入全屏；仍需补的是独立窗口/
  多显示器下的真正演讲者面（P2，§7.3）。

## 3. 我们采用的目标区域模型

### 3.1 宽屏（≥1180px 容器）

```
┌──────────────────────────────────────────────────────────────────────┐
│ top bar 44px [12:34] [pause][reset]    [prev] [3/12] [next] [*]      │
├──────────────────────────────────────────┬───────────────────────────┤
│                                          │ NEXT SLIDE                │
│                                          │ ┌───────────────────────┐ │
│                                          │ │ 16:9 thumbnail        │ │
│              current slide               │ └───────────────────────┘ │
│              (16:9 letterbox)            │ ───────────────────────── │
│                                          │ NOTES 15/1.55             │
│                                          │ ┌───────────────────────┐ │
│                                          │ │ scrollable notes      │ │
│                                          │ │                       │ │
│                                          │ └───────────────────────┘ │
├──────────────────────────────────────────┴───────────────────────────┤
│ optional: thumbnail strip 96x54, horizontal scroll                   │
└──────────────────────────────────────────────────────────────────────┘
```

### 3.2 紧凑（720–899px）/ 堆叠（<720px）

- 720–899：仍为两列，侧列 280px，下一页缩略图最大高度压到侧列 22%，备注最小 96px。
- <720：单列堆叠——当前页 16:9 在上，下面是"下一页（96×54）+ 计时"一行，
  备注占剩余高度。分栏不提供：PowerPoint 在窄窗里仍画分隔线但侧列被压到最小值，
  我们选择显式堆叠，避开 1px 拖动区；Keynote 也把 "side-by-side layout" 作为
  布局开关文档化（§10，A 级），说明堆叠是正式布局而不是降级。

### 3.3 尺寸表（目标值，C 级）

容器查询沿用 `packages/vue-pptx/src/style.css:18` 的 `container: pptx-viewer / inline-size`，
新增 1180 / 900 / 560 三个断点，已有 720 断点保持（`style.css:323`）。

| 项 | ≥1180 | 900–1179 | 720–899 | <720 | 依据 |
|---|---:|---:|---:|---:|---|
| surface padding | 16px | 16px | 12px | 8px | C（与现有 12/8 递增一致） |
| top bar 高 | 44px | 44px | 40px | 36px | C；按钮 32px + 12px 内边距 |
| 侧列宽 | 380px | `clamp(300px, 32%, 340px)` | 280px | — 堆叠 | C；PPT 约 1/3，可拖拽（A） |
| 列间距 | 12px | 12px | 10px | 8px | C |
| 当前页 | 剩余宽度，16:9 letterbox，`max-block-size: 100%` | 同 | 同 | 100% 宽 | 现有 present letterbox（`style.css:223-244`） |
| 下一页 | 侧列宽，16:9，高度 ≤ 侧列 32%，最小 96px | ≤28% | ≤22% | 96×54 行内 | C |
| 备注区 | 侧列剩余，最小 120px | 最小 110px | 最小 96px | 最小 96px（抽屉 60vh） | C |
| 备注字号 | 15px / 1.55 | 15 / 1.55 | 14 / 1.5 | 14 / 1.5 | C（现有 14/1.5，`style.css:321`） |
| 计时字号 | 15px tabular-nums | 15 | 14 | 13 | C（现有 13，`style.css:317`） |
| 控件按钮 | 32px（触摸 44px） | 32 | 32 | 32（触摸 44） | 与 `pptx-chrome.md` §2 一致 |
| 缩略图条（可选） | 96×54，8px gap | 同 | 同 | 隐藏 | C |

备注字号的四档（`sm 13/1.5`、`md 15/1.55`、`lg 19/1.5`、`xl 24/1.45`）对应
PowerPoint 的两个字号按钮、Keynote 的字体按钮、Google 的 +/−、LO 的 `G`/`S`——
四档而不是连续缩放，是为了可测试、可持久化（localStorage 由宿主决定）。

## 4. 备注区域规格

| 项 | 规格 | 依据 |
|---|---|---|
| 文本模型 | 纯文本，一行一段（`PptxSlideInfo.notes` 已是 `\n` 分段） | `packages/pptx-core/src/notes.ts:132-150` |
| 富文本 | 不支持加粗/颜色/项目符号缩进；`notesSlide` 里的字符格式被丢弃 | 现状 + core 合同 |
| 换行 | `white-space: pre-wrap`；`overflow-wrap: anywhere`（长 URL/无空格串） | `style.css:321` |
| CJK | `line-break: strict`、`word-break: normal`、不两端对齐（窄栏 CJK 两端对齐会出"河流"）；字体栈沿用 Office token，CJK 回退见 `tests/baseline/visual/manifest.json` 的 determinism | C |
| 行宽 | 宽栏时 `max-inline-size: 72ch`（与上游 `@extend-ai/react-pptx@0.6.1` 的 `.rpv-notes p { max-width: 72ch }` 一致） | A（dist 观察） |
| 滚动 | 备注区内部滚动，不滚动整个 aside；`scrollbar-gutter: stable` 防止切页时列宽跳动 | C（现状 `style.css:319-320` 缺 gutter） |
| 切页 | 切页时 `scrollTop = 0`（PowerPoint/Google 行为）；改字号时保持滚动比例 | C |
| 空备注 | 显式文案（`pptx.viewer.notes.empty`），不留空白区 | 现状 ✓ |
| 超长备注 | 不截断、不折叠、不自动缩小字号；由用户用字号按钮/滚动处理 | A（PPT/LO 都是滚动 + 字号，不是 autofit） |

## 5. 计时器规格

| 项 | 规格 | 依据 |
|---|---|---|
| 起点 | 进入 present 模式即开始（PPT "timer starts as soon as you start Presenter View"；Keynote 演示开始） | A |
| 显示 | `mm:ss`，超过 1 小时 `h:mm:ss`；`font-variant-numeric: tabular-nums` | 现状 `presenter-clock.ts:7-18` |
| 暂停/继续 | 需要（PPT 暂停按钮、Google Pause、ONLYOFFICE Pause） | A |
| 重置 | 需要（PPT reset、Google Reset、LO Restart、ONLYOFFICE Reset） | A |
| 精度 | 500ms tick + `Date.now()` 差值（不做累加），后台标签页节流不产生漂移 | 现状 `presenter-clock.ts:39-47` |
| 暂停语义 | 暂停期间不计时：`elapsed = now - startedAt - pausedTotal`，继续时不变 | C（实现要求） |
| 时钟 | v2 可选（PPT 与 Keynote 都显示当前时间；Keynote 可与计时二选一） | A |
| a11y | `role="timer"` + `aria-label`；暂停/重置为带文案的按钮；不设 `aria-live="assertive"`（每 500ms 播报不可用） | 现状 + C |

## 6. 交互清单

P0 = 0.8 必须；P1 = 紧随其后；P2 = 明确延后。来源列写"产品（A 级）"或"现状"。

| # | 交互 | 来源 | 我们 | 优先级 |
|---|---|---|---|---|
| 1 | `→` / `Space` / `PageDown`：下一页（含动画步） | PPT、Keynote、Google、LO | ✓ | P0 |
| 2 | `←` / `PageUp`：上一页/重置上一动画 | 同上 | ✓ | P0 |
| 3 | `Home` / `End`：首页/末页 | PPT、Google、LO | ✓ | P0 |
| 4 | `Esc`：先退全屏，再退 present（PPT：Esc 结束放映；LO：Esc 结束） | A | ◐ 只退全屏（`PptxViewer.vue:1570-1578`） | P0 |
| 5 | 数字 + `Enter` 跳页 | PPT `See all`、Google、LO、ONLYOFFICE | — | P1 |
| 6 | `See all slides` 缩略图面板 | PPT、Keynote 导航器、Google、LO sorter | — | P1 |
| 7 | 备注字号增大/减小（按钮 + `G`/`S`） | PPT、Keynote、Google、LO | — | P0 |
| 8 | 备注滚动（滚轮/`A`/`Z`），焦点在备注区时方向键滚动而非翻页 | LO `A`/`Z` | ◐ 只有滚轮 | P1 |
| 9 | 计时暂停/继续 | PPT、Google、ONLYOFFICE | — | P0 |
| 10 | 计时重置 | PPT、Google、LO、ONLYOFFICE | — | P0 |
| 11 | 分栏拖拽（指针）+ 键盘调整（`←/→` 在分隔条上，步长 16px，`Home`/`End` 到极值） | PPT、Google、Keynote | — | P1 |
| 12 | 黑屏 `B` / 恢复 | PPT、Keynote、Google、LO | — | P1 |
| 13 | 退出 present（按钮 + `Esc`） | 全部 | ✓ 按钮 | P0 |
| 14 | 全屏（同窗口） | 全部 | ✓ 工具栏；目标 = audience surface（#202） | P0 |
| 15 | 独立窗口 / 多显示器 | 全部 | — | P2（§7.3） |
| 16 | 笔/激光、标注 | PPT、Google、LO、ONLYOFFICE | — | P2（与 `PptxStage` 标注引擎耦合） |
| 17 | 备注内编辑（PPT 允许边讲边改备注） | PPT | — | P2（我们是只读查看器，写入不在范围） |
| 18 | 焦点顺序：top bar → 当前页 → 侧列（下一页 → 备注）→ 缩略图条 | C | — | P0 |

## 7. 状态、边界与观众面

### 7.1 状态矩阵

| 状态 | 触发 | 期望 | testid |
|---|---|---|---|
| 有下一页 | `nextSlide != null` | 缩略图 + "下一页"标签 | `pptx-presenter-next` |
| 无下一页 | 当前是最后一张（跳过隐藏页后） | 占位文案 + `data-state="end"`，不显示空缩略图 | `pptx-presenter-last` |
| 下一页是隐藏页 | `showHiddenSlides=false` | 跳过隐藏页取下一张（与放映跳页一致） | 现有测试覆盖 |
| 无备注 | `notes` 为空/未提供 | 显式空状态文案，备注区仍占位可滚 | `pptx-presenter-empty` |
| 备注超长 | 段落数 > 屏 | 区间内滚动，首段可见 | — |
| 备注仅空白行 | `split("\n").filter(trim)` | 视为空 | 现状 |
| 转场进行中 | Keynote "next slide 不立即出现" | P1：下一页缩略图加 `data-pending` 淡出，不谎报可推进 | — |
| 加载/错误 | 文档未 ready | presenter 不渲染（现状受 `state === "ready"` 约束） | 现状 |

### 7.2 主题与暗色

- 现状：present 模式根节点 `data-office-theme="dark"`（`PptxViewer.vue:12`），
  演讲者面暗色、幻灯片本身保持 authored 颜色。这与 Keynote 的 Invert Colors、
  PowerPoint 的暗色演讲者面方向一致，可作为正式规格保留。
- 不做"主题跟随系统切换演讲者面"：演示中切换主题会闪烁，五个产品都不这么做。

### 7.3 观众面与演讲者面（最重要的结构约束）

五个产品都保证"观众只看到幻灯片"。#202 之前 `enterFullscreen()` 请求的是
`.pptx-viewer` 根节点，**包含备注栏**；把它投到外接屏就是把备注给观众看。
这不是像素差异，是产品语义缺陷。

0.8 的处理：

1. **全屏目标已收窄（#202，已实现）**：`enterFullscreen()` 请求
   `div[data-testid=pptx-audience-surface]`（带 `.pptx-viewer__audience`），
   该节点只包含 stage、状态层与浏览 sorter；presenter rail（`data-pptx-presenter` /
   `data-pptx-notes`）、工具栏、侧栏、播放控制条都是它的兄弟节点。
   全屏后焦点移到 audience surface（`tabindex="-1"`），键盘/`Escape`/播放状态机
   不变；退出全屏由浏览器恢复进入前的焦点。组件测试
   `tests/component/pptx-presenter-view.test.mjs` 断言全屏子树不含备注，并锁定
   "不再请求 viewer 根节点"。
2. **presenter 面全屏不开放**：把备注投到第二块屏需要独立窗口/API（`window.open`
   + 状态同步），**不能**复用 `enterFullscreen()`；调用方若需要，等 P2。
   `mode="present" + showNotes` 是同窗口演练布局：全屏时备注位于非投影 rail，
   投影只显示舞台。
3. **导出 `PptxPresenterPanel`**（现在没导出，见 §9.2），宿主可用
   `PptxStage` + 面板在第二个窗口自行组合观众面与演讲者面。
4. **P1 结构钩子**：`PptxPresenterView` 的当前页区域与侧列拆成两个根节点
   （`data-presenter-region="stage"` / `"side"`），未来 `window.open` 方案
   只需把 stage 根节点搬到另一窗口，不必重写布局。
5. **P2**：组件级 `presenterWindow` API（`window.open` + `BroadcastChannel` 同步
   页码/计时/播放状态）单列 issue；涉及弹窗拦截、Focus、Audio 与 0.8 范围无关。

## 8. 与 OOXML 的关系

| 文件部件 | 作用 | 演讲者视图用不用 | 我们的处理 |
|---|---|---|---|
| `ppt/notesSlides/notesSlideN.xml` | 一张幻灯片的备注文本 | 用（文本来源） | `pptx-core/notes.ts` 读完转纯文本，`PptxSlideInfo.notes` |
| `ppt/notesMasters/notesMaster1.xml` | 备注页的版式与默认文本样式（`p:notesStyle`） | **不用**：那是"备注页"这一纸面的排版 | 不解析；面板用 `--office-*` token 排版 |
| `p:notesSz`（`presentation.xml`） | 备注页/讲义页纸面尺寸（ISO/IEC 29500-1：用于打印或"presented"的备注页表面） | **不用** | 不解析 |
| `p:notesMasterIdLst` | 指向 notesMaster | 不用 | 不解析 |

实证：`tests/fixtures/pptx/playback/transition-cut-fade-push-wipe.pptx` 的
`ppt/presentation.xml` 里是 `<p:notesSz cx="6858000" cy="9144000"/>`（7.5×10 in 纵向），
且该 deck 没有 `notesSlides`/`notesMasters` 部件——PowerPoint 会给每份 deck 写纸面尺寸，
但它与屏幕上的演讲者视图无关。

结论写进实现约束：**presenter 的尺寸/断点只能来自 UI 规格，不能从 `notesSz` 推导**；
`notesMaster` 的字体也不应影响面板字号（否则同一 deck 在不同机器上备注大小不同）。

## 9. 现状对照与缺口

### 9.1 现状（截至本文）

| 部件 | 事实 |
|---|---|
| `PptxPresenterPanel.vue` | `aside[data-testid=pptx-presenter]` → `OfficeCard`：meta（下一页标签 + 缩略图或"最后一页" + 时钟图标 + 计时）、`OfficeSeparator`（仅 present）、body（备注标签 + `OfficeScrollArea` + 段落/空态/`notes` 插槽） |
| 布局 | `.pptx-viewer__presenter { flex: 0 0 300px; max-inline-size: 40cqw; padding: 12px }`；≤720px → 200px / 50cqw / 8px（`style.css:287-297`、`style.css:323-326`） |
| 显隐 | `notesOpen && state === "ready" && !isSorter`（`PptxViewer.vue:902-903`） |
| 计时 | `usePptxPresenterClock` 500ms tick，仅 present+running 时走；离开即清零（`presenter-clock.ts:24-56`） |
| 备注 | `\n` 分段、去空行、14px/1.5、`pre-wrap`、`overflow-wrap: anywhere`（`style.css:318-321`） |
| present 舞台 | 容器查询 letterbox，16:9，`--pptx-slide-aspect`/`--pptx-slide-ratio`（`style.css:223-244`） |
| 控制 | 底部播放条：上一页/播放暂停/下一页/重播 + `N / M` + 全屏 + 退出演示（`PptxViewer.vue:310-346`） |
| 测试 | `tests/component/pptx-presenter-view.test.mjs` 6 项；chrome story 断言 card/scrollarea/separator/next/notes/timer 并截图（`scripts/ci/capture-pptx-chrome.py:293-341`） |
| i18n | `pptx.viewer.notes.{presenter,title,empty,next,noNext,timer}`，en + zh-CN 两套（`packages/office-interaction/src/i18n.ts:2768`、`:4148`） |
| 公共 API | `PptxPresenterPanel` **未导出**（`packages/vue-pptx/src/index.ts:19-21` 只有 `PptxThumbnail`/`PptxStage`/`PptxLinkMenu`）；`PptxViewer` 也没有转发 `notes` 插槽 → 面板的 `notes` 插槽从公共入口不可达 |

### 9.2 缺口清单

| # | 缺口 | 证据 | 影响 | 归入 |
|---|---|---|---|---|
| G1 | 观众面与演讲者面同窗口，全屏曾包含备注 | #202：`enterFullscreen()` 现请求 `pptx-audience-surface` | 已修复：全屏 DOM 不含备注/presenter rail；双窗口仍缺 | 已关闭（#202）；D（P2）保留多窗口 |
| G2 | 面板不含当前页，不能独立当演讲者面 | `PptxPresenterPanel.vue:1-40` | 无法两窗口组合 | B |
| G3 | 无控制条（上一页/下一页/跳页/黑屏/缩略图） | `PptxViewer.vue:310-346` 只有播放条 | 演讲时手要回到工具栏/键盘 | B |
| G4 | 计时只读，不能暂停/重置 | `presenter-clock.ts:24-56` | 被打断无法校准 | A |
| G5 | 备注字号固定 14px | `style.css:321` | 长备注/远距离不可读 | A |
| G6 | 备注滚动无 gutter、切页不回首段 | `style.css:319-321` | 切页瞬时列宽跳动/读到一半丢位置 | A |
| G7 | 无分栏拖拽 | `style.css:287-297` 固定 flex-basis | 不能按需放大备注 | B（P1） |
| G8 | 下一页只有缩略图/末页文案，无"转场中"与隐藏页说明 | `PptxPresenterPanel.vue:10-20` | 与放映推进不同步时像卡住 | A |
| G9 | `Esc` 只退全屏，不退 present | `PptxViewer.vue:1570-1578` | 与五个产品相反 | A |
| G10 | present 模式无数字跳页 | `PptxViewer.vue:1527-1534` | 只能顺序翻 | B（P1） |
| G11 | 面板 `notes` 插槽不可达、面板未导出 | `index.ts:19-21`、`PptxViewer.vue:288-296` | 宿主无法自定义/复用 | A |
| G12 | sorter 模式直接隐藏面板 | `PptxViewer.vue:902-903` | 与 LO "sorter 是 console 的一种 mode" 不一致 | B（P2） |
| G13 | 缺控制文案 i18n（暂停/重置/字号/跳页/黑屏） | `i18n.ts:2768-2774` | 只能硬编码英文 | A |

## 10. 参考缺席的确证（上游与私有包）

**pinned 上游**：`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`

- 该 commit 的完整文件树共 **389** 个路径，`presenter` / `speaker` / `notes` 命中 **0** 个；
  与 pptx 相关的只有 6 个：
  `apps/v4/components/extend/pptx-viewer.tsx`、`file-system-pptx-thumbnails.tsx`、
  `pptx-viewer-demo.tsx`、`pptx-viewer-docs.tsx`、`content/docs/components/pptx-viewer.mdx`、
  `public/samples/demo.pptx`。
- `apps/v4/components/extend/pptx-viewer.tsx` 共 1,419 行，`presenter` / `speaker` /
  `notes` / `fullscreen` 出现 **0** 次；导出仅 `PptxViewerPreview`（:940）与
  `PptxViewerPreviewProps`（:927）。工具体系是工具栏 + 缩略图 rail + 文件菜单，
  没有演讲者视图，也没有"备注"渲染。
- 该 commit 的 `apps/v4/package.json` 依赖 `@extend-ai/react-pptx: "0.5.0"`。

**私有包**：`@extend-ai/react-pptx`（npm 上 `0.5.0` → `0.6.1`；`dist` 未压缩前的
`unpackedSize` 23.2 MB）。

- 本次以 `npm pack @extend-ai/react-pptx@0.6.1` 取包检查（只读、不 vendor、不入库）：
  35 个文件里 `presenter` / `Presenter` 出现 **0** 次（`dist/index.js`、`dist/index.d.ts`、
  `dist/index.css`、`README.md`、`CHANGELOG.md` 全查）。
- 公开面只有：`ReactPptxViewer` / `ReactPptxViewerProps` / `usePptxViewer` /
  `usePptxPresentation` / `usePptxViewerThumbnails` / 缩略图渲染器 / 解析 API。
  `type ViewerMode = 'slide' | 'continuous'`（`dist/index.d.ts:933`）——**没有 `'present'`**。
- 备注在 0.6.1 里是**文档流里的一段**，不是演讲者视图：
  `showNotes && notes.length ? <section className="rpv-notes"><h2>Speaker notes</h2>{notes.map(n => <p>{n.text}</p>)}</section>`
  （`dist/index.js`），样式 `.rpv-notes { padding: 20px 24px; border-top: 1px solid var(--rpv-rule) }`、
  `h2 { font-size: 12px; letter-spacing: .12em; text-transform: uppercase }`、
  `p { max-width: 72ch }`。
- 该包唯一的分栏是诊断面板 `aside.rpv-diagnostics { width: 270px; flex: 0 0 270px }`——
  可作为"右侧 270px 卡片"的宽度参照，但它不是演讲者视图。

**结论**：presenter 既不在公开上游组件里，也不在私有引擎的公开面里；"upstream private"
不再是"不做"的理由，因为本文的规格不依赖上游 DOM。`office-reference` 由此成立。

## 11. `office-reference` 参考状态

### 11.1 定义

`referenceStatus: "office-reference"`：

> 该 surface 没有可 vendored 的上游渲染实现，其正确性由一个**书面规格**
> （区域、尺寸、状态、交互、a11y）加上**我们自己的回归基线**保证；
> 不参与"参考 vs 被测"的像素计数，不得出现在 unreachable 表。

与现有状态的区别：

| 状态 | 参考是什么 | 门禁方式 |
|---|---|---|
| `public-vendored` | 固定 commit 的上游组件，可在 `tools/extend-reference` 渲染 | 像素 + 几何对照，0px/0 像素 |
| `fixtures-only` | 上游没有该组合，用上游 primitive 重组 | 像素对照（重组本身有上游来源） |
| `unreachable` | 说不清正确形态 | 只进 report 的 unreachable 表，不计量 |
| **`office-reference`** | 本文档 §3–§6 的规格 | 规格 checklist（阻断）+ 自基线截图（回归）+ 组件测试 |

三条硬约束：

1. **不许自造参考**：不能用我们的组件拼一个"Office 参考"再对照自己（tautology）。
2. **不许声称像素 parity**：报告里该 surface 的字段是 `spec` / `checklist`，
   没有 `maxChannelDelta` 通过率。
3. **不许用 `unreachable` 收尾**：状态缺失或 checklist 空 → 直接失败。

### 11.2 校验方式（三层）

| 层 | 内容 | 阻断 | 落点 |
|---|---|---|---|
| L1 规格 checklist | 区域存在、最小尺寸、无溢出、状态文案、计时语义、i18n、a11y 名称 | ✅ | 扩 `scripts/ci/capture-pptx-chrome.py` 的 `present_story`（已有 10 条断言，扩到 §11.3） |
| L1 组件测试 | 行为（计时暂停/重置、字号档位、切页回顶、末页/空备注、Esc 顺序） | ✅ | `tests/component/pptx-presenter-view.test.mjs` + 新 `pptx-presenter-surface.test.mjs` |
| L2 自基线 | 4 viewport × 2 主题的 presenter 截图，与提交基线比 | 回归阻断（阈值 `maxDiffPixelRatio: 0.001`） | `tests/baseline/visual/` 新增 `pptx-present` route（门禁以它为准）；chrome story 的 presenter 截图只作为 issue 证据 |

明确**不做**：与 PowerPoint/Keynote/Google/LO/ONLYOFFICE 的截图做像素 diff。
理由：字体、渲染、DPI、私有排版都不可复现；任何"0px"声明都是伪造证据。

### 11.3 checklist（可直接变成断言）

```text
presenter-surface
  [ ] data-testid=pptx-presenter-surface 根节点存在，data-presenter-mode=present
  [ ] 区域齐备：current / next / notes / timer / controls（缺一失败）
  [ ] 当前页 16:9，且 height <= 容器高 - topbar；无横向溢出
  [ ] 侧列宽在 [280, 380]，备注区高度 >= 96
  [ ] 下一页缩略图宽度 == 侧列内宽；末页时 data-state=end 且无 <img>
  [ ] 备注区可滚动：设 scrollHeight > clientHeight 后 scrollTop > 0 有效
  [ ] 切页后备注 scrollTop == 0
  [ ] 备注字号 4 档，md=15px，切档后 px 变化符合表 §3.3
  [ ] 计时：初始 00:00；+65s 后 01:05；暂停 10s 后仍 01:05；重置回 00:00
  [ ] Esc：全屏时只退全屏；再按退 present，data-presenter-mode=browse 或面板消失
  [ ] 全屏元素 == `[data-testid=pptx-audience-surface]`，且其子树无 `[data-pptx-notes]` / `[data-pptx-presenter]`（#202）
  [ ] 键盘：→/Space/PageDown 前进，←/PageUp 后退，Home/End 首末，数字+Enter 跳页
  [ ] 所有控件有可访问名称（en + zh-CN 都不为空）
  [ ] 390×844 无横向溢出；768×1024 两列；1440×900 侧列 <= 380
  [ ] axe：0 serious / 0 critical
```

### 11.4 manifest 与门禁接线（实现 issue 的精确 seam）

| 文件 | 需要改什么 |
|---|---|
| `tools/parity-fixtures/manifest.json:714-742` | `referenceStatus: "unreachable"` → `"office-reference"`；新增 `officeReference: { doc, checklist, baseline, pendingIssue }`；删除 `unreachableReason`（或保留为 `history` 字段） |
| `scripts/ci/pixel-parity.py:1252-1281` | 采集过滤与 unreachable 表之外加 `office-reference` 分支：不进 `surfaces`，进单独 `officeReference` 列表，报告输出 `spec` 段；无 checklist 字段则退出码非 0 |
| `tests/unit/extend-parity-harness.test.mjs:97-115` | `unreachable` → 允许 `office-reference`；要求 `officeReference.doc` 文件存在、`checklist` 非空；`pendingIssue` 必填 |
| `tests/unit/extend-parity-harness.test.mjs:129-136` | 断言 reference/vue app 覆盖时排除 `office-reference`（与 `fixtures-only`/`unreachable` 同列） |
| `tools/extend-reference/src/App.tsx`、`tools/vue-reference/src/App.vue` | 只为采集脚本保留 subject 侧入口；不得为 office-reference 造 reference 渲染 |
| `docs/testing/ui-parity-ci.md:205` | 该行改为"office-reference：规格 checklist + 自基线"，并链接本文 |
| `docs/design/extend-parity-harness.md:113` | 同上，替换 unreachable 描述 |
| `tests/baseline/visual/manifest.json` + `scripts/ci/capture-visual-matrix.py` | 新增 `pptx-present` route（`240` → `264` case），同步 `tests/unit/product-visual-matrix.test.mjs` 的计数契约 |
| `scripts/ci/capture-pptx-chrome.py:293-341` | `present_story` 扩到 §11.3 的前 12 条；截图命名保持 `<viewport>-<theme>-presenter.png` |

## 12. 实现计划

三个 issue，按依赖排序。每个都能独立合并，且都保持公共 API 向后兼容。

### Issue A：面板增强（S，1–2 天）

目标：把 `PptxPresenterPanel` 从"备注栏"升级成"演讲者侧列"，自身可被宿主单独使用。

| 变更 | 细节 |
|---|---|
| 导出 | `packages/vue-pptx/src/index.ts` 增加 `PptxPresenterPanel`（低层组件，与 `PptxLinkMenu` 同级） |
| props | 新增 `layout?: "rail" \| "column"`（默认 `rail`）、`showNext?: boolean`（默认 true）、`notesScale?: "sm" \| "md" \| "lg" \| "xl"`（默认 `md`）、`timerPaused?: boolean` |
| emits | `update:notesScale`、`pause`、`resume`、`reset`、`notesScroll`（可选，供宿主持久化） |
| 计时 | `usePptxPresenterClock` 增加 `pause()`/`resume()`/`reset()`，返回 `{ elapsed, paused, pause, resume, reset }`；暂停期间不计时 |
| 备注 | 四档字号（CSS 变量 `--pptx-notes-size`/`--pptx-notes-leading`）；切页 `scrollTop = 0`；`scrollbar-gutter: stable` |
| 状态 | 末页 `data-state="end"`；空备注 `data-state="empty"`；下一页 `data-presenter-next` testid |
| i18n | 新增 `pptx.viewer.notes.{pauseTimer,resumeTimer,resetTimer,notesLarger,notesSmaller,goToSlide}`（en + zh-CN） |
| 插槽 | `PptxViewer` 转发 `#notes` 到面板（面板已有插槽，当前不可达） |
| 测试 | `tests/component/pptx-presenter-view.test.mjs` 增加：暂停 10s 仍 01:05、重置归零、字号档位生效、切页回顶、末页 `data-state`、zh/en 新文案 |
| 风险 | 计时语义变化会影响现有"clock counts"测试的假定时器写法（`Date.now` stub 需同时 stub `paused` 路径） |

验收：面板可在 `PptxStage` 旁单独渲染；无 `PptxViewer` 时也不报错；`pnpm test:component` 通过。

### Issue B：`PptxPresenterView` 布局（M，3–5 天）

目标：加入当前页区域、top bar、控制条与分栏，形成真正的演讲者面。

| 变更 | 细节 |
|---|---|
| 新组件 | `packages/vue-pptx/src/PptxPresenterView.vue`，组合：top bar（计时 + 控制）+ stage 插槽 + `PptxPresenterPanel layout="column"` + 可选缩略图条 |
| 挂载方式 | `PptxViewer` 新增 `presenterSurface?: boolean`（默认 false，保持现有 rail 行为）；true 时 present 模式渲染新布局 |
| 控制条 | 上一页/下一页/页码（可点开跳页输入）/全屏/退出/黑屏；复用 `OfficeIconButton` + `--office-*` token |
| 键盘 | `Esc` 两层退出（先全屏后 present）；`B` 黑屏；数字 + `Enter` 跳页；`G`/`S` 字号；焦点在备注区时 `↑/↓`/`PgUp/PgDn` 滚动备注而不是翻页 |
| 分栏 | 指针拖拽（`role="separator"` + `aria-orientation="vertical"` + `aria-valuenow`），键盘 `←/→` 16px、`Home`/`End` 到极值；宽度写入 `--pptx-presenter-side` 并受 §3.3 clamp 约束 |
| 断点 | 容器查询新增 1180/900/560；<720 堆叠（侧列变抽屉，`max-block-size: 60vh`） |
| DOM | 根 `data-testid="pptx-presenter-surface"`，区域 `data-presenter-region="stage\|side\|next\|notes"`，便于 checklist 与未来搬窗口 |
| 测试 | 新 `tests/component/pptx-presenter-surface.test.mjs`：区域存在、分栏键盘、Esc 顺序、390 无溢出、备注滚动优先 |
| 风险 | 与 `style.css:223` `container-type: size` 的 stage wrapper 嵌套容器查询；`PptxStage` 的 focus/keydown 与顶层键盘映射竞争（需 `event.defaultPrevented` 约定） |

验收：`presenterSurface=true` 时 `data-testid=pptx-presenter-surface` 区域齐备；
关闭时视觉与现状一致（旧 story 截图不变）。

### Issue C：验收接线（S，1–2 天）

目标：把 `office-reference` 真正接进报告与 CI，避免"文档写了、门禁没跑"。

按 §11.4 的表逐项改；额外产出 `output/ui-parity/office-reference.json`
（每个 surface 的 checklist 结果 + 自基线 diff 摘要），并在 `report.html` 增加"office-reference"段
（不混入 pass/fail 像素表）。同步更新 INDEX/三份文档的引用。

验收：`pnpm ui:parity -- --only pptx-presenter` 输出 office-reference 段且退出 0；
故意删一个区域 → 退出 1；`pnpm test:docs`、`pnpm test:unit` 通过。

### Issue D（P2，0.8 外）：多窗口 / 多显示器

`window.open` + `BroadcastChannel` 同步页码/播放/计时；弹窗拦截、焦点、全屏权限、
音频策略都要单独设计。0.8 只保留 §7.3 的结构钩子与文档化限制。

### 推荐

**新增布局 + 保留面板**，而不是"改造现有面板成演讲者视图"：

- 区域模型不同（当前页 + 控制 + 缩略图 vs 备注列表），塞进同一个组件的 props 会变成
  6 个布尔开关，测试不可读。
- `PptxPresenterPanel` 在 browse 模式（只读备注）已有 6 个测试与截图，改动风险最小化。
- 观众面/演讲者面未来要拆窗口，`presenterSurface` 的根节点边界正好是拆分点。

## 13. 验收清单（0.8 出口）

- [ ] Issue A/B/C 合并，`presenterSurface` 默认关闭时旧截图零变化
- [x] #202：`enterFullscreen()` 只投影 `pptx-audience-surface`，备注/presenter rail 不进入全屏
- [ ] §11.3 checklist 全绿（脚本化，不是人工勾选）
- [ ] 4 viewport × 2 主题的 presenter 自基线已提交并审批（`tests/baseline/visual/`）
- [ ] `docs/components/pptx-chrome.md` §4 指向本文，`docs/testing/ui-parity-ci.md` 表格更新
- [ ] `tools/parity-fixtures/manifest.json` 不再有 `pptx-presenter` 的 `unreachable`
- [ ] 变更记录：`Esc` 语义、计时暂停、新增 i18n 键（`CHANGELOG.md`）
- [ ] 未声称任何与 Office 产品的像素一致性

## 14. 截图与来源清单

只保留链接，不 vendor 图片。

| 来源 | 证明了什么 |
|---|---|
| [Start the presentation and see your notes in Presenter view](https://support.microsoft.com/en-us/powerpoint/training/start-the-presentation-and-see-your-notes-in-presenter-view) | 备注在右侧；自动换行；需要时出现垂直滚动条；备注区左下两个字号按钮；拖动垂直分隔线改分栏；可把分隔线拖到最左隐藏当前页；Windows `Display Settings → Swap Presenter View and Slide Show`；Mac 顶部 `Swap Displays` / `Use Slide Show` |
| [Use Presenter View in PowerPoint](https://support.microsoft.com/en-us/powerpoint/training/use-presenter-view-in-powerpoint) | 当前页/下一页/备注三件套；计时暂停与重置在**左上**；显示当前时间；字体图标；笔/激光/缩略图/放大镜/黑屏/END SLIDESHOW |
| [Use a screen reader to show your presentation](https://support.microsoft.com/en-us/accessibility/powerpoint/use-a-screen-reader-to-show-your-presentation-with-powerpoint) | 双屏自动进 Presenter View；单屏 `Shift+F10,R`；`Alt+A`/`Alt+Z` 读备注行；`Alt+W` 读已用时间；`Alt+Q` 读下一步 |
| [Play a presentation on a separate display in Keynote](https://support.apple.com/guide/keynote/present-on-a-separate-display-tana4da2681/mac) | presenter display 可显示"幻灯片/备注/时钟/计时"任意组合；计时在顶部、随演示开始；顶部红/绿条表示转场与动画状态；`Ready to Advance Indicator`；side-by-side 布局下转场未完时下一页不出现；可勾选显示项、拖动/缩放各项、Auto Layout；备注字段左上字体按钮与 Invert Colors；多屏逐屏配置（最多 6 屏，Primary Display） |
| [Keyboard shortcuts for Keynote on Mac](https://support.apple.com/guide/keynote/keyboard-shortcuts-tanfde4a3e6d/mac) | `B` 黑屏、`W` 白屏、`C` 显示/隐藏指针 |
| [Present slides（Google）](https://support.google.com/docs/answer/1696787?hl=en&co=GENIE.Platform%3DDesktop) | Presenter view 入口；Speaker notes；Presentation display options（需 Chrome 多显示器权限；单屏下 Presenter view 与 Full screen 互斥）；键盘 `s`/`a`/`l`/`b`/`w`/数字+Enter/Home/End |
| [New features for Slides Present mode（2019）](https://workspaceupdates.googleblog.com/2019/06/slides-present-mode-updates.html) | 拖分隔线调整备注与幻灯片缩略图大小；小屏适配动机；Tips 按钮 |
| [Adding multi-monitor support to Google Slides（2024）](https://workspaceupdates.googleblog.com/2024/09/adding-multi-monitor-support-to-google-slides.html) | "Presentation View components, such as speaker notes and the timer" 在一个显示器、幻灯片在另一个 |
| [How to use Presenter view in Google Slides（BrightCarbon，B 级）](https://www.brightcarbon.com/blog/presenter-view-google-slides/) | 计时在左侧、自动开始、Pause/Reset；页码下拉跳页；Previous/Next 带预览；备注字号 +/−；presenter 窗口可缩放移动 |
| [Using the Presenter Console（LibreOffice，MPL-2.0）](https://help.libreoffice.org/latest/en-US/text/simpress/guide/presenter_console.html) | 当前页/即将播放页/可选备注/计时；需两个显示器；`Full screen` 或 `Windowed`；控件 Previous/Next/Notes/Slide/Restart/Exchange/Close；Normal（左当前右下一页）、Notes（左当前、右备注、当前页下方下一页）、Slide sorter 三种模式；`G`/`S` 备注字号、`A`/`Z` 滚动备注、`Ctrl+1/2/3` 切模式、`Ctrl+4` 换屏 |
| LibreOffice 源码（MPL-2.0）[`sd/source/ui/presenter/`](https://github.com/LibreOffice/core/tree/libreoffice-24-2/sd/source/ui/presenter) | 演讲者面是 canvas 绘制的独立实现（`PresenterCanvas`、`PresenterHelper`、`PresenterPreviewCache`、`SlideRenderer`），不是复用 Impress 编辑视图 |
| [Previewing your presentation（ONLYOFFICE）](https://helpcenter.onlyoffice.com/docs/userguides/presentation-editor/working-with-slides/preview-presentation) | Presenter 模式独立窗口；备注在预览区下方；左侧幻灯片列表（隐藏页划掉）；计时 `hh.mm.ss` + 暂停/继续/重置 + 上一页/下一页 + 页码指示器 + 指针 + 笔 + 结束放映；单屏时窗口浮在编辑器上、`Alt+Tab` 切回 |
| [Adjustable speaker notes text size in Presenter View（ONLYOFFICE 社区，B 级）](https://community.onlyoffice.com/t/adjustable-speaker-notes-text-size-in-presenter-view/16611) | ONLYOFFICE 目前没有备注字号调节（作为"并非所有产品都有"的反例） |
| [NotesSize Class（ISO/IEC 29500-1 摘录）](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.notessize) | `p:notesSz` 定义备注页/讲义页的纸面尺寸，"intended to specify the region to which content is fitted in any special format of printout" |
| `tests/fixtures/pptx/playback/transition-cut-fade-push-wipe.pptx` | 实证 `presentation.xml` 写 `<p:notesSz cx="6858000" cy="9144000"/>`，且无 notes 部件 |
| `extend-hq/ui@70fe70c6` 的 [`pptx-viewer.tsx`](https://github.com/extend-hq/ui/blob/70fe70c6f1e1eee53954c6923f3d9f583557a57e/apps/v4/components/extend/pptx-viewer.tsx) | 上游无 presenter/notes 渲染（1,419 行内 0 次命中） |
| `@extend-ai/react-pptx@0.6.1`（`npm pack` 只读检查） | `ViewerMode = 'slide' \| 'continuous'`；`showNotes` 只渲染 `.rpv-notes` 文本段；`presenter` 0 次命中 |
