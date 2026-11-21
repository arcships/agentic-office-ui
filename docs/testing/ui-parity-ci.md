# UI 视觉 / 样式 / 无障碍门禁

> 适用批次：0.8（#165）。矩阵与采集细节见 [design/extend-parity-harness.md](../design/extend-parity-harness.md)，
> 组件对照见 [components/ui-component-map.md](../components/ui-component-map.md)。
> workflow：[`.github/workflows/ui-parity.yml`](../../.github/workflows/ui-parity.yml)。

## 1. 门禁总览

| Job | 触发 | 做什么 | 现在是否阻断 |
|---|---|---|---|
| `ui-contracts` | PR / nightly / dispatch | `pnpm lint:style`（组件 CSS 裸值：历史 baseline + parity-literal allowlist）+ `pnpm ui:baseline:verify`（基线审批） | 是（新违规、非法/过期 allowlist、未审批的基线改动） |
| `ui-parity-core` | PR / dispatch | 核心子集 × `1440x900,390x844`（`coreSubset.viewports`，快速门槛）× light/dark，参考页 vs 对照页 + 提交基线比对 | 是（`--gate-mode all`） |
| `ui-parity-nightly` | nightly / dispatch | 全矩阵（26 surface × 4 viewport × 2 theme）+ **4 个 viewport** 的提交基线比对（896 case） | 是（`--gate-mode all`） |
| `ui-a11y` | PR / nightly / dispatch | axe-core WCAG A/AA + 键盘 smoke（Tab/焦点环/可访问名/Select 开合），light + dark | 是，`tests/a11y/allowlist.json` 为空，任何发现即失败 |
| `ci.yml` `check` job | push / PR | 额外跑一次 `pnpm lint:style` | 是 |

所有 job 都会上传 artifact：

| Artifact | 内容 |
|---|---|
| `ui-parity-core-<sha>` / `ui-parity-full-<sha>` | `output/ui-parity/{report.json,report.html,SUMMARY.md,cases/**}`，含双方截图、diff、heatmap、baseline diff |
| `ui-style-guard` | `output/style-guard/summary.json` |
| `ui-a11y-<sha>` | `output/a11y/{report.json,report.html,SUMMARY.md}` |

## 2. 阈值（`tools/parity-fixtures/manifest.json` → `thresholds`）

| 规则 | 值 | 说明 |
|---|---|---|
| 几何 | `geometryTolerancePx: 0` | box、padding、gap、border 宽度、四角 radius、font-size、line-height 逐项比较，**0px 容差** |
| 像素（非文本） | 0 像素差异 → `pass` | 任何差异都要走抗锯齿规则 |
| 文本抗锯齿 | 单通道 ≤ `2/255` + 差异像素 ≤ `maxDiffPixelRatio: 0.001`（0.1%）+ 连续差异 ≤ 2×2 | 三条同时满足才是 `pass-with-antialiasing` |
| 差异区域 | `topDiffRegions: 5` | 失败时按 16px 网格聚类差异像素，输出最大的 5 个区域（坐标、尺寸、像素数、峰值通道差），写进 `SUMMARY.md`、`report.html` 和控制台 `BLOCK ...` 行 |
| 尺寸不一致 | 直接 `fail` | 另外对左上公共区域做诊断 diff |

阈值只允许收紧或通过 manifest 变更（同时更新本文档）；脚本里没有第二套硬编码。

### 2.1 Expected diffs（有界豁免）

产品有意偏离上游参考时，不允许改参考页“抹平”差异，也不允许把 state 从矩阵里删掉。在
`tools/parity-fixtures/manifest.json` 的根级 `expectedDiffs` 声明一条有界豁免，`pixel-parity.py`
会把命中的差异记为 `waived`，并在 `report.json` / `report.html` / `SUMMARY.md` 和控制台
`WAIVED ...` 行显式列出，不会静默通过。

```json
{
  "id": "button-destructive-wcag-aa-169",
  "surface": "button",
  "state": "destructive",
  "themes": ["light", "dark"],
  "viewports": ["1440x900", "1280x720", "768x1024", "390x844"],
  "maxDiffPixelRatio": 0.52,
  "issue": "#169",
  "reason": "subject 使用 WCAG AA red-600 …",
  "evidence": "tests/visual/extend/button/destructive/"
}
```

规则：

- 必须精确匹配 surface + state + theme + viewport；不支持 `*` 通配。
- 只豁免**像素差异**且实际 ratio 必须低于 `maxDiffPixelRatio`；几何、尺寸不一致、页面报错、基线 diff 仍然阻断。
- `maxDiffPixelRatio` 是硬上限；超过上限仍报 `fail`，并在 case 上写 `waiverExceeded`。
- `issue`、`reason`、`evidence` 必填；manifest 单元测试会校验 surface/state/theme/viewport 存在、
  比例不超过 0.52，并确认参考页没有被改成和产品一样。
- 报告会列出实际 diff ratio；每次 `ui:parity` 都会重新计算，不是永久白名单。

当前唯一条目是 #169 的 button destructive：subject 使用 `--office-destructive-contrast`
（red-600/white = 4.83:1），上游参考仍是 red-500（3.81:1）。

## 2.2 两套样式白名单（#175）

`pnpm lint:style` 只用于组件 CSS（`packages/*/src/**/*.{vue,css}`，token 文件除外），白名单分两类，
**不要混用**：

| 文件 | 用途 | 粒度 | 能否增长 |
|---|---|---|---|
| `tests/style/component-style-baseline.json` | 迁移前已经存在的**历史债务** | `file` + `rule` 计数 | 只能收缩；`--approve` 只能刷新/收缩，不能新增 |
| `tests/style/parity-literal-allowlist.json` | 为像素 parity 故意保留的**上游字面值** | 精确 `file` + `rule` + `property` + `value` + 有界 `count` | 只能逐条新增，必须带 owner/复核日期 |

门禁顺序是“先扣 parity literal，再比历史 baseline”：

1. stylelint 产生 warning，脚本从 `warning.text` 解析出 `property` / `value`
   （unit 规则的值是 `px`/`ms` 等，hex 规则是 `#fff`，value 规则是完整字面值）。
2. 某条 warning 精确命中 allowlist 条目（文件、规则、属性、值全部相同，且该条目的 `count` 尚未用完）→
   从“待比较集合”中扣除，并在 `output/style-guard/summary.json` 的
   `parityAllowlist.matches[].warnings[]` 里记录 `file:line:column`。
3. 剩余 warning 再与 `component-style-baseline.json` 比较；超出 baseline 计数的部分才是 `new` 并失败。
4. allowlist 条目实际匹配数少于 `count` → `parityAllowlist.stale`，控制台和证据都会提示收缩，但不阻断
   （说明债务已经减少）。条目字段非法或 `reviewBy` 已过期 → 直接失败（fail closed），
   不会把 warning 静默豁免。

条目格式（字段全部必填，除了 `source`）：

```json
{
  "id": "pptx-viewer-toolbar-ring-shadow",
  "file": "packages/vue-pptx/src/PptxViewerToolbar.vue",
  "rule": "declaration-property-unit-disallowed-list",
  "property": "box-shadow",
  "value": "px",
  "count": 16,
  "upstreamPath": "apps/v4/components/extend/pptx-viewer.tsx",
  "upstreamCommit": "70fe70c6f1e1eee53954c6923f3d9f583557a57e",
  "issue": "#161",
  "reason": "为什么这是上游像素几何，而不是没还的债",
  "reviewBy": "2027-05-21",
  "source": "upstream"
}
```

规则：

- `file` 必须是仓库内已存在的精确路径；`*`、`?`、`{}`、`[]`、前缀写法和 `..` 一律拒绝。
- `property` 必须精确（unit/value 规则），`color-no-hex` 用 `null`；`value` 必须精确，不允许通配。
- `count` 是正整数且不超过 1000：一个条目只能豁免有限条，不能写“整个文件/整个规则”。
- `upstreamPath` + `upstreamCommit` 指向固定上游来源（当前固定 commit 为
  [`docs/upstream/extend-ui.md`](../upstream/extend-ui.md) 中的 `70fe70c6…`）；`issue`、`reason`、`reviewBy` 必填。
- 新增条目必须同时说明“为什么 token 化会改变像素”，不能只写“迁移遗留”。
- 能 token 化且像素不变的，优先改回 token 并把条目的 `count` 减小或删除；`reviewBy` 到期前必须复核。

当前 seed 只有 #161/#161b 的 22 条 PPTX 字面值：`PptxViewerToolbar.vue` 16 条 box-shadow、
3 条 border-radius，`PptxSlideCard.vue` 3 条 transition。通用 baseline 已从 1,015 收缩到 **971**
（527 裸色 / 402 尺寸时长 / 42 阴影色与定时函数，61 个文件），不含任何 parity 条目；
后续目标是继续把 971 条历史债务迁回 token，而不是让 allowlist 变成第二本债务账。

新增/移除条目的流程：

```bash
# 1. 看 warning 的精确 file/property/value
pnpm lint:style

# 2. 编辑 tests/style/parity-literal-allowlist.json（只加必要的精确条目）

# 3. 确认 parity 字面值被扣除、legacy baseline 仍为 0 new
pnpm lint:style

# 4. 若 baseline 可收缩，按审批刷新（不会把 allowlist 条目写进 baseline，也不能新增计数）
pnpm lint:style:baseline --approve --reviewer <handle> --issue "#nnn" \
  --reason "迁移 X 组件的 N 条裸值"
```

单元测试 `tests/unit/style-parity-allowlist.test.mjs` 锁住：新裸值失败、allowlisted 上游字面值通过并出现在证据、
stale 被报告、过期/非法条目 fail closed、legacy baseline 仍然工作，以及真实 allowlist 的 22 条 seed。

## 3. 视觉基线审批（不允许随手 `--update`）

基线目录：`tests/visual/extend/<surface>/<state>/<theme>-<viewport>.png`。
矩阵（#174）：4 个 parity viewport 全部有基线：`1440x900`、`1280x720`、`768x1024`、`390x844`
× `light` + `dark`，`capturePaddingPx: 12`，DPR 2。25 个 blocking surface × 全部 state = **864 张**（~14 MiB）。
只有 `gateStatus: blocking` 的 surface 有基线；pending surface 一旦补上基线，`ui:baseline:verify` 会把它当成契约的一部分。

| 矩阵 | viewport | 谁在比对 | 本地实测（2026-11-21） |
|---|---|---|---|
| PR 快速门槛 | `1440x900`, `390x844` | `ui-parity-core`（`--viewports 1440x900,390x844 --baseline`） | 264 case / 174 s |
| nightly 全矩阵 | 4 个全部 | `ui-parity-nightly`（`pnpm ui:parity -- --baseline --gate-mode all`，不传 `--viewports`） | 864 case、864 次基线 diff / 578 s（含两侧采集） |
| 重新采集 | 4 个全部 | `pnpm ui:baseline capture` → 864 张写入 + 审批 | 566 s |

PR 故意只跑 2 个 viewport（分钟级反馈）；`1280x720` / `768x1024` 的基线由 nightly 全矩阵和 dispatch 比对，
缺失或漂移都会让 `ui-parity-nightly` 以 `--gate-mode all` 阻断。
`tests/unit/ui-parity-gate.test.mjs` 锁住三件事：基线矩阵 = `manifest.viewports`、
`coreSubset.viewports` 是它的子集、nightly 命令没有 `--viewports`（否则 4 个 viewport 里的两个永远不会被比）。

流程（三步都必须进 PR）：

```bash
# 1. 看现状：文件数、缺失、与 manifest 的哈希是否一致
pnpm ui:baseline:status

# 2. 采集 + 记录审批（没有 --approve 会直接拒绝，不会写任何文件）
pnpm ui:baseline capture --only pdf-toolbar \
  --approve --reviewer <github-handle> --issue "#162" \
  --reason "toolbar gap 从 6px 改为 --office-control-gap，重采基线"

# 3. 校验（CI 里是必须通过的步骤）
pnpm ui:baseline:verify
```

- `capture` 会把整棵基线树的 `treeSha256` 写进 `tests/visual/extend/manifest.json` 的 `approvals[]`，
  并重建 `APPROVALS.md`。
- `verify --require-approval` 会失败于：PNG 与 manifest 记录的 sha256 不一致、目录里有未登记的 PNG、
  blocking surface 缺基线、当前树没有被最新审批覆盖。**手改 PNG 或只 `git add` 新图都会失败。**
- 审核人看 PR 里的 PNG 变化 + `report.html`；审批理由要写清“为什么视觉变化是对的”。

## 4. 现在哪些 surface 在门禁里

`tools/parity-fixtures/manifest.json` 每个 surface 有两个字段：

- `core: true` → 进 PR 子集；`core: false` 只在 nightly 全矩阵跑。
- `gateStatus: blocking | pending`：最终门禁用 `--gate-mode all`，所有被采集的 surface 都阻断；
  `pending` 只用于登记尚未接入的 surface（`gatePendingIssue` 指向负责 issue），本地可用
  `--gate-mode ready` 临时放宽。

| Surface | core | gateStatus | 现状 |
|---|---|---|---|
| `button`, `input`, `select`, `dialog`, `tooltip` | ✅ | blocking | #157 全部 0px / 0 像素；有基线 |
| `file-upload`, `file-thumbnail` | ✅ | blocking | #163 已收敛；19 个 `pass-with-AA`（单通道 1/255，低于 0.1% 阈值）；有基线 |
| `pdf-toolbar`, `pdf-sidebar` | ✅ | blocking | #162 40/40、24/24 像素 0；有基线 |
| `docx-toolbar`, `docx-sidebar` | ✅ | blocking | #159 32/32 + 24/24 像素 0；有基线，PR 即阻断 |
| `pptx-toolbar`, `pptx-sorter` | ✅ | blocking | #161 40/40 + 24/24 像素 0；有基线，PR 即阻断 |
| `document-sidebar` | ❌ | blocking | #158 inline open/closed/skeleton + overlay open/closed，40/40 像素 0；有基线 |
| `office-empty-state`, `office-error-state`, `office-loading-state`, `office-status-bar` | ❌ | blocking | #158 共 64/64 像素 0（含长消息换行卡片）；有基线 |
| `bounding-box-citations`, `layout-blocks`, `e-signature`, `document-splits` | ❌ | blocking | #164 232/232 像素 0；有基线（nightly 阻断） |
| `xlsx-toolbar`, `xlsx-sheet-tabs`, `xlsx-status-bar` | ✅ | blocking | #160 64/64 像素 0；有基线，PR 即阻断 |
| `xlsx-ribbon` | ❌ | blocking | #170 32/32 像素 0（4 viewport × light/dark × default/hover/focus/disabled）；有基线（nightly 阻断） |
| `pptx-presenter` | ❌ | pending（#161） | 上游无 presenter 布局（`referenceStatus: unreachable`），由组件测试覆盖 |

26 个可渲染 surface 已全部 `blocking` 并提交 4 个 viewport 的基线（896 张）；16 个进 PR 子集（全部 primitive、
file、PDF、DOCX/PPTX/XLSX chrome），10 个 chrome/block surface 只在 nightly 全矩阵跑。`pptx-presenter` 是
登记在案的无上游布局 surface，不参与像素计数，也是唯一的 pending。

## 5. 最终开关

1. **当前形态（2026-11-21，含 #174/#170）**：26 个可渲染 surface 全部 `blocking`，4 个 viewport
   都有提交基线（896 张 / 9 条审批）。PR core 仍只跑 `1440x900,390x844`（264 case）；`strict` dispatch 输入已删除；nightly
   全矩阵跑 4 个 viewport 并逐 case 比对基线（896 case）。`pptx-presenter` 仍是无上游布局的
   pending surface，不参与像素计数。`#169` 的 8 个 destructive button case 为 `waived`；
   `xlsx-ribbon` 32/32（4 state × 4 viewport × 2 theme）为 `pass`。
2. **以后新增 surface 时**：
   - 参考/对照通过后，把 `gateStatus` 改成 `blocking`（删掉 `gatePendingIssue`）；
   - 需要进 PR 子集的再加 `core: true`；否则只在 nightly 全矩阵里阻断；
   - `pnpm ui:baseline capture --only <id> --approve ...` 补基线；
   - `pnpm ui:baseline:verify` + `pnpm ui:parity -- --only <id> --baseline --gate-mode all` 必须 PASS；
   - 更新本文档 §4 与 [component map](../components/ui-component-map.md)。
3. **回退到宽松模式（仅限本地排查）**：显式传 `--gate-mode ready`，pending surface 只报告。

## 6. 本地命令

```bash
pnpm ui:parity                     # 全矩阵 + 参考对照；默认 --gate-mode all（最终门禁形态）
pnpm ui:parity:report              # 同上，永远 exit 0（artifact 模式）
pnpm ui:parity:core                # PR 子集 × 4 viewport + 基线比对 + --gate-mode all（CI PR job 的 2-viewport 超集）
pnpm ui:parity:full                # 全矩阵 × 4 viewport + 基线比对 + --gate-mode all（nightly 等价命令）
pnpm ui:parity:self-check          # 参考页自己跟自己比，验证采集链路确定
pnpm ui:baseline:status            # 基线现状
pnpm ui:baseline:verify            # 基线契约（CI 用）
pnpm lint:style                    # 组件 CSS 裸值 guard（对比冻结基线）
pnpm test:a11y                     # axe-core + 键盘 smoke（light+dark；需先 build tools/vue-reference）
pnpm test:a11y -- --build          # 先构建 workspace 依赖 + tools/vue-reference 再检查（CI 用）
pnpm ui:tokens:check               # token 报告与固定上游 CSS 同步
```

常用筛选：`--only <id[,id]>`（`core` 展开为 PR 子集）、`--viewports`、`--themes`、`--no-build`、`--skip-tokens`。
要复现 PR job 的 2-viewport 参数：

```bash
pnpm ui:parity -- --only core --viewports 1440x900,390x844 --themes light,dark \
  --baseline --gate-mode all --skip-tokens
```

Python 解释器由 `scripts/ci/run-python.mjs` 解析（`PYTHON=/tmp/bb-venv/bin/python` 或 PATH 里的 venv）。

### 6.1 L2 路由对照（`pnpm ui:parity:routes`，#182）

L1 只比较组件格子；L2 比较 `pnpm dev` 的**整块产品区域**与上游参考路由。入口与产物：

```bash
pnpm ui:parity:routes                          # 构建 reference + demo，report-only 矩阵
pnpm ui:parity:routes -- --only pdf-viewer --viewports 1440x900 --themes light
pnpm ui:parity:routes -- --self-check          # reference 自己跟自己比（采集确定性）
```

配置源是 `tools/parity-fixtures/routes.json`（region/landmark 选择器、状态、排除项、分类差异、地标容差），reference 侧是 `tools/extend-reference/route.html` 原样 import 固定 commit 的上游 wrapper（R2），产物默认写到 `output/ui-parity-routes/`（`report.html` / `routes-matrix.json` / 每 case 的 reference-subject-diff-heatmap）。当前 pilot `/pdf-viewer` 是 `report-only`：chrome 与页面栅格的产品差异尚未清零，**不要把它设为 blocking**。区域定义、分类与剩余阻断项见 [ui-parity-routes.md](ui-parity-routes.md)。

## 7. 已知例外（为什么现在不是全绿）

| 例外 | 位置 | 负责 |
|---|---|---|
| `pptx-presenter` 无上游布局 | parity 报告（unreachable） | #161（组件测试覆盖） |
| `/pdf-viewer` L2 路由仍是 `report-only`（画布背景、页面栅格、暗色反转等 6 条阻断差异，另 1 条亚像素差异不阻断） | `output/ui-parity-routes/` | #182 follow-up |
| `file-upload` / `file-thumbnail` 19 个 `pass-with-AA`（单通道 1/255） | parity 报告 | #163 已知残差，低于 0.1% 阈值 |
| `button` destructive 有意偏离上游 red-500 | manifest `expectedDiffs` + `tests/visual/extend/button/destructive/` 基线 | #169；参考页保持上游 red-500，gate 报告 `WAIVED`（bounded pixel-ratio waiver） |
| 971 条存量组件 CSS 裸值 + 22 条 #161b PPTX parity literal | `tests/style/component-style-baseline.json` + `tests/style/parity-literal-allowlist.json` | #159–#164 收尾时逐步还债；parity 条目按 `reviewBy` 复核 |

`tests/a11y/allowlist.json` 在 #169 后为空：button loading 通过 `aria-label` + `aria-disabled`
保留可访问名与焦点，destructive 实心按钮使用 4.83:1 的 token 对，PDF 缩略图列表补上
`:focus-visible` token 焦点环；`pnpm test:a11y` 必须 0 violation / 0 keyboard failure。

除这些已登记项外，任何新增视觉漂移、裸值或无无障碍回归都会在 PR 阶段失败。

## 8. 发布要求

发布 checklist（[plan/release-checklist.md](../plan/release-checklist.md)）要求每个 release candidate 附上：
`ui-parity-core` 的 artifact 链接 + `ui-a11y` 结果 + `pnpm lint:style` 的 `new=0` 输出。
没有这三项，release 不允许进入打包步骤。
