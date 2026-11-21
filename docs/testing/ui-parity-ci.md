# UI 视觉 / 样式 / 无障碍门禁

> 适用批次：0.8（#165）。矩阵与采集细节见 [design/extend-parity-harness.md](../design/extend-parity-harness.md)，
> 组件对照见 [components/ui-component-map.md](../components/ui-component-map.md)。
> workflow：[`.github/workflows/ui-parity.yml`](../../.github/workflows/ui-parity.yml)。

## 1. 门禁总览

| Job | 触发 | 做什么 | 现在是否阻断 |
|---|---|---|---|
| `ui-contracts` | PR / nightly / dispatch | `pnpm lint:style`（组件 CSS 裸值）+ `pnpm ui:baseline:verify`（基线审批） | 是（新违规、未审批的基线改动） |
| `ui-parity-core` | PR / dispatch | 核心子集 × `1440x900,390x844` × light/dark，参考页 vs 对照页 + 提交基线比对 | 是（`--gate-mode all`） |
| `ui-parity-nightly` | nightly / dispatch | 全矩阵（25 surface × 4 viewport × 2 theme） | 是（`--gate-mode all`） |
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

## 3. 视觉基线审批（不允许随手 `--update`）

基线目录：`tests/visual/extend/<surface>/<state>/<theme>-<viewport>.png`。
矩阵：`1440x900` + `390x844` × `light` + `dark`，`capturePaddingPx: 12`，DPR 2。
只有 `gateStatus: blocking` 的 surface 有基线；pending surface 一旦补上基线，`ui:baseline:verify` 会把它当成契约的一部分。

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
| `pptx-presenter` | ❌ | pending（#161） | 上游无 presenter 布局（`referenceStatus: unreachable`），由组件测试覆盖 |

25 个可渲染 surface 已全部 `blocking` 并提交基线；16 个进 PR 子集（全部 primitive、file、PDF、
DOCX/PPTX/XLSX chrome），9 个 chrome/block surface 只在 nightly 全矩阵跑。`pptx-presenter` 是
登记在案的无上游布局 surface，不参与像素计数，也是唯一的 pending。

## 5. 最终开关

1. **当前形态（2026-11-20 起）**：所有可渲染 surface 已 `blocking`，`manifest.gate.prMode` 为
   `all`，workflow 两个 parity job 都直接传 `--gate-mode all`。PR 核心子集与 nightly 全矩阵
   都以严格模式阻断；`strict` dispatch 输入已删除。合并 #158b 后的全矩阵实测：
   **864 case / 845 `pass` + 19 个 file-upload AA / 0 `fail`**，基线 432 张、4 条审批。
   #169 之后 button destructive 的 8 个 matrix case（4 viewport × 2 theme）改为 `waived`；
   `pnpm ui:parity -- --only button` 实测 **104 pass + 8 waived / 0 fail**。
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
pnpm ui:parity:core                # PR 子集 + 基线比对 + --gate-mode all（CI 等价命令）
pnpm ui:parity:full                # 全矩阵 + 基线比对 + --gate-mode all（nightly 等价命令）
pnpm ui:parity:self-check          # 参考页自己跟自己比，验证采集链路确定
pnpm ui:baseline:status            # 基线现状
pnpm ui:baseline:verify            # 基线契约（CI 用）
pnpm lint:style                    # 组件 CSS 裸值 guard（对比冻结基线）
pnpm test:a11y                     # axe-core + 键盘 smoke（light+dark；需先 build tools/vue-reference）
pnpm test:a11y -- --build          # 先构建 workspace 依赖 + tools/vue-reference 再检查（CI 用）
pnpm ui:tokens:check               # token 报告与固定上游 CSS 同步
```

常用筛选：`--only <id[,id]>`（`core` 展开为 PR 子集）、`--viewports`、`--themes`、`--no-build`、`--skip-tokens`。

Python 解释器由 `scripts/ci/run-python.mjs` 解析（`PYTHON=/tmp/bb-venv/bin/python` 或 PATH 里的 venv）。

## 7. 已知例外（为什么现在不是全绿）

| 例外 | 位置 | 负责 |
|---|---|---|
| `pptx-presenter` 无上游布局 | parity 报告（unreachable） | #161（组件测试覆盖） |
| `file-upload` / `file-thumbnail` 19 个 `pass-with-AA`（单通道 1/255） | parity 报告 | #163 已知残差，低于 0.1% 阈值 |
| `button` destructive 有意偏离上游 red-500 | manifest `expectedDiffs` + `tests/visual/extend/button/destructive/` 基线 | #169；参考页保持上游 red-500，gate 报告 `WAIVED`（bounded pixel-ratio waiver） |
| 1,015 条存量组件 CSS 裸值 | `tests/style/component-style-baseline.json` | #159–#164 收尾时逐步还债 |

`tests/a11y/allowlist.json` 在 #169 后为空：button loading 通过 `aria-label` + `aria-disabled`
保留可访问名与焦点，destructive 实心按钮使用 4.83:1 的 token 对，PDF 缩略图列表补上
`:focus-visible` token 焦点环；`pnpm test:a11y` 必须 0 violation / 0 keyboard failure。

除这些已登记项外，任何新增视觉漂移、裸值或无无障碍回归都会在 PR 阶段失败。

## 8. 发布要求

发布 checklist（[plan/release-checklist.md](../plan/release-checklist.md)）要求每个 release candidate 附上：
`ui-parity-core` 的 artifact 链接 + `ui-a11y` 结果 + `pnpm lint:style` 的 `new=0` 输出。
没有这三项，release 不允许进入打包步骤。
