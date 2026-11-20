# 发布 checklist（0.8 起）

> 适用：`@arcships/*` 任一 release candidate。逐项勾选并在 release PR 里粘贴证据链接/命令输出。
> **UI 变更必须附 `ui-parity` 结果**；没有结果的 UI 改动不允许进入打包步骤。
>
> 相关文档：[UI 门禁](../testing/ui-parity-ci.md) · [验证清单](../../VERIFICATION.md) ·
> [兼容矩阵](../testing/compatibility-matrix.md) · [迁移指南](../guide/ui-style-migration.md)

## 1. 代码与公开 API

- [ ] `pnpm typecheck` 通过（`pnpm typecheck:changed` 可用于增量）。
- [ ] `pnpm build` 通过；只有改动包被重建（不用 repo-root `pnpm build` 做局部验证）。
- [ ] `node scripts/ci/verify-public-api.mjs` 通过；公开导出无意外删除/重命名。
- [ ] 变更记录写入 `RELEASE_NOTES.md` / `CHANGELOG.md`（含破坏性变更与回退办法）。

## 2. 测试门禁（不允许削弱）

- [ ] `pnpm test:unit` 通过（逐文件排查失败，不批量重跑）。
- [ ] `pnpm test:component` 通过。
- [ ] `pnpm test:docs` 通过（文档合同 + 开源归属 + UI 文档映射）。
- [ ] `pnpm test:blackbox` 通过（nightly 或本地等价环境）。
- [ ] `node scripts/ci/verify-test-registration.mjs` 通过（新测试必须注册进 `run-suite.mjs`）。
- [ ] `node scripts/ci/verify-attribution.mjs` 通过（新增/改写的 Extend 文件带来源头）。

## 3. UI 视觉与样式（本项缺失即阻塞发布）

- [ ] **`ui-parity` 结果**：附 `ui-parity-core-<sha>` artifact 链接，或本地
      `pnpm ui:parity:core` 的 `SUMMARY.md`；`gate: PASS` 且 `blocking=0`。
- [ ] 若本次改动涉及 pending surface（当前为 `document-sidebar` #158 与 unreachable 的
      `pptx-presenter` #161），附 nightly 全矩阵 artifact，并写明剩余 delta 归哪个 issue。
- [ ] 全部 surface 就绪后，确认 `pnpm ui:parity -- --gate-mode all --baseline` 通过，再把
      workflow 的 strict 开关按 [ui-parity-ci.md §5](../testing/ui-parity-ci.md) 切成最终形态。
- [ ] `pnpm ui:baseline:verify` 通过；基线变化都有审批记录（`APPROVALS.md` 新增行）。
- [ ] `pnpm lint:style` 输出 `new=0`；新增裸值要么修掉，要么有审批记录。
- [ ] `pnpm test:a11y` 通过；新增 allowlist 条目必须带 issue 与理由。
- [ ] `pnpm ui:tokens:check` 通过（token 报告与固定上游 CSS 同步）。

## 4. 打包与消费者

- [ ] `pnpm test:consumer` 通过（真实 tgz 安装）。
- [ ] `node scripts/ci/prepare-release-artifact.mjs` 生成候选包，summary 记录的 commit 与 release PR 一致。
- [ ] `node scripts/ci/p4-reproducible-pack.mjs` 复现性检查通过。

## 5. 发布与回退

- [ ] `package.json` 版本号只在 release PR 里改（不在 feature 分支里改）。
- [ ] release workflow 的必需检查包含 `ui-contracts`、`ui-parity-core`、`ui-a11y`。
- [ ] 回退办法已写清（哪个包、哪个版本、如何降级）。
- [ ] 发布后更新 `docs/testing/ui-parity-ci.md` 的 §4 状态表与 [component map](../components/ui-component-map.md)。

## 6. 证据模板

```text
commit: <sha>
ui-parity: <artifact 链接> — gate PASS, blocking 0, pending <n> (#issue…)
ui-a11y:   <artifact 链接> — PASS, new 0, allowlisted <n>
style:     pnpm lint:style — PASS, new 0, baselined <n>
unit/component/docs: PASS
consumer:  PASS
```
