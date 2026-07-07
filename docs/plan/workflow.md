# 迁移开发 Workflow

## 模型分配策略

### DOCX 模块化重做

| 难度 | 任务 | 模型 | provider | 理由 |
|---|---|---|---|---|
| ✅ done | docx-engine | deepseek-v4-pro | dimcode-api-oauth | 引擎层复制+wasm smoke |
| 🟢 低 | docx-layout, docx-canvas | deepseek-v4-pro | dimcode-api-oauth | 复制+改 import 路径,机械执行 |
| 🟡 中 | docx-viewer | deepseek-v4-pro | dimcode-api-oauth | 辅助模块复制+拆分 pretext/thumbnail |
| 🟡 中 | docx-helpers | glm-5.2 | dimcode-api-oauth | 30+ 模块拆分,需理解 editor.tsx 内部依赖 |
| 🔴 高 | docx-composables | glm-5.2 | dimcode-api-oauth | useDocxEditor 事务分发器+选区恢复,深度推理 |
| 🟡 中 | docx-render | glm-5.2 | dimcode-api-oauth | JSX→Vue h() render,语义转换 |
| 🔴 高 | docx-components | glm-5.2 | dimcode-api-oauth | 18 个 Vue 组件+虚拟化+contentEditable |
| 🟢 低 | docx-demo | deepseek-v4-pro | dimcode-api-oauth | demo 接入+构建配置 |
| 🟢 低 | docx-verify | deepseek-v4-pro | dimcode-api-oauth | 全量验证,按清单检查 |

### XLSX（旧任务体系）

| 难度 | 任务 | 模型 | provider | 理由 |
|---|---|---|---|---|
| 🔴 高 | xlsx-005 | glm-5.2 | dimcode-api-oauth | canvas 渲染管线 |
| 🟢 低 | xlsx-006, cross-integ | deepseek-v4-pro | dimcode-api-oauth | 配置/验证 |

## DOCX 执行顺序

```
Phase 1: ✅ 引擎层
  ① docx-engine → deepseek-v4-pro（已完成）

Phase 2: 布局层 + canvas
  ② docx-layout → deepseek-v4-pro
  ③ docx-canvas → deepseek-v4-pro

Phase 3: viewer 辅助模块
  ④ docx-viewer → deepseek-v4-pro（拆分 pretext-layout/thumbnail-raster）

Phase 4: editor-helpers 拆分（最大工作量）
  ⑤ docx-helpers → glm-5.2（30+ 模块,分批 5-6 个）

Phase 5: editor-ops + state
  ⑥ docx-editor-ops → deepseek-v4-pro

Phase 6: composables 改写
  ⑦ docx-composables → glm-5.2（26 个 composable）

Phase 7: render 重写
  ⑧ docx-render → glm-5.2（renderParagraphRuns → h()）

Phase 8: Vue 组件重写
  ⑨ docx-components → glm-5.2（18 个组件,先 DocxViewer 再 DocxEditor）

Phase 9: demo 接入 + 验证
  ⑩ docx-demo → deepseek-v4-pro
  ⑪ docx-verify → deepseek-v4-pro（全量 gate）
```

## 每个任务的执行流程

```
1. 创建分支 task/{id}（或继续在 task/docx-remigration 上）
2. dispatch develop agent（指定模型）
   - prompt: "你的任务是 {id}。阅读 docs/docx-migration-architecture.md 对应章节,
     按 objective 完成。硬约束:单文件 ≤1000 行。完成后运行 typecheck + build。"
3. develop 完成后 commit
4. dispatch verify agent
   - prompt: "review 任务 {id} 的开发产出。架构文档:docs/docx-migration-architecture.md。
     检查:单文件 ≤1000 行、typecheck 通过、功能对齐。写入 docs/plan/reviews/{id}-1.md。"
5. verify pass → commit → 全量 gate → 标记 done
   verify blocked → fix → re-verify
```

## 模型 dispatch 参数

### 高难度任务（glm-5.2）

```
agent create:
  providerId: dimcode-api-oauth
  modelId: glm-5.2
  forkContext: none
```

### 低难度任务（deepseek-v4-pro）

```
agent create:
  providerId: dimcode-api-oauth
  modelId: deepseek-v4-pro
  forkContext: none
```

## verify 模型分配

| 被验证任务难度 | verify 模型 | 理由 |
|---|---|---|
| 🔴 高（docx-composables, docx-components） | glm-5.2 | 需深度判断事务/选区/渲染正确性 |
| 🟡 中（docx-viewer, docx-helpers, docx-render） | glm-5.2 | 需判断拆分/render 转换语义 |
| 🟢 低（其余） | deepseek-v4-pro | 按清单检查 |

## 分支管理

```
master（主线）
  └── task/docx-remigration（当前,已完成 engine 回退+引擎层重做）
      ├── docx-layout → commit
      ├── docx-canvas → commit
      ├── docx-viewer → commit
      ├── docx-helpers → commit（分批）
      ├── docx-editor-ops → commit
      ├── docx-composables → commit（分批）
      ├── docx-render → commit
      ├── docx-components → commit（分批）
      ├── docx-demo → commit
      └── docx-verify → merge to master
```

## 中断恢复

从 `docs/plan/README.md` 的 status 字段 + git log 恢复。

## 全量 gate（每个任务 commit 后运行）

```bash
pnpm typecheck && pnpm build && \
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && \
git diff --check
```

如果 gate 失败，revert commit，任务重新 fix + re-verify。
