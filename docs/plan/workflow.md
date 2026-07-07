# 迁移开发 Workflow

## 模型分配策略

| 难度 | 任务 | 模型 | provider | 理由 |
|---|---|---|---|---|
| 🔴 高 | xlsx-005, docx-003, docx-004 | `glm-5.2` | `dimcode-api-oauth` | canvas 渲染管线 / 事务分发器 / contentEditable + 选区恢复时序，需要深度推理和长上下文（1M context） |
| 🟡 中 | xlsx-003, xlsx-004, docx-002 | `glm-5.2` | `dimcode-api-oauth` | hook 改写 / JSX→render / 类型清理，需要理解 React→Vue 语义差异 |
| 🟢 低 | xlsx-001, docx-001, docx-005, cross-001, xlsx-006, docx-006, xlsx-integ, docx-integ, cross-integ | `deepseek-v4-pro` | `dimcode-api-oauth` | 复制 / 配置 / 验证，机械执行，1M context + 0.87 $/M 性价比高 |

## 执行顺序

当前环境无 worktree，所有任务串行执行。

```
Phase 1: 引擎层（两条线串行或先做一条）
  ① xlsx-001 (glm-5.2? 不，低难度) → deepseek-v4-pro
  ② docx-001 → deepseek-v4-pro
  ③ cross-001 → deepseek-v4-pro

Phase 2: 类型清理 + helpers
  ④ docx-002 → glm-5.2

Phase 3: controller 改写
  ⑤ xlsx-003 → glm-5.2
  ⑥ docx-003 → glm-5.2

Phase 4: 集成验证（controller → engine）
  ⑦ xlsx-integ → deepseek-v4-pro
  ⑧ docx-integ → deepseek-v4-pro

Phase 5: 渲染层改写
  ⑨ xlsx-004 → glm-5.2
  ⑩ docx-005 → deepseek-v4-pro（简单只读 viewer）
  ⑪ docx-004 → glm-5.2
  ⑫ xlsx-005 → glm-5.2

Phase 6: demo 接入
  ⑬ xlsx-006 → deepseek-v4-pro
  ⑭ docx-006 → deepseek-v4-pro

Phase 7: 全链路集成验证
  ⑮ cross-integ → deepseek-v4-pro
```

## 每个任务的执行流程

```
1. 创建分支 task/{id}
2. dispatch develop agent（指定模型）
   - prompt: "你的任务文件是 docs/plan/tasks/{id}.md。阅读任务文件，只完成 objective 中定义的开发目标，不超出任务范围。context 指向的设计文档是你理解需求的来源。docs/INDEX.md 是文档总索引。完成后运行 typecheck 和任务要求的测试。"
3. develop 完成后 commit
4. dispatch verify agent（用 glm-5.2 做高难度任务的 review，deepseek-v4-pro 做低难度）
   - prompt: "review 任务 {id} 的开发产出。任务文件：docs/plan/tasks/{id}.md。开发产出：任务文件中 path 指向的路径。设计文档：任务文件中 context 指向的路径。详细阅读源码与设计文档，判断实现是否达到可交付状态。写入 docs/plan/reviews/{id}-1.md。"
5. verify pass → merge to main → 全量 test → 标记 done
   verify blocked → fix → re-verify
```

## 模型 dispatch 参数

### 高难度任务（glm-5.2）

```
agent create:
  providerId: dimcode-api-oauth
  modelId: glm-5.2
  forkContext: none  # 任务文件已自包含
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
| 🔴 高（xlsx-005, docx-003, docx-004） | `glm-5.2` | 需要深度理解判断实现是否正确 |
| 🟡 中（xlsx-003, xlsx-004, docx-002） | `glm-5.2` | 需要判断 hook 改写/JSX 转换是否语义正确 |
| 🟢 低（其余） | `deepseek-v4-pro` | 按清单检查，机械验证 |

## 分支管理

```
master（主线）
  ├── task/xlsx-001 → merge → done
  ├── task/docx-001 → merge → done
  ├── task/cross-001 → merge → done
  ├── task/docx-002 → merge → done
  ├── task/xlsx-003 → merge → done
  ├── task/docx-003 → merge → done
  ├── task/xlsx-integ → merge → done
  ├── task/docx-integ → merge → done
  ├── task/xlsx-004 → merge → done
  ├── task/docx-005 → merge → done
  ├── task/docx-004 → merge → done
  ├── task/xlsx-005 → merge → done
  ├── task/xlsx-006 → merge → done
  ├── task/docx-006 → merge → done
  └── task/cross-integ → merge → done
```

每个任务：创建分支 → develop → commit → verify → pass/merge 或 blocked/fix。

## 中断恢复

中断后从 `docs/plan/tasks/` 的 status 字段恢复：

| 任务 status | git 状态 | 恢复动作 |
|---|---|---|
| pending | 无分支 | 等待依赖完成 |
| ready | 无分支 | 创建分支，dispatch develop |
| in-progress | 分支存在，无 review | 继续 develop 或重新 dispatch |
| in-progress | 分支存在，有 review blocked | dispatch fix |
| in-progress | 分支存在，有 review pass | 执行 merge |
| in-progress | 分支不存在 | 重置为 ready，重新 dispatch |

## 全量 gate（每个任务 merge 后运行）

```bash
pnpm typecheck && pnpm build && \
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && \
git diff --check
```

如果 gate 失败，revert merge，任务重新 rebase + re-verify。
