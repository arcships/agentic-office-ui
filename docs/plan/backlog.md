# Backlog

非阻塞 findings 记录。verify 阶段产出的 non-blocking findings 追加到此文件。

---

## 2026-07-07 DOCX 迁移策略变更

- [docx-migration] 2026-07-07 — 旧方案（机械复制上游 editor.tsx 巨型文件）已回退（commit `ff149fb`）。原因：产出 editor-helpers.ts 23445 行、composables.ts 1036 行骨架、stub Vue 组件，不可维护且核心编辑/渲染能力为空。改为模块化重做，硬约束单文件 ≤1000 行。架构设计见 `docs/docx-migration-architecture.md`。旧任务 docx-001~006 + docx-integ 全部失效。

---

<!-- 格式：-->
<!-- - [task-id] 日期 — finding 描述 — 位置 -->
