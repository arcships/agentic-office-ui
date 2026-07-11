# xlsx-core-types-split Review 1

Date: 2026-07-08
Reviewer: DimCode
Status: **pass**

## Summary

将 `packages/xlsx-core/src/types.ts`（1307 行）拆分为 `types/` 目录下的 3 个子模块 + barrel。typecheck 通过，无 React 残留，无 stub/mock/fake，无循环依赖，所有导入路径正确。

## Findings

### P2 — worksheet-types.ts 实际行数高于架构估算（non-blocking）

**位置**: `packages/xlsx-core/src/types/worksheet-types.ts:1-958`
**设计文档**: `docs/xlsx-migration-architecture.md` 2.1 节，worksheet-types.ts 估算 ~250 行

实际拆分为 958 行，原因是架构估算将 cell style inputs（200+行）、controller interfaces（300+行）、thumbnails（50+行）、render props（30+行）误归入 chart/image 模块。这些类型均为工作表/控制器领域，留在 worksheet-types.ts 是正确的。

拆分后的总行数：worksheet-types.ts (958) + chart-types.ts (225) + image-types.ts (196) + index.ts (117) = 1496 行，相比原始 types.ts 的 1307 行净增 189 行（来源：跨文件 import 语句 +  barrel 文件 + section 注释）。所有单文件均 ≤1000 行，满足硬约束。

### P2 — XlsxCellAddress/XlsxCellRange 在原文件中的位置被调整（non-blocking）

**位置**: `packages/xlsx-core/src/types/worksheet-types.ts:13-21`
**设计文档**: 未规定类型声明顺序

在拆分过程中，`XlsxCellAddress` 和 `XlsxCellRange` 从原始位置（`XlsxSheetData` 之后）移至文件顶部（`XlsxThemePalette` 之前），与其他类型间增加 import 语句配合。TypeScript 类型声明顺序不影响语义，不构成功能问题。

## Verification Checklist

| 检查项 | 结果 | 备注 |
|---|---|---|
| 文件清单覆盖架构设计 | ✅ | `types/chart-types.ts`, `types/image-types.ts`, `types/worksheet-types.ts`, `types/index.ts` 全部就位 |
| 旧 types.ts 已删除 | ✅ | `packages/xlsx-core/src/types.ts` 不存在 |
| 功能对齐上游 f285a1c | ✅ | 所有数据模型类型已迁移，React 组件 prop 类型（XlsxViewerProps, XlsxViewerProviderProps）已正确移除 |
| React 残留 | ✅ | types/ 目录零 `React.` 引用，`React.CSSProperties` → `Record<string, string | number | undefined>`，`React.ReactNode` → `unknown` |
| import 路径 | ✅ | 6 个消费方（charts/images/colors/xlsx-worker/worker-client/index）均引用 `"./types"`，解析为 `./types/index.ts` |
| typecheck | ✅ | `pnpm --filter @arcships/xlsx-core typecheck` 通过，零错误 |
| stub/mock/fake | ✅ | 无残留 |
| 单文件 ≤1000 行 | ✅ | max 958（worksheet-types.ts） |
| 循环依赖 | ✅ | `chart-types.ts → image-types.ts`, `worksheet-types.ts → chart-types.ts + image-types.ts`，无反向依赖 |

## Conclusion

**pass** — 拆分符合架构设计要求，typecheck 通过，零阻塞性问题。
