# SchemaBuilder：抽取 schema 编辑器（#179）

> 适用版本：0.8.1。上游参考：`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`
> `apps/v4/components/extend/schema-builder.tsx`（2,703 行，MIT）。
> 组件入口：`@arcships/vue-ui` 的 `SchemaBuilder`；模型入口：`serializeSchemaBuilderSchema`、
> `parseJsonSchema`、`validateSchemaBuilderSchema` 等（见第 3 节）。
> 没有像素基线：本 block 不参与 `ui:parity` 门禁，差异见第 7 节。

`docs/design/viewer-editor-architecture.md` §10 指出的真正缺口不是许可证，而是**没有 schema 持久化契约**。
这个移植把它补成公开 API：一个 JSON 安全的字段树、确定性 JSON Schema 序列化、宽容导入与结构校验。

## 1. 交付范围

| 能力 | 状态 | 说明 |
|---|---|---|
| 字段行：名称 / 类型 / 必填 / 描述 | 完成 | 类型：`string`、`number`、`boolean`、`object`、`array` |
| 嵌套 object 字段 | 完成 | 递归行，`properties` |
| object 数组字段 | 完成 | `type: "array"` + `itemType: "object"` + `properties` |
| 标量数组字段 | 完成 | `itemType`：`string` / `number` / `boolean` |
| 添加 / 删除 / 上移 / 下移 | 完成 | 按钮排序，**不引入 dnd-kit** |
| JSON Schema 导入 | 完成 | `required`、`description`、嵌套 `properties`、`items`；不支持的关键字给 warning |
| JSON Schema 导出 | 完成 | “Show JSON” 面板 + “Copy JSON” 触发 `export` 事件 |
| 校验 | 完成 | 空名、同级重名、非法类型；行内 + 汇总 `role="alert"` |
| 只读 | 完成 | 隐藏所有变更控件、输入 `readonly`、类型选择禁用；JSON 导出保留 |
| 空 / 加载 / 错误状态 | 完成 | `role="status"` / `role="alert"` |
| enum / integer / null 类型 | **未做** | 导入时 `integer → number`、`null → string`，`enum` 只给 warning |
| 拖拽排序 | **未做** | 用上移/下移按钮替代，避免 dnd-kit 与嵌套拖拽可达性成本 |
| `@pierre/diffs` JSON 视图 | **未做** | 用 `<pre>` + `--office-font-mono`，避免第二套 worker pool |
| 抽取预览 / 示例值 | **未做** | 上游的预览面板依赖运行时数据，未移植 |
| 数组的数组 | **未做** | 导入时降级为 `items.type: "string"` 并给 warning |

## 2. 数据模型

```ts
type SchemaBuilderFieldType = "string" | "number" | "boolean" | "object" | "array"
type SchemaBuilderArrayItemType = "string" | "number" | "boolean" | "object"

interface SchemaBuilderField {
  id: string                 // 行 id，重命名也不变（焦点、测试、diff 都靠它）
  name: string               // JSON property name
  type: SchemaBuilderFieldType
  required?: boolean
  description?: string
  properties?: SchemaBuilderField[]   // object；array + itemType object 时为元素字段
  itemType?: SchemaBuilderArrayItemType  // array 元素类型，默认 string
}

interface SchemaBuilderSchema { properties: SchemaBuilderField[] }
```

`id` 由组件生成（`schema-field-<n>`，跳过已存在的 id）；导入器生成 `field-<n>`，因此同一文档的
round trip 结果稳定、可 diff。

## 3. API

### Props

| Prop | 类型 | 默认 | 说明 |
|---|---|---|---|
| `schema` | `SchemaBuilderSchema` | — | 必填。组件保留乐观本地副本并 `emit("update:schema")`；宿主传入新对象即可替换/重置 |
| `title` | `string` | 文案 `Schema` | 标题 |
| `description` | `string` | — | 标题下的说明 |
| `readOnly` | `boolean` | `false` | 关闭所有变更入口（含导入） |
| `loading` | `boolean` | `false` | 加载状态 |
| `error` | `string` | — | 错误状态（替换编辑器） |
| `emptyMessage` | `string` | 文案 `No fields yet.` | 空态文案 |
| `className` | `string` | `""` | 追加到根元素 |

### 事件

| 事件 | 载荷 | 触发 |
|---|---|---|
| `update:schema` | `SchemaBuilderSchema` | 任何编辑、导入 |
| `schema-change` | `SchemaBuilderSchema` | 与 `update:schema` 同时触发 |
| `validation-change` | `SchemaBuilderIssue[]` | 校验结果变化（挂载时即触发一次） |
| `export` | `SerializedSchema` | 点击 “Copy JSON”（只导出，不改模型） |

### 模板引用

```ts
const builder = ref<InstanceType<typeof SchemaBuilder> | null>(null)
builder.value?.toJSONSchema() // SerializedSchema
builder.value?.issues()       // SchemaBuilderIssue[]
```

### 公开模型函数

```ts
import {
  serializeSchemaBuilderSchema,  // 字段树 → JSON Schema（draft 2020-12 子集）
  parseJsonSchema,               // unknown → { schema, warnings }；结构错误抛 SchemaBuilderParseError
  parseSchemaBuilderJsonText,    // 同上，先 JSON.parse
  validateSchemaBuilderSchema,   // 结构校验，返回 issue 列表
  findSchemaBuilderField,        // 按 id 定位 + 同级列表
  createSchemaBuilderField,
  nextSchemaBuilderFieldName,
  createEmptySchemaBuilderSchema,
  SCHEMA_BUILDER_FIELD_TYPES,
  SCHEMA_BUILDER_ARRAY_ITEM_TYPES,
  SchemaBuilderParseError,
} from "@arcships/vue-ui"
```

导出示例（`serializeSchemaBuilderSchema`）：

```json
{
  "type": "object",
  "properties": {
    "vendor_name": { "type": "string", "description": "Legal name of the supplier." },
    "remit_to": { "type": "object", "properties": { "street": { "type": "string" } } },
    "line_items": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": { "sku": { "type": "string" } },
        "required": ["sku"]
      }
    }
  },
  "required": ["vendor_name"]
}
```

序列化是确定性的：同级按数组顺序，空名被跳过，`description` 为空时不输出，`required` 只列非空名。

### 校验

`SchemaBuilderIssue` 的 `code` 是稳定值，文案由 i18n 目录供给：

| code | 含义 |
|---|---|
| `empty-name` | 字段名为空或全空白 |
| `duplicate-name` | 同一 `properties` 作用域内重名（JSON property 语义） |
| `invalid-type` | `type` 或数组 `itemType` 不在允许集合 |

每条 issue 带 `fieldId`、`path`（点号路径，空名用 `?`）与 `name`；`invalid-type` 额外带 `type`。
行内显示在对应字段下方，汇总渲染在编辑器尾部（`role="alert"`，`data-issue-code` / `data-issue-path`）。

### 导入警告

`parseJsonSchema` 不抛异常处理“能读懂但不完全支持”的文档，而是返回 warning：
`unsupported-type`、`unsupported-item-type`、`unsupported-item-shape`、`missing-name`。
组件把它们显示在导入面板里（`role="status"`），导入仍然生效。
结构性错误（根不是对象、`properties` 不是对象、字段值不是对象）抛 `SchemaBuilderParseError`，
组件按 `code` 映射到目录文案。

## 4. 状态与 ARIA

- 根元素：`<section>`，`data-slot="schema-builder"`，`data-read-only`。
- 标题层级：`<h2>`；按钮：`OfficeButton` / `OfficeIconButton`；输入：`OfficeInput` / `OfficeSelect` / `OfficeTextarea`。
- 每个字段：`data-slot="schema-builder-field"` + `data-field-id` + `data-field-type`。
- 名称输入 `aria-invalid`（空名/重名/非法类型都会标记），类型选择 `aria-invalid`（仅类型问题）。
- 必填使用真实 `<input type="checkbox">` + 可见标签（不是 switch）。
- 移动/删除按钮始终有 `aria-label`（`Move {name} up` / `Remove {name}`），首项上移、末项下移 `disabled`。
- 只读时变更按钮不渲染（不是禁用），输入 `readonly`，选择器 `disabled`。

## 5. 键盘与焦点

- Tab 顺序：标题按钮 → 字段名 → 类型 → 必填 → 描述 → 行内操作 → 下一行（嵌套行同理）。
- 类型选择继承 Reka Select 的键盘模型（Enter/Space 打开、方向键移动、Enter 选定、Esc 关闭）。
- 新增字段后焦点落到新行的名称输入（`nextTick` 后 `focus()`）。
- 排序没有拖拽，因此没有额外的键盘别名：按钮本身可 Tab + Enter/Space。

## 6. 响应式

没有像素基线，按 390 / 768 / 1440 手工核对（demo `/schema-builder`）：

- `≤640px`：内边距收紧，工具栏换行，名称输入独占一行，行内操作按钮不靠右。
- `768–1439px`：名称弹性、类型/必填/操作同一行。
- `≥1440px`：默认布局，描述与名称同宽展开。

## 7. 与上游的差异（明确登记）

1. **拖拽排序 → 按钮排序**：上游用 `@dnd-kit/core` + `sortable`（MIT）；这里不引入，避免嵌套拖拽的键盘可达性问题与额外依赖。
2. **类型集合**：上游 5 个标量 + `object`/`array`/`enum`；这里 5 个类型。`integer` 在导入时映射为 `number`（JSON Schema 的 `integer` 更严格，但 builder 不区分），`null` 映射为 `string`，`enum` 只给 warning 并丢弃候选值。
3. **JSON 视图**：上游 `@pierre/diffs`（Apache-2.0）做虚拟化高亮；这里 `<pre>` + token 字体。
4. **没有抽取预览 / 示例值 / 抽取运行器**。
5. **没有 `@pierre/diffs` worker pool、`next-themes` 主题探测**：主题走 `--office-*` token（`OfficeThemeProvider`）。
6. **DOM 合同**：上游行标记是 `data-schema-builder-property-id` / `-row-id`；这里是 `data-field-id` / `data-field-type`（没有像素门禁，不承诺同名）。
7. **校验语义**：上游没有同级重名校验；这里把空名/重名/非法类型做成公开 issue code。

## 8. 测试与证据

```bash
pnpm --filter=@arcships/vue-ui build
node --test tests/component/schema-builder.test.mjs
```

`tests/component/schema-builder.test.mjs`（13 个用例，happy-dom）覆盖：

- 纯模型：序列化（required/嵌套/数组）、`serialize → parse → serialize` 稳定、`integer`/数组降级 warning、
  `invalid-root`/`invalid-properties` 抛错、校验三种 code、字段名生成、按 id 定位；
- 渲染：7 行（5 个根字段 + 2 个嵌套）、类型文本、必填勾选、描述值；
- 编辑：名称/描述/必填 → `update:schema`；类型改为 object 时自动建 `properties`；
- 结构：新增根字段（`field_1`、`field_2`）、对象行加嵌套字段、删除、上移/下移与首项禁用；
- 校验：空名/重名/非法类型的行内与汇总提示、`validation-change` 载荷；
- JSON：导出面板结构、“Copy JSON” 触发 `export`；
- 导入：合法文档生效（含 `integer` warning）、非法 JSON / 结构错误 / 根数组报错且模型不变、成功导入关闭面板；
- 只读：按钮消失、输入只读、选择器禁用、导出仍可用；
- 状态：空/自定义空态/加载/错误；
- 依赖守卫：源码不含 `dnd-kit` / `@pierre/diffs`，`package.json` 无这两个依赖。

Demo：`apps/demo` 的 `/schema-builder`（`SchemaBuilderPage.vue`），fixtures 在
`apps/demo/src/lib/schema-builder-fixture.ts`（含一份可直接导入的 JSON Schema 样例）。

## 9. 归属

- `SchemaBuilder.vue`、`SchemaBuilderFieldRow.vue`、`packages/vue-ui/src/schema-builder.ts` 顶部保留 Extend MIT 头；
  对应行登记在 `docs/upstream/attribution.json` 与 `docs/upstream/extend-ui.md`（状态 `adapted`）。
- 未引入 `@dnd-kit/*`、`@pierre/diffs`、`next-themes`；图标走 `OfficeIcon`（`@lucide/vue`）。
