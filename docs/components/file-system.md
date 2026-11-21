# FileSystem：Finder 风格文件浏览（#179）

> 适用版本：0.8.1。上游参考：`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`
> `apps/v4/components/extend/file-system.tsx`（5,381 行，MIT）。
> 组件入口：`@arcships/vue-ui` 的 `FileSystem`；纯模型入口：`buildFileSystemIndex` 等（见下文）。
> 没有像素基线：本 block 不参与 `ui:parity` 门禁，差异见「与上游的差异」。

## 1. 交付范围

| 能力 | 状态 | 说明 |
|---|---|---|
| 面包屑导航 | 完成 | 根标签 + 祖先文件夹按钮，`aria-current="page"`，点击任意层级回到该目录 |
| 前进/后退历史 | 完成 | 内部历史栈；`Back`/`Forward` 在首/末项禁用 |
| 列表 / 网格视图 | 完成 | 受控 `view` 或内部状态；`aria-pressed` 分段按钮 |
| 缩略图 | 完成 | 图片条目经 `FileThumbnail` 渲染；文件夹与无预览文件用 `OfficeIcon` glyph |
| 选择与打开 | 完成 | 单击选择、双击打开；`selection-change`、`open-folder`、`open-file` |
| 键盘可达性 | 完成 | 见第 5 节 |
| 空 / 加载 / 错误状态 | 完成 | `role="status"` / `role="alert"` |
| 只读模式 | 完成 | 隐藏删除操作，其余浏览能力保留 |
| 删除（宿主驱动） | 完成 | `delete-item` 事件；组件从不修改 `items` |
| 排序 / 筛选 / 搜索 | **未做** | 上游有排序、文件类型筛选与搜索字段；宿主可自行过滤 `items` |
| columns / gallery 视图 | **未做** | 只移植 Finder 的列表与网格 |
| 拖放移动 | **未做** | 上游文件导航本身也没有拖放移动，本移植同样不做 |
| 重命名 / 新建文件夹 | **未做** | 需要命名对话框，未在本 issue 范围内 |
| 内嵌 PDF/DOCX/PPTX/XLSX 预览对话框 | **未做** | 组件只发 `open-file`，预览由宿主决定 |
| `DocumentPreviewCard` 悬浮预览 | **未做** | 仍为 `reference-only`，见 `docs/upstream/extend-ui.md` |

## 2. 数据模型

显式嵌套树，而不是上游按路径推断的扁平 manifest：

```ts
interface FileSystemItem {
  id: string            // 稳定 id，也是选择/导航事件的值
  name: string          // 显示名，不参与解析
  kind: "folder" | "file"
  children?: FileSystemItem[]   // 仅 folder 使用
  size?: number         // 字节，列表视图显示
  mime?: string         // 驱动 glyph 与图片判断
  thumbnailUrl?: string // 宿主生成的预览图（缩略图服务、首页渲染、data URL）
  url?: string          // 打开/下载 URL，原样放进 open-file
  updatedAt?: string    // ISO 时间，列表视图显示
}
```

为什么改模型：宿主本来就拥有文件夹结构（S3 listing、上传清单、网盘目录）；按路径字符串推断会丢空文件夹，也会让同名子项变得有歧义。空文件夹（`children: []`）会被保留并显示空态。

`buildFileSystemIndex()` 输出 `nodes`（id → 节点）、`children`（文件夹 id → 子 id，已排序）、`duplicateIds` 与 `warnings`。排序规则：文件夹在前、文件在后，同组内 `localeCompare(..., { numeric: true, sensitivity: "base" })`。

## 3. API

### Props

| Prop | 类型 | 默认 | 说明 |
|---|---|---|---|
| `items` | `FileSystemItem[]` | — | 根条目，必填 |
| `currentFolderId` | `string \| null` | 受控 | 当前文件夹 id；`null` 为根。传入即为受控 |
| `defaultFolderId` | `string \| null` | `null` | 非受控初始文件夹 |
| `view` | `"list" \| "grid"` | 受控 | 传入即为受控 |
| `defaultView` | `"list" \| "grid"` | `"grid"` | 非受控初始视图 |
| `selectedId` | `string \| null` | 受控 | 受控选择 |
| `title` | `string` | 文案 `Files` | 根面包屑标签 |
| `readOnly` | `boolean` | `false` | 隐藏删除按钮并忽略 Delete 键 |
| `loading` | `boolean` | `false` | 显示加载状态 |
| `error` | `string` | — | 显示 `role="alert"` 错误并替换内容 |
| `emptyMessage` | `string` | 文案 `This folder is empty.` | 空态文案 |
| `className` | `string` | `""` | 追加到根元素 |

悬空/已删除的 `currentFolderId` 会解析到根，而不是渲染坏面包屑。

### 事件

| 事件 | 载荷 | 触发 |
|---|---|---|
| `update:currentFolderId` | `string \| null` | 面包屑、打开文件夹、前进/后退、`navigate()` |
| `update:view` | `"list" \| "grid"` | 视图切换 |
| `update:selectedId` | `string \| null` | 选择变化（含导航后清空） |
| `selection-change` | `FileSystemItem \| null` | 与 `update:selectedId` 同时触发，带完整条目 |
| `open-folder` | `FileSystemItem` | 双击、Enter/Space、ArrowRight 打开文件夹 |
| `open-file` | `FileSystemItem` | 双击、Enter/Space 打开文件 |
| `delete-item` | `FileSystemItem` | 删除按钮或 Delete 键（`readOnly` 时不触发） |

### 模板引用

```ts
const fs = ref<InstanceType<typeof FileSystem> | null>(null)
fs.value?.navigate("folder-invoices") // 等价于点击面包屑
fs.value?.currentFolder()
```

## 4. 状态与 ARIA

- 根元素：`role="group"` + `aria-label`（标题），`data-slot="file-system"`，`data-view`、`data-current-folder-id`（根为空串）、`data-read-only`。
- 工具栏：`data-slot="file-system-toolbar"`；前进/后退是带 `aria-label` 的 `OfficeIconButton`，无历史时原生 `disabled`。
- 面包屑：`<nav aria-label>` + `<ol>`，每级是 `<button>`，当前级 `aria-current="page"`。
- 视图切换：`role="group"` 内两个 `aria-pressed` 按钮。
- 列表/网格容器：`role="list"`，每项 `role="listitem"`；条目主体是真实 `<button>`，选中时 `aria-current="true"`。
- 空/加载：`role="status"`；错误：`role="alert"`。
- 行内按钮使用 roving `tabindex`：只有当前项（选中项，否则第一项）是 `0`，其余 `-1`。

## 5. 键盘

焦点在条目主体按钮上：

| 键 | 行为 |
|---|---|
| `ArrowDown` / `ArrowUp` | 在可见条目间移动选择并把焦点带到目标行 |
| `Home` / `End` | 选择第一项 / 最后一项 |
| `ArrowRight` | 打开文件夹（文件忽略） |
| `ArrowLeft` / `Backspace` | 回到上一级文件夹（根目录忽略） |
| `Enter` / `Space` | 打开当前条目（文件夹导航、文件触发 `open-file`） |
| `Delete` | 触发 `delete-item`（`readOnly` 时忽略） |

网格视图没有二维方向键移动：`ArrowDown`/`ArrowUp` 仍是线性切换，`ArrowRight` 进入文件夹、`ArrowLeft` 返回上级。
组件拿不到可靠的网格列数，这一点与 Finder 不同，已在第 7 节登记。

## 6. 响应式

没有像素基线，按 390 / 768 / 1440 手工核对（demo `/file-system`）：

- `≤640px`：根高度 420px，工具栏间距收紧，列表行隐藏右侧大小/日期，网格最小列宽 120px。
- `768–1023px`：网格最小列宽 132px，工具栏保持单行。
- `≥1024px`：默认布局，网格 `minmax(148px, 1fr)` 自适应。

## 7. 与上游的差异（明确登记）

1. **文件树来源**：上游用 `@pierre/trees` 的 `FileTree`（Apache-2.0）绘制侧边树；本移植不引入该依赖，用自建索引 + 列表/网格渲染。视觉与 DOM 结构不同，**不做像素对齐**。
2. **视图集合**：只有 `list` / `grid`（上游四个：icons / list / columns / gallery）。
3. **工具栏**：没有排序选择、文件类型筛选、搜索输入（上游有）。
4. **没有内嵌预览**：上游在文件打开对话框里懒加载 PDF/DOCX/PPTX/XLSX viewer；这里只发事件。
5. **没有重命名/新建/拖放**：删除是唯一的写入型操作，且也是宿主驱动的。
6. **键盘语义**：上游依赖 `@pierre/trees` 的键盘模型；这里是列表语义（见第 5 节），网格视图也没有二维方向键移动。
7. **选择模型**：单选。上游列表/网格同样是单选，但没有 `aria-current` 合同。
8. **图标**：默认轨道 `OfficeIcon`（`@lucide/vue`），未使用 Hugeicons，因此没有 `tests/icons/allowlist.json` 条目。

## 8. 测试与证据

```bash
pnpm --filter=@arcships/vue-ui build
node --test tests/component/file-system.test.mjs
```

`tests/component/file-system.test.mjs`（14 个用例，happy-dom）覆盖：

- 纯模型：排序（文件夹优先）、祖先链、大小格式化、预览 URL、重复 id；
- 渲染：根面包屑、默认网格、`role`/`aria` 合同；
- 交互：视图切换、单击选择、双击导航、面包屑回根、前进/后退栈；
- 打开文件事件、`FileThumbnail` 预览图；
- 键盘：方向键/Home/End/Enter/Space/ArrowRight/ArrowLeft/Backspace/Delete；
- 状态：空（含自定义文案）、加载、错误；
- 只读：隐藏删除、Delete 无效、无事件；
- 受控 `currentFolderId`/`selectedId` 与悬空 id 回退；
- 依赖与断点守卫：源码不出现 `@pierre/trees`、`package.json` 无该依赖、`640/768/1023px` 断点仍在。

Demo：`apps/demo` 的 `/file-system`（`FileSystemPage.vue`），fixtures 在
`apps/demo/src/lib/file-system-fixture.ts`。

## 9. 归属

- `packages/vue-ui/src/components/FileSystem.vue`、`packages/vue-ui/src/file-system.ts` 顶部保留 Extend MIT 头；
  对应行登记在 `docs/upstream/attribution.json` 与 `docs/upstream/extend-ui.md`（状态 `adapted`）。
- 未引入 `@pierre/trees`；`FileThumbnail` 的 Lucide 图标属于既有默认轨道。
