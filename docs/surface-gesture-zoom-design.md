# Surface 手势缩放设计

> 状态：待评审，未实现
>
> 范围：`DocxDocumentSurface`、`PdfSurface`、`PptxStage`、`XlsxSheetSurface` 和宿主 `OfficeFilePreview`

## 1. 结论

手势识别和指针锚点保持放在各格式 Surface；`OfficeFilePreview` 继续持有统一 zoom state、toolbar 和切文件重置逻辑。

```text
OfficeFilePreview
├─ zoom state + toolbar
└─ Surface
   ├─ 接收 Ctrl+wheel / trackpad pinch
   ├─ 保持指针下的内容位置
   └─ 请求宿主更新 zoom
```

不在 file panel 最外层实现通用监听。DOCX、PDF、PPTX 的页面布局不同，XLSX 还是带冻结区的二维 canvas grid，只有 Surface 知道应调整哪个滚动容器和内容坐标。

## 2. 统一合同

公开 zoom 使用倍率，`1 = 100%`：

```ts
interface SurfaceZoomProps {
  /** 受控倍率；传入后才启用手势缩放。 */
  zoom?: number

  /** 默认 true。 */
  enableGestureZoom?: boolean
}

interface SurfaceZoomEmits {
  (event: "update:zoom", nextZoom: number): void
}
```

推荐接入：

```vue
<FormatSurface v-model:zoom="zoom" />
```

第一阶段统一范围为 `0.5–2`，按钮步进保持 `0.25`，手势保留 `0.01` 精度。这个范围是四种格式当前共同安全区间，不要求同时改造 PDF 高倍率渲染。

兼容规则：

- 未传 `zoom` 时保持现有行为，不拦截缩放手势；
- `zoom` 与旧接口同时存在时，以 `zoom` 为准；
- `zoom` 与 `fitWidth` 同时存在时，以 `zoom` 为准；
- 旧 `zoomScale`、`defaultZoom` 和 expose `zoom` 暂时保留，迁移完成后再单独讨论删除。

## 3. 手势行为

### 3.1 `Ctrl + wheel`

- Surface 根节点使用 `{ passive: false }` 监听 wheel；
- 仅当 `enableGestureZoom !== false`、传入有效 `zoom` 且 `event.ctrlKey` 时处理；
- 识别为缩放后调用 `preventDefault()`；普通 wheel 继续原生滚动；
- 连续事件从最近一次请求值累计，避免受控 prop 回写稍慢时缩放停滞；
- 到达上下限后仍阻止浏览器页面缩放，但不重复 emit 相同值。

缩放倍率使用连续计算：

```ts
nextZoom = clamp(zoom * Math.exp(-deltaY * 0.005), 0.5, 2)
```

### 3.2 WebKit gesture

WebKit 场景补充监听 `gesturestart`、`gesturechange`、`gestureend`：

- start 时记录初始 zoom 和锚点；
- change 时按 `startZoom * event.scale` 更新；
- gesture active 期间忽略重复到达的 `ctrl+wheel` 计算；
- 暂不增加基于时间窗口的去重，只有真实浏览器测试证明需要时再补。

### 3.3 指针锚点

Surface 在发出 `update:zoom` 前记录指针下的逻辑位置。新 zoom 实际生效后，通过 `nextTick + requestAnimationFrame` 恢复。

```text
捕获锚点
  -> emit update:zoom
  -> 等待真实 zoom 生效
  -> 恢复滚动位置
```

只对手势发起的缩放恢复指针锚点。toolbar 或重置按钮改变 zoom 时，不复用旧手势坐标。

每次缩放使用递增 token。新手势、切文件、切 sheet 或组件卸载都会使旧 token 失效，防止异步回调修改新状态。

## 4. 格式差异

| 格式 | zoom 适配 | 锚点 | 实现位置 |
|---|---|---|---|
| DOCX | `zoom * 100 -> zoomScale` | 页码 + 页面内比例 | `DocxViewerRoot` 读写内部滚动容器 |
| PDF | 倍率直接使用 | 页码 + 页面内比例 | `PdfSurface` |
| PPTX | `zoom * 100 -> setZoom()` | slide index + slide 内比例 | `PptxStage` + `usePptxDocument` adapter |
| XLSX | `zoom * 100 -> controller.setZoomScale()` | cell + cell 内比例 | `XlsxSheetSurface` 委托 `XlsxGrid` |

### DOCX / PDF

根据指针下的 page 保存页码和页内比例。缩放后定位同一页，再修正 `scrollTop` / `scrollLeft`。

PDF 页面 slot 在 bitmap 重绘期间必须保持稳定尺寸，否则锚点会随 placeholder 跳动。本轮只修正 slot 几何，不同时引入整套可见页渲染改造。

### PPTX

PPTX 的真实 zoom owner 是异步 session。`setZoom()` 完成并回写实际 zoom 后再恢复锚点。

Stage 锚点逻辑和 renderer 必须使用同一个显式 scroll container，不通过查找最近的 `overflow:auto` 元素猜测。播放模式默认关闭手势缩放。

### XLSX

XLSX 使用 worksheet cell 作为锚点，不能套用外层通用比例公式：

- body 区保持同一 cell 和 cell 内位置；
- 冻结行列对应轴不调整 scroll；
- row header 只调整 Y，column header 只调整 X；
- 保持 active cell、selection 和编辑 input 的 keyboard focus。

## 5. 宿主接入

`OfficeFilePreview` 继续持有：

```ts
const zoom = ref(1)
const MIN_ZOOM = 0.5
const MAX_ZOOM = 2
const ZOOM_STEP = 0.25
```

四个格式统一使用 `v-model:zoom`：

```vue
<DocxDocumentSurface v-model:zoom="zoom" />
<XlsxPreview v-model:zoom="zoom" />
<PdfSurface v-model:zoom="zoom" />
<PptxSurfacePreview v-model:zoom="zoom" />
```

接入后：

- 删除 `applyFormatZoom()` 的格式 switch；
- toolbar 只修改和展示统一 zoom；
- `XlsxPreview` 转发 zoom，不额外 expose 缩放命令；
- `PptxSurfacePreview` 负责倍率与 session 百分比的异步转换；
- `[cwd, path]` 变化时由宿主重置为 `1`。

## 6. 验收

功能验收只保留以下六项：

1. 四个 Surface 都能通过 `v-model:zoom` 改变实际内容，toolbar 百分比一致。
2. `Ctrl+wheel` 和支持的 pinch 能缩放；普通 wheel 只滚动。
3. DOCX、PDF、PPTX 缩放后保持同一 page/slide 及页内位置。
4. XLSX 缩放后保持同一 cell；冻结行列场景通过。
5. 快速连续缩放、异步 PPTX 缩放和切文件不会产生迟到滚动修正。
6. Chromium/Electron 四种格式通过；Firefox 的 `Ctrl+wheel` 和 WebKit gesture 各用一个代表 Surface 验证。

测试分层：

- 共享单元测试：clamp、连续 delta、WebKit 去重、过期 token；
- 每个 Surface：普通 wheel + 一个锚点集成测试；
- XLSX：额外增加冻结行列测试；
- 浏览器测试按上述第 6 条执行，不做四格式乘三浏览器的全组合。

## 7. 实施顺序

1. 实现统一合同和共享手势计算。
2. 分别接入四个 Surface，并补最小集成测试。
3. 更新 demo，验证浏览器行为。
4. 发布 Surface 版本后修改 DimCode 宿主，删除过渡 switch。

## 8. 后续事项

以下内容不阻塞第一阶段：

- 把统一范围从 `50%–200%` 扩展到 `25%–400%`；
- PDF 高倍率下的可见页渲染和内存优化；
- 旧 zoom API 的移除版本；
- 未被真实测试触发的额外 WebKit 去重和虚拟列表恢复逻辑。
