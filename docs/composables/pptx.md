# PPTX 组合函数

## `usePptxDocument`

```ts
const stage = ref<HTMLElement | null>(null)
const document = usePptxDocument(stage, {
  source: () => file.value,
  initialSlide: 0,
  session: {
    fitMode: "contain",
    renderMode: "list",
    listOptions: { windowed: true },
  },
})
```

它持有唯一文档会话，公开：

- `state`、`error`、`document`、`capability`；
- `activeIndex` 和 `zoomPercent`；
- `open()`、`close()`、`goTo()`、`nextSlide()`、`previousSlide()` 和 `setZoom()`；
- `slideError`（当前页渲染失败记录）和 `clearSlideError()`；
- 高级逃生出口 `getSession()`。

`state` 为 `idle`、`waiting-for-stage`、`loading`、`ready`、`cancelled`、`error` 或 `disposed`。舞台尚未挂载时进入 `waiting-for-stage`。快速更换舞台或来源时销毁旧会话，只保留最新加载结果。

`openTimeoutMs`（默认 30000，`0` 关闭）为舞台等待和加载设置超时；超时后进入 `error` 并报告 `TIMEOUT`。被新请求取代或取消的加载（`ABORTED`、`STALE_RESULT`）进入 `cancelled`，不会无限停留在 `loading`。

`goTo()` 失败时仍然抛错，但单页渲染失败不再把 `state` 切到 `error`：失败记录在 `slideError` 中，上一页保持可见，宿主可以重试或跳过；只有加载失败才会进入 `error`。

`open(source, { initialSlide })` 可以覆盖本次打开的初始页。舞台元素变化或同一来源重新打开时，组合函数保留当前 `activeIndex`。

`renderMode: "list"` 纵向连续渲染全部幻灯片，适合普通 Surface 浏览；`goTo()`、`nextSlide()` 和 `previousSlide()` 在该模式下滚动定位页面。`renderMode: "slide"` 只渲染当前页，供播放控制器使用。`PptxPreviewSession` 默认使用列表模式，`PptxDocumentSession` 默认使用单页模式；自定义组合时建议显式指定，避免把播放控制器接到列表 Surface。

## `usePptxPlayback`

```ts
const playback = usePptxPlayback(document, {
  autoplay: false,
  skipHiddenSlides: true,
  onEvent(event) {
    if (event.type === "error") console.error(event.error)
  },
})
```

`enabled` 和 `skipHiddenSlides` 可以传值、`ref` 或 getter；值变化时重建播放控制器，并从当前页继续。

它必须接收 `usePptxDocument` 的返回值，并为当前文档会话持有唯一播放控制器。公开状态包括 `controller`、`snapshot`、`status`、`capability`、`lastWarning` 和 `lastError`。

播放用的文档组合函数应配置 `session: { renderMode: "slide" }`。只有单页舞台执行原文件动画和页面切换；列表模式用于静态浏览，不执行播放时间轴。

公开方法包括 `next()`、`previous()`、`activateObject()`、`play()`、`pause()`、`resume()`、`reset()`、`goToSlide()` 和 `resumeBlockedMedia()`。

## 点击规则

```ts
async function onStageClick(event: MouseEvent) {
  const object = (event.target as Element | null)
    ?.closest<HTMLElement>("[data-pptx-object-key]")
  const handled = object?.dataset.pptxObjectKey
    ? await playback.activateObject(object.dataset.pptxObjectKey)
    : false
  if (!handled) await playback.next()
}
```

对象动作优先。只有对象没有处理点击时才执行普通下一步，否则一次点击可能同时触发对象动作和下一动画。

演示过程中使用 `playback.goToSlide()`；`document.goTo()` 只用于没有播放控制器的纵向浏览模式。不要混用两套跳页方法。
