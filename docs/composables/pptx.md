# PPTX 组合函数

## `usePptxDocument`

```ts
const stage = ref<HTMLElement | null>(null)
const document = usePptxDocument(stage, {
  source: () => file.value,
  initialSlide: 0,
  session: { fitMode: "contain" },
})
```

它持有唯一文档会话，公开：

- `state`、`error`、`document`、`capability`；
- `activeIndex` 和 `zoomPercent`；
- `open()`、`close()`、`goTo()`、`nextSlide()`、`previousSlide()` 和 `setZoom()`；
- 高级逃生出口 `getSession()`。

舞台尚未挂载时进入 `waiting-for-stage`。快速更换舞台或来源时销毁旧会话，只保留最新加载结果。

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

它必须接收 `usePptxDocument` 的返回值，并为当前文档会话持有唯一播放控制器。公开状态包括 `controller`、`snapshot`、`status`、`capability`、`lastWarning` 和 `lastError`。

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

演示过程中使用 `playback.goToSlide()`；`document.goTo()` 只用于没有播放控制器的浏览模式。不要混用两套翻页方法。
