# @arcships/vue-pptx

## 当前状态

本包提供静态浏览和 `mode="present"` 演示模式，从 `0.3.0` 起公开发布。演示模式使用 `@arcships/pptx-core/browser` 的文档会话和播放控制器，不自行解析动画。

```bash
pnpm add @arcships/pptx-core@0.3.0 @arcships/vue-pptx@0.3.0
```

```ts
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
```

## 组件边界

Vue 包负责：

- 浏览和演示界面；
- 控制条、键盘、鼠标、触摸和全屏；
- 加载、媒体恢复、错误和能力提示；
- 将公开控制器暴露给组件调用方。

Vue 包不负责：

- 解包 PPTX；
- 解析动画 XML；
- 计算时间节点；
- 合并动画属性；
- 保存独立于控制器的第二份播放状态。

现有 `PptxViewer` 默认浏览行为保持兼容。传入 `mode="present"` 后可使用下一步、上一步、暂停、继续、重播、跳页、媒体恢复和全屏；模板引用也暴露同一组公开方法。
