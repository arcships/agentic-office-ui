# 查看器键盘导航

DOCX、PDF 和 PPTX 查看器共用同一套键盘映射。映射由 `@arcships/vue-ui` 的 `useViewerKeys()` 实现，`resolveViewerKeyCommand()` 是可单测的纯函数版本。快捷键只挂在组件根节点上，因此焦点在哪个查看器里，快捷键就只作用于哪个查看器。

统一规则：

- 焦点位于 `input`、`textarea`、`select` 或 `contenteditable` 中时，查看器不抢任何快捷键；工具栏页码、搜索框、PDF 表单控件保持浏览器原生键盘行为。
- `Shift`/`Alt` 加方向键仍用于选择文本或打开菜单，不触发翻页。
- 只有查看器实际处理了按键时才调用 `preventDefault()`；已到第一页/最后一页时，按键仍留给浏览器滚动。
- 查看器根节点是 `tabindex="0"` 的焦点目标，并带 `:focus-visible` 焦点环，键盘用户可以看见焦点位置；DOCX、PDF、PPTX 和 XLSX 查看器根节点都应用了该焦点环。

## 共享映射

| 按键 | 命令 | 说明 |
|---|---|---|
| `PageDown` / `PageUp` | 下一页 / 上一页 | 演示模式同样有效 |
| `Home` / `End` | 第一页 / 最后一页 | 演示模式同样有效 |
| `Ctrl/Cmd` + `+` 或 `=` | 放大 | 按组件的缩放档位前进 |
| `Ctrl/Cmd` + `-` | 缩小 | 按组件的缩放档位后退 |
| `Ctrl/Cmd` + `0` | 重置缩放 | 回到 100% |
| `Ctrl/Cmd` + `F` | 聚焦搜索 | 没有搜索界面的 Surface 不启用 |
| `Escape` | 取消 | PDF 优先清搜索，再清选择；PPTX 浏览模式清搜索，演示模式退出全屏 |
| 演示模式：`↓` / `→` / `Space` | 下一页 | 焦点在按钮等交互控件上时让给控件 |
| 演示模式：`↑` / `←` | 上一页 | |

浏览模式不占用方向键：PPTX 浏览模式把 `↑`/`↓`、`PageUp`/`PageDown`、`Home`/`End` 和 `Space` 交给幻灯片列表滚动，不会切换幻灯片。

## 各组件

### `DocxViewer`

- `PageDown`/`PageUp`：向前/向后翻一页，页码与可见页范围保持同步。
- `Home`/`End`：跳到第一页/最后一页。
- `Ctrl/Cmd` + `+`/`=`、`-`、`0`：按 50/75/100/125/150/175/200 档位缩放或重置到 100%。
- `Ctrl/Cmd` + `F`：聚焦工具栏搜索框（`showToolbar` 为 `false` 时不启用）。
- `Escape`：关闭搜索并清空查询，把焦点还给查看器根节点。

### `PdfViewer`

与 `DocxViewer` 相同的翻页、`Home`/`End`、缩放和 `Ctrl/Cmd + F` 映射。`Escape` 先取消搜索，再清除当前文本选择；`Ctrl/Cmd + P` 仍保留为打印。

### `PdfSurface`

- `PageDown`/`PageUp`、`Home`/`End`：按页滚动连续页面；连续按 `PageDown` 会以最后一次键盘目标页为基准继续前进。
- `Ctrl/Cmd` + `+`/`=`、`-`、`0`：在 50%–200% 之间缩放或重置到 100%；受控 `zoom` 通过 `update:zoom` 通知宿主。
- `Escape`：取消当前选择或区域选择。
- Surface 没有搜索栏，因此不占用 `Ctrl/Cmd + F`。

大纲面板（`showOutline`，文档有书签时默认显示）使用 WAI-ARIA tree 键盘语义：`↓`/`↑` 在可见条目间移动焦点，`→` 展开当前条目或进入第一个子条目，`←` 收起当前条目或回到父条目，`Home`/`End` 跳到第一个/最后一个可见条目，`Enter`/`Space` 激活条目并跳到目标页。焦点在大纲条目上时，这些键不会传给页面的翻页快捷键。

### `PptxViewer`

- 浏览模式：`PageDown`/`PageUp` 按视口高度滚动幻灯片列表，`↑`/`↓` 按行滚动，`Home`/`End` 到列表顶部/底部，`Space` 向下翻一屏；这些键不会切换当前幻灯片。`Ctrl/Cmd` 缩放、`Ctrl/Cmd + F` 搜索和 `Escape` 关闭搜索与其它格式一致。
- 演示模式：`PageDown`/`PageUp`、`Home`/`End`、`↓`/`→`/`Space` 和 `↑`/`←` 交给播放控制器执行上一步/下一步或跳到首尾页；`Escape` 退出全屏。播放按钮等交互控件获得焦点时，`Space` 仍触发按钮本身。

## 自定义查看器

```ts
import { useViewerKeys } from "@arcships/vue-ui"

const keyboard = useViewerKeys({
  pageNavigation: true,
  zoom: true,
  search: () => showToolbar.value,
  escape: true,
  mode: () => (isPresent.value ? "present" : "browse"),
  onCommand(command) {
    // return false when the command cannot be applied at this boundary
    return applyViewerCommand(command)
  },
})
```

```vue
<div tabindex="0" @keydown="keyboard.onKeydown">
  <!-- document surface -->
</div>
```

宿主只需实现 `onCommand`；根节点作用域、可编辑目标保护、修饰键规则和 `preventDefault()` 时机都由共享组合函数统一处理。XLSX 单元格网格的键盘导航仍由 [#23](https://github.com/arcships/agentic-office-ui/issues/23) 单独跟踪。
