# PPTX 播放开发准备与入口

## 1. 当前状态

- 纵向连续静态预览、窗口化页面挂载和 Surface 交互事件已经实现；
- 三段播放探索和三个定向验证已经完成；
- 正式设计、接口、兼容范围和路线图已经确定；
- 正式播放代码、Vue 演示模式、三浏览器黑盒、外部消费、性能和资源释放检查已经完成；
- 13 份正式素材和浏览器事件基准已经入库；PowerPoint 画面对照只留给发布前独立兼容复核；
- 两个 PPTX 包已经公开；后续变更仍须通过工作区外压缩包消费和三浏览器验收。

## 2. 仓库分区

| 目录 | 用途 | 规则 |
|---|---|---|
| `packages/pptx-core/` | 正式类型、解析、时间、控制器和浏览器适配 | 不导入探索台代码 |
| `packages/vue-pptx/` | 正式静态浏览与演示组件 | 不解析 XML，不另建播放状态 |
| `apps/pptx-playback-lab/` | 已完成的探索代码和人工复现实验 | 保持私有，不作为产品依赖 |
| `tests/fixtures/pptx/playback/` | PowerPoint 原生验收素材和清单 | 敏感文件不得提交 |
| `tests/unit/` | 对象身份、时间树、属性轨道和纯净性测试 | 不依赖页面元素 |
| `tests/component/` | Vue 生命周期和控制器接入测试 | 使用可控会话替身 |
| `docs/` | 正式设计、接口、兼容范围和路线图 | 冲突时以正式实现设计为准 |

## 3. 常用命令

```bash
# 查看已完成的探索结果
pnpm dev:pptx-lab

# 只检查 PPTX 正式包和探索台
pnpm typecheck:pptx

# 构建两个正式包、探索台和正式 demo
pnpm build:pptx

# 构建、检查仓库边界并运行 PPTX 单元和组件测试
pnpm test:pptx

# 运行 Chromium、Firefox、WebKit 黑盒与内存检查
pnpm test:pptx:blackbox

# 重新采集 13 份正式素材的 Chromium 事件基准
pnpm capture:pptx-baselines
```

完整仓库仍使用 `pnpm typecheck`、`pnpm build` 和原有测试套件。`test:pptx` 只用于缩短日常开发反馈时间，不能替代正式浏览器和发布验收。

## 4. 第三方渲染器补丁

当前固定 `@aiden0z/pptx-renderer@1.2.4`。需要增加稳定对象标记时：

1. 使用 `pnpm patch @aiden0z/pptx-renderer@1.2.4` 创建临时修改目录；
2. 只修改对象和段落标记所需位置；
3. 使用 `pnpm patch-commit <临时目录>` 生成仓库补丁；
4. 确认补丁登记、锁文件和全新安装结果；
5. 运行 `pnpm test:pptx`；
6. 同时准备上游改动，但正式开发不等待上游发布。

没有实际补丁前不创建空补丁文件，也不修改第三方包缓存。

## 5. 素材规则

`tests/fixtures/pptx/playback/manifest.json` 是正式最小素材清单。13 份素材均为 `ready`，每份都包含：

- 桌面版 PowerPoint 版本；
- 来源和使用许可；
- SHA-256；
- 预期对象、单击次数和事件顺序；
- 浏览器解析、时间安排、事件和能力报告基准。

仓库还包含一份代码回归用的组合素材 `tests/fixtures/pptx/playback-controlled.pptx`。它覆盖普通单击、对象单击、有限和无限重复、自动往返、透明度强调、直线运动、旋转、对象擦除、淡化、页面推进和擦除、自动换页、隐藏页、内部跳页、外部链接、强身份 Morph、音视频命令、媒体裁剪、循环、音量、媒体书签、同属性冲突和文字逐段播放；SHA-256 为 `0e98ea7bd47b8c5150a25d11836d2da1431a894b7634897c750d405f7a99f19a`。可用下面的命令重新生成：

```bash
python3 scripts/fixtures/generate-pptx-playback-fixtures.py
```

这份文件只负责稳定回归，不替代清单中由 PowerPoint 创建并人工核对的正式素材。

业务文件、临时派生文件和没有授权的公开样例不能直接复制到仓库。它们只记录哈希、页数、大小和验证结果。

## 6. 后续扩展入口

新增动画或播放能力时，按下面的顺序扩展：

1. 在 `pptx-core` 增加解析结果和能力报告记录；
2. 在时间安排和属性轨道中定义确定的执行规则；
3. 增加正式素材和浏览器事件基准；需要扩大严格兼容声明时，再安排独立 PowerPoint 画面对照；
4. 补单元、组件和三浏览器黑盒检查；
5. 保持纵向静态预览、Surface 事件接口和未支持内容的静态显示行为不回退。
