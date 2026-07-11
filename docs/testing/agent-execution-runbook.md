# Agent 端到端与黑盒验收执行手册

> 文档状态：当前执行手册。
>
> 本文说明 Agent 如何运行测试和保存证据；测试范围与断言以[端到端与黑盒验收测试方案](../end-to-end-blackbox-test-plan.md)为准。

## 1. 执行原则

1. 黑盒执行者是验证者，不是本轮实现者；发现产品缺陷后记录证据，不在同一次运行中修改产品代码。
2. 只通过正式构建页面、公开包入口、用户控件、下载文件和公开浏览器行为判断结果。
3. 必须先完成页面侦察，再选择稳定定位方式和执行操作。
4. 动态页面必须等待网络空闲或明确业务就绪状态后再读取 DOM、截图和判断。
5. 每个用例使用新的浏览器 context；需要验证同一会话恢复时才复用。
6. 不得把 DOM 节点存在、方法存在、页面写着 `READY` 当作功能通过。
7. 不得忽略 Vue warning、page error、未处理 Promise、关键资源 404、Worker 回退和错误 MIME。
8. 不得覆盖用户已有未提交改动，不运行破坏性 Git 命令。
9. 失败最多重试一次，首次证据必须保留。
10. 结束时停止自己启动的 preview、Chrome、临时服务器和 Worker，删除临时目录。

## 2. 开始前必读

执行者按顺序阅读：

1. [端到端与黑盒验收测试方案](../end-to-end-blackbox-test-plan.md)；
2. [稳定化整改路线图](../plan/stabilization-roadmap.md)中本任务和对应 `BB-*` 套件；
3. [架构审查与目标设计](../architecture-review-and-target-design.md)中相关边界；
4. 需要视觉判断时再读[视觉验收交接](../visual-acceptance-handoff.md)；
5. 需要组件细节时再读[组件浏览器验证计划](../component-browser-verification-plan.md)。

执行前必须知道：套件 ID、允许操作范围、基线提交、证据目录、使用材料、浏览器矩阵、预计结束条件。

## 3. 工作区只读侦察

先运行：

```bash
pwd
git status --short --branch
git rev-parse HEAD
node --version
pnpm --version
python --version
python -c "import playwright; print(playwright.__file__)"
```

记录以下信息：

- 当前绝对路径和提交号；
- 当前分支；
- 所有已修改和未跟踪文件；
- 操作系统、CPU 架构、可用内存和磁盘；
- Node、pnpm、Python、Playwright 和浏览器版本；
- 计划使用的端口是否已被占用；
- 是否存在本地残留 `dist`。

若工作区有用户改动：

- 不清理、不重置、不覆盖；
- 在 `environment.json` 中记录；
- 若测试必须基于干净提交，使用临时 clone 或工作树之外的目录；
- 临时 clone 只能证明已提交状态，不能证明未提交修改。

## 4. 当前环境的已知限制

本机 `webapp-testing` 技能说明中提到 `scripts/with_server.py`，但实际路径：

```text
/Users/eric8810/.codex/skills/webapp-testing/scripts/with_server.py
```

当前不存在。因此，执行者不得假定该脚本可用。优先使用产品提供的服务器管理工具；没有时，分别启动正式 preview 和浏览器，并保存进程或会话 ID，在结束阶段主动停止。

Python Playwright 模块可能已经安装，但它下载的 Chromium 可能不存在。启动失败时按第 7 节使用系统 Chrome/CDP；不要未经允许下载大型浏览器文件。

P1 持续集成入口使用仓库声明的固定依赖和浏览器版本。持续集成环境运行：

```bash
python -m pip install --requirement requirements-ci.txt
python -m playwright install --with-deps chromium
pnpm install --frozen-lockfile
pnpm check
```

本地已有可用系统 Chrome 时，`tests/blackbox/routes_smoke.py` 会在缺少配套 Chromium 后回退使用系统 Chrome。只有执行已登记的 P1-CI-01 或候选验证任务时，才按上述命令安装固定浏览器；安装行为和版本必须写入证据。

P1 分套执行入口为：

```bash
pnpm test:unit
pnpm test:component
pnpm test:blackbox
pnpm test:consumer
pnpm test:docs
```

`test:blackbox` 会自行管理正式 preview 和延迟服务器；`test:consumer` 会自行打包、创建工作区外消费项目并清理。不要在运行这两个命令时另占它们登记的端口。

DOCX/XLSX 界面或交互发生变化时，必须额外确认 `tests/blackbox/ux_parity.py` 已由 `test:blackbox` 实际执行。它不是可选截图任务；失败、阻断或重试后通过都会阻止把对应路线图任务标为完成。

失败时先读对应套件的 `summary.json`，再读首个 `FAIL` 步骤日志。浏览器用例继续查看该 attempt 下的 `failure.png`、`trace.zip`、`violations.json` 和网络文件；消费用例查看 `external-tgz-consumer/{typescript,vue-typescript,vite-build,browser}.log`。父套件将失败后的步骤标为 `NOT_RUN`，不要把它误记为新的产品失败。

## 5. 建立证据目录

运行 ID 建议使用 UTC 或带时区时间：

```text
output/acceptance/<commit>/<suite-id>/<YYYYMMDDTHHMMSS+0800>/
```

建立以下结构：

```text
summary.json
environment.json
commands.log
console.json
page-errors.json
network.json
notes.md
screenshots/
downloads/
traces/
performance/
```

不同 Agent 不得共用同一个运行目录。所有截图和下载文件名以用例 ID 开头，例如：

```text
DOCX-004-legal-contract-page-1.png
PACK-003-consumer-network.json
SEC-003-parent-marker.json
```

## 6. 构建前检查

当前可运行命令如下；逐条运行并保存完整输出和退出码：

```bash
pnpm install --frozen-lockfile
pnpm typecheck
pnpm build
node packages/vue-docx/tests/verify-integration.mjs
node packages/vue-xlsx/test/structure.mjs
pnpm audit --prod --audit-level moderate
git diff --check
```

解释结果时必须附带已知边界：类型和构建通过不能替代黑盒；DOCX 既有集成检查也不能替代正式 Worker 断言；发布结论必须同时看 `test:blackbox` 与 `test:consumer`。当前 Vue 类型检查覆盖所有 Vue 包和 demo，XLSX 结构检查不得再产生 Vue 生命周期 warning。

若任何命令改写了生成目录，测试结束后只清理自己创建且明确属于生成物的内容；不处理用户文件。

## 7. 启动正式 preview 和浏览器

### 7.1 启动正式 preview

必须先执行 `pnpm build`，然后启动：

```bash
pnpm --filter demo preview --host 127.0.0.1 --port 4173
```

正式发布结论必须来自 preview，不能用开发服务器代替。`DOCX-000` 和 `PACK-000` 另外要求验证 Vite 开发模式；这两项是兼容性子用例，应在独立端口、独立证据目录执行，然后仍回到正式 preview 完成套件结论。

当前工作区开发模式命令是：

```bash
pnpm dev
```

真实 tgz 消费项目使用自身安装的 Vite 和不含源码别名的独立配置，按第 11.2 节命令执行。用以下请求确认工作区正式 preview 入口可访问：

```bash
curl --fail --silent --show-error http://127.0.0.1:4173/ > /dev/null
```

如果并行执行，不同 Agent 使用不同端口，例如 4173、4174、4175，并在结果中记录。

### 7.2 首选：Playwright 自带 Chromium

先尝试最小启动：

```python
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("http://127.0.0.1:4173/#/")
    page.wait_for_load_state("networkidle")
    print(page.title())
    browser.close()
```

若报“Executable doesn't exist”，不要把产品判为失败；记录环境缺失，并进入 CDP 备用流程。

### 7.3 备用：系统 Chrome + CDP

macOS 常见 Chrome 路径：

```text
/Applications/Google Chrome.app/Contents/MacOS/Google Chrome
```

用独立临时用户目录和独立端口启动：

```bash
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --headless=new \
  --no-sandbox \
  --disable-gpu \
  --remote-debugging-port=9222 \
  --user-data-dir=/tmp/agentic-office-blackbox-chrome \
  about:blank
```

读取 WebSocket 地址：

```bash
curl --silent http://127.0.0.1:9222/json/version
```

从结果取 `webSocketDebuggerUrl`，直接传给：

```python
browser = p.chromium.connect_over_cdp(
    "ws://127.0.0.1:9222/devtools/browser/<实际标识>"
)
```

当前环境通过 HTTP 地址连接 CDP 可能受到代理影响而返回 400；使用上一步得到的完整 `ws://` 地址更稳定。

若系统没有 Chrome，记录环境阻断。只有任务权限允许并且网络可用时，才安装 Playwright 浏览器；不要把安装操作夹在候选版本证据中。

### 7.4 Firefox 与 WebKit

系统 Chrome 备用流程只能补足 Chromium 检查，不能替代 Firefox 和 WebKit。候选版本必须分别启动这两个浏览器，并使用与 Chromium 相同的路由、材料、断言和证据规则：

```python
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    for name, browser_type in (("firefox", p.firefox), ("webkit", p.webkit)):
        browser = browser_type.launch(headless=True)
        page = browser.new_page()
        page.goto("http://127.0.0.1:4173/#/")
        page.wait_for_load_state("networkidle")
        print(name, page.title())
        browser.close()
```

若固定版本的浏览器文件尚未准备好，将对应浏览器结果记为 `BLOCKED`，不要记为产品失败，也不能把候选版本判为通过。经任务负责人允许并且网络可用时，可在验收开始前安装：

```bash
python -m playwright install firefox webkit
```

持续集成环境应预装并固定浏览器版本；正式证据生成期间不得临时升级浏览器。每次提交可以只把 Chromium 作为快速门禁，但候选版本的 `BB-RELEASE` 必须包含 Chromium、Firefox 和 WebKit。

### 7.5 候选制品与同批 tgz

P4 候选只使用以下入口：

```bash
CI_RUN_ID=<run-id> pnpm test:release
```

执行顺序不能调换：先生成两次隔离构建证据，再从第一轮固定出唯一五个 tgz；外部消费与完整兼容矩阵都读取这份清单；全部通过后才生成 `candidate/`。不得在消费测试、浏览器矩阵或上传前重新运行 `npm pack` 替换文件。

关键路径如下：

```text
output/acceptance/<commit>/BB-RELEASE/<run-id>/
  release/p4-reproducible-pack/summary.json
  candidate-source/real-tgz-manifests/summary.json
  consumer/external-tgz-consumer/summary.json
  matrix/compatibility-matrix/summary.json
  candidate/candidate-manifest.json
  candidate/tgz/*.tgz
```

验收人逐项核对：候选清单的提交号；五包统一版本；外部消费、矩阵和候选目录中的 tgz SHA-256；两轮有效内容 SHA；Worker/WASM 文件 SHA；每个套件摘要。`candidate/prepare-release-artifact.mjs verify candidate` 只能作为自动复核，不能替代人工检查这些字段。

发布流程只下载 `release-candidate-<commit>` 制品并校验。五包先进入本次运行专用的临时标签，全部完整后再提升正式标签；提升中途失败时脚本恢复已经改动的旧标签并返回非零。npm 不提供跨包事务，所以任何恢复失败都必须立即暂停发布并保留日志，不能报告成功。

实现者运行 `pnpm test:release` 只属于开发后自测。候选发布前仍需全新会话从当前四份核心文档开始，复跑 `BB-RELEASE` 并独立签字。

## 8. 每个页面的侦察流程

对每个路由先执行侦察，再做交互：

1. 创建新 context，并设置视口、语言和时区；
2. 在导航前注册 console、pageerror、requestfailed、response、download 监听；
3. 打开目标路由；
4. 等待 `domcontentloaded`；
5. 等待 `networkidle`；若页面有长期连接，则等待明确的页面业务状态；
6. 截取全页图；
7. 读取可见标题、按钮、输入、选择框、Canvas、iframe 和错误区域；
8. 确认稳定定位方式后再操作；
9. 操作完成后再次等待状态稳定并截图；
10. 检查控制台、网络、下载和最终页面；
11. 关闭 context，确认资源释放。

推荐定位方式顺序：

1. `get_by_role` + 可访问名称；
2. `get_by_label`；
3. 稳定的 `data-*` 属性；
4. 明确 CSS 类；
5. 文字定位只用于稳定、唯一的文案。

不要依赖构建哈希类名、元素序号或坐标点击，除非用例本身在验证 Canvas 坐标。

## 9. 必须收集的浏览器事件

最小监听示例：

```python
console_messages = []
page_errors = []
network = []

page.on("console", lambda msg: console_messages.append({
    "type": msg.type,
    "text": msg.text,
}))
page.on("pageerror", lambda error: page_errors.append(str(error)))
page.on("requestfailed", lambda request: network.append({
    "url": request.url,
    "failure": str(request.failure),
}))
page.on("response", lambda response: network.append({
    "url": response.url,
    "status": response.status,
    "contentType": response.headers.get("content-type", ""),
}))
```

每个用例结束后至少检查：

- `console.error`、Vue warning 和未预期 warning；
- `pageerror`；
- `requestfailed`；
- HTTP 4xx/5xx；
- Worker 和 WASM 的 URL、MIME、状态；
- 下载是否出现、文件名、字节数、SHA-256；
- 页面是否仍在 loading；
- 失败后是否能恢复正常材料。

预期损坏材料可以产生受控错误，但错误必须是页面可见、内容明确且不伴随旧内容残留或未处理异常。

### 9.1 临时故障服务器

`SEC-003`、`SEC-004` 和 `BB-RACE` 需要可控制 MIME、重复响应和延迟的同源服务器。仓库脚本为 `tests/blackbox/fault_server.py`，竞态入口为 `tests/blackbox/race_workflows.py`，并已接入 `pnpm test:blackbox`。服务器只服务正式构建的 `apps/demo/dist`，由套件启动和停止；不要再把临时副本复制到证据目录执行。

```bash
pnpm build
pnpm test:blackbox
```

只排查竞态套件时可直接运行：

```bash
BLACKBOX_EVIDENCE_DIR="output/acceptance/<commit>/BB-RACE/<run-id>" \
node scripts/ci/run-python.mjs tests/blackbox/race_workflows.py
```

套件会分别为 DOCX 和 XLSX 启动仅绑定 `127.0.0.1` 的延迟服务器，结束后主动停止。若端口或浏览器无法启动，结果写为 `BLOCKED`；不能自行改产品代码绕过。

## 10. 套件执行顺序

推荐按以下顺序执行，前一层失败时避免产生误导性后续结果：

1. `BB-P0-ROUTES`；
2. `BB-DOCX-WORKER`；
3. `BB-SEC-URL`；
4. `BB-RACE`；
5. `BB-PACK-CONSUMER`；
6. `BB-PERF-XLSX`；
7. `BB-STRESS`；
8. `BB-RELEASE`。

P0 路由失败时仍可执行为定位问题所需的更小用例，但不能继续生成“发布通过”摘要。

## 11. 文件上传和下载

上传使用仓库中的真实材料绝对路径，上传前记录 SHA-256。不要只点击 sample 按钮就假设文件内容正确。

下载用例必须：

1. 在点击前注册 download 监听；
2. 保存到本次运行的 `downloads/`；
3. 记录建议文件名、实际文件名、MIME、字节数和 SHA-256；
4. 对 DOCX 用 `python-docx`、XLSX 用可独立打开的工作簿库、PDF 用 PDF 解析器重新打开；
5. 对导出再导入用新的浏览器 context，避免复用内存状态；
6. 不执行下载文件中的公式或宏。

CSV 安全用例只检查原始字节，不能在真实 Excel 中自动打开危险公式。

### 11.1 独立解析工具

当前人工验证可以使用 `uv` 临时环境，但必须记录实际版本：

```bash
uv run \
  --with python-docx \
  --with openpyxl \
  --with pypdf \
  python - <<'PY'
from importlib.metadata import version
for name in ("python-docx", "openpyxl", "pypdf"):
    print(name, version(name))
PY
```

候选版本自动化必须使用 `test-data/manifest.json` 和 `requirements-ci.txt` 登记的固定环境；临时解析器版本不同只能用于侦察，不能更新基线。

### 11.2 自动 tgz 消费流程

首选入口：

```bash
pnpm test:consumer
```

该命令先运行 `scripts/ci/pack-manifests.mjs`，再由 `scripts/ci/pack-consumer.mjs` 把 `tests/consumer/template/` 复制到系统临时目录。它检查五包 tgz/资源 SHA、普通目录安装、无工作区别名、无复制 Worker/WASM、TypeScript、Vue 类型、三份 CSS、Vite 正式构建和浏览器资源请求。临时目录在成功或失败后都会清理。下面的人工命令只用于排查自动测试自身，不能另行冒充通过结果。

```bash
pnpm install --frozen-lockfile
pnpm build

PACK_ROOT="$(mktemp -d)"
PACK_DIR="$PACK_ROOT/packs"
mkdir -p "$PACK_DIR"
for pkg in packages/docx-core packages/xlsx-core packages/vue-docx packages/vue-xlsx packages/vue-extend; do
  (cd "$pkg" && npm pack --json --pack-destination "$PACK_DIR")
done

CONSUMER="$PACK_ROOT/consumer"
cp -R apps/demo "$CONSUMER"
rm -rf "$CONSUMER/node_modules" "$CONSUMER/dist"
find "$CONSUMER/public" -maxdepth 1 -type f \
  \( -name '*.wasm' -o -name '*worker*.js' \) -delete
cd "$CONSUMER"
npm pkg delete \
  'dependencies.@arcships/docx-core' \
  'dependencies.@arcships/xlsx-core' \
  'dependencies.@arcships/vue-docx' \
  'dependencies.@arcships/vue-xlsx' \
  'dependencies.@arcships/vue-extend'
pnpm add "$PACK_DIR"/*.tgz
pnpm add --save-dev vue-tsc@3.3.6

node --input-type=module - <<'NODE'
import { writeFileSync } from "node:fs"
writeFileSync(
  "vite.consumer.config.mjs",
  `import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"

export default defineConfig({
  plugins: [vue()],
  resolve: { dedupe: ["vue"] },
  worker: { format: "es" },
})
`,
)
NODE

pnpm exec vue-tsc --noEmit
pnpm exec vite --config vite.consumer.config.mjs build
pnpm exec vite --config vite.consumer.config.mjs --host 127.0.0.1 --port 4181
pnpm exec vite --config vite.consumer.config.mjs preview --host 127.0.0.1 --port 4182
```

`find` 命令会删除复制自 demo 的 WASM/Worker 公共资源，防止它们替 tgz 掩盖缺失文件。不得再从工作区或包的 `dist` 手工复制资源到消费项目；浏览器实际请求的资源必须由 tgz 的公开入口被构建工具发现并输出。若临时项目仍使用硬编码根路径，或 tgz 没有可用的公开资源入口，`PACK-003` 应判产品失败。

临时 `vite.consumer.config.mjs` 明确去掉工作区源码别名；不得使用复制来的 `vite.config.ts` 构建或启动，否则测试的不是 tgz。自动测试使用仓库内固定最小模板 `tests/consumer/template/`，候选发布以该模板为准。

消费测试结束后回到原工作区、停止 4181/4182 服务并删除 `PACK_ROOT`；tgz 的文件清单与 SHA-256 已复制到证据目录后再删除。

### 11.3 当前公开 Worker/WASM 入口

消费项目只能使用下面的公开入口，不能猜测 dist 内部文件名或硬编码 demo 根路径。Vite 配置仍必须保留 worker: { format: "es" }。不要排除整个核心包，否则 CommonJS 依赖会绕过 Vite 的正常预构建；应用入口通过下面公开 worker?worker&url 导入注入 Worker URL。

~~~ts
import docxWorkerUrl from "@arcships/docx-core/worker?worker&url"
import docxWasmUrl from "@arcships/docx-core/assets/docx_wasm_bg.wasm?url"
import { createDocxRuntime } from "@arcships/docx-core/runtime"

const docxRuntime = createDocxRuntime({
  wasmUrl: docxWasmUrl,
  workerUrl: docxWorkerUrl,
})
~~~

应用从依赖预构建中运行时，应通过公开 worker?worker&url 入口注入 DOCX Worker URL；不要猜测内部 dist 文件名。需要自建 CDN 时，才通过 Runtime 的 workerUrl 和 wasmUrl 覆盖。

~~~ts
import xlsxWorkerUrl from "@arcships/xlsx-core/worker?worker&url"
import xlsxWasmUrl from "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm?url"
import { XlsxWorkerClient } from "@arcships/xlsx-core/runtime"

const xlsxWorker = new XlsxWorkerClient({
  wasmSource: xlsxWasmUrl,
  workerUrl: xlsxWorkerUrl,
})
~~~

Vue XLSX Viewer 的只读 Worker 路径使用同一包内默认资源。依赖预构建时还要将同一个公开 xlsxWorkerUrl 传给 controller 的 workerUrl：设置 readOnly: true 和 useWorker: true 后，浏览器必须实际请求 Worker 与 WASM，不能把 not-worker 当作通过。两个核心包也公开 ./worker 和 ./assets/*.wasm，供有明确 CDN 部署需求的宿主构建工具引用。

三个 Vue 包的样式只能从公开 CSS 子路径引入：

~~~ts
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-extend/style.css"
~~~

不得从 `dist` 路径导入样式。`@arcships/vue-extend/dist/index.css` 仅为已有应用的兼容映射，新代码一律使用 `./style.css`。本地 demo 若保留源码别名，别名只能匹配包根入口，不能把 CSS 子路径错误改写到 `src/index.ts` 下。

## 12. 安全和压力用例的保护措施

- 所有服务只绑定 `127.0.0.1`；
- 恶意 URL 只设置 `window.__blackboxMarker`，不发外网请求；
- 压缩炸弹使用低配置阈值和小文件，不制造真实大膨胀；
- 图片炸弹使用小文件加超限元数据，不分配巨大图像；
- 超时通过延迟服务器模拟，不创建无限循环；
- 每个安全用例使用独立 context 和用户目录；
- 用例完成后检查临时服务器、Chrome、Worker 和文件是否清理。

若产品尚未提供可注入低阈值，记录测试阻断并创建实现任务；不要为了执行压力用例制造危险文件。

## 13. 视觉、键盘和可访问性

视觉用例对四种视口分别保存：初始、成功、交互后、错误、恢复五类状态。文档/Canvas 动态区域与 UI 外壳分开判断。

DOCX/XLSX 还必须执行上游一致性检查：

1. 在 `environment.json` 记录官方上游仓库、固定提交 `f2ff2f90954acfed1f60b7fd070a9491d4113047` 和本地提交；
2. 分别保存上游与本地的同尺寸、同状态截图，至少覆盖 `1440×900` 和 `1280×720`；
3. 不只看颜色和间距，还要真实操作页码、缩放、缩略图、选区、格式、公式栏、编辑、撤销、冻结和只读；
4. 记录每个上游已有操作在本地的对应入口、结果和差异；没有对应入口直接记为 `FAIL`；
5. 实际产品界面必须位于首屏，运行配置、诊断和 JSON 快照放在产品界面之后；
6. 执行 `node scripts/ci/run-python.mjs tests/blackbox/ux_parity.py`，结果只能按 `PASS/FAIL/BLOCKED/FLAKY` 原样登记。

内部实现可以不同，但不能以 Vue 与 React 框架不同为理由删减用户能看到和能完成的核心路径。仅靠静态截图相似也不算通过，交互后的工作簿或文档状态必须一致。

键盘检查必须真实发送按键，至少覆盖：

- Tab、Shift+Tab；
- Enter、Space；
- Escape；
- 方向键；
- DOCX/XLSX 已声明快捷键；
- 焦点在弹层、iframe、Canvas 和编辑区之间的进入与退出。

自动扫描通过不代表无障碍通过。执行者还要检查可访问名称、焦点可见、语义角色、禁用/选中/展开状态、错误提示关联和 200% 缩放。

## 14. 性能采集

性能用例记录：

- 构建产物原始和 gzip 大小；
- 首次导航的请求瀑布；
- 从选择文件到页面可操作的时间；
- 主线程长任务；
- 大表滚动帧耗时；
- Worker 数量和寿命；
- 操作前、峰值、关闭 context 后的内存；
- 重复切换后对象 URL、监听器和缓存数量。

`P3-PERF-BASELINE-01` 固定使用 `1440×900` 视口、设备像素比 1、同一机器、同一材料和同一浏览器。第一次执行连续采集三轮，同时报告每个指标的三次原始值、中位数和最差值，只形成待批准基线；因为没有旧基线，这一轮不能给性能 PASS。基线批准并写入机器可读配置后，后续运行才执行“相对基线回退超过 10% 则失败并人工分析”。最终硬预算建立后，以机器可读预算为准。

固定环境、命令、机器配置路径、指标计算和批准流程见[性能基线执行说明](performance-baseline.md)。

## 15. 失败和重试

### 15.1 判定

| 结果 | 条件 |
|---|---|
| PASS | 所有断言满足，证据完整，无隐藏异常 |
| FAIL | 产品行为与断言不符，或有未知异常/网络失败 |
| BLOCKED | 环境或前置缺失，且有明确证据证明无法执行 |
| FLAKY | 首次失败、唯一一次重试通过 |

`BLOCKED` 不是 PASS。候选版本存在 BLOCKED 的 P0/P1 用例时不得发布。

### 15.2 重试规则

1. 单个失败最多重试一次；
2. 重试必须使用新 context，必要时重启 preview；
3. 保留首次和重试两套证据；
4. 重试通过仍登记 FLAKY；
5. 第二次仍失败就停止重试，建立缺陷；
6. 不得通过增加固定长等待掩盖竞态，必须等待明确状态。

## 16. 多 Agent 分工

四个并发槽位的推荐分工：

| 角色 | 套件 | 注意 |
|---|---|---|
| 主 Agent | 汇总、P0 路由、发布结论 | 不与子 Agent 共用证据目录 |
| Agent A | DOCX Worker、Viewer、Editor、竞态 | 独立 preview 端口和浏览器用户目录 |
| Agent B | XLSX、包体、滚动、导出、CSV | 不修改 DOCX/PDF |
| Agent C | PDF、Components、URL、安全和压力 | 所有危险输入仅 localhost |

真实发布包消费可以由最先完成的 Agent 在全新临时目录执行。最终结论必须由未参与实现的 Agent 复核。

### 16.1 分派提示模板

```text
你负责执行 {SUITE_ID}，只做黑盒验证，不修改产品代码。

先完整阅读：
- docs/end-to-end-blackbox-test-plan.md
- docs/testing/agent-execution-runbook.md
- docs/plan/stabilization-roadmap.md 中相关任务

基线提交：{COMMIT}
浏览器/端口：{BROWSER_AND_PORT}
用例范围：{CASE_IDS}
材料：{FIXTURES}
证据目录：{EVIDENCE_DIR}
禁止操作：{FORBIDDEN_ACTIONS}

从正式构建开始。记录命令退出码、console、pageerror、network、截图、下载和环境。失败最多重试一次并保留首次证据。遇到产品缺陷只报告，不修改代码。输出 PASS/FAIL/BLOCKED/FLAKY 和每个用例证据路径。
```

### 16.2 子 Agent 汇报模板

```text
套件：
提交：
环境：
开始/结束时间：
用例总数：
PASS / FAIL / BLOCKED / FLAKY：
P0 / P1 / P2 / P3 缺陷：
首次失败与重试：
关键 console/pageerror/network：
下载与校验值：
证据目录：
未执行项及原因：
```

## 17. 结束和清理

执行完成后：

1. 关闭所有 page、context 和 browser；
2. 停止自己启动的 preview 和临时 HTTP 服务；
3. 停止 CDP Chrome；
4. 确认测试端口已释放；
5. 删除自己的临时 Chrome 用户目录和临时消费项目；
6. 保留证据目录和下载样本；
7. 再运行 `git status --short`，确认没有覆盖用户改动；
8. 生成 `summary.json` 和人类可读摘要。

如果进程由工具会话启动，使用该会话的停止机制；不要依靠关闭终端窗口。禁止使用会影响其他 Agent 的全局杀进程命令。

## 18. 最终摘要模板

```md
# 黑盒验收结果

- 提交：
- 工作区状态：
- 套件：
- 执行者 / 复核者：
- 系统 / 浏览器：
- 材料校验值：
- 开始 / 结束时间：

## 结果

| 用例 | 结果 | 严重级别 | 证据 | 缺陷 |
|---|---|---|---|---|

## 控制台与网络

- 未预期 warning/error：
- page error：
- request failed：
- HTTP 4xx/5xx：
- Worker/WASM：

## 下载与发布包

- 文件名、大小、SHA-256：
- 独立解析/导入结果：
- tgz 与候选版本是否一致：

## 结论

PASS / FAIL / BLOCKED

未通过或未执行的 P0/P1：
剩余 P2/P3：
建议下一步：
```

## 19. 允许宣布完成的条件

只有同时满足以下条件，执行者才能写“黑盒验收通过”：

- 对应 `BB-*` 套件的全部必跑用例都有证据；
- P0/P1 为零；
- 没有未解释的控制台异常、关键网络失败和静默回退；
- 正式构建、真实 Worker/WASM、真实下载或发布包按用例得到验证；
- 首次失败和重试没有被隐藏；
- 证据可由另一名 Agent 重复执行；
- 清理完成，工作区用户改动未受影响。

缺少任何一项时，只能报告当前进度或 BLOCKED，不能宣布完成。
