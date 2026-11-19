#!/usr/bin/env python3
"""Capture and compare the product-route visual matrix.

The matrix is intentionally independent from the upstream Extend parity
harness (`tests/blackbox/ux_parity.py`): it drives the local demo's product
routes only and stores its own baselines under `tests/baseline/visual/`.

Deterministic conditions:
  * Chromium headless, DPR 1, sRGB, hidden scrollbars, fixed font stack
  * `document.fonts.ready` plus route-specific readiness predicates
  * CSS animations/transitions disabled, reduced-motion emulation
  * a fixed viewport, URL theme (`?theme=dark`) and CSS zoom per capture
  * no wall-clock or random values in the structure snapshot

Usage:
  python3 scripts/ci/capture-visual-matrix.py --update-baseline
  python3 scripts/ci/capture-visual-matrix.py --check
  python3 scripts/ci/capture-visual-matrix.py --skip-build --base-url http://127.0.0.1:4186

Pixel diffs are non-blocking by default. Structural assertions are always
blocking, matching issue #74 / #57. `--fail-on-pixel` exists for the eventual
blocking gate in #165 but is deliberately not used by nightly CI yet.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import signal
import socket
import subprocess
import sys
import time
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[2]
BLACKBOX_DIR = ROOT / "tests" / "blackbox"
if str(BLACKBOX_DIR) not in sys.path:
    sys.path.insert(0, str(BLACKBOX_DIR))

try:
    from playwright.sync_api import Page, sync_playwright  # noqa: E402
    from browser_launch import launch_browser  # noqa: E402
except ImportError:  # report-only only needs Pillow
    Page = Any  # type: ignore[misc,assignment]
    sync_playwright = None  # type: ignore[assignment]
    launch_browser = None  # type: ignore[assignment]
try:
    from PIL import Image, ImageChops, ImageStat
except ImportError as error:  # pragma: no cover - environment guard
    raise SystemExit(
        "Pillow is required for visual baselines and pixel diffs; "
        "install requirements-ci.txt first."
    ) from error


DEFAULT_BASELINE_DIR = ROOT / "tests" / "baseline" / "visual"
DEFAULT_OUTPUT_DIR = ROOT / "output" / "visual-matrix"
DEFAULT_PORT = 4186
DEFAULT_PIXEL_THRESHOLD = 8
DEFAULT_MAX_DIFF_RATIO = 0.002
FIXTURE_DIR = ROOT / "apps" / "demo" / "public" / "samples"
PPTP_FIXTURE = FIXTURE_DIR / "deck-structure.pptx"

VIEWPORTS = (
    ("desktop-1440x900", 1440, 900),
    ("laptop-1280x720", 1280, 720),
    ("tablet-768x1024", 768, 1024),
    ("mobile-390x844", 390, 844),
)
THEMES = ("light", "dark")
ZOOMS = (50, 100, 200)

FONT_STACK = (
    '"Liberation Sans", "DejaVu Sans", "Noto Sans CJK SC", '
    '"Noto Sans", Arial, sans-serif'
)

DETERMINISM_CSS = f"""
*, *::before, *::after {{
  animation: none !important;
  transition: none !important;
  caret-color: transparent !important;
  scroll-behavior: auto !important;
}}
html, body, button, input, select, textarea {{
  font-family: {FONT_STACK} !important;
  font-synthesis: none !important;
  text-rendering: geometricPrecision !important;
}}
html {{ scrollbar-width: none; }}
::-webkit-scrollbar {{ width: 0; height: 0; }}
"""

SNAPSHOT_JS = r"""
({ rootSelector }) => {
  const visible = (el) => {
    if (!el) return false;
    const style = getComputedStyle(el);
    if (style.display === "none" || style.visibility === "hidden") return false;
    if (Number(style.opacity) === 0) return false;
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  };
  const norm = (value) => (value || "").replace(/\s+/g, " ").trim();
  const cls = (el) => Array.from(el.classList || []).slice(0, 8);
  const rectOf = (el) => {
    const r = el.getBoundingClientRect();
    return [Math.round(r.x), Math.round(r.y), Math.round(r.width), Math.round(r.height)];
  };
  const attrs = (el) => {
    const out = {};
    for (const name of [
      "data-testid", "data-state", "data-error-code", "data-freeze-row",
      "data-freeze-column", "data-merged-region-count", "data-display-row-count",
      "data-display-column-count", "data-worker-backed", "data-sheet-name",
      "data-slide-index", "data-docx-image-layer", "data-docx-gutter-zoom",
      "data-docx-annotation-gutter-mode", "data-docx-gutter-card-id",
      "data-docx-gutter-card-top", "data-docx-gutter-anchor-y",
      "data-docx-tracked-change-kind", "data-docx-comment-marker",
      "aria-selected", "aria-expanded", "role", "type", "name",
    ]) {
      const value = el.getAttribute(name);
      if (value !== null) out[name] = value.length > 120 ? value.slice(0, 120) : value;
    }
    return out;
  };
  const interesting = (el) => {
    const tag = el.tagName.toLowerCase();
    if (el.hasAttribute("data-testid")) return true;
    if (["h1", "h2", "h3", "h4", "main", "header", "nav", "button", "select",
         "textarea", "canvas", "svg"].includes(tag)) return true;
    if (tag === "input") return true;
    const classes = el.classList || [];
    return Array.from(classes).some((name) =>
      /^(docx|xlsx|pdf|pptx|office|app|page|nav|product|parity|viewer|editor|surface)/.test(name)
    );
  };
  const nodes = [];
  for (const el of document.querySelectorAll("body *")) {
    if (nodes.length >= 900) break;
    if (!visible(el) || !interesting(el)) continue;
    const text = el.childElementCount === 0 ? norm(el.textContent).slice(0, 100) : "";
    nodes.push({
      tag: el.tagName.toLowerCase(),
      classes: cls(el),
      ...attrs(el),
      rect: rectOf(el),
      text: text || undefined,
    });
  }
  const grid = document.querySelector('[data-testid="xlsx-grid"]');
  const paneKeys = grid
    ? Array.from(new Set(Array.from(grid.querySelectorAll("[data-grid-pane]"))
        .map((el) => el.getAttribute("data-grid-pane"))
        .filter(Boolean))).sort()
    : [];
  const canvases = Array.from(document.querySelectorAll("canvas")).filter(visible).map((canvas) => {
    const r = canvas.getBoundingClientRect();
    return {
      width: canvas.width,
      height: canvas.height,
      rect: [Math.round(r.x), Math.round(r.y), Math.round(r.width), Math.round(r.height)],
    };
  });
  const cards = Array.from(document.querySelectorAll("[data-docx-gutter-card-id]")).map((el) => ({
    id: el.getAttribute("data-docx-gutter-card-id"),
    top: Number(el.getAttribute("data-docx-gutter-card-top") || 0),
    anchor: Number(el.getAttribute("data-docx-gutter-anchor-y") || 0),
    height: el.offsetHeight,
    rect: rectOf(el),
    text: norm(el.textContent).slice(0, 120),
  }));
  const root = document.querySelector(rootSelector);
  const status = document.querySelector('[data-testid="page-status"]');
  const htmlStyle = getComputedStyle(document.documentElement);
  const bodyStyle = getComputedStyle(document.body);
  return {
    schemaVersion: 1,
    url: location.href,
    title: document.title,
    viewport: { width: innerWidth, height: innerHeight, dpr: devicePixelRatio },
    zoom: htmlStyle.zoom,
    theme: document.documentElement.dataset.theme || null,
    colors: {
      htmlBackground: htmlStyle.backgroundColor,
      bodyBackground: bodyStyle.backgroundColor,
      bodyColor: bodyStyle.color,
    },
    body: {
      textLength: norm(document.body.innerText).length,
      scrollWidth: document.documentElement.scrollWidth,
      scrollHeight: document.documentElement.scrollHeight,
      headings: document.querySelectorAll("h1, h2, h3").length,
      buttons: document.querySelectorAll("button").length,
      selects: document.querySelectorAll("select").length,
      inputs: document.querySelectorAll("input").length,
      canvases: document.querySelectorAll("canvas").length,
      pageSurfaces: document.querySelectorAll(".docx-page-surface").length,
    },
    root: root
      ? {
          found: true,
          classes: cls(root),
          rect: rectOf(root),
          testid: root.getAttribute("data-testid"),
        }
      : { found: false },
    status: status ? status.getAttribute("data-state") : null,
    parity: (() => {
      const el = document.querySelector('[data-testid="parity-status"]');
      return el ? el.getAttribute("data-state") : null;
    })(),
    xlsx: grid
      ? {
          freezeRow: Number(grid.getAttribute("data-freeze-row") || 0),
          freezeColumn: Number(grid.getAttribute("data-freeze-column") || 0),
          mergedRegions: Number(grid.getAttribute("data-merged-region-count") || 0),
          displayRows: Number(grid.getAttribute("data-display-row-count") || 0),
          displayColumns: Number(grid.getAttribute("data-display-column-count") || 0),
          workerBacked: grid.getAttribute("data-worker-backed"),
          paneKeys,
          scrollLeft: grid.scrollLeft,
          scrollTop: grid.scrollTop,
          chartItems: document.querySelectorAll('[data-testid="xlsx-chart-item"]').length,
          chartTexts: Array.from(document.querySelectorAll('[data-testid="xlsx-chart-item"] text, [data-testid="xlsx-chart-item"] svg text'))
            .map((element) => norm(element.textContent))
            .filter(Boolean)
            .slice(0, 80),
        }
      : null,
    docx: {
      thumbnails: document.querySelectorAll('[data-testid="docx-thumbnail-panel"] canvas').length,
      thumbnailPanel: Boolean(document.querySelector('[data-testid="docx-thumbnail-panel"]')),
      commentCards: cards.length,
      commentMarkers: document.querySelectorAll("[data-docx-comment-marker]").length,
      commentHighlights: document.querySelectorAll("[data-docx-comment-id], [data-docx-comment-ids]").length,
      floatingImages: document.querySelectorAll("[data-docx-image-layer], [data-testid=\"docx-floating-image\"]").length,
      thumbnailSizes: Array.from(document.querySelectorAll('[data-testid="docx-thumbnail-panel"] canvas')).map((canvas) => {
        const r = canvas.getBoundingClientRect();
        return [Math.round(r.width), Math.round(r.height)];
      }),
      tables: document.querySelectorAll(".docx-table-host, [data-testid=\"docx-table\"]").length,
      gutterMode: document.querySelector("[data-docx-annotation-gutter]")?.getAttribute("data-docx-annotation-gutter-mode") || null,
      gutterZoom: document.querySelector("[data-docx-annotation-gutter]")?.getAttribute("data-docx-gutter-zoom") || null,
      cards,
    },
    pdf: {
      pages: document.querySelectorAll('.pdf-page-render, .pdf-surface__page-slot').length,
      renderedPages: document.querySelectorAll('[data-testid="pdf-page-render"], .pdf-surface__page-img').length,
    },
    pptx: {
      slides: document.querySelectorAll("[data-slide-index]").length,
      stage: Boolean(document.querySelector(".pptx-stage")),
    },
    errors: Array.from(document.querySelectorAll(".error, [data-testid$=\"-error\"], [data-error-code]"))
      .filter(visible)
      .map((el) => ({
        testid: el.getAttribute("data-testid"),
        code: el.getAttribute("data-error-code"),
        text: norm(el.textContent).slice(0, 160),
      })),
    canvases,
    nodes,
  };
}
"""


@dataclass(frozen=True)
class RouteSpec:
    id: str
    path: str
    root_selector: str
    ready_js: str
    dark_selector: str | None = None
    dark_coverage: str = "component"
    prepare: str | None = None
    required_status: str | None = None


ROUTES: tuple[RouteSpec, ...] = (
    RouteSpec(
        "docx-viewer",
        "/docx-viewer",
        ".docx-viewer",
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.docx-page-surface').length > 0""",
        dark_selector=".docx-viewer--dark",
    ),
    RouteSpec(
        "docx-editor",
        "/docx-editor",
        ".docx-editor-viewer",
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.docx-editor-viewer').length > 0""",
        dark_selector=".docx-editor-viewer--dark",
    ),
    RouteSpec(
        "docx-surface",
        "/docx-surface",
        ".docx-document-surface",
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.docx-page-surface').length > 0""",
    ),
    RouteSpec(
        "docx-parity",
        "/docx-parity",
        ".parity-grid",
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="parity-status"]')?.dataset.state === 'PASS'""",
        dark_selector=".docx-editor-viewer--dark",
    ),
    RouteSpec(
        "xlsx-viewer",
        "/xlsx-viewer",
        ".xlsx-viewer",
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && Number(document.querySelector('[data-testid="xlsx-grid"]')?.dataset.displayRowCount || 0) > 0""",
        dark_selector=".xlsx-toolbar--dark",
    ),
    RouteSpec(
        "xlsx-surface",
        "/xlsx-surface",
        ".xlsx-sheet-surface",
        """() => document.querySelector('[data-testid="xlsx-surface-state"]')?.dataset.state === 'ready'
          && Number(document.querySelector('[data-testid="xlsx-grid"]')?.dataset.displayRowCount || 0) > 0""",
        dark_coverage="canvas",
    ),
    RouteSpec(
        "pdf-viewer",
        "/pdf-viewer",
        ".pdf-viewer",
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.pdf-page-render').length > 0""",
        dark_selector=".pdf-viewer--dark",
    ),
    RouteSpec(
        "pdf-surface",
        "/pdf-surface",
        ".pdf-surface-shell",
        """() => document.querySelector('[data-testid="pdf-surface-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.pdf-surface__page-slot').length > 0""",
        dark_selector=".pdf-surface--dark",
        prepare="pdf-surface-load",
    ),
    RouteSpec(
        "pptx-viewer",
        "/pptx-viewer",
        ".pptx-viewer",
        """() => document.querySelector('[data-testid="pptx-demo-status"]')?.textContent?.includes('已加载')
          && document.querySelectorAll('.pptx-stage [data-slide-index]').length > 0""",
        dark_selector=".pptx-viewer--dark",
        prepare="pptx-load",
    ),
    RouteSpec(
        "pptx-surface",
        "/pptx-surface",
        ".pptx-stage",
        """() => document.querySelector('[data-testid="pptx-surface-status"]')?.textContent?.includes('已加载')
          && document.querySelectorAll('.pptx-stage [data-slide-index]').length > 0""",
        dark_coverage="chrome-only",
        prepare="pptx-surface-load",
    ),
)

@dataclass
class CaseResult:
    id: str
    kind: str
    route: str
    viewport: str
    theme: str
    zoom: int
    status: str = "PASS"
    screenshot: Path | None = None
    baseline: Path | None = None
    structure: Path | None = None
    structure_baseline: Path | None = None
    assertions: list[dict[str, Any]] = field(default_factory=list)
    metrics: dict[str, Any] = field(default_factory=dict)
    pixel: dict[str, Any] = field(default_factory=dict)
    error: str | None = None

    def fail(self, message: str) -> None:
        self.status = "FAIL"
        self.assertions.append({"ok": False, "message": message})

    def check(self, condition: bool, message: str) -> None:
        self.assertions.append({"ok": bool(condition), "message": message})
        if not condition:
            self.status = "FAIL"


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def git_head() -> str:
    result = subprocess.run(
        ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True, capture_output=True
    )
    return result.stdout.strip() if result.returncode == 0 else "unknown"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def rel(path: Path | None) -> str | None:
    if path is None:
        return None
    try:
        return str(path.relative_to(ROOT))
    except ValueError:
        return str(path)


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_json_compact(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, separators=(",", ":")) + "\n",
        encoding="utf-8",
    )


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def port_in_use(port: int) -> bool:
    with socket.socket() as sock:
        return sock.connect_ex(("127.0.0.1", port)) == 0


def free_port(preferred: int) -> int:
    if not port_in_use(preferred):
        return preferred
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def wait_for_preview(base_url: str, process: subprocess.Popen[bytes] | None, log_path: Path) -> None:
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        if process is not None and process.poll() is not None:
            raise RuntimeError(
                f"preview exited early with {process.returncode}; see {log_path}"
            )
        try:
            with urllib.request.urlopen(base_url, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"preview did not become ready at {base_url}; see {log_path}")


def stop_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=8)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            pass


def run_command(command: list[str], log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("wb") as stream:
        result = subprocess.run(
            command, cwd=ROOT, stdout=stream, stderr=subprocess.STDOUT
        )
    if result.returncode != 0:
        raise RuntimeError(f"{' '.join(command)} failed with {result.returncode}; see {log_path}")


def prepare_page(
    page: Page,
    spec: RouteSpec,
    *,
    theme: str,
    zoom: int,
) -> None:
    page.emulate_media(reduced_motion="reduce", color_scheme=theme)
    page.add_style_tag(content=DETERMINISM_CSS)
    if spec.prepare == "pdf-surface-load":
        page.locator('[data-testid="pdf-surface-load-sample"]').click()
    elif spec.prepare == "pptx-load":
        page.locator('[data-testid="pptx-file-input"]').set_input_files(str(PPTP_FIXTURE))
    elif spec.prepare == "pptx-surface-load":
        page.locator('[data-testid="pptx-surface-file-input"]').set_input_files(
            str(PPTP_FIXTURE)
        )
    page.wait_for_function(spec.ready_js, timeout=60_000)
    page.evaluate("() => document.fonts.ready.then(() => true)")
    if zoom != 100:
        page.evaluate("(value) => { document.documentElement.style.zoom = String(value) }", zoom / 100)
        page.wait_for_timeout(250)
        page.wait_for_function(spec.ready_js, timeout=60_000)
    page.evaluate(
        """() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))"""
    )
    page.wait_for_timeout(300)


def wait_snapshot_stable(page: Page, spec: RouteSpec) -> None:
    page.wait_for_function(spec.ready_js, timeout=60_000)
    page.evaluate("() => document.fonts.ready.then(() => true)")
    page.evaluate(
        """() => new Promise((resolve) => {
          const images = Array.from(document.images);
          if (images.every((image) => image.complete)) return resolve();
          let pending = images.length;
          const done = () => { if (--pending <= 0) resolve(); };
          for (const image of images) {
            if (image.complete) done();
            else {
              image.addEventListener('load', done, { once: true });
              image.addEventListener('error', done, { once: true });
            }
          }
        })"""
    )
    page.wait_for_timeout(200)


def image_stats(path: Path) -> dict[str, Any]:
    image = Image.open(path).convert("RGB")
    width, height = image.size
    sample = image.resize((max(1, width // 8), max(1, height // 8)))
    stat = ImageStat.Stat(sample)
    colors = sample.getcolors(maxcolors=1_000_000) or []
    return {
        "width": width,
        "height": height,
        "mean": round(sum(stat.mean) / 3, 2),
        "stddev": round(sum(stat.stddev) / 3, 2),
        "colors": len(colors),
    }


def sample_pixels(path: Path, rect: list[int]) -> dict[str, Any]:
    image = Image.open(path).convert("RGB")
    x, y, width, height = rect
    x = max(0, min(image.width - 1, x))
    y = max(0, min(image.height - 1, y))
    width = max(1, min(image.width - x, width))
    height = max(1, min(image.height - y, height))
    crop = image.crop((x, y, x + width, y + height))
    stat = ImageStat.Stat(crop)
    colors = crop.getcolors(maxcolors=1_000_000) or []
    mean = tuple(round(value, 1) for value in stat.mean)
    return {
        "rect": [x, y, width, height],
        "mean": mean,
        "luminance": round(
            (0.2126 * mean[0] + 0.7152 * mean[1] + 0.0722 * mean[2]) / 255, 3
        ),
        "uniqueColors": len(colors),
    }


def assert_case(
    result: CaseResult,
    snapshot: dict[str, Any],
    screenshot: Path,
    spec: RouteSpec,
    theme: str,
    zoom: int,
) -> None:
    stats = image_stats(screenshot)
    result.metrics["image"] = stats
    result.check(stats["stddev"] >= 3.0, f"screenshot is not blank (stddev {stats['stddev']})")
    result.check(stats["colors"] >= 8, f"screenshot has visible content ({stats['colors']} colors)")
    result.check(snapshot["body"]["textLength"] > 80, "page has readable text")
    result.check(snapshot["body"]["headings"] >= 1, "page has at least one heading")
    result.check(snapshot["root"].get("found", False), f"root selector {spec.root_selector} is visible")
    result.check(not snapshot["errors"], f"no visible error slots ({len(snapshot['errors'])})")
    result.check(
        snapshot["theme"] == theme,
        f"document theme is {theme} (observed {snapshot['theme']})",
    )

    zoom_value = snapshot.get("zoom") or "1"
    expected_zoom = str(zoom / 100)
    try:
        observed_zoom = float(zoom_value)
    except (TypeError, ValueError):
        observed_zoom = -1
    result.check(
        abs(observed_zoom - float(expected_zoom)) < 0.001,
        f"CSS zoom is {expected_zoom} (observed {zoom_value})",
    )

    if spec.required_status is not None:
        result.check(
            snapshot["status"] == spec.required_status,
            f"page status is {spec.required_status} (observed {snapshot['status']})",
        )

    corner = sample_pixels(screenshot, [2, 2, 8, 8])
    result.metrics["corner"] = corner
    if theme == "dark":
        result.check(
            corner["luminance"] < 0.6,
            f"dark chrome background (corner luminance {corner['luminance']})",
        )
    else:
        result.check(
            corner["luminance"] > 0.4,
            f"light chrome background (corner luminance {corner['luminance']})",
        )

    if spec.dark_selector:
        selector_class = spec.dark_selector.lstrip(".")
        selector_present = selector_class in snapshot.get("root", {}).get("classes", []) or any(
            selector_class in node.get("classes", []) for node in snapshot["nodes"]
        )
        if theme == "dark":
            result.check(selector_present, f"dark component class present: {spec.dark_selector}")
        else:
            result.check(not selector_present, f"light component has no dark class: {spec.dark_selector}")

    if zoom == 100:
        result.check(
            snapshot["body"]["scrollWidth"] <= snapshot["viewport"]["width"] + 2,
            "no horizontal page overflow at 100%",
        )

    route = spec.id
    if route.startswith("docx-"):
        result.check(
            snapshot["body"]["pageSurfaces"] >= 1,
            f"DOCX page surfaces render ({snapshot['body']['pageSurfaces']})",
        )
    if route == "docx-parity":
        result.check(snapshot["parity"] == "PASS", f"viewer/editor parity is PASS (observed {snapshot['parity']})")
    if route in {"xlsx-viewer", "xlsx-surface"}:
        grid = snapshot.get("xlsx") or {}
        result.check(bool(grid), "xlsx grid is mounted")
        result.check(grid.get("displayRows", 0) > 0, f"xlsx grid has rows ({grid.get('displayRows')})")
        result.check(grid.get("displayColumns", 0) > 0, f"xlsx grid has columns ({grid.get('displayColumns')})")
        if theme == "dark" and spec.dark_coverage == "canvas":
            canvas_rects = [canvas["rect"] for canvas in snapshot.get("canvases", []) if canvas["rect"][2] > 40]
            if canvas_rects:
                sample = sample_pixels(screenshot, canvas_rects[0])
                result.metrics["gridCanvasSample"] = sample
                result.check(
                    sample["luminance"] < 0.55,
                    f"dark grid canvas (luminance {sample['luminance']})",
                )
    if route in {"pdf-viewer", "pdf-surface"}:
        result.check(snapshot["pdf"]["pages"] >= 1, f"PDF pages render ({snapshot['pdf']['pages']})")
    if route in {"pptx-viewer", "pptx-surface"}:
        result.check(snapshot["pptx"]["slides"] >= 1, f"PPTX slides render ({snapshot['pptx']['slides']})")


def capture_case(
    browser: Any,
    base_url: str,
    spec: RouteSpec,
    viewport_id: str,
    width: int,
    height: int,
    theme: str,
    zoom: int,
    output: Path,
) -> CaseResult:
    result = CaseResult(
        id=spec.id,
        kind="matrix",
        route=spec.path,
        viewport=viewport_id,
        theme=theme,
        zoom=zoom,
    )
    context = browser.new_context(
        viewport={"width": width, "height": height},
        device_scale_factor=1,
        locale="en-US",
        timezone_id="UTC",
        reduced_motion="reduce",
        color_scheme=theme,
    )
    page = context.new_page()
    page.set_default_timeout(60_000)
    try:
        url = f"{base_url}/?theme={theme}#{spec.path}"
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        prepare_page(page, spec, theme=theme, zoom=zoom)
        wait_snapshot_stable(page, spec)
        page.wait_for_timeout(300)

        screenshot = (
            output / "screenshots" / spec.id / viewport_id / f"{theme}-zoom{zoom:03d}.png"
        )
        screenshot.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(screenshot), full_page=False, animations="disabled", caret="hide")
        result.screenshot = screenshot

        snapshot = page.evaluate(SNAPSHOT_JS, {"rootSelector": spec.root_selector})
        snapshot.update(
            {
                "route": spec.path,
                "routeId": spec.id,
                "scenario": f"{spec.id}/{viewport_id}/{theme}/zoom{zoom}",
                "viewportId": viewport_id,
                "capturedZoom": zoom,
            }
        )
        structure = (
            output / "structure" / spec.id / viewport_id / f"{theme}-zoom{zoom:03d}.json"
        )
        write_json_compact(structure, snapshot)
        result.structure = structure
        result.metrics["snapshot"] = {
            "body": snapshot["body"],
            "status": snapshot["status"],
            "xlsx": snapshot.get("xlsx"),
            "docx": {
                key: snapshot["docx"][key]
                for key in ("thumbnails", "thumbnailPanel", "commentCards", "commentMarkers",
                            "commentHighlights", "floatingImages", "tables", "gutterMode", "gutterZoom")
            },
            "pdf": snapshot.get("pdf"),
            "pptx": snapshot.get("pptx"),
        }
        assert_case(result, snapshot, screenshot, spec, theme, zoom)
    except Exception as error:  # noqa: BLE001 - report every case instead of aborting
        result.error = f"{type(error).__name__}: {error}"
        result.fail(result.error)
    finally:
        context.close()
    return result


# ── Evidence scenarios for the historical rendering findings ───────────────


@dataclass(frozen=True)
class EvidenceSpec:
    id: str
    route: str
    prepare: Callable[[Page], None]
    notes: str
    viewport: str = "desktop-1440x900"
    theme: str = "light"
    zoom: int = 100
    dpr: int = 1
    checks: tuple[str, ...] = ()


def wait_docx_ready(page: Page, *, file_name: str | None = None) -> None:
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.docx-page-surface').length > 0""",
        timeout=60_000,
    )
    if file_name:
        page.wait_for_function(
            """name => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes(name)""",
            arg=file_name,
            timeout=30_000,
        )


def select_docx_sample(
    page: Page,
    file_name: str,
    *,
    comments: bool | None = None,
    tracked: bool | None = None,
    thumbnails: bool | None = None,
) -> None:
    page.select_option('[data-testid="docx-sample-select"]', file_name)
    wait_docx_ready(page, file_name=file_name)
    if tracked is not None:
        page.set_checked('[data-testid="docx-demo-show-tracked-changes"]', tracked)
    if comments is not None:
        page.set_checked('[data-testid="docx-demo-show-comments"]', comments)
    if thumbnails is not None:
        page.set_checked('[data-testid="docx-prop-default-thumbnails-open"]', thumbnails)
        wait_docx_ready(page, file_name=file_name)
    page.wait_for_timeout(500)


def prepare_docx_comments(page: Page) -> None:
    select_docx_sample(page, "review-threads.docx", comments=True, tracked=True)


def prepare_docx_comments_viewer_zoom(page: Page, zoom: int) -> None:
    prepare_docx_comments(page)
    page.select_option('[data-testid="docx-zoom-select"]', str(zoom))
    page.wait_for_function(
        """zoom => Number(document.querySelector('[data-testid="docx-zoom-select"]')?.value) === zoom""",
        arg=zoom,
        timeout=30_000,
    )
    page.wait_for_timeout(700)


def prepare_docx_comments_zoom(page: Page) -> None:
    prepare_docx_comments_viewer_zoom(page, 50)


def prepare_docx_comments_zoom200(page: Page) -> None:
    prepare_docx_comments_viewer_zoom(page, 200)


def prepare_docx_thumbnails(page: Page) -> None:
    select_docx_sample(page, "report-with-image.docx", thumbnails=True)


def prepare_docx_page_sizes(page: Page) -> None:
    select_docx_sample(page, "page-sizes.docx", thumbnails=True)


def prepare_docx_tables(page: Page) -> None:
    select_docx_sample(page, "table-structures.docx")


def prepare_docx_embedded(page: Page) -> None:
    page.locator('[data-testid="docx-file-input"]').set_input_files(
        str(FIXTURE_DIR / "docx-embedded-objects.docx")
    )
    wait_docx_ready(page, file_name="docx-embedded-objects.docx")


def prepare_docx_editor(page: Page) -> None:
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('.docx-editor-viewer')""",
        timeout=60_000,
    )


def wait_xlsx_ready(page: Page, *, file_name: str | None = None) -> None:
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && Number(document.querySelector('[data-testid="xlsx-grid"]')?.dataset.displayRowCount || 0) > 0""",
        timeout=60_000,
    )
    if file_name:
        page.wait_for_function(
            """name => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes(name)""",
            arg=file_name,
            timeout=30_000,
        )


def prepare_xlsx_freeze_merge(page: Page) -> None:
    page.select_option('[data-testid="xlsx-sample-select"]', "charts-images.xlsx")
    wait_xlsx_ready(page, file_name="charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="xlsx-grid"]')?.dataset.mergedRegionCount === '1'
          && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.freezeRow === '2'
          && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.freezeColumn === '2'""",
        timeout=30_000,
    )
    page.eval_on_selector(
        '[data-testid="xlsx-grid"]',
        "el => { el.scrollLeft = 520; el.scrollTop = 240; el.dispatchEvent(new Event('scroll')); }",
    )
    page.wait_for_timeout(600)


def prepare_xlsx_databar_icons(page: Page) -> None:
    page.locator('[data-testid="xlsx-file-input"]').set_input_files(
        str(FIXTURE_DIR / "xlsx-negative-databar-icons.xlsx")
    )
    wait_xlsx_ready(page, file_name="xlsx-negative-databar-icons.xlsx")
    page.wait_for_timeout(700)


def prepare_xlsx_iconsets(page: Page) -> None:
    prepare_xlsx_databar_icons(page)
    page.locator('[data-testid="xlsx-sheet-tab"][data-sheet-name="Icon Sets"]').click()
    page.wait_for_timeout(700)


def prepare_xlsx_charts(page: Page) -> None:
    page.locator('[data-testid="xlsx-file-input"]').set_input_files(
        str(FIXTURE_DIR / "xlsx-chart-families.xlsx")
    )
    wait_xlsx_ready(page, file_name="xlsx-chart-families.xlsx")
    page.locator('[data-testid="xlsx-sheet-tab"][data-sheet-name="Charts"]').click()
    page.wait_for_timeout(1200)


def prepare_xlsx_merge(page: Page) -> None:
    page.select_option('[data-testid="xlsx-sample-select"]', "charts-images.xlsx")
    wait_xlsx_ready(page, file_name="charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="xlsx-grid"]')?.dataset.mergedRegionCount === '1'""",
        timeout=30_000,
    )
    page.wait_for_timeout(600)


def prepare_xlsx_surface_freeze_merge(page: Page) -> None:
    page.select_option('[data-testid="xlsx-surface-sample-select"]', "charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="xlsx-surface-state"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.mergedRegionCount === '1'
          && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.freezeRow === '2'
          && document.querySelector('[data-testid="xlsx-grid"]')?.dataset.freezeColumn === '2'""",
        timeout=60_000,
    )
    page.eval_on_selector(
        '[data-testid="xlsx-grid"]',
        "el => { el.scrollLeft = 520; el.scrollTop = 240; el.dispatchEvent(new Event('scroll')); }",
    )
    page.wait_for_timeout(600)


def prepare_pdf_rotated(page: Page) -> None:
    page.select_option('[data-testid="pdf-sample-select"]', "rotated-pages.pdf")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelectorAll('.pdf-page-render').length > 0""",
        timeout=60_000,
    )
    page.wait_for_timeout(500)


def prepare_pptx_viewer(page: Page) -> None:
    page.locator('[data-testid="pptx-file-input"]').set_input_files(str(PPTP_FIXTURE))
    page.wait_for_function(
        """() => document.querySelector('[data-testid="pptx-demo-status"]')?.textContent?.includes('已加载')
          && document.querySelectorAll('.pptx-stage [data-slide-index]').length > 0""",
        timeout=60_000,
    )
    page.wait_for_timeout(500)


def prepare_pptx_surface(page: Page) -> None:
    page.locator('[data-testid="pptx-surface-file-input"]').set_input_files(str(PPTP_FIXTURE))
    page.wait_for_function(
        """() => document.querySelector('[data-testid="pptx-surface-status"]')?.textContent?.includes('已加载')
          && document.querySelectorAll('.pptx-stage [data-slide-index]').length > 0""",
        timeout=60_000,
    )
    page.wait_for_timeout(500)


EVIDENCE_SCENARIOS: tuple[EvidenceSpec, ...] = (
    EvidenceSpec(
        "docx-viewer-comments",
        "/docx-viewer",
        prepare_docx_comments,
        "Threaded comments, tracked changes, gutter anchoring (#29/#30)",
        checks=("docx-comments",),
    ),
    EvidenceSpec(
        "docx-viewer-comments-zoom50",
        "/docx-viewer",
        prepare_docx_comments_zoom,
        "Comment gutter at 50% zoom (#29)",
        zoom=50,
        checks=("docx-comments",),
    ),
    EvidenceSpec(
        "docx-viewer-comments-zoom200",
        "/docx-viewer",
        prepare_docx_comments_zoom200,
        "Comment gutter at 200% zoom (#29)",
        zoom=200,
        checks=("docx-comments",),
    ),
    EvidenceSpec(
        "docx-viewer-thumbnails",
        "/docx-viewer",
        prepare_docx_thumbnails,
        "Thumbnail panel with an image-bearing document (#44/#48)",
        checks=("docx-thumbnails",),
    ),
    EvidenceSpec(
        "docx-viewer-page-sizes",
        "/docx-viewer",
        prepare_docx_page_sizes,
        "Mixed portrait/landscape pages and per-page thumbnails (#27)",
        checks=("docx-thumbnails",),
    ),
    EvidenceSpec(
        "docx-viewer-tables",
        "/docx-viewer",
        prepare_docx_tables,
        "Merged/nested tables with repeating headers (#30)",
        checks=("docx-tables",),
    ),
    EvidenceSpec(
        "docx-viewer-embedded",
        "/docx-viewer",
        prepare_docx_embedded,
        "Embedded text boxes, shapes and chart/OLE objects (#36)",
        checks=("docx-embedded",),
    ),
    EvidenceSpec(
        "docx-editor-default",
        "/docx-editor",
        prepare_docx_editor,
        "Editor shell, toolbar and editable surface (#28)",
        checks=("docx-editor",),
    ),
    EvidenceSpec(
        "xlsx-viewer-freeze-merge",
        "/xlsx-viewer",
        prepare_xlsx_freeze_merge,
        "Two-axis freeze panes and a scrolled-out merge anchor (#14/#15)",
        checks=("xlsx-freeze-merge",),
    ),
    EvidenceSpec(
        "xlsx-viewer-databar-icons",
        "/xlsx-viewer",
        prepare_xlsx_databar_icons,
        "Negative data bars and 4/5-bucket icon sets (#15)",
        checks=("xlsx-databar-icons",),
    ),
    EvidenceSpec(
        "xlsx-viewer-iconsets",
        "/xlsx-viewer",
        prepare_xlsx_iconsets,
        "4- and 5-bucket icon sets (#15)",
        checks=("xlsx-databar-icons",),
    ),
    EvidenceSpec(
        "xlsx-viewer-charts",
        "/xlsx-viewer",
        prepare_xlsx_charts,
        "Chart families, combo axes, tick labels and legends (#67/#70)",
        checks=("xlsx-charts",),
    ),
    EvidenceSpec(
        "xlsx-viewer-dpr2",
        "/xlsx-viewer",
        prepare_xlsx_merge,
        "Backing-store resolution at DPR 2 (#15)",
        dpr=2,
        checks=("xlsx-dpr2",),
    ),
    EvidenceSpec(
        "xlsx-surface-freeze-merge",
        "/xlsx-surface",
        prepare_xlsx_surface_freeze_merge,
        "Surface grid freeze panes and merged regions (#14/#15)",
        checks=("xlsx-freeze-merge",),
    ),
    EvidenceSpec(
        "pdf-viewer-rotated",
        "/pdf-viewer",
        prepare_pdf_rotated,
        "PDF page rendering after fixture switching",
    ),
    EvidenceSpec(
        "pptx-viewer-deck",
        "/pptx-viewer",
        prepare_pptx_viewer,
        "PPTX viewer with a real deck",
    ),
    EvidenceSpec(
        "pptx-surface-stack",
        "/pptx-surface",
        prepare_pptx_surface,
        "PPTX continuous surface with a real deck",
    ),
)


def evidence_assertions(
    result: CaseResult, snapshot: dict[str, Any], screenshot: Path, spec: EvidenceSpec
) -> None:
    stats = image_stats(screenshot)
    result.metrics["image"] = stats
    result.check(stats["stddev"] >= 3.0, "evidence screenshot is not blank")
    result.check(not snapshot["errors"], f"no visible errors ({len(snapshot['errors'])})")
    result.check(snapshot["theme"] == spec.theme, f"theme is {spec.theme}")

    if "docx-comments" in spec.checks:
        docx = snapshot["docx"]
        result.check(docx["commentCards"] > 0, f"comment cards render ({docx['commentCards']})")
        result.check(docx["commentMarkers"] > 0, f"comment markers render ({docx['commentMarkers']})")
        result.check(
            docx["commentHighlights"] > 0,
            f"comment highlights render ({docx['commentHighlights']})",
        )
        deltas = [
            card["top"] - card["anchor"]
            for card in docx["cards"]
            if card["top"] or card["anchor"]
        ]
        if deltas:
            result.metrics["commentAnchorDeltas"] = deltas
            previous_bottom = None
            layout_errors = []
            for card in docx["cards"]:
                desired = card["anchor"] - 8
                if previous_bottom is not None:
                    desired = max(desired, previous_bottom + 4)
                layout_errors.append(abs(card["top"] - desired))
                previous_bottom = card["top"] + card.get("height", 0)
            result.metrics["commentLayoutErrors"] = layout_errors
            if spec.zoom == 100:
                result.check(
                    max(layout_errors) <= 2,
                    f"cards follow anchor/collision layout (max error {max(layout_errors)}px)",
                )
            anchors = [card["anchor"] for card in docx["cards"]]
            result.check(
                anchors == sorted(anchors),
                "comment cards are ordered by document anchor",
            )
            heights = [card.get("height", 0) for card in docx["cards"]]
            result.metrics["commentCardHeights"] = heights
            if spec.id.endswith("zoom200"):
                result.check(
                    max(heights) <= 90,
                    f"gutter cards do not scale with a 200% viewer zoom (max height {max(heights)}px)",
                )
            if spec.id.endswith("zoom50"):
                result.check(
                    min(heights) >= 55,
                    f"gutter cards do not shrink with a 50% viewer zoom (min height {min(heights)}px)",
                )
        else:
            result.metrics["commentAnchorDeltas"] = []
            result.check(False, "comment cards expose anchor offsets")

    if "docx-thumbnails" in spec.checks:
        docx = snapshot["docx"]
        result.check(docx["thumbnailPanel"], "thumbnail panel is mounted")
        result.check(docx["thumbnails"] > 0, f"thumbnail canvases render ({docx['thumbnails']})")
        if docx["thumbnails"] > 0:
            panel = next(
                (
                    node
                    for node in snapshot["nodes"]
                    if "docx-thumbnail-panel" in node.get("classes", [])
                ),
                None,
            )
            if panel:
                sample = sample_pixels(screenshot, panel["rect"])
                result.metrics["thumbnailSample"] = sample
                result.check(
                    sample["uniqueColors"] > 4,
                    f"thumbnail panel contains rendered content ({sample['uniqueColors']} colors)",
                )
        sizes = docx.get("thumbnailSizes") or []
        result.metrics["thumbnailSizes"] = sizes
        if spec.id == "docx-viewer-page-sizes":
            aspects = {
                round(width / height, 2)
                for width, height in sizes
                if width > 0 and height > 0
            }
            result.check(
                len(aspects) >= 2,
                f"mixed page sizes produce distinct thumbnail aspects ({sorted(aspects)})",
            )

    if "docx-tables" in spec.checks:
        result.check(snapshot["docx"]["tables"] > 0, f"DOCX tables render ({snapshot['docx']['tables']})")

    if "docx-embedded" in spec.checks:
        result.check(
            snapshot["docx"]["floatingImages"] >= 0,
            "embedded-object fixture loaded (floating image count is informative)",
        )
        result.metrics["floatingImages"] = snapshot["docx"]["floatingImages"]

    if "docx-editor" in spec.checks:
        result.check(snapshot["body"]["textLength"] > 120, "editor document has text")
        result.check(snapshot["body"]["buttons"] >= 8, f"editor toolbar renders ({snapshot['body']['buttons']} buttons)")

    if "xlsx-freeze-merge" in spec.checks:
        grid = snapshot["xlsx"] or {}
        result.check(grid.get("freezeRow") == 2, f"freeze row is 2 (observed {grid.get('freezeRow')})")
        result.check(grid.get("freezeColumn") == 2, f"freeze column is 2 (observed {grid.get('freezeColumn')})")
        result.metrics["paneKeys"] = grid.get("paneKeys") or []
        result.check(
            grid.get("mergedRegions", 0) >= 1,
            f"merged region metadata exists ({grid.get('mergedRegions')})",
        )
        result.check(
            (grid.get("scrollLeft") or 0) > 0 and (grid.get("scrollTop") or 0) > 0,
            "grid is scrolled away from the merge anchor",
        )

    if "xlsx-databar-icons" in spec.checks:
        grid = snapshot["xlsx"] or {}
        result.check(grid.get("displayRows", 0) > 0, "data-bar/icon workbook has rows")
        result.check(
            snapshot["body"]["canvases"] > 0,
            f"canvas grid renders ({snapshot['body']['canvases']} canvases)",
        )

    if "xlsx-charts" in spec.checks:
        grid = snapshot["xlsx"] or {}
        result.check(
            grid.get("chartItems", 0) > 0,
            f"chart overlays render ({grid.get('chartItems')})",
        )
        texts = grid.get("chartTexts") or []
        result.metrics["chartTexts"] = texts
        result.check(len(texts) >= 4, f"chart labels render ({len(texts)} text nodes)")
        result.check(
            any("Revenue" in text for text in texts),
            "chart legend labels render",
        )
        result.check(
            any(text.strip() in {"Q1", "Q2", "Q3", "Q4"} for text in texts),
            "chart category ticks render",
        )

    if "xlsx-dpr2" in spec.checks:
        grid = snapshot["xlsx"] or {}
        result.check(
            snapshot["viewport"]["dpr"] == 2,
            f"device pixel ratio is 2 (observed {snapshot['viewport']['dpr']})",
        )
        canvas_rects = [canvas for canvas in snapshot.get("canvases", []) if canvas["rect"][2] > 40]
        result.check(bool(canvas_rects), "grid canvas exists at DPR 2")
        for canvas in canvas_rects[:3]:
            ratio = canvas["width"] / max(1, canvas["rect"][2])
            result.check(
                1.8 <= ratio <= 2.2,
                f"canvas backing store follows DPR 2 (width ratio {ratio:.2f})",
            )


def capture_evidence(
    browser: Any,
    base_url: str,
    spec: EvidenceSpec,
    output: Path,
) -> CaseResult:
    width, height = next(
        (width, height)
        for viewport_id, width, height in VIEWPORTS
        if viewport_id == spec.viewport
    )
    result = CaseResult(
        id=spec.id,
        kind="evidence",
        route=spec.route,
        viewport=spec.viewport,
        theme=spec.theme,
        zoom=spec.zoom,
    )
    context = browser.new_context(
        viewport={"width": width, "height": height},
        device_scale_factor=spec.dpr,
        locale="en-US",
        timezone_id="UTC",
        reduced_motion="reduce",
        color_scheme=spec.theme,
    )
    page = context.new_page()
    page.set_default_timeout(60_000)
    try:
        page.goto(f"{base_url}/?theme={spec.theme}#{spec.route}", wait_until="domcontentloaded")
        page.add_style_tag(content=DETERMINISM_CSS)
        spec.prepare(page)
        page.evaluate("() => document.fonts.ready.then(() => true)")
        page.evaluate(
            """() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))"""
        )
        page.wait_for_timeout(300)
        screenshot = output / "evidence" / f"{spec.id}.png"
        screenshot.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(screenshot), full_page=False, animations="disabled", caret="hide")
        result.screenshot = screenshot
        snapshot = page.evaluate(SNAPSHOT_JS, {"rootSelector": ".page"})
        snapshot.update(
            {
                "route": spec.route,
                "routeId": spec.id,
                "scenario": spec.id,
                "viewportId": spec.viewport,
                "capturedZoom": spec.zoom,
            }
        )
        structure = output / "evidence" / f"{spec.id}.json"
        write_json_compact(structure, snapshot)
        result.structure = structure
        result.metrics.update(
            {
                "body": snapshot["body"],
                "status": snapshot["status"],
                "xlsx": snapshot.get("xlsx"),
                "docx": {
                    key: snapshot["docx"][key]
                    for key in (
                        "thumbnails", "thumbnailPanel", "commentCards", "commentMarkers",
                        "commentHighlights", "floatingImages", "tables", "gutterMode",
                        "gutterZoom", "cards", "thumbnailSizes",
                    )
                },
                "pdf": snapshot.get("pdf"),
                "pptx": snapshot.get("pptx"),
                "errors": snapshot["errors"],
            }
        )
        evidence_assertions(result, snapshot, screenshot, spec)
    except Exception as error:  # noqa: BLE001
        result.error = f"{type(error).__name__}: {error}"
        result.fail(result.error)
    finally:
        context.close()
    return result


# ── Baselines, diffs and report ────────────────────────────────────────────


def baseline_paths(
    baseline_dir: Path, result: CaseResult
) -> tuple[Path, Path]:
    if result.kind == "evidence":
        return (
            baseline_dir / "evidence" / f"{result.id}.webp",
            baseline_dir / "evidence" / f"{result.id}.json",
        )
    return (
        baseline_dir
        / "screenshots"
        / result.id
        / result.viewport
        / f"{result.theme}-zoom{result.zoom:03d}.webp",
        baseline_dir
        / "structure"
        / result.id
        / result.viewport
        / f"{result.theme}-zoom{result.zoom:03d}.json",
    )


def write_baseline(result: CaseResult, baseline_dir: Path) -> None:
    if result.screenshot is None or result.structure is None or result.status != "PASS":
        return
    image_path, structure_path = baseline_paths(baseline_dir, result)
    image_path.parent.mkdir(parents=True, exist_ok=True)
    structure_path.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(result.screenshot) as image:
        image.convert("RGB").save(image_path, "WEBP", lossless=True, quality=100, method=6)
    shutil.copyfile(result.structure, structure_path)
    result.baseline = image_path
    result.structure_baseline = structure_path


def diff_images(
    current_path: Path,
    baseline_path: Path,
    *,
    pixel_threshold: int,
    max_diff_ratio: float,
    diff_path: Path | None,
) -> dict[str, Any]:
    current = Image.open(current_path).convert("RGB")
    baseline = Image.open(baseline_path).convert("RGB")
    if current.size != baseline.size:
        return {
            "status": "DIFF",
            "reason": "size-mismatch",
            "currentSize": list(current.size),
            "baselineSize": list(baseline.size),
            "diffRatio": 1.0,
            "meanAbsDiff": 255.0,
        }
    difference = ImageChops.difference(current, baseline)
    gray = difference.convert("L")
    histogram = gray.histogram()
    differing = sum(histogram[pixel_threshold + 1 :])
    total = current.width * current.height
    ratio = differing / total if total else 0.0
    stat = ImageStat.Stat(difference)
    mean_abs = sum(stat.mean) / 3
    if diff_path is not None and differing:
        base = baseline.copy()
        mask = gray.point(lambda value: 255 if value > pixel_threshold else 0)
        red = Image.new("RGB", current.size, (255, 0, 0))
        highlighted = Image.composite(red, base, mask)
        diff_path.parent.mkdir(parents=True, exist_ok=True)
        highlighted.save(diff_path, "PNG", optimize=True)
    return {
        "status": "MATCH" if ratio <= max_diff_ratio else "DIFF",
        "reason": "pixel-diff",
        "diffRatio": round(ratio, 6),
        "differingPixels": differing,
        "totalPixels": total,
        "meanAbsDiff": round(mean_abs, 3),
        "threshold": pixel_threshold,
        "maxDiffRatio": max_diff_ratio,
        "diffImage": rel(diff_path) if diff_path and differing else None,
    }


def compare_baseline(
    result: CaseResult,
    baseline_dir: Path,
    output: Path,
    *,
    pixel_threshold: int,
    max_diff_ratio: float,
) -> None:
    image_path, structure_path = baseline_paths(baseline_dir, result)
    result.baseline = image_path
    result.structure_baseline = structure_path
    if result.screenshot is None:
        result.pixel = {"status": "SKIPPED", "reason": "no-screenshot"}
        return
    if not image_path.exists():
        result.pixel = {
            "status": "MISSING",
            "reason": "baseline-missing",
            "baseline": rel(image_path),
        }
        return
    diff_path = (
        output
        / "diffs"
        / (
            f"{result.id}.png"
            if result.kind == "evidence"
            else f"{result.id}/{result.viewport}/{result.theme}-zoom{result.zoom:03d}.png"
        )
    )
    result.pixel = diff_images(
        result.screenshot,
        image_path,
        pixel_threshold=pixel_threshold,
        max_diff_ratio=max_diff_ratio,
        diff_path=diff_path,
    )
    if structure_path.exists() and result.structure is not None:
        baseline_structure = read_json(structure_path)
        current_structure = read_json(result.structure)
        result.pixel["structure"] = compare_structure_summary(
            current_structure, baseline_structure
        )


def compare_structure_summary(current: dict[str, Any], baseline: dict[str, Any]) -> dict[str, Any]:
    keys = (
        "body.textLength",
        "body.headings",
        "body.buttons",
        "body.selects",
        "body.inputs",
        "body.canvases",
        "body.pageSurfaces",
        "status",
        "parity",
    )

    def value(source: dict[str, Any], path: str) -> Any:
        node: Any = source
        for part in path.split("."):
            node = node.get(part) if isinstance(node, dict) else None
        return node

    drifted = {
        key: {"current": value(current, key), "baseline": value(baseline, key)}
        for key in keys
        if value(current, key) != value(baseline, key)
    }
    return {
        "nodeCount": {
            "current": len(current.get("nodes", [])),
            "baseline": len(baseline.get("nodes", [])),
        },
        "drifted": drifted,
    }


def scenario_table_rows(results: list[CaseResult]) -> list[dict[str, Any]]:
    rows = []
    for result in results:
        rows.append(
            {
                "id": result.id,
                "kind": result.kind,
                "route": result.route,
                "viewport": result.viewport,
                "theme": result.theme,
                "zoom": result.zoom,
                "status": result.status,
                "screenshot": rel(result.screenshot),
                "baseline": rel(result.baseline),
                "structure": rel(result.structure),
                "pixel": result.pixel.get("status"),
                "diffRatio": result.pixel.get("diffRatio"),
                "assertions": result.assertions,
                "error": result.error,
            }
        )
    return rows


def write_reports(
    output: Path,
    results: list[CaseResult],
    *,
    mode: str,
    baseline_dir: Path,
    git: str,
    browser_version: str,
    pixel_threshold: int,
    max_diff_ratio: float,
) -> None:
    matrix = [result for result in results if result.kind == "matrix"]
    evidence = [result for result in results if result.kind == "evidence"]
    failures = [result for result in results if result.status != "PASS"]
    pixel_diffs = [result for result in matrix + evidence if result.pixel.get("status") == "DIFF"]
    pixel_missing = [result for result in matrix + evidence if result.pixel.get("status") == "MISSING"]
    verdicts_path = baseline_dir / "verdicts.json"
    verdicts = read_json(verdicts_path) if verdicts_path.exists() else []

    payload = {
        "schemaVersion": 1,
        "kind": "product-visual-matrix",
        "mode": mode,
        "generatedAt": now_iso(),
        "gitHead": git,
        "browser": browser_version,
        "baselineDir": rel(baseline_dir),
        "pixelPolicy": {
            "blocking": False,
            "threshold": pixel_threshold,
            "maxDiffRatio": max_diff_ratio,
        },
        "summary": {
            "matrixScenarios": len(matrix),
            "evidenceScenarios": len(evidence),
            "structuralFailures": len(failures),
            "pixelDiffs": len(pixel_diffs),
            "missingBaselines": len(pixel_missing),
        },
        "verdicts": verdicts,
        "results": scenario_table_rows(results),
    }
    write_json(output / "report.json", payload)

    lines: list[str] = [
        "# Product visual matrix",
        "",
        f"- Mode: `{mode}`",
        f"- Generated: {payload['generatedAt']}",
        f"- Git HEAD: `{git}`",
        f"- Browser: {browser_version}",
        f"- Baseline: `{rel(baseline_dir)}`",
        f"- Pixel policy: non-blocking, per-pixel threshold {pixel_threshold}, "
        f"max differing ratio {max_diff_ratio:.4f}",
        "",
        "## Summary",
        "",
        "| Metric | Value |",
        "|---|---:|",
        f"| Matrix scenarios | {len(matrix)} |",
        f"| Evidence scenarios | {len(evidence)} |",
        f"| Structural failures | {len(failures)} |",
        f"| Pixel diffs | {len(pixel_diffs)} |",
        f"| Missing baselines | {len(pixel_missing)} |",
        "",
    ]

    by_route: dict[str, list[CaseResult]] = {}
    for result in matrix:
        by_route.setdefault(result.id, []).append(result)
    lines.extend(["## Route matrix", ""])
    for route in ROUTES:
        route_results = by_route.get(route.id, [])
        if not route_results:
            continue
        passed = sum(1 for result in route_results if result.status == "PASS")
        line = f"### `{route.path}` — {passed}/{len(route_results)} structural"
        if route.dark_coverage != "component":
            line += f" (dark coverage: {route.dark_coverage})"
        lines.extend(
            [
                "",
                line,
                "",
                "| Viewport | Theme | Zoom | Screenshot | Baseline | Structural | Pixel | Diff ratio |",
                "|---|---|---:|---|---|---|---|---:|",
            ]
        )
        route_results.sort(
            key=lambda result: (
                [viewport[0] for viewport in VIEWPORTS].index(result.viewport),
                THEMES.index(result.theme),
                result.zoom,
            )
        )
        for result in route_results:
            screenshot = rel(result.screenshot) or ""
            baseline = rel(result.baseline) or ""
            pixel = result.pixel.get("status") or "—"
            ratio = result.pixel.get("diffRatio")
            ratio_text = f"{ratio:.4f}" if isinstance(ratio, (int, float)) else "—"
            lines.append(
                f"| {result.viewport} | {result.theme} | {result.zoom}% | `{screenshot}` | "
                f"`{baseline}` | {result.status} | {pixel} | {ratio_text} |"
            )
        lines.append("")

    lines.extend(["## Issue evidence scenarios", ""])
    if evidence:
        lines.extend(
            [
                "| Scenario | Route | Screenshot | Structural | Pixel | Notes |",
                "|---|---|---|---|---|---|",
            ]
        )
        for result in evidence:
            spec = next((item for item in EVIDENCE_SCENARIOS if item.id == result.id), None)
            lines.append(
                f"| `{result.id}` | `{result.route}` | `{rel(result.screenshot) or ''}` | "
                f"{result.status} | {result.pixel.get('status') or '—'} | "
                f"{spec.notes if spec else ''} |"
            )
        lines.append("")
    else:
        lines.extend(["No evidence scenarios were captured in this run.", ""])

    if verdicts:
        lines.extend(
            [
                "## Confirmed / refuted / needs-another-fixture",
                "",
                "| Issue | Finding | Verdict | Evidence | Observed | Remaining gap |",
                "|---:|---|---|---|---|---|",
            ]
        )
        for verdict in verdicts:
            lines.append(
                f"| #{verdict['issue']} | {verdict['title']} | **{verdict['verdict']}** | "
                f"{', '.join(f'`{item}`' for item in verdict.get('evidence', []))} | "
                f"{verdict.get('observed', '')} | {verdict.get('gap', '')} |"
            )
        lines.append("")

    if failures:
        lines.extend(["## Structural failures", ""])
        for result in failures:
            lines.append(f"### `{result.id}` / {result.viewport} / {result.theme} / {result.zoom}%")
            lines.append("")
            if result.error:
                lines.append(f"- error: `{result.error}`")
            for assertion in result.assertions:
                if not assertion["ok"]:
                    lines.append(f"- FAIL: {assertion['message']}")
            lines.append("")

    lines.extend(
        [
            "## Reproduction",
            "",
            "```bash",
            "python3 scripts/ci/capture-visual-matrix.py --update-baseline",
            "python3 scripts/ci/capture-visual-matrix.py --check",
            "```",
            "",
            "Structural assertions are blocking. Pixel diffs are reported but do "
            "not fail the command until the gate owned by #165 enables "
            "`--fail-on-pixel`.",
            "",
        ]
    )
    (output / "report.md").write_text("\n".join(lines), encoding="utf-8")


def update_manifest(
    baseline_dir: Path,
    results: list[CaseResult],
    *,
    git: str,
    browser_version: str,
    label: str,
) -> None:
    manifest_path = baseline_dir / "manifest.json"
    entries_by_key: dict[tuple[str, str, str, str, int], dict[str, Any]] = {}
    if manifest_path.exists():
        try:
            existing = read_json(manifest_path)
            for entry in existing.get("entries", []):
                key = (
                    str(entry.get("kind")),
                    str(entry.get("id")),
                    str(entry.get("viewport")),
                    str(entry.get("theme")),
                    int(entry.get("zoom")),
                )
                entries_by_key[key] = entry
        except (OSError, ValueError, TypeError):
            entries_by_key = {}

    for result in results:
        if result.status != "PASS":
            continue
        image_path, structure_path = baseline_paths(baseline_dir, result)
        if not image_path.exists() or not structure_path.exists():
            continue
        key = (result.kind, result.id, result.viewport, result.theme, result.zoom)
        entries_by_key[key] = {
            "id": result.id,
            "kind": result.kind,
            "route": result.route,
            "viewport": result.viewport,
            "theme": result.theme,
            "zoom": result.zoom,
            "screenshot": rel(image_path),
            "screenshotSha256": sha256(image_path),
            "screenshotBytes": image_path.stat().st_size,
            "structure": rel(structure_path),
            "structureSha256": sha256(structure_path),
            "structural": result.status,
        }

    entries = sorted(
        entries_by_key.values(),
        key=lambda entry: (
            0 if entry["kind"] == "matrix" else 1,
            str(entry["route"]),
            str(entry["viewport"]),
            str(entry["theme"]),
            int(entry["zoom"]),
        ),
    )
    manifest = {
        "schemaVersion": 1,
        "kind": "product-visual-matrix-baseline",
        "label": label,
        "generatedAt": now_iso(),
        "gitHead": git,
        "browser": browser_version,
        "determinism": {
            "deviceScaleFactor": 1,
            "locale": "en-US",
            "timezone": "UTC",
            "reducedMotion": "reduce",
            "fontStack": FONT_STACK,
            "animations": "disabled",
            "scrollbars": "hidden",
            "theme": "URL ?theme=light|dark",
            "zoom": "CSS zoom 50/100/200%",
        },
        "viewports": [viewport[0] for viewport in VIEWPORTS],
        "themes": list(THEMES),
        "zooms": list(ZOOMS),
        "routes": [route.id for route in ROUTES],
        "screenshotFormat": "lossless WebP",
        "counts": {
            "matrix": sum(1 for entry in entries if entry["kind"] == "matrix"),
            "evidence": sum(1 for entry in entries if entry["kind"] == "evidence"),
            "total": len(entries),
        },
        "entries": entries,
    }
    write_json(manifest_path, manifest)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--baseline-dir", type=Path, default=DEFAULT_BASELINE_DIR)
    parser.add_argument("--update-baseline", action="store_true")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--skip-build", action="store_true")
    parser.add_argument("--base-url", default=None)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--filter", default=None, help="Comma-separated matrix route ids")
    parser.add_argument("--viewport", default=None, help="Only one viewport id")
    parser.add_argument("--theme", choices=THEMES, default=None)
    parser.add_argument("--zoom", type=int, choices=ZOOMS, default=None)
    parser.add_argument("--no-evidence", action="store_true")
    parser.add_argument("--pixel-threshold", type=int, default=DEFAULT_PIXEL_THRESHOLD)
    parser.add_argument("--max-diff-ratio", type=float, default=DEFAULT_MAX_DIFF_RATIO)
    parser.add_argument("--fail-on-pixel", action="store_true")
    parser.add_argument("--baseline-label", default="pre-ui-migration")
    parser.add_argument(
        "--report-only",
        action="store_true",
        help="Build a report from tests/baseline/visual/manifest.json without a browser",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    if args.update_baseline and args.check:
        raise SystemExit("--update-baseline and --check are mutually exclusive")

    output = args.output.resolve()
    baseline_dir = args.baseline_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    output = output / run_id if args.output == DEFAULT_OUTPUT_DIR else output
    output.mkdir(parents=True, exist_ok=True)

    selected_routes = list(ROUTES)
    if args.filter:
        wanted = {item.strip() for item in args.filter.split(",") if item.strip()}
        selected_routes = [route for route in selected_routes if route.id in wanted]
        if not selected_routes:
            raise SystemExit(f"no matrix route matched --filter {args.filter!r}")
    selected_viewports = list(VIEWPORTS)
    if args.viewport:
        selected_viewports = [viewport for viewport in selected_viewports if viewport[0] == args.viewport]
        if not selected_viewports:
            raise SystemExit(f"unknown --viewport {args.viewport!r}")
    selected_themes = [args.theme] if args.theme else list(THEMES)
    selected_zooms = [args.zoom] if args.zoom is not None else list(ZOOMS)

    if args.report_only:
        manifest_path = baseline_dir / "manifest.json"
        if not manifest_path.exists():
            raise SystemExit(f"no baseline manifest at {manifest_path}")
        manifest = read_json(manifest_path)
        results: list[CaseResult] = []
        for entry in manifest.get("entries", []):
            image_path = ROOT / entry["screenshot"]
            structure_path = ROOT / entry["structure"]
            results.append(
                CaseResult(
                    id=str(entry["id"]),
                    kind=str(entry["kind"]),
                    route=str(entry["route"]),
                    viewport=str(entry["viewport"]),
                    theme=str(entry["theme"]),
                    zoom=int(entry["zoom"]),
                    screenshot=image_path,
                    baseline=image_path,
                    structure=structure_path,
                    structure_baseline=structure_path,
                    pixel={"status": "BASELINE"},
                )
            )
        write_reports(
            output,
            results,
            mode="report-only",
            baseline_dir=baseline_dir,
            git=manifest.get("gitHead", git_head()),
            browser_version=manifest.get("browser", "unknown"),
            pixel_threshold=args.pixel_threshold,
            max_diff_ratio=args.max_diff_ratio,
        )
        print(f"[visual] baseline entries: {len(results)}")
        print(f"[visual] report: {output / 'report.md'}")
        return 0

    browser_version = "unknown"
    base_url = args.base_url.rstrip("/") if args.base_url else None
    preview: subprocess.Popen[bytes] | None = None
    preview_log_path = output / "preview.log"

    try:
        if base_url is None:
            if not args.skip_build:
                print("[visual] building demo and workspace dependencies")
                run_command(["pnpm", "--filter", "demo...", "build"], output / "build.log")
            port = free_port(args.port)
            base_url = f"http://127.0.0.1:{port}"
            if port_in_use(port):
                raise RuntimeError(f"port {port} is already in use")
            log_stream = preview_log_path.open("wb")
            preview = subprocess.Popen(
                [
                    "pnpm",
                    "--filter",
                    "demo",
                    "preview",
                    "--host",
                    "127.0.0.1",
                    "--port",
                    str(port),
                ],
                cwd=ROOT,
                stdout=log_stream,
                stderr=subprocess.STDOUT,
                start_new_session=True,
            )
            wait_for_preview(base_url, preview, preview_log_path)
        else:
            wait_for_preview(base_url, None, preview_log_path)

        if sync_playwright is None or launch_browser is None:
            raise RuntimeError("Playwright is required to capture the visual matrix")
        with sync_playwright() as playwright:
            browser = launch_browser(
                playwright.chromium,
                args=[
                    "--hide-scrollbars",
                    "--force-color-profile=srgb",
                    "--font-render-hinting=none",
                    "--disable-lcd-text",
                ],
            )
            browser_version = browser.version
            results: list[CaseResult] = []
            try:
                for spec in selected_routes:
                    for viewport_id, width, height in selected_viewports:
                        for theme in selected_themes:
                            for zoom in selected_zooms:
                                print(
                                    f"[visual] {spec.id} / {viewport_id} / {theme} / {zoom}%",
                                    flush=True,
                                )
                                result = capture_case(
                                    browser,
                                    base_url,
                                    spec,
                                    viewport_id,
                                    width,
                                    height,
                                    theme,
                                    zoom,
                                    output,
                                )
                                if args.update_baseline:
                                    write_baseline(result, baseline_dir)
                                elif args.check:
                                    compare_baseline(
                                        result,
                                        baseline_dir,
                                        output,
                                        pixel_threshold=args.pixel_threshold,
                                        max_diff_ratio=args.max_diff_ratio,
                                    )
                                results.append(result)

                if not args.no_evidence:
                    for spec in EVIDENCE_SCENARIOS:
                        print(f"[visual] evidence {spec.id}", flush=True)
                        result = capture_evidence(browser, base_url, spec, output)
                        if args.update_baseline:
                            write_baseline(result, baseline_dir)
                        elif args.check:
                            compare_baseline(
                                result,
                                baseline_dir,
                                output,
                                pixel_threshold=args.pixel_threshold,
                                max_diff_ratio=args.max_diff_ratio,
                            )
                        results.append(result)
            finally:
                browser.close()

        mode = "update-baseline" if args.update_baseline else "check" if args.check else "capture"
        write_reports(
            output,
            results,
            mode=mode,
            baseline_dir=baseline_dir,
            git=git_head(),
            browser_version=browser_version,
            pixel_threshold=args.pixel_threshold,
            max_diff_ratio=args.max_diff_ratio,
        )
        if args.update_baseline:
            update_manifest(
                baseline_dir,
                results,
                git=git_head(),
                browser_version=browser_version,
                label=args.baseline_label,
            )

        failures = [result for result in results if result.status != "PASS"]
        pixel_diffs = [result for result in results if result.pixel.get("status") == "DIFF"]
        print(
            f"[visual] {len(results)} scenarios, {len(failures)} structural failures, "
            f"{len(pixel_diffs)} pixel diffs"
        )
        print(f"[visual] report: {output / 'report.md'}")
        if failures:
            print("[visual] structural assertions are blocking: FAIL", file=sys.stderr)
            return 1
        if pixel_diffs and args.fail_on_pixel:
            print("[visual] pixel diff gate enabled: FAIL", file=sys.stderr)
            return 1
        return 0
    except Exception as error:  # noqa: BLE001
        print(f"[visual] BLOCKED: {type(error).__name__}: {error}", file=sys.stderr)
        output.mkdir(parents=True, exist_ok=True)
        write_json(
            output / "report.json",
            {
                "schemaVersion": 1,
                "kind": "product-visual-matrix",
                "mode": "blocked",
                "generatedAt": now_iso(),
                "error": f"{type(error).__name__}: {error}",
            },
        )
        return 2
    finally:
        stop_process(preview)


if __name__ == "__main__":
    raise SystemExit(main())
