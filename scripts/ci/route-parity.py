#!/usr/bin/env python3
"""L2 product-route parity: demo route region vs pinned upstream reference route.

L1 (`scripts/ci/pixel-parity.py`) compares isolated component cells that the
harness composes itself. L2 compares the *route*: the product region of a demo
page against the same region rendered by the pinned upstream wrapper, with the
same fixture, state, theme and viewport.

    pnpm ui:parity:routes                       # report-only (default)
    pnpm ui:parity:routes -- --only pdf-viewer --viewports 1440x900 --themes light

Inputs
    tools/parity-fixtures/routes.json   region selectors, landmarks, states,
                                        exclusions and classified differences
    tools/extend-reference/dist         reference app (route.html entries)
    apps/demo/dist                      subject app (hash routes)

Outputs (default `output/ui-parity-routes/`)
    report.json / report.html / SUMMARY.md / routes-matrix.json
    cases/<route>/<viewport>/<theme>/{reference,subject,diff,heatmap,compare}.png
    cases/<route>/<viewport>/<theme>/{reference-chrome,subject-chrome,diff-chrome,heatmap-chrome}.png

Exit code: 0 for pass/report-only, 1 for a blocking failure (`--gate-mode all`)
or for any `gateStatus: blocking` route (`--gate-mode ready`). A snapshot is
never silently green: `--report-only` is the documented pilot mode until the
route's classified product differences are fixed.
"""

from __future__ import annotations

import argparse
import functools
import hashlib
import http.server
import importlib.util
import json
import sys
import threading
import time
import traceback
import urllib.parse
from pathlib import Path

from PIL import Image, ImageChops
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[2]
ROUTES_FILE = ROOT / "tools/parity-fixtures/routes.json"


def load_pixel_parity():
    """Reuse the L1 diff engine so both layers share thresholds + heatmaps."""
    path = Path(__file__).with_name("pixel-parity.py")
    spec = importlib.util.spec_from_file_location("pixel_parity", path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


PP = load_pixel_parity()


def log(message: str) -> None:
    print(message, flush=True)


def rel(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(ROOT))
    except ValueError:
        return str(path)


def parse_csv(value: str | None) -> list[str] | None:
    if not value:
        return None
    return [item.strip() for item in value.split(",") if item.strip()]


# --------------------------------------------------------------------------
# Static servers
# --------------------------------------------------------------------------


class RouteHandler(http.server.SimpleHTTPRequestHandler):
    """Serve a built app, mapping `/fixtures/<name>` to the shared corpus.

    Both sides must load identical bytes; the subject app serves its own
    `/samples/<name>` copy from `apps/demo/public`, while the reference app has
    no fixture bundle, so the harness maps its `/fixtures/` prefix onto the
    repository files and asserts the sha256 of both responses.
    """

    fixtures: dict[str, Path] = {}

    def log_message(self, format: str, *args) -> None:  # noqa: A002
        pass

    def translate_path(self, path: str) -> str:
        request = urllib.parse.urlsplit(path)
        if request.path.startswith("/fixtures/"):
            name = urllib.parse.unquote(request.path[len("/fixtures/") :])
            target = self.fixtures.get(name)
            if target is not None:
                return str(target)
        return super().translate_path(path)


def serve(directory: Path, fixtures: dict[str, Path]) -> tuple[http.server.ThreadingHTTPServer, str]:
    if not (directory / "index.html").exists():
        raise SystemExit(f"{rel(directory)}/index.html is missing; build the app first")
    handler = type(
        "BoundRouteHandler",
        (RouteHandler,),
        {"fixtures": fixtures},
    )
    server = http.server.ThreadingHTTPServer(
        ("127.0.0.1", 0),
        functools.partial(handler, directory=str(directory)),
    )
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, f"http://127.0.0.1:{server.server_address[1]}"


def fetch_sha256(url: str) -> str:
    import urllib.request

    with urllib.request.urlopen(url, timeout=30) as response:  # noqa: S310 - local harness
        return hashlib.sha256(response.read()).hexdigest()


# --------------------------------------------------------------------------
# Page driving
# --------------------------------------------------------------------------


def wait_route_ready(page, spec: dict, timeout: int = 60_000) -> None:
    for selector in spec.get("ready") or []:
        page.wait_for_selector(selector, state="attached", timeout=timeout)
    ready_function = spec.get("readyFunction")
    if ready_function:
        page.wait_for_function(ready_function, timeout=timeout)
    PP.settle_images(page)
    page.evaluate(
        """async () => {
          await document.fonts.ready;
        }"""
    )
    page.wait_for_timeout(int(spec.get("settleMs", 250)))


def run_prepare(page, actions: list[dict]) -> None:
    for action in actions or []:
        kind = action.get("action")
        selector = action.get("selector")
        if kind == "click":
            page.locator(selector).first.click()
        elif kind == "hover":
            page.locator(selector).first.hover()
        elif kind == "evaluate":
            page.evaluate(action["script"])
        else:
            raise RuntimeError(f"unknown prepare action {kind!r}")
        page.wait_for_timeout(int(action.get("settleMs", 250)))


def locate_region(page, selector: str) -> dict:
    """Resolve the region's document box.

    The capture is cropped out of one full-page screenshot at integer device
    pixels, so a region that sits on a half CSS pixel (the demo page has a 0.5px
    line box above the viewer) still starts exactly on its own first device
    pixel. That keeps both sides phase-aligned without scrolling or nudging the
    layout.
    """
    target = page.locator(selector).first
    target.wait_for(state="visible", timeout=30_000)
    box = target.bounding_box()
    if box is None:
        raise RuntimeError(f"region {selector} has no box")
    scroll = page.evaluate("() => ({ x: window.scrollX, y: window.scrollY })")
    viewport_box = {key: round(box[key], 3) for key in ("x", "y", "width", "height")}
    document = {
        "x": round(box["x"] + scroll["x"], 3),
        "y": round(box["y"] + scroll["y"], 3),
        "width": round(box["width"], 3),
        "height": round(box["height"], 3),
    }
    viewport = page.evaluate("() => ({ width: window.innerWidth, height: window.innerHeight })")
    return {"document": document, "viewportBox": viewport_box, "viewport": viewport}


STYLE_KEYS = [
    "backgroundColor",
    "color",
    "fontSize",
    "lineHeight",
    "paddingTop",
    "paddingRight",
    "paddingBottom",
    "paddingLeft",
    "borderTopWidth",
    "borderTopColor",
    "borderTopLeftRadius",
]

METRIC_SCRIPT = """(el) => {
  const rect = el.getBoundingClientRect();
  const style = getComputedStyle(el);
  const result = { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  const styleKeys = %s;
  result.style = {};
  for (const key of styleKeys) result.style[key] = style[key];
  return result;
}""" % json.dumps(STYLE_KEYS)


def collect_landmarks(page, selectors: dict[str, str], region_box: dict) -> dict:
    landmarks = {}
    for name, selector in (selectors or {}).items():
        locator = page.locator(selector).first
        if locator.count() == 0:
            landmarks[name] = {"missing": True, "selector": selector}
            continue
        try:
            metrics = locator.evaluate(METRIC_SCRIPT)
        except Exception as error:  # noqa: BLE001 - report the landmark, don't abort
            landmarks[name] = {"missing": True, "selector": selector, "error": str(error)}
            continue
        landmarks[name] = {
            "selector": selector,
            "box": {
                key: round(metrics[key], 3)
                for key in ("x", "y", "width", "height")
            },
            "offset": {
                "x": round(metrics["x"] - region_box["x"], 3),
                "y": round(metrics["y"] - region_box["y"], 3),
            },
            "style": metrics["style"],
        }
    return landmarks


def apply_visibility(page, selectors: list[str], visible: bool) -> None:
    if not selectors:
        return
    page.evaluate(
        """([selectors, visible]) => {
          for (const selector of selectors) {
            for (const element of document.querySelectorAll(selector)) {
              element.style.visibility = visible ? '' : 'hidden';
            }
          }
        }""",
        [selectors, visible],
    )
    page.wait_for_timeout(120)


def capture_region(page, region: dict, destination: Path, device_scale_factor: int) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    box = region["document"]
    full = destination.with_name(destination.stem + "-fullpage.png")
    page.screenshot(path=str(full), full_page=True, animations="disabled")
    with Image.open(full) as screenshot:
        left = round(box["x"] * device_scale_factor)
        top = round(box["y"] * device_scale_factor)
        right = round((box["x"] + box["width"]) * device_scale_factor)
        bottom = round((box["y"] + box["height"]) * device_scale_factor)
        screenshot.crop((left, top, right, bottom)).save(destination)
    full.unlink(missing_ok=True)


# --------------------------------------------------------------------------
# Comparison
# --------------------------------------------------------------------------


def dominant_colors(path: Path, top: int = 3) -> list[dict]:
    """Most common colours in a capture (evidence for chrome deltas)."""
    with Image.open(path) as image:
        rgb = image.convert("RGB")
        total = rgb.size[0] * rgb.size[1]
        counts = rgb.getcolors(maxcolors=1 << 24) or []
    counts.sort(reverse=True)
    return [
        {
            "rgb": list(color),
            "hex": "#%02x%02x%02x" % color,
            "share": round(count / total * 100, 2),
        }
        for count, color in counts[:top]
    ]


def diff_bands(reference_path: Path, subject_path: Path, split_device_px: int) -> dict:
    """Share of differing pixels above/below a device-pixel row (toolbar split)."""
    with Image.open(reference_path) as reference_image, Image.open(subject_path) as subject_image:
        reference = reference_image.convert("RGB")
        subject = subject_image.convert("RGB")
        diff = ImageChops.difference(reference, subject).convert("L")
        width, height = diff.size

        def ratio(box) -> float:
            crop = diff.crop(box)
            histogram = crop.histogram()
            total = box[2] - box[0]
            total *= box[3] - box[1]
            changed = sum(histogram[3:])
            return round(changed / total, 6) if total else 0.0

        split = max(0, min(height, split_device_px))
        return {
            "splitDevicePx": split,
            "aboveSplitRatio": ratio((0, 0, width, split)),
            "belowSplitRatio": ratio((0, split, width, height)),
        }


def compare_geometry(
    reference: dict, subject: dict, tolerance: float, per_landmark: dict
) -> dict:
    rows = []
    worst = 0.0
    failures = []
    for name, reference_landmark in reference.items():
        subject_landmark = subject.get(name)
        if subject_landmark is None:
            rows.append({"landmark": name, "status": "missing-subject"})
            failures.append(name)
            continue
        if reference_landmark.get("missing") or subject_landmark.get("missing"):
            rows.append(
                {
                    "landmark": name,
                    "status": "missing",
                    "referenceSelector": reference_landmark.get("selector"),
                    "subjectSelector": subject_landmark.get("selector"),
                }
            )
            failures.append(name)
            continue
        allowed = per_landmark.get(name, tolerance)
        deltas = {}
        for key in ("x", "y", "width", "height"):
            deltas[key] = round(
                subject_landmark["box"][key] - reference_landmark["box"][key], 3
            )
        # Compare the box relative to each region so a page-level y offset does
        # not leak into the landmark comparison.
        offset_delta = {
            "x": round(subject_landmark["offset"]["x"] - reference_landmark["offset"]["x"], 3),
            "y": round(subject_landmark["offset"]["y"] - reference_landmark["offset"]["y"], 3),
        }
        size_delta = {
            "width": deltas["width"],
            "height": deltas["height"],
        }
        worst_delta = max(
            [abs(value) for value in offset_delta.values()] + [abs(value) for value in size_delta.values()]
        )
        worst = max(worst, worst_delta)
        status = "pass" if worst_delta <= allowed else "fail"
        if status == "fail":
            failures.append(name)
        rows.append(
            {
                "landmark": name,
                "status": status,
                "tolerancePx": allowed,
                "deltaOffset": offset_delta,
                "deltaSize": size_delta,
                "reference": reference_landmark["box"],
                "subject": subject_landmark["box"],
                "referenceStyle": reference_landmark.get("style"),
                "subjectStyle": subject_landmark.get("style"),
            }
        )
    return {
        "status": "fail" if failures else "pass",
        "worstDeltaPx": round(worst, 3),
        "failures": failures,
        "landmarks": rows,
    }


def compare_pixels(
    reference_path: Path,
    subject_path: Path,
    case_dir: Path,
    thresholds: dict,
    out: Path,
    suffix: str = "",
) -> dict:
    diff_path = case_dir / f"diff{suffix}.png"
    heatmap_path = case_dir / f"heatmap{suffix}.png"
    compare_path = case_dir / f"compare{suffix}.png"
    result = PP.compare_images(
        reference_path, subject_path, diff_path, heatmap_path, compare_path, thresholds
    )
    result["screenshot"] = PP.artifact_rel(diff_path, out) if diff_path.exists() else None
    result["heatmap"] = PP.artifact_rel(heatmap_path, out) if heatmap_path.exists() else None
    result["compare"] = PP.artifact_rel(compare_path, out) if compare_path.exists() else None
    case_dir_rel = PP.artifact_rel(case_dir, out)
    result["caseDir"] = case_dir_rel
    return result


# --------------------------------------------------------------------------
# Case execution
# --------------------------------------------------------------------------


def build_url(origin: str, template: str, theme: str) -> str:
    return origin + template.replace("{theme}", theme)


def route_gate(route: dict) -> dict:
    """Resolve which checks block this route.

    A route can only block on the product shell while its page bitmap is
    produced by a different renderer version than the pinned reference. The
    full-page pixel diff then stays in the report as a diagnostic; chrome-only
    pixels, geometry and console errors remain blocking by default.
    """
    configured = route.get("gate") or {}
    return {
        "fullPagePixels": bool(configured.get("fullPagePixels", True)),
        "chromePixels": bool(configured.get("chromePixels", True)),
        "geometry": bool(configured.get("geometry", True)),
        "consoleErrors": bool(configured.get("consoleErrors", True)),
    }


def run_case(
    *,
    browser,
    route: dict,
    viewport: dict,
    theme: str,
    reference_origin: str,
    subject_origin: str,
    out: Path,
    thresholds: dict,
    device_scale_factor: int,
    self_check: bool,
) -> dict:
    case_dir = out / "cases" / route["id"] / viewport["id"] / theme
    case = {
        "route": route["id"],
        "viewport": viewport["id"],
        "theme": theme,
        "status": "error",
        "errors": [],
    }
    context = browser.new_context(
        viewport={"width": viewport["width"], "height": viewport["height"]},
        device_scale_factor=2,
        reduced_motion="reduce",
    )
    reference_page = context.new_page()
    subject_page = None
    try:
        reference_errors = PP.prepare_page(reference_page, "reference")
        reference_page.goto(
            build_url(reference_origin, route["reference"]["url"], theme),
            wait_until="load",
        )
        wait_route_ready(reference_page, route["reference"])
        run_prepare(reference_page, route["reference"].get("prepare") or [])
        reference_region = locate_region(
            reference_page, route["reference"]["regionSelector"]
        )
        reference_landmarks = collect_landmarks(
            reference_page,
            route["reference"].get("landmarks") or {},
            reference_region["viewportBox"],
        )
        capture_region(reference_page, reference_region, case_dir / "reference.png", device_scale_factor)
        apply_visibility(
            reference_page, (route["reference"].get("masks") or {}).get("chrome") or [], False
        )
        capture_region(
            reference_page, reference_region, case_dir / "reference-chrome.png", device_scale_factor
        )
        case["reference"] = {
            "region": reference_region,
            "landmarks": reference_landmarks,
            "screenshot": PP.artifact_rel(case_dir / "reference.png", out),
        }

        subject_errors: list[str] = []
        if not self_check:
            subject_page = context.new_page()
            subject_errors = PP.prepare_page(subject_page, "subject")
            subject_page.goto(
                build_url(subject_origin, route["subject"]["url"], theme),
                wait_until="load",
            )
            wait_route_ready(subject_page, route["subject"])
            run_prepare(subject_page, route["subject"].get("prepare") or [])
            subject_region = locate_region(subject_page, route["subject"]["regionSelector"])
            subject_landmarks = collect_landmarks(
                subject_page,
                route["subject"].get("landmarks") or {},
                subject_region["viewportBox"],
            )
            capture_region(subject_page, subject_region, case_dir / "subject.png", device_scale_factor)
            apply_visibility(
                subject_page, (route["subject"].get("masks") or {}).get("chrome") or [], False
            )
            capture_region(subject_page, subject_region, case_dir / "subject-chrome.png", device_scale_factor)
            case["subject"] = {
                "region": subject_region,
                "landmarks": subject_landmarks,
                "screenshot": PP.artifact_rel(case_dir / "subject.png", out),
            }

        case["errors"] = reference_errors + subject_errors

        subject_capture = case_dir / "subject.png"
        if self_check:
            # Determinism gate: capture the reference twice and diff it with
            # itself. Any diff means the route is not reproducible.
            second = context.new_page()
            try:
                PP.prepare_page(second, "reference")
                second.goto(
                    build_url(reference_origin, route["reference"]["url"], theme),
                    wait_until="load",
                )
                wait_route_ready(second, route["reference"])
                run_prepare(second, route["reference"].get("prepare") or [])
                second_region = locate_region(
                    second, route["reference"]["regionSelector"]
                )
                capture_region(second, second_region, case_dir / "reference-2.png", device_scale_factor)
            finally:
                second.close()
            subject_capture = case_dir / "reference-2.png"
            case["selfCheck"] = True

        geometry = compare_geometry(
            reference_landmarks,
            case["subject"]["landmarks"] if not self_check else reference_landmarks,
            thresholds["geometryTolerancePx"],
            route.get("landmarkTolerance") or {},
        )
        # Self-check compares the same page twice; landmark deltas are 0 by
        # construction, so only the pixel gate is meaningful.
        case["geometry"] = geometry

        pixels = compare_pixels(
            case_dir / "reference.png", subject_capture, case_dir, thresholds, out
        )
        chrome = compare_pixels(
            case_dir / "reference-chrome.png",
            case_dir / ("reference-chrome.png" if self_check else "subject-chrome.png"),
            case_dir,
            thresholds,
            out,
            suffix="-chrome",
        )
        case["pixels"] = pixels
        case["chrome"] = chrome
        toolbar = next(
            (
                landmark
                for landmark in (case.get("geometry") or {}).get("landmarks", [])
                if landmark.get("landmark") == "toolbar" and "reference" in landmark
            ),
            None,
        )
        split_px = (
            round(toolbar["reference"]["height"] * device_scale_factor)
            if toolbar
            else 0
        )
        pixels["bands"] = diff_bands(
            case_dir / "reference.png", subject_capture, split_px
        )
        pixels["dominantColors"] = {
            "reference": dominant_colors(case_dir / "reference.png"),
            "subject": dominant_colors(subject_capture),
        }
        chrome["dominantColors"] = {
            "reference": dominant_colors(case_dir / "reference-chrome.png"),
            "subject": dominant_colors(
                case_dir / ("reference-chrome.png" if self_check else "subject-chrome.png")
            ),
        }

        gate = route_gate(route)
        checks = {
            "pixels": pixels["status"] == "fail",
            "chrome-pixels": chrome["status"] == "fail",
            "geometry": geometry["status"] == "fail",
            "console-errors": bool(case["errors"]),
        }
        gating = {
            "pixels": gate["fullPagePixels"],
            "chrome-pixels": gate["chromePixels"],
            "geometry": gate["geometry"],
            "console-errors": gate["consoleErrors"],
        }
        failures = [name for name, failed in checks.items() if failed and gating[name]]
        diagnostics = [name for name, failed in checks.items() if failed and not gating[name]]
        case["gate"] = gate
        case["failures"] = failures
        if diagnostics:
            case["diagnosticFailures"] = diagnostics
        if failures:
            case["status"] = "fail"
        elif diagnostics:
            # The check ran and is reported; the route contract does not gate it.
            case["status"] = "pass-with-diagnostic"
        elif (
            pixels["status"] == "pass-with-antialiasing"
            or chrome["status"] == "pass-with-antialiasing"
        ):
            case["status"] = "pass-with-antialiasing"
        else:
            case["status"] = "pass"
    except Exception as error:  # noqa: BLE001 - report the case, keep the matrix
        case["status"] = "error"
        case["errors"] = case.get("errors", []) + [f"{type(error).__name__}: {error}"]
        case["traceback"] = traceback.format_exc()
    finally:
        if subject_page is not None:
            subject_page.close()
        reference_page.close()
        context.close()
    return case


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------


def summarize_route(route: dict, cases: list[dict]) -> dict:
    blocking_cases = [case for case in cases if case["status"] in ("fail", "error")]
    worst_pixels = max(
        (case.get("pixels", {}).get("diffRatio", 0) for case in cases), default=0
    )
    worst_chrome = max(
        (case.get("chrome", {}).get("diffRatio", 0) for case in cases), default=0
    )
    worst_geometry = max(
        (
            case.get("geometry", {}).get("worstDeltaPx", 0)
            for case in cases
            if case.get("geometry")
        ),
        default=0,
    )
    return {
        "id": route["id"],
        "title": route["title"],
        "gateStatus": route.get("gateStatus", "report-only"),
        "blocking": route.get("gateStatus") == "blocking",
        "gate": route_gate(route),
        "status": "fail" if blocking_cases else "pass",
        "caseCount": len(cases),
        "failedCases": [
            f"{case['viewport']}/{case['theme']}: {case['status']}"
            for case in blocking_cases
        ],
        "diagnosticCases": [
            f"{case['viewport']}/{case['theme']}: {', '.join(case['diagnosticFailures'])}"
            for case in cases
            if case.get("diagnosticFailures")
        ],
        "worstDiffRatio": round(worst_pixels, 6),
        "worstChromeDiffRatio": round(worst_chrome, 6),
        "worstGeometryDeltaPx": round(worst_geometry, 3),
        "region": route.get("region"),
        "referenceRoute": route.get("referenceRoute"),
        "subjectRoute": route.get("subjectRoute"),
        "exclusions": route.get("exclusions") or [],
        "differences": route.get("differences") or [],
    }


def diagnostic_suffix(case: dict) -> str:
    diagnostics = case.get("diagnosticFailures") or []
    return f" (diagnostic: {', '.join(diagnostics)})" if diagnostics else ""


def write_summary(report: dict, destination: Path) -> None:
    lines = [
        "# L2 route parity report",
        "",
        f"- harness: `{report['harness']}`",
        f"- upstream pin: `{report['referenceCommit']}`",
        f"- device scale factor: {report['deviceScaleFactor']}",
        f"- thresholds: geometry {report['thresholds']['geometryTolerancePx']}px, "
        f"pixels {report['thresholds']['maxDiffPixelRatio'] * 100:.2f}%, "
        f"channel delta {report['thresholds']['maxChannelDelta']}",
        "",
        "## Route matrix",
        "",
        "| route | gate | status | cases | worst pixels | worst chrome | worst geometry |",
        "|---|---|---|---|---|---|---|",
    ]
    for route in report["routes"]:
        lines.append(
            f"| `{route['id']}` | {route['gateStatus']} | {route['status']} | "
            f"{route['caseCount']} | {route['worstDiffRatio'] * 100:.3f}% | "
            f"{route['worstChromeDiffRatio'] * 100:.3f}% | "
            f"{route['worstGeometryDeltaPx']:.2f}px |"
        )
    lines += ["", "## Route gates", ""]
    for route in report["routes"]:
        gate = route.get("gate") or {}
        enabled = [
            label
            for key, label in (
                ("fullPagePixels", "full-page pixels"),
                ("chromePixels", "chrome-only pixels"),
                ("geometry", "landmarks"),
                ("consoleErrors", "console errors"),
            )
            if gate.get(key)
        ]
        lines.append(
            f"- `{route['id']}` ({route['gateStatus']}): blocking checks "
            f"{', '.join(enabled) or '—'}"
        )
        if gate and not gate.get("fullPagePixels", True):
            lines.append(
                "  full-page pixels are reported as diagnostic because the page "
                "bitmap comes from a different pinned renderer version than the reference"
            )
    lines += ["", "## Cases", ""]
    for route in report["routes"]:
        lines.append(f"### `{route['id']}`")
        lines.append("")
        lines.append("| viewport | theme | status | pixels | chrome | geometry | failures |")
        lines.append("|---|---|---|---|---|---|---|")
        for case in route["cases"]:
            pixels = case.get("pixels") or {}
            chrome = case.get("chrome") or {}
            geometry = case.get("geometry") or {}
            lines.append(
                f"| {case['viewport']} | {case['theme']} | {case['status']} | "
                f"{(pixels.get('diffRatio') or 0) * 100:.3f}% | "
                f"{(chrome.get('diffRatio') or 0) * 100:.3f}% | "
                f"{geometry.get('worstDeltaPx', 0):.2f}px | "
                f"{', '.join(case.get('failures') or []) or '—'}"
                f"{diagnostic_suffix(case)} |"
            )
        lines.append("")
        lines.append("Dominant chrome colours (background evidence):")
        lines.append("")
        for case in route["cases"]:
            colors = (case.get("chrome") or {}).get("dominantColors") or {}
            reference = (colors.get("reference") or [{}])[0]
            subject = (colors.get("subject") or [{}])[0]
            lines.append(
                f"- `{case['viewport']}/{case['theme']}`: reference "
                f"`{reference.get('hex', '—')}` {reference.get('share', 0)}% vs subject "
                f"`{subject.get('hex', '—')}` {subject.get('share', 0)}%"
            )
        lines.append("")
        if route["differences"]:
            lines.append("Classified differences (product block, scaffolding excluded):")
            lines.append("")
            lines.append("| id | category | blocking | status |")
            lines.append("|---|---|---|---|")
            for difference in route["differences"]:
                lines.append(
                    f"| `{difference['id']}` | {difference['category']} | "
                    f"{'yes' if difference['blocking'] else 'no'} | {difference['status']} |"
                )
            lines.append("")
        if route["exclusions"]:
            lines.append("Excluded scaffolding (never compared):")
            lines.append("")
            for exclusion in route["exclusions"]:
                lines.append(
                    f"- `{exclusion['scope']}` `{exclusion['selector']}` — {exclusion['reason']}"
                )
            lines.append("")
    lines += [
        "## Landmark geometry",
        "",
        "Deltas are `subject - reference` for each landmark box relative to its region origin.",
        "",
        "| route | case | landmark | Δx | Δy | Δw | Δh | tolerance | status |",
        "|---|---|---|---|---|---|---|---|---|",
    ]
    for route in report["routes"]:
        for case in route["cases"]:
            for landmark in (case.get("geometry") or {}).get("landmarks", []):
                if "deltaOffset" not in landmark:
                    continue
                lines.append(
                    f"| `{route['id']}` | {case['viewport']}/{case['theme']} | "
                    f"{landmark['landmark']} | {landmark['deltaOffset']['x']:+.2f} | "
                    f"{landmark['deltaOffset']['y']:+.2f} | {landmark['deltaSize']['width']:+.2f} | "
                    f"{landmark['deltaSize']['height']:+.2f} | {landmark['tolerancePx']}px | "
                    f"{landmark['status']} |"
                )
    lines.append("")
    destination.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_html(report: dict, destination: Path) -> None:
    import html as html_module

    def escape(value) -> str:
        return html_module.escape(str(value))

    def image_cell(relative: str | None, label: str) -> str:
        if not relative:
            return "<td>—</td>"
        return (
            f'<td><a href="{escape(relative)}">'
            f'<img src="{escape(relative)}" alt="{escape(label)}"></a></td>'
        )

    rows = []
    for route in report["routes"]:
        for case in route["cases"]:
            image_cells = []
            for key in ("reference", "subject"):
                screenshot = (case.get(key) or {}).get("screenshot")
                image_cells.append(image_cell(screenshot, key))
            pixels = case.get("pixels") or {}
            chrome = case.get("chrome") or {}
            diff = pixels.get("screenshot")
            heatmap = pixels.get("heatmap")
            image_cells.append(image_cell(diff, "diff"))
            image_cells.append(image_cell(heatmap, "heatmap"))
            rows.append(
                "<tr>"
                f"<td>{escape(route['id'])}</td>"
                f"<td>{escape(case['viewport'])}</td>"
                f"<td>{escape(case['theme'])}</td>"
                f"<td class='{escape(case['status'])}'>{escape(case['status'])}</td>"
                f"<td>{(pixels.get('diffRatio') or 0) * 100:.3f}%</td>"
                f"<td>{(chrome.get('diffRatio') or 0) * 100:.3f}%</td>"
                f"<td>{(case.get('geometry') or {}).get('worstDeltaPx', 0):.2f}px</td>"
                + "".join(image_cells)
                + "</tr>"
            )

    differences = []
    for route in report["routes"]:
        for difference in route["differences"]:
            differences.append(
                "<tr>"
                f"<td>{escape(route['id'])}</td>"
                f"<td><code>{escape(difference['id'])}</code></td>"
                f"<td>{escape(difference['category'])}</td>"
                f"<td>{'yes' if difference['blocking'] else 'no'}</td>"
                f"<td>{escape(difference['status'])}</td>"
                f"<td>{escape(difference.get('evidence', ''))}</td>"
                "</tr>"
            )

    exclusions = []
    for route in report["routes"]:
        for exclusion in route["exclusions"]:
            exclusions.append(
                "<tr>"
                f"<td>{escape(route['id'])}</td>"
                f"<td>{escape(exclusion['scope'])}</td>"
                f"<td><code>{escape(exclusion['selector'])}</code></td>"
                f"<td>{escape(exclusion['category'])}</td>"
                f"<td>{escape(exclusion['reason'])}</td>"
                "</tr>"
            )

    document = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>L2 route parity report</title>
<style>
  body {{ font-family: ui-sans-serif, system-ui, sans-serif; margin: 24px; color: #111; }}
  h1 {{ font-size: 20px; }}
  table {{ border-collapse: collapse; margin-bottom: 24px; }}
  th, td {{ border: 1px solid #ddd; padding: 6px 8px; font-size: 12px; vertical-align: top; }}
  th {{ background: #f6f6f6; text-align: left; }}
  img {{ max-width: 260px; display: block; }}
  .pass {{ color: #067647; font-weight: 600; }}
  .pass-with-antialiasing {{ color: #93370d; font-weight: 600; }}
  .pass-with-diagnostic {{ color: #175cd3; font-weight: 600; }}
  .fail {{ color: #b42318; font-weight: 600; }}
  .error {{ color: #b54708; font-weight: 600; }}
  code {{ font-size: 11px; }}
</style>
</head>
<body>
<h1>L2 route parity report</h1>
<p>Upstream pin <code>{escape(report['referenceCommit'])}</code> · DPR {report['deviceScaleFactor']} ·
geometry {report['thresholds']['geometryTolerancePx']}px · pixels
{report['thresholds']['maxDiffPixelRatio'] * 100:.2f}%</p>
<h2>Cases</h2>
<table>
<thead><tr><th>route</th><th>viewport</th><th>theme</th><th>status</th><th>pixels</th><th>chrome</th><th>geometry</th><th>reference</th><th>subject</th><th>diff</th><th>heatmap</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
<h2>Classified differences</h2>
<table>
<thead><tr><th>route</th><th>id</th><th>category</th><th>blocking</th><th>status</th><th>evidence</th></tr></thead>
<tbody>{''.join(differences)}</tbody>
</table>
<h2>Excluded scaffolding</h2>
<table>
<thead><tr><th>route</th><th>scope</th><th>selector</th><th>category</th><th>reason</th></tr></thead>
<tbody>{''.join(exclusions)}</tbody>
</table>
</body>
</html>
"""
    destination.write_text(document, encoding="utf-8")


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--routes", default=str(ROUTES_FILE))
    parser.add_argument("--out", default=str(ROOT / "output/ui-parity-routes"))
    parser.add_argument("--only", default=None, help="comma separated route ids")
    parser.add_argument("--viewports", default=None, help="comma separated viewport ids")
    parser.add_argument("--themes", default=None, help="comma separated themes")
    parser.add_argument(
        "--gate-mode",
        choices=["all", "ready", "off"],
        default="off",
        help=(
            "off (default, report-only) never fails; ready fails only routes with "
            "gateStatus=blocking; all fails any route regression"
        ),
    )
    parser.add_argument(
        "--self-check",
        action="store_true",
        help="capture the reference twice and diff it with itself (harness determinism)",
    )
    parser.add_argument(
        "--skip-fixture-check",
        action="store_true",
        help="do not compare the served fixture bytes with routes.json",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = json.loads(Path(args.routes).read_text(encoding="utf-8"))
    thresholds = config["thresholds"]

    only = parse_csv(args.only)
    viewport_filter = parse_csv(args.viewports)
    theme_filter = parse_csv(args.themes)

    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)

    references_dir = ROOT / config["apps"]["reference"]["dist"]
    subject_dir = ROOT / config["apps"]["subject"]["dist"]

    fixtures: dict[str, Path] = {}
    fixture_report = []
    for name, fixture in config.get("fixtures", {}).items():
        path = ROOT / fixture["path"]
        if not path.exists():
            log(f"fixture {name} is missing at {rel(path)}")
            return 2
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest != fixture["sha256"]:
            log(
                f"fixture {name} sha256 mismatch: {digest} != {fixture['sha256']} "
                f"({rel(path)} changed; update routes.json deliberately)"
            )
            return 2
        fixtures[path.name] = path
        fixture_report.append(
            {
                "id": name,
                "path": fixture["path"],
                "sha256": digest,
                "bytes": fixture["bytes"],
            }
        )

    reference_server, reference_origin = serve(references_dir, fixtures)
    subject_server, subject_origin = serve(subject_dir, fixtures)

    viewports = [
        viewport
        for viewport in config["viewports"]
        if not viewport_filter or viewport["id"] in viewport_filter
    ]
    themes = [theme for theme in config["themes"] if not theme_filter or theme in theme_filter]
    routes = [
        route
        for route in config["routes"]
        if not only or route["id"] in only
    ]
    if not routes:
        log(f"no route matches --only {args.only}")
        return 2

    report = {
        "schemaVersion": 1,
        "harness": config["harness"],
        "referenceCommit": config["referenceCommit"],
        "referenceUrl": config["referenceUrl"],
        "deviceScaleFactor": config["deviceScaleFactor"],
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "selfCheck": args.self_check,
        "gateMode": args.gate_mode,
        "thresholds": thresholds,
        "fixtures": fixture_report,
        "viewports": viewports,
        "themes": themes,
        "routes": [],
    }

    fixture_errors: list[str] = []
    if not args.skip_fixture_check:
        for name, fixture in config.get("fixtures", {}).items():
            path = ROOT / fixture["path"]
            reference_url = f"{reference_origin}/fixtures/{path.name}"
            subject_url = f"{subject_origin}/samples/{path.name}"
            try:
                reference_digest = fetch_sha256(reference_url)
                subject_digest = fetch_sha256(subject_url)
            except Exception as error:  # noqa: BLE001 - report as a matrix error
                fixture_errors.append(f"{name}: {error}")
                continue
            if reference_digest != subject_digest:
                fixture_errors.append(
                    f"{name}: reference {reference_digest[:12]} != subject {subject_digest[:12]}"
                )
            elif reference_digest != fixture["sha256"]:
                fixture_errors.append(
                    f"{name}: served {reference_digest[:12]} != pinned {fixture['sha256'][:12]}"
                )

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch()
            try:
                for route in routes:
                    log(f"▸ route {route['id']}")
                    cases = []
                    for viewport in viewports:
                        for theme in themes:
                            case = run_case(
                                browser=browser,
                                route=route,
                                viewport=viewport,
                                theme=theme,
                                reference_origin=reference_origin,
                                subject_origin=subject_origin,
                                out=out,
                                thresholds=thresholds,
                                device_scale_factor=config["deviceScaleFactor"],
                                self_check=args.self_check,
                            )
                            log(
                                f"  {viewport['id']}/{theme}: {case['status']}"
                                f" pixels={(case.get('pixels') or {}).get('diffRatio', 0) * 100:.3f}%"
                                f" chrome={(case.get('chrome') or {}).get('diffRatio', 0) * 100:.3f}%"
                                f" geometry={(case.get('geometry') or {}).get('worstDeltaPx', 0):.2f}px"
                            )
                            if case.get("errors"):
                                for error in case["errors"][:5]:
                                    log(f"    ! {error}")
                            cases.append(case)
                    summary = summarize_route(route, cases)
                    summary["cases"] = cases
                    report["routes"].append(summary)
            finally:
                browser.close()
    finally:
        reference_server.shutdown()
        subject_server.shutdown()

    if fixture_errors:
        report["fixtureErrors"] = fixture_errors
        for error in fixture_errors:
            log(f"fixture error: {error}")

    matrix_path = out / "routes-matrix.json"
    matrix_path.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    (out / "report.json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    write_summary(report, out / "SUMMARY.md")
    write_html(report, out / "report.html")

    blocking_failures = []
    for route in report["routes"]:
        if route["status"] != "fail":
            continue
        if args.gate_mode == "all":
            blocking_failures.append(route["id"])
        elif args.gate_mode == "ready" and route["blocking"]:
            blocking_failures.append(route["id"])

    log("")
    log(f"routes: {', '.join(route['id'] for route in report['routes'])}")
    log(f"matrix: {rel(matrix_path)}")
    log(f"report: {rel(out / 'report.html')}")
    if fixture_errors:
        log("fixture bytes did not match routes.json; treat the matrix as diagnostic")
        return 1
    if blocking_failures:
        log(f"blocking route failures ({args.gate_mode}): {', '.join(blocking_failures)}")
        return 1
    failed = [route["id"] for route in report["routes"] if route["status"] == "fail"]
    if failed:
        log(f"report-only route failures: {', '.join(failed)} (gate mode {args.gate_mode})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
