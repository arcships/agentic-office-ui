#!/usr/bin/env python3
"""Extend UI <-> @arcships/vue-ui pixel-parity harness.

Serves the built reference app (tools/extend-reference/dist) and the built Vue
comparison app (tools/vue-reference/dist) from local threads, drives Chromium at
DPR 2, and writes screenshots, diff images, a heatmap, geometry deltas and a
JSON/HTML report.

Usage (normally invoked by scripts/ci/pixel-parity.mjs):

    python3 scripts/ci/pixel-parity.py \
        --manifest tools/parity-fixtures/manifest.json \
        --reference-dir tools/extend-reference/dist \
        --subject-dir tools/vue-reference/dist \
        --tokens docs/design/extend-tokens.json \
        --out output/ui-parity

Thresholds (manifest `thresholds`):
  geometryTolerancePx   allowed delta for every measured box/padding/gap/radius
  maxChannelDelta       per-channel delta treated as text anti-aliasing (2/255)
  maxDiffPixelRatio     share of differing pixels allowed inside the AA rule
  maxContiguousDiffPx   longest allowed run of differing pixels (2 => 2x2 box)
  topDiffRegions        number of clustered diff regions reported per failure

Modes:
  default                  live reference-vs-subject parity, every failure blocks
  --gate-mode ready        only `gateStatus: blocking` surfaces block (CI)
  --report-only            alias for `--gate-mode off` (artifact mode)
  --baseline               also compare the subject with tests/visual/extend/*
  --write-baseline         write the subject captures into the baseline tree
                           (`pnpm ui:baseline` owns approval + manifest hashes)

Baselines are named `<surface>/<state>/<theme>-<viewport>.png` and only the
`manifest.baselines` viewport/theme combinations are compared.
"""

from __future__ import annotations

import argparse
import functools
import html
import http.server
import importlib.metadata
import json
import platform
import shutil
import socketserver
import sys
import threading
import time
import traceback
from pathlib import Path

from playwright.sync_api import sync_playwright
from PIL import Image, ImageChops, ImageFilter

Image.MAX_IMAGE_PIXELS = None

ROOT = Path(__file__).resolve().parents[2]

GEOMETRY_KEYS = [
    "width",
    "height",
    "paddingTop",
    "paddingRight",
    "paddingBottom",
    "paddingLeft",
    "rowGap",
    "columnGap",
    "borderTopWidth",
    "borderRightWidth",
    "borderBottomWidth",
    "borderLeftWidth",
    "borderTopLeftRadius",
    "borderTopRightRadius",
    "borderBottomRightRadius",
    "borderBottomLeftRadius",
    "fontSize",
    "lineHeight",
]

STYLE_KEYS = [
    "color",
    "backgroundColor",
    "fontFamily",
    "fontWeight",
    "letterSpacing",
    "boxShadow",
    "outlineWidth",
    "outlineStyle",
    "outlineOffset",
    "opacity",
    "display",
]

METRIC_SCRIPT = """
(el) => {
  const s = getComputedStyle(el);
  const r = el.getBoundingClientRect();
  const pick = (keys) => Object.fromEntries(keys.map((key) => [key, s[key]]));
  return {
    box: { x: r.x, y: r.y, width: r.width, height: r.height },
    offset: { width: el.offsetWidth, height: el.offsetHeight },
    metrics: pick(%s),
    style: pick(%s),
  };
}
""" % (
    json.dumps(GEOMETRY_KEYS),
    json.dumps(STYLE_KEYS),
)


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):  # noqa: A002 - stdlib signature
        pass


class Server:
    def __init__(self, directory: Path):
        handler = functools.partial(QuietHandler, directory=str(directory))
        self.httpd = socketserver.ThreadingTCPServer(("127.0.0.1", 0), handler)
        self.httpd.daemon_threads = True
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()
        self.url = f"http://127.0.0.1:{self.httpd.server_address[1]}"

    def close(self):
        self.httpd.shutdown()
        self.httpd.server_close()


def log(message: str) -> None:
    print(message, flush=True)


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(ROOT))
    except ValueError:
        return str(path)


def parse_csv(value: str | None) -> list[str] | None:
    if not value:
        return None
    return [item.strip() for item in value.split(",") if item.strip()]


# --------------------------------------------------------------------------
# Playwright page setup
# --------------------------------------------------------------------------


def prepare_page(page, side: str) -> list[str]:
    """Freeze animations and fonts before any screenshot."""
    errors: list[str] = []
    page.on(
        "pageerror",
        lambda error: errors.append(f"{side} pageerror: {error}"),
    )
    page.on(
        "console",
        lambda message: errors.append(f"{side} console.{message.type}: {message.text}")
        if message.type == "error" and "favicon" not in message.text
        else None,
    )
    page.emulate_media(reduced_motion="reduce")
    page.add_init_script(
        """
        document.addEventListener('DOMContentLoaded', () => {
          const style = document.createElement('style');
          style.setAttribute('data-parity-harness', 'animations-off');
          style.textContent = `*, *::before, *::after {
            animation: none !important;
            transition: none !important;
            caret-color: transparent !important;
            scroll-behavior: auto !important;
          }`;
          document.head.appendChild(style);
        });
        """
    )
    return errors


def wait_ready(page, surface: str) -> None:
    page.wait_for_selector(
        f"[data-parity-surface='{surface}'] [data-parity-cell]",
        state="attached",
        timeout=30_000,
    )
    page.wait_for_function("document.fonts.status === 'loaded'", timeout=30_000)
    page.evaluate("() => document.fonts.ready")
    page.wait_for_timeout(120)


def cell_locator(page, state: str):
    return page.locator(f"[data-parity-cell='{state}']")


def target_locator(page, cell, selector: str | None):
    if selector:
        target = cell.locator(selector)
    else:
        target = cell.locator("[data-parity-host] > *").first
    return target


def apply_interaction(page, target, interaction: dict, cell=None) -> None:
    kind = interaction.get("type", "none")
    selector = interaction.get("selector")
    element = resolve_interaction_target(target, selector, cell)

    if kind == "none":
        return
    if kind == "hover":
        element.hover(timeout=5_000)
        page.wait_for_timeout(80)
        return
    if kind == "focus":
        element.focus(timeout=5_000)
        page.wait_for_timeout(80)
        return
    if kind == "press":
        box = element.bounding_box()
        if box is None:
            raise RuntimeError("press: element has no box")
        page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        page.mouse.down()
        page.wait_for_timeout(60)
        return
    if kind == "dragenter":
        scope = target.first.element_handle()
        page.evaluate(
            """
            ({ root, selector }) => {
              const element = root.matches(selector) ? root : root.querySelector(selector);
              if (!element) throw new Error('dragenter: no ' + selector);
              const transfer = new DataTransfer();
              transfer.items.add(new File(['parity'], 'sample.pdf', { type: 'application/pdf' }));
              for (const type of ['dragenter', 'dragover']) {
                element.dispatchEvent(new DragEvent(type, {
                  bubbles: true,
                  cancelable: true,
                  dataTransfer: transfer,
                }));
              }
            }
            """,
            {"root": scope, "selector": selector},
        )
        page.wait_for_timeout(120)
        return
    if kind == "set-input-files":
        payload = interaction["fixture"]
        element.set_input_files(
            {
                "name": payload["name"],
                "mimeType": payload["mimeType"],
                "buffer": payload["content"].encode("utf-8"),
            }
        )
        page.wait_for_timeout(150)
        return
    raise RuntimeError(f"unsupported interaction: {kind}")


def resolve_interaction_target(target, selector: str | None, cell=None):
    """A selector may match the measured target itself (a Button cell whose
    target is the <button>), a descendant (the FileUpload dropzone), or a
    sibling inside the same cell (a Tooltip bubble is measured while its
    trigger is hovered)."""
    first = target.first
    if not selector:
        return first
    if target.count() == 0:
        if cell is not None:
            return cell.locator(selector).first
        return first
    if first.evaluate("(el, selector) => el.matches(selector)", selector):
        return first
    descendant = target.locator(selector).first
    if descendant.count():
        return descendant
    if cell is not None:
        in_cell = cell.locator(selector).first
        if in_cell.count():
            return in_cell
    return descendant


def release_interaction(page, interaction: dict) -> None:
    if interaction.get("type") == "press":
        page.mouse.up()


def capture(page, cell, target, destination: Path, padding: int) -> dict:
    scroll = page.evaluate("() => ({ x: window.scrollX, y: window.scrollY })")
    document = page.evaluate(
        "() => ({ width: document.documentElement.scrollWidth, height: document.documentElement.scrollHeight })"
    )
    box = target.bounding_box()
    if box is None:
        raise RuntimeError("target has no bounding box")
    metrics = target.evaluate(METRIC_SCRIPT)
    metrics["box"] = box

    x = max(0.0, box["x"] + scroll["x"] - padding)
    y = max(0.0, box["y"] + scroll["y"] - padding)
    right = min(document["width"], box["x"] + scroll["x"] + box["width"] + padding)
    bottom = min(document["height"], box["y"] + scroll["y"] + box["height"] + padding)
    clip = {
        "x": round(x, 2),
        "y": round(y, 2),
        "width": round(max(1.0, right - x), 2),
        "height": round(max(1.0, bottom - y), 2),
    }
    destination.parent.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(destination), clip=clip, full_page=True, animations="disabled")
    return {"screenshot": clip, "metrics": metrics}


# --------------------------------------------------------------------------
# Image comparison
# --------------------------------------------------------------------------


def shifted_and(mask: Image.Image, dx: int, dy: int) -> Image.Image:
    width, height = mask.size
    shifted = Image.new("L", (width, height), 0)
    shifted.paste(mask, (dx, dy))
    return ImageChops.multiply(mask, shifted)


def longest_run(mask: Image.Image, limit: int = 8) -> int:
    best = 1
    for dx, dy in ((1, 0), (0, 1), (1, 1), (1, -1)):
        current = mask
        length = 1
        while length < limit:
            candidate = shifted_and(current, dx, dy)
            if candidate.getbbox() is None:
                break
            current = candidate
            length += 1
        best = max(best, length)
    return best if mask.getbbox() is not None else 0


REGION_CELL_PX = 16


def diff_regions(
    max_diff: Image.Image,
    mask: Image.Image,
    limit: int,
    grid: int = REGION_CELL_PX,
) -> list[dict]:
    """Cluster the differing pixels into coarse regions, worst first.

    A 16px grid keeps this cheap on a 4000px-tall capture: cells that contain
    any differing pixel are connected (8-neighbourhood), then each cluster is
    measured against the real pixel mask so the reported pixel count and peak
    channel delta are exact.
    """
    if limit <= 0 or mask.getbbox() is None:
        return []
    width, height = mask.size
    columns = max(1, (width + grid - 1) // grid)
    rows = max(1, (height + grid - 1) // grid)
    cells = mask.resize((columns, rows), Image.BOX).load()
    seen = [[False] * columns for _ in range(rows)]
    clusters: list[tuple[int, int, int, int]] = []
    for start_y in range(rows):
        for start_x in range(columns):
            if seen[start_y][start_x] or not cells[start_x, start_y]:
                continue
            seen[start_y][start_x] = True
            stack = [(start_x, start_y)]
            min_x = max_x = start_x
            min_y = max_y = start_y
            while stack:
                cell_x, cell_y = stack.pop()
                min_x, max_x = min(min_x, cell_x), max(max_x, cell_x)
                min_y, max_y = min(min_y, cell_y), max(max_y, cell_y)
                for dx, dy in (
                    (1, 0), (-1, 0), (0, 1), (0, -1),
                    (1, 1), (1, -1), (-1, 1), (-1, -1),
                ):
                    next_x, next_y = cell_x + dx, cell_y + dy
                    if (
                        0 <= next_x < columns
                        and 0 <= next_y < rows
                        and not seen[next_y][next_x]
                        and cells[next_x, next_y]
                    ):
                        seen[next_y][next_x] = True
                        stack.append((next_x, next_y))
            clusters.append((min_x, min_y, max_x, max_y))
    if not clusters:
        # A 16x16 BOX reduction of a single differing pixel rounds to 1, so
        # this is a safety net rather than the normal path.
        bbox = mask.getbbox()
        clusters = [
            (
                bbox[0] // grid,
                bbox[1] // grid,
                max(0, (bbox[2] - 1) // grid),
                max(0, (bbox[3] - 1) // grid),
            )
        ]

    regions = []
    for min_x, min_y, max_x, max_y in clusters:
        box = (
            min_x * grid,
            min_y * grid,
            min(width, (max_x + 1) * grid),
            min(height, (max_y + 1) * grid),
        )
        pixels = mask.crop(box).histogram()[255]
        peak = max_diff.crop(box).getextrema()[1]
        regions.append(
            {
                "x": box[0],
                "y": box[1],
                "width": box[2] - box[0],
                "height": box[3] - box[1],
                "diffPixels": pixels,
                "maxChannelDelta": peak,
            }
        )
    regions.sort(key=lambda region: region["diffPixels"], reverse=True)
    return regions[:limit]


def diff_metrics(reference: Image.Image, subject: Image.Image, thresholds: dict) -> dict:
    diff = ImageChops.difference(reference, subject)
    red, green, blue, alpha = diff.split()
    max_diff = ImageChops.lighter(ImageChops.lighter(red, green), blue)
    max_diff = ImageChops.lighter(max_diff, alpha)

    histogram = max_diff.histogram()
    total_pixels = reference.size[0] * reference.size[1]
    diff_pixels = total_pixels - histogram[0]
    ratio = diff_pixels / total_pixels if total_pixels else 0.0
    max_channel = max(index for index, count in enumerate(histogram) if count)
    mask = max_diff.point(lambda value: 255 if value else 0)
    run = longest_run(mask)
    has_3x3 = mask.filter(ImageFilter.MinFilter(3)).getbbox() is not None

    reasons = []
    if diff_pixels:
        if max_channel > thresholds["maxChannelDelta"]:
            reasons.append(
                f"max channel delta {max_channel} > {thresholds['maxChannelDelta']}"
            )
        if ratio > thresholds["maxDiffPixelRatio"]:
            reasons.append(
                f"diff ratio {ratio:.5f} > {thresholds['maxDiffPixelRatio']}"
            )
        if run > 2 or has_3x3:
            reasons.append(
                f"contiguous diff run {run}px (3x3 block: {has_3x3}) exceeds the 2x2 anti-aliasing allowance"
            )

    status = "pass" if diff_pixels == 0 else ("fail" if reasons else "pass-with-antialiasing")
    regions = diff_regions(
        max_diff,
        mask,
        int(thresholds.get("topDiffRegions", 5)),
    ) if diff_pixels else []
    return {
        "status": status,
        "reasons": reasons,
        "diffPixels": diff_pixels,
        "diffRatio": round(ratio, 6),
        "maxChannelDelta": max_channel,
        "maxContiguousRunPx": run,
        "has3x3Block": has_3x3,
        "regions": regions,
        "regionGridPx": REGION_CELL_PX,
        "width": reference.size[0],
        "height": reference.size[1],
        "_diff": max_diff,
        "_mask": mask,
        "_subject": subject,
    }


def write_diff_artifacts(result: dict, diff_path: Path | None, heatmap_path: Path | None, thresholds: dict) -> None:
    if result["status"] == "pass":
        return
    max_diff = result["_diff"]
    if diff_path is not None:
        diff_path.parent.mkdir(parents=True, exist_ok=True)
        max_diff.save(diff_path)
    if result["status"] == "fail" and heatmap_path is not None:
        heatmap_path.parent.mkdir(parents=True, exist_ok=True)
        subject = result["_subject"]
        amplified = max_diff.point(
            lambda value: min(255, value * thresholds["heatmapAmplification"])
        )
        dimmed = subject.convert("RGB").point(lambda value: value // 4)
        red_overlay = Image.merge(
            "RGB",
            (amplified, Image.new("L", subject.size, 0), Image.new("L", subject.size, 0)),
        )
        Image.composite(red_overlay, dimmed, result["_mask"]).save(heatmap_path)


def strip_private(result: dict) -> dict:
    return {key: value for key, value in result.items() if not key.startswith("_")}


def compare_images(
    reference_path: Path,
    subject_path: Path,
    diff_path: Path | None,
    heatmap_path: Path | None,
    compare_path: Path | None,
    thresholds: dict,
) -> dict:
    reference = Image.open(reference_path).convert("RGBA")
    subject = Image.open(subject_path).convert("RGBA")
    if reference.size != subject.size:
        if compare_path is not None:
            compare_path.parent.mkdir(parents=True, exist_ok=True)
            gap = 16
            width = max(reference.size[0], subject.size[0])
            height = reference.size[1] + gap + subject.size[1]
            canvas = Image.new("RGBA", (width, height), (255, 255, 255, 255))
            canvas.paste(reference, (0, 0))
            canvas.paste(subject, (0, reference.size[1] + gap))
            canvas.save(compare_path)

        # Sizes differ, so the pixel gate cannot pass; still diff the shared
        # top-left region so colour, border and type deltas are visible. The
        # overlap result is diagnostic only and never turns a case green.
        width = min(reference.size[0], subject.size[0])
        height = min(reference.size[1], subject.size[1])
        overlap = diff_metrics(
            reference.crop((0, 0, width, height)),
            subject.crop((0, 0, width, height)),
            thresholds,
        )
        overlap_path = None
        overlap_heatmap = None
        if diff_path is not None:
            overlap_path = diff_path.with_name("overlap-diff.png")
            overlap_heatmap = diff_path.with_name("overlap-heatmap.png")
        write_diff_artifacts(overlap, overlap_path, overlap_heatmap, thresholds)
        return {
            "status": "fail",
            "reasons": [
                f"capture size mismatch: reference {reference.size[0]}x{reference.size[1]} vs subject {subject.size[0]}x{subject.size[1]}",
            ],
            "sizeMismatch": True,
            "referenceSize": list(reference.size),
            "subjectSize": list(subject.size),
            "overlap": strip_private(overlap),
            "overlapDiffScreenshot": overlap_path.name if overlap_path and overlap_path.exists() else None,
            "overlapHeatmapScreenshot": overlap_heatmap.name if overlap_heatmap and overlap_heatmap.exists() else None,
        }

    result = diff_metrics(reference, subject, thresholds)
    write_diff_artifacts(result, diff_path, heatmap_path, thresholds)
    return strip_private(result)


# --------------------------------------------------------------------------
# Committed baselines (tests/visual/extend)
# --------------------------------------------------------------------------


def baseline_relpath(surface_id: str, state_id: str, theme: str, viewport_id: str) -> str:
    """`<component>/<state>/<theme>-<viewport>.png` — the committed naming scheme."""
    return f"{surface_id}/{state_id}/{theme}-{viewport_id}.png"


def compare_baseline(
    baseline_path: Path,
    subject_path: Path,
    case_dir: Path,
    out: Path,
    thresholds: dict,
) -> dict:
    if not baseline_path.exists():
        return {"status": "missing", "baseline": rel(baseline_path)}
    result = compare_images(
        baseline_path,
        subject_path,
        case_dir / "baseline-diff.png",
        case_dir / "baseline-heatmap.png",
        None,
        thresholds,
    )
    status = result.pop("status")
    entry = {"status": status, "baseline": rel(baseline_path), **result}
    diff_file = case_dir / "baseline-diff.png"
    heatmap_file = case_dir / "baseline-heatmap.png"
    if diff_file.exists():
        entry["screenshot"] = artifact_rel(diff_file, out)
    if heatmap_file.exists():
        entry["heatmap"] = artifact_rel(heatmap_file, out)
    return entry


def write_baseline(baseline_path: Path, subject_path: Path) -> dict:
    baseline_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(subject_path, baseline_path)
    return {"status": "written", "baseline": rel(baseline_path)}


def compare_metrics(reference: dict, subject: dict, tolerance: float) -> dict:
    deltas = []
    for key in GEOMETRY_KEYS:
        left = reference["metrics"].get(key)
        right = subject["metrics"].get(key)
        delta = numeric_delta(left, right)
        if delta is None:
            if str(left) != str(right):
                deltas.append({"key": key, "reference": left, "subject": right, "failing": True})
            continue
        if delta > tolerance:
            deltas.append(
                {
                    "key": key,
                    "reference": left,
                    "subject": right,
                    "delta": round(delta, 4),
                    "failing": True,
                }
            )
    style_deltas = []
    for key in STYLE_KEYS:
        left = reference["style"].get(key)
        right = subject["style"].get(key)
        if str(left) != str(right):
            style_deltas.append({"key": key, "reference": left, "subject": right})
    return {
        "status": "fail" if any(item["failing"] for item in deltas) else "pass",
        "deltas": deltas,
        "styleDeltas": style_deltas,
        "boxDelta": {
            "width": round(subject["box"]["width"] - reference["box"]["width"], 4),
            "height": round(subject["box"]["height"] - reference["box"]["height"], 4),
        },
    }


def numeric_delta(left, right) -> float | None:
    try:
        return abs(float(str(left).replace("px", "")) - float(str(right).replace("px", "")))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------
# Report rendering
# --------------------------------------------------------------------------


def write_html_report(report: dict, destination: Path) -> None:
    rows = []
    for case in report["cases"]:
        images = ""
        for label, key in (("reference", "reference"), ("subject", "subject")):
            screenshot = case.get(key, {}).get("screenshot")
            if screenshot:
                images += (
                    f'<figure><img src="{html.escape(screenshot)}" loading="lazy" />'
                    f"<figcaption>{label}</figcaption></figure>"
                )
        if case.get("diff", {}).get("screenshot"):
            images += (
                f'<figure><img src="{html.escape(case["diff"]["screenshot"])}" loading="lazy" />'
                f"<figcaption>diff</figcaption></figure>"
            )
        if case.get("heatmap"):
            images += (
                f'<figure><img src="{html.escape(case["heatmap"])}" loading="lazy" />'
                f"<figcaption>heatmap</figcaption></figure>"
            )
        if case.get("compare"):
            images += (
                f'<figure><img src="{html.escape(case["compare"])}" loading="lazy" />'
                f"<figcaption>reference / subject</figcaption></figure>"
            )
        baseline = case.get("baseline") or {}
        if baseline.get("screenshot"):
            images += (
                f'<figure><img src="{html.escape(baseline["screenshot"])}" loading="lazy" />'
                f"<figcaption>baseline diff</figcaption></figure>"
            )
        if baseline.get("heatmap"):
            images += (
                f'<figure><img src="{html.escape(baseline["heatmap"])}" loading="lazy" />'
                f"<figcaption>baseline heatmap</figcaption></figure>"
            )
        for key, caption in (
            ("overlapDiffScreenshot", "overlap diff"),
            ("overlapHeatmapScreenshot", "overlap heatmap"),
        ):
            path = case.get("diff", {}).get(key)
            if path:
                images += (
                    f'<figure><img src="{html.escape(path)}" loading="lazy" />'
                    f"<figcaption>{caption}</figcaption></figure>"
                )
        failures = []
        if case.get("diff", {}).get("reasons"):
            failures.extend(case["diff"]["reasons"])
        for delta in case.get("geometry", {}).get("deltas", []):
            failures.append(
                f"geometry {delta['key']}: {delta['reference']} -> {delta['subject']}"
            )
        failures.extend(case.get("errors", []))
        if baseline.get("status") == "missing":
            failures.append(f"baseline missing: {baseline.get('baseline')}")
        elif baseline.get("status") == "fail":
            for region in (baseline.get("regions") or [])[:3]:
                failures.append(
                    f"baseline diff region {region['width']}x{region['height']} "
                    f"at ({region['x']}, {region['y']}) — {region['diffPixels']} px"
                )
        for region in (case.get("diff", {}).get("regions") or [])[:3]:
            failures.append(
                f"diff region {region['width']}x{region['height']} "
                f"at ({region['x']}, {region['y']}) — {region['diffPixels']} px"
            )
        if baseline.get("status"):
            failures.append(f"baseline: {baseline['status']}")
        rows.append(
            """
            <tr class="status-{status}">
              <td>{surface}</td><td>{state}</td><td>{viewport}</td><td>{theme}</td>
              <td><strong>{status}</strong><br/><small>{detail}</small></td>
              <td class="images">{images}</td>
            </tr>
            """.format(
                status=html.escape(case["status"]),
                surface=html.escape(case["surface"]),
                state=html.escape(case["state"]),
                viewport=html.escape(case["viewport"]),
                theme=html.escape(case["theme"]),
                detail=html.escape(
                    "; ".join(failures[:4])
                    if failures
                    else (
                        f"{case['diff'].get('diffPixels', 0)} px diff"
                        if case.get("diff")
                        else case.get("reason", "")
                    )
                ),
                images=images,
            )
        )

    token_rows = []
    for entry in report["tokens"].get("tokenLayerMismatches", []):
        token_rows.append(
            f"<tr><td>{html.escape(entry['group'])}</td><td>{html.escape(entry['token'])}</td>"
            f"<td>{html.escape(str(entry['extend'].get('raw')))}</td>"
            f"<td>{html.escape(str(entry['office'].get('raw')))}</td></tr>"
        )
    color_rows = []
    for entry in report["tokens"].get("colorMismatches", []):
        color_rows.append(
            f"<tr><td>{html.escape(entry['token'])}</td>"
            f"<td>{html.escape(str(entry['extendLight']))}</td>"
            f"<td>{html.escape(str(entry['officeLight']))}</td>"
            f"<td>{html.escape(str(entry['status']))}</td></tr>"
        )

    summary = report["summary"]
    document = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Extend UI pixel parity</title>
<style>
  body {{ font-family: ui-sans-serif, system-ui, sans-serif; margin: 24px; color: #111; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th, td {{ border: 1px solid #ddd; padding: 6px; vertical-align: top; }}
  th {{ background: #f5f5f5; text-align: left; }}
  .images {{ display: flex; gap: 8px; flex-wrap: wrap; }}
  figure {{ margin: 0; }}
  img {{ max-width: 360px; max-height: 220px; border: 1px solid #eee; background: #fff; }}
  figcaption {{ font-size: 11px; color: #666; }}
  tr.status-fail td {{ background: #fff3f3; }}
  tr.status-pass td {{ background: #f6fff6; }}
  tr.status-pass-with-antialiasing td {{ background: #fffdf0; }}
  tr.status-baseline-only td {{ background: #f4f6ff; }}
  code {{ background: #f3f3f3; padding: 1px 4px; }}
</style>
</head>
<body>
<h1>Extend UI &harr; @arcships/vue-ui pixel parity</h1>
<p>
  reference <code>{html.escape(report["reference"]["repository"])}@{html.escape(report["reference"]["commit"])}</code>
  · chromium <code>{html.escape(report["environment"]["browserVersion"])}</code>
  · DPR {report["environment"]["deviceScaleFactor"]}
  · {html.escape(report["generatedAt"])}
</p>
<p>
  <strong>{summary["passed"]}</strong> pass ·
  <strong>{summary["failed"]}</strong> fail ·
  {summary["baselineOnly"]} reference-only ·
  {summary["total"]} cases
</p>
<p>
  gate <code>{html.escape(summary["gate"])}</code> ·
  mode <code>{html.escape(str(summary.get("gateMode")))}</code> ·
  baseline <code>{html.escape(str(summary.get("baselineMode")))}</code> ·
  blocking <strong>{summary.get("blockingFailures", 0)}</strong> ·
  pending {summary.get("pendingFailures", 0)}
</p>
<h2>Token diff (Extend &rarr; --office-*)</h2>
<table>
  <thead><tr><th>Group</th><th>Token</th><th>Extend</th><th>Office (#156)</th></tr></thead>
  <tbody>{''.join(token_rows) or '<tr><td colspan="4">no token mismatches</td></tr>'}</tbody>
</table>
<p></p>
<table>
  <thead><tr><th>Color token</th><th>Extend light</th><th>Office light</th><th>Status</th></tr></thead>
  <tbody>{''.join(color_rows) or '<tr><td colspan="4">no color gaps or mismatches</td></tr>'}</tbody>
</table>
<h2>Cases</h2>
<table>
  <thead><tr><th>Surface</th><th>State</th><th>Viewport</th><th>Theme</th><th>Status</th><th>Screens</th></tr></thead>
  <tbody>{''.join(rows)}</tbody>
</table>
</body>
</html>
"""
    destination.write_text(document, encoding="utf-8")


def write_summary_markdown(report: dict, destination: Path) -> None:
    summary = report["summary"]
    lines = [
        "# Extend UI pixel parity report",
        "",
        f"- reference: `{report['reference']['repository']}@{report['reference']['commit']}`",
        f"- generated: `{report['generatedAt']}`",
        f"- chromium: `{report['environment']['browserVersion']}` at DPR {report['environment']['deviceScaleFactor']}",
        f"- cases: {summary['total']} (pass {summary['passed']}, pass-with-AA {summary['passedWithAntialiasing']}, fail {summary['failed']}, reference-only {summary['baselineOnly']})",
        f"- gate: `{summary['gate']}` (mode `{summary.get('gateMode')}`, baseline `{summary.get('baselineMode')}`)",
        f"- blocking failures: {summary.get('blockingFailures', 0)}; pending (reported only): {summary.get('pendingFailures', 0)}",
        f"- baseline: {summary.get('baselineWritten', 0)} written, {summary.get('baselineFailed', 0)} failed, {summary.get('baselineMissing', 0)} missing",
        "",
        "## Per surface",
        "",
        "| Surface | Subject | Core | Gate | Pending | Pass | Fail | Reference-only |",
        "|---|---|---|---|---|---|---|---|",
    ]
    for surface in report["surfaces"]:
        counts = surface["counts"]
        lines.append(
            f"| {surface['id']} | {surface['subjectComponent'] or '—'} "
            f"| {'yes' if surface.get('core') else 'no'} | {surface.get('gateStatus') or '—'} "
            f"| {surface['pendingIssue'] or '—'} "
            f"| {counts['passed'] + counts['passedWithAntialiasing']} | {counts['failed']} | {counts['baselineOnly']} |"
        )
    lines.extend(
        [
            "",
            "## Token diff",
            "",
            f"- token layer: {report['tokens']['tokenLayerSummary']}",
            f"- color layer: {report['tokens']['colorSummary']}",
            "",
            "| Group | Token | Extend | Office |",
            "|---|---|---|---|",
        ]
    )
    for entry in report["tokens"]["tokenLayerMismatches"]:
        lines.append(
            f"| {entry['group']} | `{entry['token']}` | `{entry['extend'].get('raw')}` | `{entry['office'].get('raw')}` |"
        )
    for entry in report["tokens"]["colorMismatches"]:
        lines.append(
            f"| color | `{entry['token']}` | `{entry['extendLight']}` | `{entry['officeLight']}` ({entry['status']}) |"
        )
    lines.extend(
        [
            "",
            "## Failing cases",
            "",
        ]
    )
    failing = [
        case
        for case in report["cases"]
        if case["status"] == "fail"
        or case["status"] == "error"
        or (case.get("baseline") or {}).get("status") in {"fail", "missing"}
    ]
    if not failing:
        lines.append("None.")
    for case in failing[:200]:
        reasons = case.get("diff", {}).get("reasons", []) + [
            f"geometry {delta['key']} {delta['reference']} -> {delta['subject']}"
            for delta in case.get("geometry", {}).get("deltas", [])
        ]
        baseline = case.get("baseline") or {}
        if baseline.get("status") == "missing":
            reasons.append(f"baseline missing: `{baseline.get('baseline')}`")
        elif baseline.get("status") == "fail":
            reasons.append(
                "baseline diff "
                + ", ".join(
                    f"{region['width']}x{region['height']} at ({region['x']}, {region['y']}) "
                    f"= {region['diffPixels']} px"
                    for region in (baseline.get("regions") or [])[:5]
                )
            )
        regions = case.get("diff", {}).get("regions") or []
        if regions:
            reasons.append(
                "top diff regions "
                + ", ".join(
                    f"{region['width']}x{region['height']} at ({region['x']}, {region['y']})"
                    for region in regions[:5]
                )
            )
        lines.append(
            f"- `{case['surface']}` / `{case['state']}` / `{case['viewport']}` / `{case['theme']}`: "
            + ("; ".join(reasons) or "subject missing")
        )
    destination.write_text("\n".join(lines) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", default=str(ROOT / "tools/parity-fixtures/manifest.json"))
    parser.add_argument("--reference-dir", default=str(ROOT / "tools/extend-reference/dist"))
    parser.add_argument("--subject-dir", default=str(ROOT / "tools/vue-reference/dist"))
    parser.add_argument("--tokens", default=str(ROOT / "docs/design/extend-tokens.json"))
    parser.add_argument("--out", default=str(ROOT / "output/ui-parity"))
    parser.add_argument("--only", default=None, help="comma separated surface ids (`core` expands to the PR subset)")
    parser.add_argument("--viewports", default=None, help="comma separated viewport ids")
    parser.add_argument("--themes", default=None, help="comma separated themes")
    parser.add_argument("--report-only", action="store_true", help="always exit 0")
    parser.add_argument(
        "--gate-mode",
        choices=["all", "ready", "off"],
        default=None,
        help=(
            "which parity failures fail the run: all (default), ready (only "
            "gateStatus=blocking surfaces), off (artifact mode)"
        ),
    )
    parser.add_argument(
        "--baseline",
        action="store_true",
        help="compare the subject captures with the committed baselines",
    )
    parser.add_argument(
        "--write-baseline",
        action="store_true",
        help="write the subject captures into the baseline directory (use pnpm ui:baseline instead)",
    )
    parser.add_argument(
        "--baseline-dir",
        default=None,
        help="baseline directory (default: manifest.baselines.directory)",
    )
    parser.add_argument(
        "--self-check",
        action="store_true",
        help="capture the reference page twice and compare it with itself (harness determinism gate)",
    )
    parser.add_argument("--capture-padding", type=int, default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    fixtures = json.loads(
        (Path(args.manifest).parent / "fixtures.json").read_text(encoding="utf-8")
    )

    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    cases_dir = out / "cases"
    cases_dir.mkdir(parents=True, exist_ok=True)

    only = parse_csv(args.only)
    viewport_filter = parse_csv(args.viewports)
    theme_filter = parse_csv(args.themes)

    if only and "core" in only:
        remaining = [item for item in only if item != "core"]
        only = remaining + [
            surface["id"] for surface in manifest["surfaces"] if surface.get("core")
        ]

    baseline_config = manifest.get("baselines") or {}
    if args.baseline and args.write_baseline:
        log("--baseline and --write-baseline are mutually exclusive")
        return 2
    baseline_mode = "write" if args.write_baseline else ("compare" if args.baseline else "off")
    if args.self_check:
        # The self-check compares the reference page with itself; committed
        # subject baselines do not apply.
        baseline_mode = "off"
    baseline_dir = None
    baseline_combos: set[tuple[str, str]] = set()
    if baseline_mode != "off":
        baseline_dir = Path(
            args.baseline_dir or baseline_config.get("directory") or "tests/visual/extend"
        )
        if not baseline_dir.is_absolute():
            baseline_dir = ROOT / baseline_dir
        baseline_combos = {
            (viewport_id, theme)
            for viewport_id in baseline_config.get("viewports", [v["id"] for v in manifest["viewports"]])
            for theme in baseline_config.get("themes", manifest["themes"])
        }

    gate_mode = args.gate_mode or ("off" if args.report_only else "all")
    if baseline_mode == "write" and args.gate_mode is None:
        gate_mode = "off"

    padding = args.capture_padding
    if padding is None:
        padding = (
            baseline_config.get("capturePaddingPx", 12)
            if baseline_mode != "off"
            else 12
        )

    viewports = [
        viewport
        for viewport in manifest["viewports"]
        if not viewport_filter or viewport["id"] in viewport_filter
    ]
    themes = [theme for theme in manifest["themes"] if not theme_filter or theme in theme_filter]
    surfaces = [
        surface
        for surface in manifest["surfaces"]
        if surface.get("referenceStatus") != "fixtures-only"
        and (not only or surface["id"] in only)
        and surface.get("states")
    ]
    if not surfaces:
        log("no surfaces selected")
        return 2

    thresholds = manifest["thresholds"]
    if "topDiffRegions" not in thresholds:
        thresholds["topDiffRegions"] = (manifest.get("gate") or {}).get("topDiffRegions", 5)
    reference_server = Server(Path(args.reference_dir).resolve())
    subject_server = Server(Path(args.subject_dir).resolve())

    started = time.time()
    cases: list[dict] = []
    token_report = {}
    if Path(args.tokens).exists():
        token_report = json.loads(Path(args.tokens).read_text(encoding="utf-8"))

    self_check = args.self_check
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            browser_version = browser.version
            log(f"chromium {browser_version}")
            for viewport in viewports:
                context = browser.new_context(
                    viewport={"width": viewport["width"], "height": viewport["height"]},
                    device_scale_factor=manifest["deviceScaleFactor"],
                    reduced_motion="reduce",
                    color_scheme="light",
                    forced_colors="none",
                )
                for theme in themes:
                    for surface in surfaces:
                        log(
                            f"surface {surface['id']} @ {viewport['id']} / {theme}"
                        )
                        case_batch = run_surface(
                                context=context,
                                surface=surface,
                                fixtures=fixtures,
                                theme=theme,
                                viewport=viewport,
                                self_check=self_check,
                                reference_url=f"{reference_server.url}/?surface={surface['id']}&theme={theme}",
                                subject_url=(
                                    f"{reference_server.url}/?surface={surface['id']}&theme={theme}"
                                    if self_check
                                    else f"{subject_server.url}/?surface={surface['id']}&theme={theme}"
                                ),
                                cases_dir=cases_dir,
                                out=out,
                                thresholds=thresholds,
                                padding=padding,
                                baseline_dir=baseline_dir,
                                baseline_mode=baseline_mode,
                                baseline_combos=baseline_combos,
                            )
                        cases.extend(case_batch)
                context.close()
            browser.close()
    finally:
        reference_server.close()
        subject_server.close()

    surfaces_by_id = {surface["id"]: surface for surface in manifest["surfaces"]}

    def surface_is_blocking(case: dict) -> bool:
        return surfaces_by_id.get(case["surface"], {}).get("gateStatus") == "blocking"

    blocking: list[dict] = []
    pending: list[dict] = []
    for case in cases:
        live_failed = case["status"] in {"fail", "error"}
        baseline_failed = case.get("baseline", {}).get("status") in {"fail", "missing"}
        failed = live_failed or (baseline_failed and surface_is_blocking(case))
        if not failed:
            continue
        if gate_mode == "all" or (gate_mode == "ready" and surface_is_blocking(case)):
            blocking.append(case)
        else:
            pending.append(case)

    summary = {
        "total": len(cases),
        "passed": sum(1 for case in cases if case["status"] == "pass"),
        "passedWithAntialiasing": sum(
            1 for case in cases if case["status"] == "pass-with-antialiasing"
        ),
        "failed": sum(1 for case in cases if case["status"] == "fail"),
        "baselineOnly": sum(1 for case in cases if case["status"] == "baseline-only"),
        "errors": sum(1 for case in cases if case["status"] == "error"),
        "baselineFailed": sum(
            1 for case in cases if case.get("baseline", {}).get("status") == "fail"
        ),
        "baselineMissing": sum(
            1 for case in cases if case.get("baseline", {}).get("status") == "missing"
        ),
        "baselineWritten": sum(
            1 for case in cases if case.get("baseline", {}).get("status") == "written"
        ),
        "blockingFailures": len(blocking),
        "pendingFailures": len(pending),
        "gateMode": gate_mode,
        "baselineMode": baseline_mode,
        "durationMs": int((time.time() - started) * 1000),
    }
    summary["gate"] = "FAIL" if blocking else "PASS"

    tokens = build_token_summary(token_report)
    surface_summaries = build_surface_summaries(surfaces, cases)

    report = {
        "schemaVersion": 1,
        "harness": manifest.get("harness", "extend-ui-pixel-parity"),
        "mode": "self-check" if args.self_check else "parity",
        "gateMode": gate_mode,
        "baselineMode": baseline_mode,
        "baselineDirectory": rel(baseline_dir) if baseline_dir is not None else None,
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "reference": {
            "repository": manifest.get("referenceUrl"),
            "commit": manifest.get("referenceCommit"),
        },
        "environment": {
            "browser": "chromium",
            "browserVersion": browser_version,
            "deviceScaleFactor": manifest["deviceScaleFactor"],
            "playwright": safe_package_version("playwright"),
            "python": sys.version.split()[0],
            "platform": platform.platform(),
            "animations": "disabled",
            "fonts": "document.fonts.ready before capture",
        },
        "thresholds": thresholds,
        "viewports": viewports,
        "themes": themes,
        "surfaces": surface_summaries,
        "summary": summary,
        "tokens": tokens,
        "cases": cases,
    }

    (out / "report.json").write_text(
        json.dumps(report, indent=2) + "\n", encoding="utf-8"
    )
    write_html_report(report, out / "report.html")
    write_summary_markdown(report, out / "SUMMARY.md")

    log("")
    log(
        f"pixel parity: {summary['gate']} — {summary['passed']} pass, "
        f"{summary['passedWithAntialiasing']} pass-with-AA, {summary['failed']} fail, "
        f"{summary['baselineOnly']} reference-only ({summary['durationMs']} ms)"
    )
    log(
        f"gate: {summary['gate']} mode={gate_mode} "
        f"baseline={baseline_mode} blocking={len(blocking)} pending={len(pending)}"
    )
    if baseline_mode != "off":
        log(
            f"baseline: {summary['baselineWritten']} written, "
            f"{summary['baselineFailed']} failed, {summary['baselineMissing']} missing"
        )
    for case in blocking[:20]:
        reasons = list(case.get("diff", {}).get("reasons", []))
        baseline = case.get("baseline")
        if baseline:
            if baseline["status"] == "missing":
                reasons.append(f"baseline missing: {baseline.get('baseline')}")
            elif baseline["status"] == "fail":
                regions = baseline.get("regions") or []
                detail = ", ".join(
                    f"{region['width']}x{region['height']}@{region['x']},{region['y']} "
                    f"({region['diffPixels']} px)"
                    for region in regions[:3]
                )
                reasons.append(f"baseline diff: {detail or 'size mismatch'}")
        reasons.extend(case.get("errors", [])[:2])
        log(
            f"  BLOCK {case['surface']}/{case['state']}/{case['viewport']}/{case['theme']}: "
            + ("; ".join(reasons[:3]) or case["status"])
        )
    if pending:
        log(f"  {len(pending)} pending failure(s) reported but not gating")
    log(f"report: {rel(out / 'report.json')} / {rel(out / 'report.html')}")
    if blocking:
        return 1
    if baseline_mode == "write" and summary["errors"]:
        return 1
    return 0


def safe_package_version(name: str) -> str | None:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return None


def build_token_summary(token_report: dict) -> dict:
    if not token_report:
        return {
            "available": False,
            "tokenLayerSummary": "missing docs/design/extend-tokens.json",
            "colorSummary": "missing docs/design/extend-tokens.json",
            "tokenLayerMismatches": [],
            "colorMismatches": [],
        }
    layer = token_report.get("tokenLayer", {})
    entries = layer.get("entries", [])
    mismatches = [
        entry
        for entry in entries
        if entry.get("match") is False and entry.get("office", {}).get("raw") is not None
    ]
    mapping = token_report.get("mapping", [])
    color_mismatches = []
    for entry in mapping:
        light = entry.get("lightDelta", {})
        dark = entry.get("darkDelta", {})
        if not light.get("comparable"):
            status = "missing in --office-* layer"
        elif not (light.get("matches") and dark.get("matches")):
            status = (
                f"delta light {light.get('maxChannelDelta')}/255, dark {dark.get('maxChannelDelta')}/255"
            )
        else:
            continue
        color_mismatches.append(
            {
                "token": entry["extendToken"],
                "officeToken": entry.get("officeToken"),
                "extendLight": (entry.get("extendLight") or {}).get("resolved"),
                "officeLight": (entry.get("officeLight") or {}).get("resolved"),
                "status": status,
            }
        )
    return {
        "available": True,
        "source": rel(ROOT / "docs/design/extend-tokens.json"),
        "referenceCommit": token_report.get("provenance", {}).get("commit"),
        "tokenLayerSummary": (
            f"{layer.get('summary', {}).get('matched', 0)} match / "
            f"{layer.get('summary', {}).get('mismatched', 0)} differ / "
            f"{layer.get('summary', {}).get('missing', 0)} missing "
            f"of {layer.get('summary', {}).get('total', 0)} compared tokens"
        ),
        "colorSummary": (
            f"{len(mapping) - len(color_mismatches)} of {len(mapping)} semantic colors match "
            f"the --office-* layer within 2/255"
        ),
        "tokenLayerMismatches": mismatches,
        "colorMismatches": color_mismatches,
    }


def build_surface_summaries(surfaces: list[dict], cases: list[dict]) -> list[dict]:
    summaries = []
    for surface in surfaces:
        surface_cases = [case for case in cases if case["surface"] == surface["id"]]
        summaries.append(
            {
                "id": surface["id"],
                "title": surface["title"],
                "subjectComponent": surface.get("subjectComponent"),
                "subjectStatus": surface.get("subjectStatus"),
                "pendingIssue": surface.get("pendingIssue"),
                "core": bool(surface.get("core")),
                "gateStatus": surface.get("gateStatus"),
                "gatePendingIssue": surface.get("gatePendingIssue"),
                "referenceComponent": surface.get("referenceComponent"),
                "notes": surface.get("notes"),
                "counts": {
                    "total": len(surface_cases),
                    "passed": sum(1 for case in surface_cases if case["status"] == "pass"),
                    "passedWithAntialiasing": sum(
                        1
                        for case in surface_cases
                        if case["status"] == "pass-with-antialiasing"
                    ),
                    "failed": sum(1 for case in surface_cases if case["status"] == "fail"),
                    "baselineOnly": sum(
                        1 for case in surface_cases if case["status"] == "baseline-only"
                    ),
                    "errors": sum(1 for case in surface_cases if case["status"] == "error"),
                },
            }
        )
    return summaries


def run_surface(
    *,
    context,
    surface: dict,
    fixtures: dict,
    theme: str,
    viewport: dict,
    reference_url: str,
    subject_url: str,
    self_check: bool,
    cases_dir: Path,
    out: Path,
    thresholds: dict,
    padding: int,
    baseline_dir: Path | None = None,
    baseline_mode: str = "off",
    baseline_combos: set | None = None,
) -> list[dict]:
    cases: list[dict] = []
    subject_status = "available" if self_check else surface.get("subjectStatus", "available")
    subject_selector = (
        surface.get("referenceTargetSelector") if self_check else surface.get("subjectTargetSelector")
    )
    reference_selector = surface.get("referenceTargetSelector")
    states = resolve_states(surface, fixtures)

    for state in states:
        case_dir = (
            cases_dir
            / surface["id"]
            / viewport["id"]
            / theme
            / state["id"]
        )
        case: dict = {
            "surface": surface["id"],
            "state": state["id"],
            "matrixState": state.get("matrixState", state["id"]),
            "theme": theme,
            "viewport": viewport["id"],
            "pendingIssue": surface.get("pendingIssue"),
            "status": "fail",
            "errors": [],
        }
        reference_page = context.new_page()
        subject_page = None
        try:
            reference_errors = prepare_page(reference_page, "reference")
            reference_page.goto(reference_url, wait_until="load")
            wait_ready(reference_page, surface["id"])
            reference_capture = capture_side(
                page=reference_page,
                out=out,
                state=state,
                target_selector=reference_selector,
                destination=case_dir / "reference.png",
                padding=padding,
                side_errors=reference_errors,
            )
            case["reference"] = reference_capture

            if subject_status != "available":
                case["status"] = "baseline-only"
                case["reason"] = (
                    f"no Vue subject yet: {surface.get('pendingIssue') or 'unscheduled'}"
                )
                cases.append(case)
                continue

            subject_page = context.new_page()
            subject_errors = prepare_page(subject_page, "subject")
            subject_page.goto(subject_url, wait_until="load")
            wait_ready(subject_page, surface["id"])
            subject_capture = capture_side(
                page=subject_page,
                out=out,
                state=state,
                target_selector=subject_selector,
                destination=case_dir / "subject.png",
                padding=padding,
                side_errors=subject_errors,
            )
            case["subject"] = subject_capture

            if (
                baseline_dir is not None
                and baseline_mode != "off"
                and (viewport["id"], theme) in (baseline_combos or set())
            ):
                baseline_file = baseline_dir / baseline_relpath(
                    surface["id"], state["id"], theme, viewport["id"]
                )
                if baseline_mode == "write":
                    case["baseline"] = write_baseline(
                        baseline_file, case_dir / "subject.png"
                    )
                else:
                    case["baseline"] = compare_baseline(
                        baseline_file,
                        case_dir / "subject.png",
                        case_dir,
                        out,
                        thresholds,
                    )

            geometry = compare_metrics(
                reference_capture["metrics"],
                subject_capture["metrics"],
                thresholds["geometryTolerancePx"],
            )
            case["geometry"] = geometry

            diff = compare_images(
                case_dir / "reference.png",
                case_dir / "subject.png",
                case_dir / "diff.png",
                case_dir / "heatmap.png",
                case_dir / "compare.png",
                thresholds,
            )
            prev_diff = diff.pop("status")
            diff_file = case_dir / "diff.png"
            heatmap_file = case_dir / "heatmap.png"
            for key in ("overlapDiffScreenshot", "overlapHeatmapScreenshot"):
                if diff.get(key):
                    diff[key] = artifact_rel(case_dir / diff[key], out)
            case["diff"] = {
                **diff,
                "screenshot": artifact_rel(diff_file, out) if diff_file.exists() else None,
            }
            if heatmap_file.exists():
                case["heatmap"] = artifact_rel(heatmap_file, out)
            compare_file = case_dir / "compare.png"
            if compare_file.exists():
                case["compare"] = artifact_rel(compare_file, out)

            failures = []
            if prev_diff == "fail":
                failures.append("pixel")
            if geometry["status"] == "fail":
                failures.append("geometry")
            if failures:
                case["status"] = "fail"
            elif prev_diff == "pass-with-antialiasing":
                case["status"] = "pass-with-antialiasing"
            else:
                case["status"] = "pass"
            case["errors"] = reference_capture["errors"] + subject_capture["errors"]
            if case["errors"]:
                case["status"] = "error"
        except Exception as error:  # noqa: BLE001 - report, don't abort the matrix
            case["status"] = "error"
            case["errors"] = [f"{type(error).__name__}: {error}"]
            traceback.print_exc()
        finally:
            if subject_page is not None:
                subject_page.close()
            reference_page.close()

        cases.append(case)
    return cases


def capture_side(
    *,
    page,
    out: Path,
    state: dict,
    target_selector: str | None,
    destination: Path,
    padding: int,
    side_errors: list[str],
) -> dict:
    cell = cell_locator(page, state["id"])
    cell.wait_for(state="attached", timeout=10_000)
    target = target_locator(page, cell, target_selector)
    interaction = state.get("interaction") or {"type": "none"}
    # Some targets only mount after the interaction (a tooltip bubble appears
    # on hover). Resolve the interaction from the cell in that case, then wait
    # for the measured target before capturing.
    if target.count() == 0 and not interaction.get("selector"):
        raise RuntimeError(f"no parity target for cell {state['id']}")
    apply_interaction(page, target, interaction, cell)
    try:
        target.first.wait_for(state="attached", timeout=5_000)
        captured = capture(page, cell, target, destination, padding)
    finally:
        release_interaction(page, interaction)
    captured["screenshot"] = artifact_rel(destination, out)
    captured["errors"] = side_errors
    return captured


def artifact_rel(path: Path, out: Path) -> str:
    try:
        return str(path.resolve().relative_to(out.resolve()))
    except ValueError:
        return str(path)


def resolve_states(surface: dict, fixtures: dict) -> list[dict]:
    """Expand `set-input-files` dotted fixture references before driving the UI."""
    resolved = []
    for state in surface["states"]:
        interaction = dict(state.get("interaction") or {"type": "none"})
        if interaction.get("type") == "set-input-files":
            node = fixtures
            for part in interaction["file"].split("."):
                node = node[part]
            interaction["fixture"] = node
        resolved.append({**state, "interaction": interaction})
    return resolved


if __name__ == "__main__":
    sys.exit(main())
