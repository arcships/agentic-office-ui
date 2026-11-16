"""Pixel baseline vs LibreOffice/poppler (V4).

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright --with pillow python tests/blackbox/pixel_baseline.py \
        --output verification/evidence/2026-11-16/pixel/summary.json

For each case:
1. render the same file with LibreOffice → PDF → `pdftoppm` (96 dpi);
2. screenshot the corresponding page/slide in the local viewer;
3. resize both to a common width, compare grayscale RMSE and the ratio of
   pixels differing by more than 32 levels;
4. save `local.png`, `oracle.png`, `diff.png` next to the summary.

Text rasterisation differs between engines, so thresholds are intentionally
loose; the value is in tracking drift, not absolute equality.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
THRESHOLDS = {"docx": 0.30, "pptx": 0.30, "pdf": 0.20}


def render_pdf_page(soffice: str, pdftoppm: str, source: Path, outdir: Path, page: int = 1) -> Path | None:
    profile = outdir / "lo"
    subprocess.run(
        [soffice, f"-env:UserInstallation=file://{profile.resolve()}", "--headless", "--convert-to", "pdf", "--outdir", str(outdir.resolve()), str(source)],
        capture_output=True,
        text=True,
        timeout=180,
    )
    pdf = outdir / f"{source.stem}.pdf"
    if not pdf.exists():
        return None
    prefix = outdir / f"{source.stem}-page"
    subprocess.run(
        [pdftoppm, "-png", "-r", "96", "-f", str(page), "-l", str(page), str(pdf), str(prefix)],
        capture_output=True,
        text=True,
        timeout=120,
    )
    pngs = sorted(outdir.glob(f"{source.stem}-page*.png"))
    return pngs[0] if pngs else None


def compare(local: Path, oracle: Path, outdir: Path, name: str) -> dict:
    a = Image.open(local).convert("L")
    b = Image.open(oracle).convert("L")
    width = 800
    a = a.resize((width, max(1, round(a.height * width / a.width))))
    b = b.resize((width, max(1, round(b.height * width / b.width))))
    height = min(a.height, b.height)
    a = a.crop((0, 0, width, height))
    b = b.crop((0, 0, width, height))
    diff = ImageChops.difference(a, b)
    hist = diff.histogram()
    total = width * height
    mean = sum(i * count for i, count in enumerate(hist)) / total
    ratio = sum(count for i, count in enumerate(hist) if i > 32) / total
    ssim = None
    try:
        import numpy as np

        x = np.asarray(a, dtype=float)
        y = np.asarray(b, dtype=float)
        mu_x, mu_y = x.mean(), y.mean()
        var_x, var_y = x.var(), y.var()
        cov = ((x - mu_x) * (y - mu_y)).mean()
        c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
        ssim = round(float(((2 * mu_x * mu_y + c1) * (2 * cov + c2)) / ((mu_x**2 + mu_y**2 + c1) * (var_x + var_y + c2))), 4)
    except Exception:  # noqa: BLE001
        ssim = None
    outdir.mkdir(parents=True, exist_ok=True)
    a.save(outdir / f"{name}-local.png")
    b.save(outdir / f"{name}-oracle.png")
    diff.save(outdir / f"{name}-diff.png")
    return {"rmse": round(mean, 2), "diff_ratio": round(ratio, 4), "ssim": ssim, "height": height}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()
    from playwright.sync_api import sync_playwright

    soffice, pdftoppm = shutil.which("soffice"), shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        print("SKIP: soffice/pdftoppm not available")
        return 0

    outdir = Path(args.output).parent if args.output else Path(tempfile.mkdtemp(prefix="pixel-"))
    outdir.mkdir(parents=True, exist_ok=True)
    cases = {
        "docx": {
            "file": ROOT / "apps" / "demo" / "public" / "samples" / "demo.docx",
            "route": "/#/docx-viewer",
            "upload": 'input[type="file"][accept*="docx"]',
            "shot": "[data-docx-page-surface]",
            "wait": "[data-docx-page-count]",
            "pages": [1, 2],
        },
        "pptx": {
            "file": ROOT / "test-data" / "corpus" / "pptx-deck-basic.pptx",
            "route": "/#/pptx-viewer",
            "upload": 'input[type="file"][accept*="pptx"]',
            "shot": ".pptx-viewer__stage",
            "wait": '[data-testid="pptx-viewer"][data-state="ready"]',
            "pages": [1, 2],
        },
        "pdf": {
            "file": ROOT / "test-data" / "corpus" / "pdf-rotated-90.pdf",
            "route": "/#/pdf-surface",
            "upload": 'input[type="file"][accept*="pdf"]',
            "shot": ".pdf-surface__page-img",
            "wait": ".pdf-surface__page-img",
            "pages": [1, 2],
        },
    }

    report: dict = {"suite": "pixel-baseline", "date": "2026-11-16", "cases": {}, "thresholds": THRESHOLDS, "failures": []}
    with sync_playwright() as pw:
        chromium = "/usr/bin/chromium"
        browser = pw.chromium.launch(
            executable_path=chromium if Path(chromium).exists() else None,
            args=["--no-sandbox", "--disable-gpu"],
        )
        for name, case in cases.items():
            page = browser.new_page(viewport={"width": 1440, "height": 900})
            page.goto(LOCAL_URL + case["route"], wait_until="networkidle")
            page.wait_for_timeout(800)
            page.locator(case["upload"]).first.set_input_files(str(case["file"]))
            page.wait_for_selector(case["wait"], timeout=30000)
            page.wait_for_timeout(5000)
            for page_no in case["pages"]:
                key = f"{name}-{page_no}"
                oracle = render_pdf_page(soffice, pdftoppm, case["file"], outdir / key, page=page_no)
                if not oracle:
                    report["cases"][key] = {"error": "oracle render failed"}
                    report["failures"].append(f"{key}: oracle render failed")
                    continue
                if name == "docx" and page_no > 1:
                    page.fill('[data-testid="docx-page-current"]', str(page_no))
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(1500)
                if name == "pptx" and page_no > 1:
                    for _ in range(page_no - 1):
                        page.click('[aria-label="下一页"]')
                        page.wait_for_timeout(1200)
                if name in {"docx", "pdf"}:
                    target = page.locator(case["shot"]).nth(page_no - 1)
                    target.scroll_into_view_if_needed()
                    page.wait_for_timeout(800)
                else:
                    target = page.locator(case["shot"]).first
                local = outdir / key / "local.png"
                target.screenshot(path=str(local))
                metrics = compare(local, oracle, outdir / key, key)
                metrics["threshold"] = THRESHOLDS[name]
                metrics["ok"] = metrics["diff_ratio"] <= THRESHOLDS[name]
                if not metrics["ok"]:
                    report["failures"].append(f"{key}: diff_ratio {metrics['diff_ratio']} > {THRESHOLDS[name]}")
                report["cases"][key] = metrics
            page.close()
        browser.close()
    report["verdict"] = "FAIL" if report["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"cases": report["cases"], "failures": report["failures"], "verdict": report["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
