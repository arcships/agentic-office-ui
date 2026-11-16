"""WebKit scenario subset via WebKitWebDriver (system WebKitGTK).

WebDriver has no print-media emulation, so print stays Chromium/Firefox-only;
this covers DOCX zoom/dark, XLSX zoom, PPTX zoom and PDF page slots.

    WebKitWebDriver --port=4444
    python tests/blackbox/webkit_webdriver_scenarios.py --output ...
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from webkit_webdriver import Driver  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = "http://127.0.0.1:4173"
DOCX = ROOT / "apps" / "demo" / "public" / "samples" / "demo.docx"
XLSX = ROOT / "test-data" / "corpus" / "xlsx-charts.xlsx"
PDF = ROOT / "test-data" / "corpus" / "pdf-rotated-90.pdf"
PPTX = ROOT / "test-data" / "corpus" / "pptx-deck-basic.pptx"


def load(driver: Driver, route: str, selector: str, path: Path, wait_script: str | None = None) -> None:
    driver.navigate(LOCAL_URL + route)
    time.sleep(1.2)
    driver.send_file(selector, path)
    deadline = time.time() + 25
    if wait_script:
        while time.time() < deadline and not driver.execute(wait_script):
            time.sleep(0.5)
    time.sleep(2)


def docx(driver: Driver) -> dict:
    load(driver, "/#/docx-viewer", 'input[type="file"][accept*="docx"]', DOCX,
         'return document.querySelector("[data-docx-page-count]") !== null')
    result = driver.execute(
        """return {
          pages: Number(document.querySelector('[data-docx-page-count]')?.getAttribute('data-docx-page-count') || 0),
          height: Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0),
        }"""
    )
    for zoom in ("200", "50"):
        driver.execute(
            f"""const s = document.querySelector('[data-testid="docx-zoom-select"]');
               s.value = '{zoom}'; s.dispatchEvent(new Event('change', {{bubbles:true}})); return true;"""
        )
        time.sleep(1)
        result[f"zoom_{zoom}"] = driver.execute(
            "return Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height || 0)"
        )
    driver.execute(
        """const s = document.querySelector('[data-testid="docx-zoom-select"]');
           s.value = '100'; s.dispatchEvent(new Event('change', {bubbles:true}));
           document.querySelector('[data-testid="docx-theme-toggle"]')?.click(); return true;"""
    )
    time.sleep(0.8)
    result["dark"] = driver.execute(
        "return getComputedStyle(document.querySelector('[data-docx-page-surface]')).backgroundColor"
    )
    return result


def xlsx(driver: Driver) -> dict:
    load(driver, "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', XLSX,
         'return document.querySelector("[data-testid=\\"page-status\\"]")?.textContent?.trim() === "ready"')
    result = {}
    for level in ("200", "50"):
        driver.execute(
            f"""const s = document.querySelector('select[aria-label="Zoom level"]');
               s.value = '{level}'; s.dispatchEvent(new Event('change', {{bubbles:true}})); return true;"""
        )
        time.sleep(0.8)
        result[f"zoom_{level}"] = driver.execute(
            """const c = document.querySelector('canvas.xlsx-grid__body');
               return { value: document.querySelector('select[aria-label="Zoom level"]')?.value,
                        canvas: c ? { w: c.width, h: c.height } : null };"""
        )
    return result


def pptx(driver: Driver) -> dict:
    load(driver, "/#/pptx-viewer", 'input[type="file"][accept*="pptx"]', PPTX,
         'return document.querySelector("[data-testid=\\"pptx-viewer\\"]")?.getAttribute("data-state") === "ready"')
    before = driver.execute("return document.querySelectorAll('.pptx-viewer__zoom button')[1]?.textContent?.trim()")
    driver.execute("const b=document.querySelectorAll('.pptx-viewer__zoom button'); b[2]?.click(); b[2]?.click(); return true;")
    time.sleep(0.8)
    after = driver.execute("return document.querySelectorAll('.pptx-viewer__zoom button')[1]?.textContent?.trim()")
    return {"before": before, "after": after}


def pdf(driver: Driver) -> dict:
    load(driver, "/#/pdf-surface", 'input[type="file"][accept*="pdf"]', PDF,
         'return document.querySelectorAll(".pdf-surface__page-slot").length > 0')
    return driver.execute(
        "return { slots: document.querySelectorAll('.pdf-surface__page-slot').length, images: document.querySelectorAll('.pdf-surface__page-img').length }"
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()
    driver = Driver("http://127.0.0.1:4444")
    caps = driver.start()
    if not caps:
        print("SKIP: no WebKitWebDriver at 127.0.0.1:4444")
        return 0
    report: dict = {"suite": "webkit-scenarios", "browser": caps.get("browserVersion"), "cases": {}, "failures": []}
    try:
        for name, fn in {"docx": docx, "xlsx": xlsx, "pptx": pptx, "pdf": pdf}.items():
            try:
                report["cases"][name] = fn(driver)
            except Exception as exc:  # noqa: BLE001
                report["cases"][name] = {"error": f"{type(exc).__name__}: {exc}"}
                report["failures"].append(name)
    finally:
        driver.stop()
    report["print"] = "not supported by WebDriver (Chromium/Firefox only)"
    report["verdict"] = "FAIL" if report["failures"] else "PASS"
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
