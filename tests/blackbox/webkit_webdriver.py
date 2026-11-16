"""WebKitGTK smoke via the system WebKitWebDriver (Omarchy/Arch).

Playwright's bundled WebKit needs Ubuntu libraries; Arch/Omarchy ships
WebKitGTK + WebKitWebDriver instead. This script drives the system browser over
the W3C WebDriver protocol, so no Playwright download is required.

    # start the driver once (opens a MiniBrowser window on $DISPLAY)
    WebKitWebDriver --port=4444

    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    python tests/blackbox/webkit_webdriver.py \
        --output verification/evidence/2026-11-16/cross-browser/webkit.json

If no driver is listening the script prints SKIP and exits 0.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")
DRIVER = os.environ.get("WEBKIT_WEBDRIVER_URL", "http://127.0.0.1:4444").rstrip("/")
DOCX = ROOT / "apps" / "demo" / "public" / "samples" / "demo.docx"
XLSX = ROOT / "test-data" / "corpus" / "xlsx-charts.xlsx"
PDF = ROOT / "test-data" / "corpus" / "pdf-rotated-90.pdf"
PPTX = ROOT / "test-data" / "corpus" / "pptx-deck-basic.pptx"
ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"


class Driver:
    def __init__(self, base: str) -> None:
        self.base = base
        self.session: str | None = None

    def call(self, method: str, path: str, data: dict | None = None) -> dict:
        request = urllib.request.Request(
            self.base + path,
            method=method,
            data=json.dumps(data).encode() if data is not None else None,
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read())

    def start(self) -> dict | None:
        try:
            payload = self.call("POST", "/session", {"capabilities": {"alwaysMatch": {"browserName": "MiniBrowser"}}})
        except (urllib.error.URLError, TimeoutError):
            return None
        self.session = payload["value"]["sessionId"]
        return payload["value"]["capabilities"]

    def stop(self) -> None:
        if self.session:
            try:
                self.call("DELETE", f"/session/{self.session}")
            except Exception:  # noqa: BLE001
                pass

    def navigate(self, url: str) -> None:
        self.call("POST", f"/session/{self.session}/url", {"url": url})

    def find(self, selector: str) -> str | None:
        try:
            payload = self.call(
                "POST",
                f"/session/{self.session}/element",
                {"using": "css selector", "value": selector},
            )
        except urllib.error.HTTPError:
            return None
        return payload["value"].get(ELEMENT_KEY)

    def send_file(self, selector: str, path: Path) -> bool:
        element = self.find(selector)
        if not element:
            return False
        self.call("POST", f"/session/{self.session}/element/{element}/value", {"text": str(path)})
        return True

    def execute(self, script: str) -> object:
        return self.call("POST", f"/session/{self.session}/execute/sync", {"script": script, "args": []})["value"]

    def screenshot(self, path: Path) -> None:
        data = self.call("GET", f"/session/{self.session}/screenshot")["value"]
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(base64.b64decode(data))


def wait_for(driver: Driver, script: str, timeout: float = 20.0) -> object:
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = driver.execute(script)
        if last:
            return last
        time.sleep(0.5)
    return last


def run_case(driver: Driver, name: str, route: str, selector: str, path: Path, check: str, metrics: str, outdir: Path) -> dict:
    driver.navigate(LOCAL_URL + route)
    time.sleep(1.5)
    uploaded = driver.send_file(selector, path)
    wait_for(driver, check, 25)
    time.sleep(2)
    data = driver.execute(metrics)
    shot = outdir / f"webkit-{name}.png"
    try:
        driver.screenshot(shot)
        data["screenshot"] = str(shot)
    except Exception:  # noqa: BLE001
        data["screenshot"] = None
    data["uploaded"] = uploaded
    return data


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output")
    args = parser.parse_args()

    driver = Driver(DRIVER)
    capabilities = driver.start()
    if not capabilities:
        print(f"SKIP: no WebKitWebDriver at {DRIVER}")
        return 0

    outdir = Path(args.output).parent if args.output else Path(".")
    report: dict = {
        "suite": "cross-browser-webkit",
        "date": "2026-11-16",
        "driver": DRIVER,
        "browser": capabilities.get("browserVersion"),
        "cases": {},
        "failures": [],
    }
    try:
        report["cases"]["docx"] = run_case(
            driver, "docx", "/#/docx-viewer", 'input[type="file"][accept*="docx"]', DOCX,
            'return document.querySelector("[data-docx-page-count]") !== null',
            'return {state: document.querySelector("[data-testid=\\"docx-viewer\\"]")?.getAttribute("data-state"), pages: Number(document.querySelector("[data-docx-page-count]")?.getAttribute("data-docx-page-count") || 0), firstPageHeight: Math.round(document.querySelector("[data-docx-page-surface]")?.getBoundingClientRect().height || 0)}',
            outdir,
        )
        report["cases"]["xlsx"] = run_case(
            driver, "xlsx", "/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]', XLSX,
            'return document.querySelector("[data-testid=\\"page-status\\"]")?.textContent?.trim() === "ready"',
            'return {status: document.querySelector("[data-testid=\\"page-status\\"]")?.textContent?.trim(), canvases: document.querySelectorAll("canvas").length, charts: document.querySelectorAll("[data-testid=\\"xlsx-chart-item\\"]").length}',
            outdir,
        )
        report["cases"]["pdf"] = run_case(
            driver, "pdf", "/#/pdf-surface", 'input[type="file"][accept*="pdf"]', PDF,
            'return document.querySelectorAll(".pdf-surface__page-slot").length > 0',
            'return {slots: document.querySelectorAll(".pdf-surface__page-slot").length, images: document.querySelectorAll(".pdf-surface__page-img").length}',
            outdir,
        )
        report["cases"]["pptx"] = run_case(
            driver, "pptx", "/#/pptx-viewer", 'input[type="file"][accept*="pptx"]', PPTX,
            'return document.querySelector("[data-testid=\\"pptx-viewer\\"]")?.getAttribute("data-state") === "ready"',
            'return {state: document.querySelector("[data-testid=\\"pptx-viewer\\"]")?.getAttribute("data-state"), thumbs: document.querySelectorAll(".pptx-thumbnail").length}',
            outdir,
        )
    finally:
        driver.stop()

    for name, case in report["cases"].items():
        if not case.get("uploaded"):
            report["failures"].append(f"{name}: file input not found")
        if case.get("state") not in (None, "ready") and name in {"docx", "pptx"}:
            report["failures"].append(f"{name}: state {case.get('state')}")
        if name == "xlsx" and case.get("status") != "ready":
            report["failures"].append(f"{name}: status {case.get('status')}")
        if name == "pdf" and not case.get("slots"):
            report["failures"].append(f"{name}: no page slots")
    report["verdict"] = "FAIL" if report["failures"] else "PASS"

    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
