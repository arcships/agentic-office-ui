"""Real-world fixture verification (V1).

Two modes:

    # oracle only (LibreOffice/pdfinfo page counts, no browser)
    uv run python tests/blackbox/real_files.py

    # oracle + local viewer smoke (needs a preview at LOCAL_DEMO_URL)
    LOCAL_DEMO_URL=http://127.0.0.1:4173 \
    uv run --with playwright python tests/blackbox/real_files.py --viewer \
        --output verification/evidence/2026-11-16/real-files/summary.json

Fixtures are listed in `test-data/real/manifest.json` with sha256, feature tags,
source repo and the oracle page count recorded at collection time.

Viewer mode uploads each fixture through the demo's file input and records the
public state, page/slide counts, console errors and load failures. Files whose
manifest has `"expect": "loads-with-repair-or-error"` are allowed to fail to
load (they are deliberately damaged test documents).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REAL = ROOT / "test-data" / "real"
MANIFEST = REAL / "manifest.json"
LOCAL_URL = os.environ.get("LOCAL_DEMO_URL", "http://127.0.0.1:4173").strip().rstrip("/")

ROUTES = {
    "docx": ("/#/docx-viewer", 'input[type="file"][accept*="docx"]'),
    "xlsx": ("/#/xlsx-viewer", 'input[type="file"][accept*="xlsx"]'),
    "pptx": ("/#/pptx-viewer", 'input[type="file"][accept*="pptx"]'),
    "pdf": ("/#/pdf-surface", 'input[type="file"][accept*="pdf"]'),
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def pdf_pages(pdf: Path, pdfinfo: str) -> int | None:
    result = subprocess.run([pdfinfo, str(pdf)], capture_output=True, text=True, timeout=30)
    for line in result.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":", 1)[1].strip())
    return None


def convert(soffice: str, profile: str, source: Path, outdir: Path) -> Path | None:
    subprocess.run(
        [soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(source)],
        capture_output=True,
        text=True,
        timeout=180,
    )
    pdf = outdir / f"{source.stem}.pdf"
    return pdf if pdf.exists() else None


def oracle(entries: list[dict]) -> tuple[dict, list[str]]:
    soffice, pdfinfo = shutil.which("soffice"), shutil.which("pdfinfo")
    results: dict[str, dict] = {}
    failures: list[str] = []
    if not soffice or not pdfinfo:
        return {"skipped": "soffice/pdfinfo not available"}, ["oracle skipped"]
    with tempfile.TemporaryDirectory(prefix="real-oracle-") as tmp:
        profile = str(Path(tmp) / "lo")
        for entry in entries:
            name, fmt = entry["file"], entry["format"]
            path = REAL / name
            if fmt == "pdf":
                pages = pdf_pages(path, pdfinfo)
            else:
                pdf = convert(soffice, profile, path, Path(tmp))
                pages = pdf_pages(pdf, pdfinfo) if pdf else None
            expected = entry.get("oracle_pages")
            ok = pages == expected
            results[name] = {"format": fmt, "pages": pages, "expected": expected, "ok": ok}
            if not ok:
                failures.append(f"{name}: oracle pages {pages} != {expected}")
    return results, failures


def viewer(entries: list[dict]) -> tuple[dict, list[str]]:
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:  # pragma: no cover - optional dependency
        return {"skipped": f"playwright unavailable: {exc}"}, []

    results: dict[str, dict] = {}
    failures: list[str] = []
    chromium = os.environ.get("REAL_FILES_CHROMIUM", "/usr/bin/chromium")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            executable_path=chromium if Path(chromium).exists() else None,
            args=["--no-sandbox", "--disable-gpu"],
        )
        for entry in entries:
            name, fmt = entry["file"], entry["format"]
            route, selector = ROUTES[fmt]
            page = browser.new_page(viewport={"width": 1440, "height": 900})
            errors: list[str] = []
            page.on("pageerror", lambda exc: errors.append(str(exc)))
            page.goto(LOCAL_URL + route, wait_until="networkidle")
            page.wait_for_timeout(1000)
            load_error = None
            try:
                page.locator(selector).first.set_input_files(str(REAL / name))
                page.wait_for_timeout(6000)
                state = page.evaluate(
                    """() => {
                      const viewer = document.querySelector('[data-testid="docx-viewer"],[data-testid="pptx-viewer"]');
                      const status = document.querySelector('[data-testid="page-status"]');
                      const slots = document.querySelectorAll('.pdf-surface__page-slot').length;
                      return {
                        viewerState: viewer?.getAttribute('data-state') ?? null,
                        status: status?.textContent?.trim() ?? null,
                        pageCount: document.querySelector('[data-docx-page-count]')?.getAttribute('data-docx-page-count') ?? null,
                        firstPageHeight: Math.round(document.querySelector('[data-docx-page-surface]')?.getBoundingClientRect().height ?? 0),
                        sheetTabs: document.querySelectorAll('[data-testid="xlsx-sheet-tab"]').length,
                        charts: document.querySelectorAll('[data-testid="xlsx-chart-item"]').length,
                        thumbnails: document.querySelectorAll('.pptx-thumbnail').length,
                        pdfSlots: slots,
                        bodyText: document.body.innerText.replace(/\\s+/g, ' ').slice(0, 160),
                      }
                    }"""
                )
                if fmt == "docx" and state["viewerState"] != "ready":
                    load_error = f"state={state['viewerState']}"
                if fmt == "xlsx" and (state["status"] or "").lower() not in {"ready", "idle"}:
                    load_error = f"status={state['status']}"
                if fmt == "pptx" and state["viewerState"] != "ready":
                    load_error = f"state={state['viewerState']}"
                if fmt == "pdf" and state["pdfSlots"] == 0:
                    load_error = "no pdf page slots"
            except Exception as exc:
                state = {}
                load_error = str(exc)
            expected_failure = entry.get("expect") == "loads-with-repair-or-error"
            result = {"state": state, "errors": errors[:5], "load_error": load_error, "expected_failure": expected_failure}
            results[name] = result
            if load_error and not expected_failure:
                failures.append(f"{name}: {load_error}")
            page.close()
        browser.close()
    return results, failures


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--viewer", action="store_true", help="also run the local viewer smoke")
    parser.add_argument("--output", help="write the JSON summary to this path")
    args = parser.parse_args()

    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    entries = manifest["fixtures"]
    failures: list[str] = []

    for entry in entries:
        path = REAL / entry["file"]
        if not path.exists():
            failures.append(f"{entry['file']}: missing")
        elif sha256(path) != entry["sha256"]:
            failures.append(f"{entry['file']}: sha256 mismatch")

    oracle_results, oracle_failures = oracle(entries)
    failures.extend(oracle_failures)
    viewer_results = {"skipped": "use --viewer"} if not args.viewer else None
    if args.viewer:
        viewer_results, viewer_failures = viewer(entries)
        failures.extend(viewer_failures)

    payload = {
        "suite": "real-files",
        "date": "2026-11-16",
        "manifest": str(MANIFEST.relative_to(ROOT)),
        "fixtures": len(entries),
        "oracle": oracle_results,
        "viewer": viewer_results,
        "failures": failures,
        "verdict": "FAIL" if failures else "PASS",
    }
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"fixtures": len(entries), "oracle_failures": len(oracle_failures), "failures": failures[:10], "verdict": payload["verdict"]}, ensure_ascii=False, indent=2))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
