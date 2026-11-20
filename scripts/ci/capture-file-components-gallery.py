"""Issue #163 — deterministic FileUpload / FileThumbnail story screenshots.

The story page lives at `/#/file-components` in the verification demo. This
helper captures the fixed states the issue asks for at 390 / 768 / 1440 in
light and dark, and fails on console errors. It is a local evidence helper;
the pixel gate remains `pnpm ui:parity -- --only file-upload,file-thumbnail`.

    timeout 180s pnpm --filter demo dev -- --host 127.0.0.1 --port 4173
    python3 scripts/ci/capture-file-components-gallery.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

OUT = Path(os.environ.get("FILE_COMPONENTS_OUT", "output/parity/file-components"))
URL = os.environ.get("FILE_COMPONENTS_URL", "http://127.0.0.1:4173")
VIEWPORTS = [("390", 390, 844), ("768", 768, 1024), ("1440", 1440, 900)]
THEMES = ["light", "dark"]


def scenario_url(theme: str) -> str:
    return f"{URL}/?theme={theme}#/file-components"


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        for label, width, height in VIEWPORTS:
            for theme in THEMES:
                page = browser.new_page(
                    viewport={"width": width, "height": height},
                    device_scale_factor=2,
                )
                page.emulate_media(reduced_motion="reduce")
                page.on(
                    "console",
                    lambda message: errors.append(f"{label}/{theme}: {message.text}")
                    if message.type == "error"
                    else None,
                )
                page.on(
                    "pageerror",
                    lambda error: errors.append(f"{label}/{theme}: {error}"),
                )
                page.goto(scenario_url(theme), wait_until="networkidle")
                page.wait_for_selector('[data-testid="file-components-story"]')
                page.wait_for_timeout(250)

                prefix = OUT / f"{label}-{theme}"
                page.locator('[data-testid="file-components-story"]').screenshot(
                    path=str(prefix.with_name(f"{prefix.name}-story.png"))
                )
                for test_id, name in [
                    ("section-upload", "upload"),
                    ("section-thumbnail", "thumbnail"),
                    ("document-sidebar-seam", "sidebar-seam"),
                ]:
                    page.locator(f'[data-testid="{test_id}"]').screenshot(
                        path=str(prefix.with_name(f"{prefix.name}-{name}.png"))
                    )
                page.close()
        browser.close()

    if errors:
        print("console errors:", file=sys.stderr)
        for error in errors:
            print(f"  {error}", file=sys.stderr)
        return 1
    print(f"screenshots: {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
