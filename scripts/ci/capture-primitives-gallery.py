"""Issue #157 — deterministic primitive gallery screenshots.

The #155 pixel-parity harness is not present in this worktree, so this script
captures element-level evidence for Button/Input/Select/Dialog/Tooltip and the
whole gallery in light/dark. It deliberately fails on console errors.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

OUT = Path("/tmp/b08-157/output/parity")
URL = os.environ.get("PRIMITIVES_URL", "http://127.0.0.1:4173/#/primitives")


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1280, "height": 960}, device_scale_factor=1)
        page.on("console", lambda message: errors.append(message.text) if message.type == "error" else None)
        page.on("pageerror", lambda error: errors.append(str(error)))

        page.goto(URL, wait_until="networkidle")
        page.wait_for_selector('[data-testid="primitives-gallery"]')
        page.wait_for_timeout(300)

        page.screenshot(path=str(OUT / "primitives-gallery-light.png"), full_page=True)

        # Interaction smoke checks in a real browser (the component suite
        # covers the same contract in happy-dom).
        page.get_by_role("button", name="Open menu").click()
        page.wait_for_selector('[role="menu"]')
        page.get_by_role("menuitem", name="Copy").click()
        page.wait_for_timeout(120)
        assert "Copy" in page.locator('[data-testid="action-log"]').inner_text()

        page.locator('[data-testid="context-menu-target"]').click(button="right")
        page.wait_for_selector('[role="menu"]')
        page.keyboard.press("Escape")
        page.wait_for_timeout(120)

        page.locator('[data-testid="sheet-trigger"]').click()
        sheet = page.locator('[data-slot="sheet-content"]')
        sheet.wait_for()
        page.wait_for_timeout(200)
        page.keyboard.press("Escape")
        page.wait_for_timeout(150)

        page.get_by_role("tab", name="Properties").click()
        page.wait_for_timeout(150)
        assert page.get_by_role("tabpanel").inner_text().startswith("Title, author")

        toggle = page.get_by_role("button", name="Show grid")
        assert toggle.get_attribute("aria-pressed") == "true"
        toggle.click()
        page.wait_for_timeout(100)
        assert toggle.get_attribute("aria-pressed") == "false"
        toggle.click()
        page.wait_for_timeout(100)

        button = page.locator('[data-testid="button-default"]')
        button.screenshot(path=str(OUT / "button-default.png"))

        def padded_screenshot(locator, name: str, padding: int = 8) -> None:
            locator.scroll_into_view_if_needed()
            box = locator.bounding_box()
            assert box, f"{name}: no bounding box"
            page.screenshot(
                path=str(OUT / name),
                clip={
                    "x": max(box["x"] - padding, 0),
                    "y": max(box["y"] - padding, 0),
                    "width": box["width"] + padding * 2,
                    "height": box["height"] + padding * 2,
                },
            )

        button.hover()
        page.wait_for_timeout(120)
        padded_screenshot(button, "button-hover.png")
        page.mouse.move(0, 0)
        page.wait_for_timeout(120)
        button.focus()
        page.wait_for_timeout(120)
        padded_screenshot(button, "button-focus.png")
        page.locator('[data-testid="button-disabled"]').screenshot(
            path=str(OUT / "button-disabled.png")
        )
        page.locator('[data-testid="button-loading"]').screenshot(
            path=str(OUT / "button-loading.png")
        )

        page.locator('[data-testid="input-default"]').screenshot(
            path=str(OUT / "input-default.png")
        )
        page.locator("#gallery-select").screenshot(
            path=str(OUT / "select-trigger.png")
        )
        page.locator("#gallery-select").click()
        page.wait_for_selector('[role="listbox"]')
        page.wait_for_timeout(150)
        page.locator('[role="listbox"]').screenshot(path=str(OUT / "select-content.png"))
        page.keyboard.press("Escape")
        page.wait_for_timeout(120)

        page.locator('[data-testid="dialog-trigger"]').click()
        dialog = page.locator('[data-slot="dialog-content"]')
        dialog.wait_for()
        page.wait_for_timeout(200)
        dialog.screenshot(path=str(OUT / "dialog-light.png"))
        page.keyboard.press("Escape")
        page.wait_for_timeout(150)

        tooltip_trigger = page.locator('[data-testid="tooltip-trigger"]')
        tooltip_trigger.hover()
        tooltip = page.locator('[data-slot="tooltip-content"]')
        tooltip.wait_for()
        page.wait_for_timeout(150)
        tooltip.screenshot(path=str(OUT / "tooltip.png"))

        # Dark theme through the gallery control, then re-capture the gallery.
        page.mouse.move(0, 0)
        page.wait_for_timeout(250)
        page.locator('[data-testid="gallery-theme"]').select_option("dark")
        page.wait_for_timeout(250)
        page.screenshot(path=str(OUT / "primitives-gallery-dark.png"), full_page=True)
        page.locator('[data-testid="button-default"]').screenshot(
            path=str(OUT / "button-dark.png")
        )
        page.mouse.move(0, 0)
        page.locator('[data-testid="dialog-trigger"]').click()
        dialog.wait_for()
        page.wait_for_timeout(200)
        dialog.screenshot(path=str(OUT / "dialog-dark.png"))
        page.keyboard.press("Escape")
        page.wait_for_timeout(150)

        page.locator("text=Show success").click()
        toast = page.locator('[data-slot="toast"]').first
        toast.wait_for()
        page.wait_for_timeout(250)
        toast.screenshot(path=str(OUT / "toast-success.png"))

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
