from __future__ import annotations

from pathlib import Path

from playwright.sync_api import Error as PlaywrightError


CHROMIUM_EXECUTABLES = (
    Path("/usr/bin/chromium"),
    Path("/usr/bin/chromium-browser"),
    Path("/usr/bin/google-chrome"),
    Path("/usr/bin/google-chrome-stable"),
    Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
)


def launch_browser(browser_type, *, headless: bool = True, args: list[str] | None = None):
    """Launch the bundled browser, falling back to an installed Chromium.

    CI installs the pinned Playwright bundle, while developer machines often
    only have a system Chromium. The fallback keeps the black-box scripts
    runnable without changing what CI exercises.
    """
    launch_options = {"headless": headless, "args": list(args or [])}
    failures: list[str] = []
    try:
        return browser_type.launch(**launch_options)
    except PlaywrightError as error:
        failures.append(f"bundled: {error}")
    if browser_type.name == "chromium":
        for candidate in CHROMIUM_EXECUTABLES:
            if not candidate.exists():
                continue
            try:
                return browser_type.launch(
                    **launch_options,
                    executable_path=str(candidate),
                )
            except PlaywrightError as error:
                failures.append(f"{candidate}: {error}")
    raise RuntimeError("no browser could launch: " + " | ".join(failures))
