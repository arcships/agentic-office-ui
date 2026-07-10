from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import Page


DEFAULT_ALLOWLIST = Path(__file__).with_name("console_allowlist.json")


class BrowserEvidence:
    """Collect and enforce browser failures through one public black-box policy."""

    def __init__(self, page: Page, allowlist_path: Path = DEFAULT_ALLOWLIST):
        self.allowlist = json.loads(allowlist_path.read_text(encoding="utf-8"))
        self.events: dict[str, list[dict[str, object]]] = {
            key: []
            for key in ("console", "pageErrors", "requestFailures", "responses", "downloads")
        }
        page.on(
            "console",
            lambda item: self.events["console"].append(
                {"type": item.type, "text": item.text}
            ),
        )
        page.on(
            "pageerror",
            lambda item: self.events["pageErrors"].append({"message": str(item)}),
        )
        page.on(
            "requestfailed",
            lambda item: self.events["requestFailures"].append(
                {"url": item.url, "failure": item.failure or ""}
            ),
        )
        page.on(
            "response",
            lambda item: self.events["responses"].append(
                {
                    "url": item.url,
                    "status": item.status,
                    "contentType": item.headers.get("content-type", ""),
                }
            ),
        )
        page.on(
            "download",
            lambda item: self.events["downloads"].append(
                {"url": item.url, "suggestedFilename": item.suggested_filename}
            ),
        )

    def _allowed(self, category: str, item: dict[str, object]) -> bool:
        rendered = json.dumps(item, ensure_ascii=False, sort_keys=True)
        return any(pattern in rendered for pattern in self.allowlist.get(category, []))

    def violations(self) -> dict[str, list[dict[str, object]]]:
        candidates = {
            "console": [
                item
                for item in self.events["console"]
                if item["type"] in {"warning", "error"}
            ],
            "pageErrors": list(self.events["pageErrors"]),
            "requestFailures": list(self.events["requestFailures"]),
            "responses": [
                item for item in self.events["responses"] if int(item["status"]) >= 400
            ],
        }
        return {
            category: [
                item for item in items if not self._allowed(category, item)
            ]
            for category, items in candidates.items()
        }

    def assert_clean(self) -> None:
        violations = self.violations()
        if any(violations.values()):
            raise AssertionError(
                "browser evidence policy violations: "
                + json.dumps(violations, ensure_ascii=False, sort_keys=True)
            )

    def save(self, directory: Path) -> None:
        directory.mkdir(parents=True, exist_ok=True)
        (directory / "console.json").write_text(
            json.dumps(self.events["console"], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (directory / "page-errors.json").write_text(
            json.dumps(self.events["pageErrors"], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (directory / "network.json").write_text(
            json.dumps(
                {
                    "requestFailures": self.events["requestFailures"],
                    "responses": self.events["responses"],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        (directory / "downloads.json").write_text(
            json.dumps(self.events["downloads"], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (directory / "violations.json").write_text(
            json.dumps(self.violations(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
