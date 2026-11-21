#!/usr/bin/env python3
"""axe-core + keyboard smoke for the core Extend-aligned components (#165).

Serves the built parity comparison app (`tools/vue-reference/dist`), loads
axe-core into the page and runs WCAG A/AA rules against each configured
surface. Keyboard checks drive real Tab/Arrow/Escape key presses so focus
reachability and focus-ring regressions fail the same gate.

Usage (normally invoked by scripts/ci/a11y-check.mjs):

    python3 scripts/ci/a11y-check.py \
        --dir tools/vue-reference/dist \
        --axe node_modules/axe-core/axe.min.js \
        --config tests/a11y/config.json \
        --allowlist tests/a11y/allowlist.json \
        --out output/a11y

Exit codes:
  0  no unlisted violations and every keyboard check passed
  1  a new axe violation, a failed keyboard check, or a harness error
  2  configuration problem (missing build, missing axe-core)
"""

from __future__ import annotations

import argparse
import functools
import html
import http.server
import json
import socketserver
import sys
import threading
import time
import traceback
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[2]

AXE_RUNNER = """
async ({ selector, tags, disabled }) => {
  const root = document.querySelector(selector);
  if (!root) throw new Error(`no root for ${selector}`);
  const options = {
    runOnly: { type: 'tag', values: tags },
    resultTypes: ['violations'],
  };
  if (disabled.length) options.rules = Object.fromEntries(disabled.map((id) => [id, { enabled: false }]));
  const results = await window.axe.run(root, options);
  return results.violations.map((violation) => ({
    id: violation.id,
    impact: violation.impact,
    help: violation.help,
    helpUrl: violation.helpUrl,
    nodeCount: violation.nodes.length,
    nodes: violation.nodes.slice(0, 5).map((node) => ({
      target: node.target,
      html: node.html.slice(0, 240),
      failureSummary: node.failureSummary,
    })),
  }));
}
"""

ACCESSIBLE_NAME = """
() => {
  const element = document.activeElement;
  if (!element || element === document.body) return { name: '', reason: 'nothing focused' };
  const normalize = (value) => (value || '').replace(/\\s+/g, ' ').trim();
  const labelledby = element.getAttribute('aria-labelledby');
  if (labelledby) {
    const text = labelledby
      .split(/\\s+/)
      .map((id) => document.getElementById(id)?.textContent || '')
      .join(' ');
    if (normalize(text)) return { name: normalize(text), reason: 'aria-labelledby' };
  }
  const ariaLabel = element.getAttribute('aria-label');
  if (normalize(ariaLabel)) return { name: normalize(ariaLabel), reason: 'aria-label' };
  if (element.labels && element.labels.length) {
    const text = Array.from(element.labels).map((label) => label.textContent).join(' ');
    if (normalize(text)) return { name: normalize(text), reason: 'label' };
  }
  for (const attribute of ['alt', 'title']) {
    const value = element.getAttribute(attribute);
    if (normalize(value)) return { name: normalize(value), reason: attribute };
  }
  const text = normalize(element.innerText || element.textContent);
  if (text) return { name: text.slice(0, 120), reason: 'text content' };
  const value = element.value;
  if (normalize(value)) return { name: normalize(value), reason: 'value' };
  return { name: '', reason: 'no accessible name' };
}
"""

FOCUS_RING = """
() => {
  // The ring is often painted by a wrapper (`:has(:focus-visible)` on the
  // input control), not by the focused element itself, so walk up to the
  // parity cell and report the first level that paints something.
  let element = document.activeElement;
  if (!element || element === document.body) return { visible: false, reason: 'nothing focused' };
  const root = element.closest('[data-parity-cell]') || document.body;
  let current = element;
  for (let depth = 0; current && depth < 5; depth += 1) {
    const style = getComputedStyle(current);
    const outline = style.outlineStyle !== 'none' && parseFloat(style.outlineWidth) > 0;
    const ring = style.boxShadow && style.boxShadow !== 'none';
    if (outline || ring) {
      return {
        visible: true,
        reason: `${current.tagName.toLowerCase()}${current.className ? '.' + String(current.className).split(' ')[0] : ''}: outline ${style.outlineWidth} ${style.outlineStyle}; box-shadow ${style.boxShadow}`,
        tag: element.tagName.toLowerCase(),
        role: element.getAttribute('role'),
      };
    }
    if (current === root) break;
    current = current.parentElement;
  }
  return {
    visible: false,
    reason: 'no outline or box-shadow on the control or its ancestors',
    tag: element.tagName.toLowerCase(),
    role: element.getAttribute('role'),
  };
}
"""

ACTIVE_INSIDE = """
(selector) => {
  const element = document.activeElement;
  if (!element) return false;
  return Boolean(element.closest(selector));
}
"""

SELECT_OPEN = """
(selector) => {
  const cell = document.querySelector(selector);
  if (!cell) return { opened: false, reason: 'no cell' };
  const listbox = cell.querySelector('[role="listbox"]')
    || document.querySelector('[data-parity-surface] [role="listbox"]')
    || document.querySelector('[role="listbox"]');
  return { opened: Boolean(listbox && listbox.offsetParent !== null) };
}
"""


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


def parse_csv(value: str | None) -> list[str] | None:
    if not value:
        return None
    return [item.strip() for item in value.split(",") if item.strip()]


def wait_ready(page, surface: str) -> None:
    page.wait_for_selector(
        f"[data-parity-surface='{surface}'] [data-parity-cell]",
        state="attached",
        timeout=30_000,
    )
    page.wait_for_function("document.fonts.status === 'loaded'", timeout=30_000)
    page.evaluate("() => document.fonts.ready")
    page.wait_for_timeout(120)


def cell_selector(surface: str, state: str) -> str:
    return f"[data-parity-surface='{surface}'] [data-parity-cell='{state}']"


def run_keyboard_checks(page, surface: str, state: str, checks: list[str]) -> list[dict]:
    results = []
    selector = cell_selector(surface, state)
    needs_focus = any(
        check in checks for check in ("focusable", "focus-indicator", "accessible-name")
    )
    reached = False
    if needs_focus:
        page.evaluate("() => document.body.focus()")
        for _ in range(30):
            page.keyboard.press("Tab")
            if page.evaluate(ACTIVE_INSIDE, selector):
                reached = True
                break
    for check in checks:
        entry = {"check": check, "status": "pass", "detail": None}
        try:
            if check == "focusable":
                if not reached:
                    entry["status"] = "fail"
                    entry["detail"] = "Tab never reached a control inside the surface"
                else:
                    entry["detail"] = "Tab reached a control inside the surface"
            elif check == "focus-indicator":
                ring = page.evaluate(FOCUS_RING)
                entry["detail"] = ring["reason"]
                if not reached:
                    entry["status"] = "fail"
                    entry["detail"] = "no control inside the surface could be focused"
                elif not ring["visible"]:
                    entry["status"] = "fail"
                    entry["detail"] = f"no visible focus indicator ({ring['reason']})"
            elif check == "accessible-name":
                name = page.evaluate(ACCESSIBLE_NAME)
                entry["detail"] = f"{name['reason']}: {name['name']!r}"
                if not reached:
                    entry["status"] = "fail"
                    entry["detail"] = "no control inside the surface could be focused"
                elif not name["name"]:
                    entry["status"] = "fail"
                    entry["detail"] = "focused control has no accessible name"
            elif check == "select-opens-on-arrow":
                trigger = page.locator(f"{selector} [role='combobox']").first
                trigger.focus()
                page.keyboard.press("ArrowDown")
                page.wait_for_timeout(150)
                opened = page.evaluate(SELECT_OPEN, selector)
                if not opened.get("opened"):
                    entry["status"] = "fail"
                    entry["detail"] = "ArrowDown did not open a listbox"
                else:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(150)
                    closed = page.evaluate(SELECT_OPEN, selector)
                    if closed.get("opened"):
                        entry["status"] = "fail"
                        entry["detail"] = "Escape did not close the listbox"
                    else:
                        entry["detail"] = "ArrowDown opened and Escape closed the listbox"
            else:
                entry["status"] = "skip"
                entry["detail"] = f"unknown check {check}"
        except Exception as error:  # noqa: BLE001 - report the check, keep going
            entry["status"] = "fail"
            entry["detail"] = f"{type(error).__name__}: {error}"
        results.append(entry)
    return results


def is_allowlisted(entry: dict, violation: dict, surface: str) -> bool:
    for rule in entry.get("rules", [entry]):
        if rule.get("surface") not in (surface, "*"):
            continue
        if rule.get("rule") != violation["id"]:
            continue
        return True
    return False


def is_keyboard_allowlisted(entries: list[dict], surface: str, check: str) -> bool:
    for rule in entries:
        if rule.get("surface") in (surface, "*") and rule.get("check") == check:
            return True
    return False


def surface_report(
    *,
    context,
    surface: dict,
    config: dict,
    base_url: str,
    theme: str,
    viewport: dict,
    axe_path: Path,
) -> dict:
    page = context.new_page()
    errors: list[str] = []
    page.on("pageerror", lambda error: errors.append(f"pageerror: {error}"))
    report = {
        "surface": surface["id"],
        "theme": theme,
        "viewport": viewport["id"],
        "axe": [],
        "keyboard": [],
        "errors": errors,
    }
    try:
        page.goto(f"{base_url}/?surface={surface['id']}&theme={theme}", wait_until="load")
        wait_ready(page, surface["id"])
        page.add_script_tag(path=str(axe_path))
        axe_config = config.get("axe", {})
        violations = page.evaluate(
            AXE_RUNNER,
            {
                "selector": f"[data-parity-surface='{surface['id']}']",
                "tags": axe_config.get("tags", ["wcag2a", "wcag2aa"]),
                "disabled": axe_config.get("disabledRules", []),
            },
        )
        report["axe"] = violations
        report["keyboard"] = run_keyboard_checks(
            page, surface["id"], surface["keyboardState"], surface.get("checks", [])
        )
    except Exception as error:  # noqa: BLE001 - one surface must not abort the run
        report["errors"].append(f"{type(error).__name__}: {error}")
        traceback.print_exc()
    finally:
        page.close()
    return report


def write_summary(report: dict, destination: Path) -> None:
    lines = [
        "# Core component accessibility report",
        "",
        f"- generated: `{report['generatedAt']}`",
        f"- chromium: `{report['environment']['browserVersion']}`",
        f"- axe-core: `{report['environment']['axeVersion']}`",
        f"- result: `{report['summary']['result']}`",
        f"- violations: {report['summary']['violations']} "
        f"({report['summary']['allowlisted']} allowlisted, {report['summary']['new']} new)",
        f"- keyboard checks: {report['summary']['keyboardFailed']} failed",
        "",
        "## Surfaces",
        "",
        "| Surface | axe | new violations | keyboard |",
        "|---|---|---|---|",
    ]
    for surface in report["surfaces"]:
        failed_checks = [item for item in surface["keyboard"] if item["status"] == "fail"]
        lines.append(
            f"| {surface['surface']} | {len(surface['axe'])} | "
            f"{sum(1 for item in surface['axe'] if item['status'] == 'new')} | "
            f"{'FAIL ' + ', '.join(item['check'] for item in failed_checks) if failed_checks else 'pass'} |"
        )
    problems = report["summary"]["failures"]
    lines.extend(["", "## Failures", ""])
    if not problems:
        lines.append("None.")
    for problem in problems:
        lines.append(f"- {problem}")
    destination.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_html_report(report: dict, destination: Path) -> None:
    rows = []
    for surface in report["surfaces"]:
        for violation in surface["axe"]:
            rows.append(
                "<tr>"
                f"<td>{html.escape(surface['surface'])}</td>"
                f"<td>{html.escape(violation['id'])}</td>"
                f"<td>{html.escape(str(violation.get('impact')))}</td>"
                f"<td>{html.escape(violation['status'])}</td>"
                f"<td>{violation.get('nodeCount', 0)}</td>"
                f"<td><code>{html.escape(str(violation['nodes'][0]['target']))}</code></td>"
                "</tr>"
            )
        for check in surface["keyboard"]:
            rows.append(
                "<tr>"
                f"<td>{html.escape(surface['surface'])}</td>"
                f"<td>keyboard:{html.escape(check['check'])}</td>"
                f"<td>—</td>"
                f"<td>{html.escape(check['status'])}</td>"
                f"<td>—</td>"
                f"<td>{html.escape(str(check.get('detail')))}</td>"
                "</tr>"
            )
    document = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Core component accessibility</title>
<style>
  body {{ font-family: ui-sans-serif, system-ui, sans-serif; margin: 24px; color: #111; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th, td {{ border: 1px solid #ddd; padding: 6px; vertical-align: top; text-align: left; }}
  th {{ background: #f5f5f5; }}
  code {{ background: #f3f3f3; padding: 1px 4px; }}
</style>
</head>
<body>
<h1>Core component accessibility (axe-core + keyboard smoke)</h1>
<p><strong>{report['summary']['result']}</strong> · {report['summary']['new']} new violation(s) ·
{report['summary']['keyboardFailed']} keyboard failure(s) · {report['generatedAt']}</p>
<table>
  <thead><tr><th>Surface</th><th>Rule</th><th>Impact</th><th>Status</th><th>Nodes</th><th>Target / detail</th></tr></thead>
  <tbody>{''.join(rows)}</tbody>
</table>
</body>
</html>
"""
    destination.write_text(document, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dir", default=str(ROOT / "tools/vue-reference/dist"))
    parser.add_argument("--axe", default=str(ROOT / "node_modules/axe-core/axe.min.js"))
    parser.add_argument("--config", default=str(ROOT / "tests/a11y/config.json"))
    parser.add_argument("--allowlist", default=str(ROOT / "tests/a11y/allowlist.json"))
    parser.add_argument("--out", default=str(ROOT / "output/a11y"))
    parser.add_argument("--only", default=None, help="comma separated surface ids")
    parser.add_argument("--theme", default=None, help="override the configured themes (comma separated)")
    parser.add_argument("--report-only", action="store_true", help="always exit 0")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    build_dir = Path(args.dir)
    if not (build_dir / "index.html").exists():
        print(
            f"a11y: {build_dir}/index.html is missing; build the comparison app with\n"
            "  pnpm --filter @aou/vue-reference build",
            file=sys.stderr,
        )
        return 2
    axe_path = Path(args.axe)
    if not axe_path.exists():
        print(f"a11y: axe-core is missing at {axe_path}; run pnpm install", file=sys.stderr)
        return 2

    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    allowlist = json.loads(Path(args.allowlist).read_text(encoding="utf-8"))
    allow_entries = [entry for entry in allowlist.get("entries", []) if entry.get("surface")]

    only = parse_csv(args.only)
    surfaces = [
        surface
        for surface in config["surfaces"]
        if not only or surface["id"] in only
    ]
    if not surfaces:
        print("a11y: no surfaces selected", file=sys.stderr)
        return 2

    themes = parse_csv(args.theme) or config.get("themes") or [config.get("theme", "light")]
    viewports = config.get("viewports", ["1440x900"])
    viewport_specs = [
        {
            "id": viewport,
            "width": int(viewport.split("x")[0]),
            "height": int(viewport.split("x")[1]),
        }
        for viewport in viewports
    ]

    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    server = Server(build_dir.resolve())
    started = time.time()
    reports: list[dict] = []
    browser_version = "unknown"
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            browser_version = browser.version
            for viewport in viewport_specs:
                context = browser.new_context(
                    viewport={"width": viewport["width"], "height": viewport["height"]},
                    reduced_motion="reduce",
                    color_scheme="light",
                )
                for theme in themes:
                    for surface in surfaces:
                        print(
                            f"a11y: {surface['id']} @ {viewport['id']} / {theme}",
                            flush=True,
                        )
                        reports.append(
                            surface_report(
                                context=context,
                                surface=surface,
                                config=config,
                                base_url=server.url,
                                theme=theme,
                                viewport=viewport,
                                axe_path=axe_path,
                            )
                        )
                context.close()
            browser.close()
    finally:
        server.close()

    failures: list[str] = []
    total_violations = 0
    allowlisted = 0
    new = 0
    keyboard_failed = 0
    keyboard_allowlisted = 0
    used_allow_entries = set()
    for surface in reports:
        for violation in surface["axe"]:
            total_violations += 1
            if is_allowlisted(
                {"rules": allow_entries}, violation, surface["surface"]
            ):
                violation["status"] = "allowlisted"
                allowlisted += 1
                used_allow_entries.add((surface["surface"], violation["id"]))
            else:
                violation["status"] = "new"
                new += 1
                failures.append(
                    f"{surface['surface']}: axe {violation['id']} ({violation.get('impact')}) "
                    f"— {violation['help']} [{violation['nodes'][0]['target']}]"
                )
        for check in surface["keyboard"]:
            if check["status"] == "fail":
                if is_keyboard_allowlisted(allow_entries, surface["surface"], check["check"]):
                    check["status"] = "allowlisted"
                    keyboard_allowlisted += 1
                    used_allow_entries.add((surface["surface"], check["check"]))
                    continue
                keyboard_failed += 1
                failures.append(
                    f"{surface['surface']}: keyboard {check['check']} — {check.get('detail')}"
                )
        for error in surface["errors"]:
            failures.append(f"{surface['surface']}: {error}")

    stale_allowlist = sorted(
        {
            (entry["surface"], entry.get("rule") or entry.get("check"))
            for entry in allow_entries
            if (entry["surface"], entry.get("rule") or entry.get("check")) not in used_allow_entries
            and entry.get("surface") != "*"
        }
    )

    report = {
        "schemaVersion": 1,
        "harness": "core-component-a11y",
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "environment": {
            "browser": "chromium",
            "browserVersion": browser_version,
            "axeVersion": json.loads(
                (axe_path.parent / "package.json").read_text(encoding="utf-8")
            ).get("version"),
            "theme": ",".join(themes),
            "themes": themes,
            "viewports": viewport_specs,
        },
        "axe": config.get("axe", {}),
        "summary": {
            "result": "FAIL" if failures else "PASS",
            "violations": total_violations,
            "allowlisted": allowlisted,
            "new": new,
            "keyboardFailed": keyboard_failed,
            "keyboardAllowlisted": keyboard_allowlisted,
            "durationMs": int((time.time() - started) * 1000),
            "failures": failures,
            "staleAllowlist": [
                {"surface": surface, "rule": rule} for surface, rule in stale_allowlist
            ],
        },
        "surfaces": reports,
    }
    (out / "report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    write_summary(report, out / "SUMMARY.md")
    write_html_report(report, out / "report.html")

    print("")
    print(
        f"a11y: {report['summary']['result']} — {total_violations} violation(s) "
        f"({allowlisted} allowlisted, {new} new), {keyboard_failed} keyboard failure(s) "
        f"({keyboard_allowlisted} allowlisted) ({report['summary']['durationMs']} ms)"
    )
    for failure in failures[:20]:
        print(f"  {failure}")
    if stale_allowlist:
        print(
            f"  {len(stale_allowlist)} allowlist entr(ies) no longer match a violation; "
            "remove them from tests/a11y/allowlist.json"
        )
    print(f"report: {out / 'report.html'}")
    if failures and not args.report_only:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
