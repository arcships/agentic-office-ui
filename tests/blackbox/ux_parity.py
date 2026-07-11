from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import getpass
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import platform
import re
import shlex
import shutil
import signal
import socket
import subprocess
import sys
import time
import traceback
import urllib.request

from playwright.sync_api import Browser, Error as PlaywrightError, Page, expect
from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
VIEWPORT = {"width": 1440, "height": 900}
DEFAULT_PORT = 4185
DEFAULT_TIMEOUT_MS = 30_000
UPSTREAM_DEFAULT_PORT = 4195
UPSTREAM_REFERENCE_ROOT = Path("/tmp/extend-hq-ui-reference")
UPSTREAM_REFERENCE_COMMIT = "f2ff2f90954acfed1f60b7fd070a9491d4113047"
UPSTREAM_REPOSITORY = "https://github.com/extend-hq/ui"
UPSTREAM_VIEWPORTS = (
    ("desktop-1440x900", {"width": 1440, "height": 900}),
    ("laptop-1280x720", {"width": 1280, "height": 720}),
)

LOCAL_MATERIAL_PATHS = (
    Path("apps/demo/public/samples/demo.docx"),
    Path("apps/demo/public/samples/legal-contract.docx"),
    Path("apps/demo/public/samples/financial-model.xlsx"),
    Path("apps/demo/public/samples/sales-table.xlsx"),
    Path("apps/demo/public/samples/charts-images.xlsx"),
)

UPSTREAM_COMPONENTS: dict[str, dict[str, object]] = {
    "docx-viewer": {
        "label": "DOCX Viewer",
        "route": "/ui/docs/components/docx-viewer",
        "readySelector": '[data-slot="component-preview"]',
        "sources": (
            {
                "path": "apps/v4/components/ui/docx-viewer.tsx",
                "entry": "DocxViewerPreview",
            },
            {
                "path": "apps/v4/components/docx-viewer-docs.tsx",
                "entry": "DocxViewerDemo",
            },
            {
                "path": "apps/v4/content/docs/components/docx-viewer.mdx",
                "entry": "documentation route",
            },
        ),
    },
    "docx-editor": {
        "label": "DOCX Editor",
        "route": "/ui/docs/components/docx-editor",
        "readySelector": '[data-slot="component-preview"]',
        "sources": (
            {
                "path": "apps/v4/components/ui/docx-editor.tsx",
                "entry": "DocxEditorPreview",
            },
            {
                "path": "apps/v4/components/docx-editor-docs.tsx",
                "entry": "DocxEditorDemo",
            },
            {
                "path": "apps/v4/content/docs/components/docx-editor.mdx",
                "entry": "documentation route",
            },
        ),
    },
    "xlsx-viewer": {
        "label": "XLSX Viewer",
        "route": "/ui/docs/components/xlsx-viewer",
        "readySelector": '[data-slot="component-preview"]',
        "sources": (
            {
                "path": "apps/v4/components/ui/xlsx-viewer.tsx",
                "entry": "XlsxViewerPreview",
            },
            {
                "path": "apps/v4/components/xlsx-viewer-docs.tsx",
                "entry": "XlsxViewerDemo",
            },
            {
                "path": "apps/v4/content/docs/components/xlsx-viewer.mdx",
                "entry": "documentation route",
            },
        ),
    },
    "xlsx-editor": {
        "label": "XLSX Editor",
        "route": "/ui/docs/components/xlsx-editor",
        "readySelector": '[data-slot="component-preview"]',
        "sources": (
            {
                "path": "apps/v4/components/ui/xlsx-editor.tsx",
                "entry": "XlsxEditorPreview",
            },
            {
                "path": "apps/v4/components/xlsx-editor-docs.tsx",
                "entry": "XlsxEditorDemo",
            },
            {
                "path": "apps/v4/content/docs/components/xlsx-editor.mdx",
                "entry": "documentation route",
            },
        ),
    },
}

UPSTREAM_CASES: dict[str, dict[str, object]] = {
    "UX-DOCX-001": {
        "components": ("docx-viewer",),
        "focus": "查看器外壳、分页、缩放、缩略图、下载与首屏布局",
    },
    "UX-DOCX-002": {
        "components": ("docx-editor",),
        "focus": "真实选区、格式修改、撤销与重做",
    },
    "UX-DOCX-003": {
        "components": ("docx-editor",),
        "focus": "回车拆段、段落合并与只读阻写",
    },
    "UX-DOCX-004": {
        "components": ("docx-editor",),
        "focus": "编辑器工具栏、表格、图片与响应式布局",
    },
    "UX-XLSX-001": {
        "components": ("xlsx-viewer", "xlsx-editor"),
        "focus": "工作区、公式栏、样式、合并与首屏布局",
    },
    "UX-XLSX-002": {
        "components": ("xlsx-viewer", "xlsx-editor"),
        "focus": "选择、编辑、尺寸调整与常用工具栏",
    },
    "UX-XLSX-003": {
        "components": ("xlsx-viewer",),
        "focus": "冻结区域、图表、图片与滚动位置",
    },
    "UX-XLSX-004": {
        "components": ("xlsx-viewer", "xlsx-editor"),
        "focus": "右键命令、搜索、只读、表排序与图表工作表",
    },
}


@dataclass(frozen=True)
class Settings:
    base_url: str
    evidence_dir: Path
    port: int
    timeout_ms: int
    manage_preview: bool


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def now_iso() -> str:
    return datetime.now().astimezone().isoformat()


def append_command_log(evidence_dir: Path, message: str) -> None:
    log_path = evidence_dir / "commands.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as stream:
        stream.write(f"[{now_iso()}] {message.rstrip()}\n")


def run_logged_command(
    evidence_dir: Path,
    command: list[str],
    *,
    cwd: Path,
    timeout: float | None = None,
) -> subprocess.CompletedProcess[str]:
    rendered = shlex.join(command)
    append_command_log(evidence_dir, f"RUN cwd={cwd} command={rendered}")
    try:
        completed = subprocess.run(
            command,
            cwd=cwd,
            text=True,
            capture_output=True,
            timeout=timeout,
            check=False,
        )
    except Exception as error:
        append_command_log(
            evidence_dir,
            f"ERROR cwd={cwd} command={rendered} error={error!r}",
        )
        raise
    append_command_log(
        evidence_dir,
        f"EXIT code={completed.returncode} cwd={cwd} command={rendered}",
    )
    if completed.stdout:
        append_command_log(evidence_dir, f"STDOUT command={rendered}\n{completed.stdout}")
    if completed.stderr:
        append_command_log(evidence_dir, f"STDERR command={rendered}\n{completed.stderr}")
    return completed


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_evidence(path: Path, *, label: str | None = None) -> dict[str, object]:
    resolved = path.resolve()
    exists = resolved.is_file()
    return {
        "path": label or str(path),
        "absolutePath": str(resolved),
        "exists": exists,
        "bytes": resolved.stat().st_size if exists else None,
        "sha256": sha256_file(resolved) if exists else None,
    }


def package_version(package: str) -> str | None:
    try:
        return importlib.metadata.version(package)
    except importlib.metadata.PackageNotFoundError:
        return None


def collect_git_state(evidence_dir: Path) -> dict[str, object]:
    commands = {
        "head": ["git", "rev-parse", "HEAD"],
        "branch": ["git", "branch", "--show-current"],
        "statusShort": ["git", "status", "--short"],
        "diffStat": ["git", "diff", "--stat"],
    }
    results = {
        key: run_logged_command(evidence_dir, command, cwd=ROOT)
        for key, command in commands.items()
    }
    status = results["statusShort"].stdout.rstrip()
    state = {
        "capturedAt": now_iso(),
        "head": results["head"].stdout.strip() or None,
        "branch": results["branch"].stdout.strip() or None,
        "dirty": bool(status),
        "statusShort": status.splitlines(),
        "diffStat": results["diffStat"].stdout.rstrip().splitlines(),
        "commands": {
            key: {
                "command": shlex.join(commands[key]),
                "exitCode": result.returncode,
                "stderr": result.stderr.rstrip(),
            }
            for key, result in results.items()
        },
    }
    write_json(evidence_dir / "git-state.json", state)
    (evidence_dir / "git-status.txt").write_text(
        status + ("\n" if status else ""), encoding="utf-8"
    )
    return state


def collect_materials() -> dict[str, object]:
    local = [
        file_evidence(ROOT / relative_path, label=str(relative_path))
        for relative_path in LOCAL_MATERIAL_PATHS
    ]
    upstream_fixture_paths = (
        Path("apps/v4/public/samples/demo.docx"),
        Path("apps/v4/public/samples/crazy-chart-zoo.xlsx"),
    )
    upstream = [
        file_evidence(
            UPSTREAM_REFERENCE_ROOT / relative_path,
            label=str(relative_path),
        )
        for relative_path in upstream_fixture_paths
    ]
    return {
        "schemaVersion": 1,
        "capturedAt": now_iso(),
        "localFixtures": local,
        "upstreamFixtures": upstream,
    }


def upstream_git_value(
    evidence_dir: Path, *arguments: str
) -> subprocess.CompletedProcess[str]:
    return run_logged_command(
        evidence_dir,
        ["git", *arguments],
        cwd=UPSTREAM_REFERENCE_ROOT,
    )


def build_upstream_reference(evidence_dir: Path) -> dict[str, object]:
    root_exists = UPSTREAM_REFERENCE_ROOT.is_dir()
    actual_commit: str | None = None
    upstream_status: list[str] = []
    git_error: str | None = None
    if root_exists:
        try:
            head = upstream_git_value(evidence_dir, "rev-parse", "HEAD")
            status = upstream_git_value(evidence_dir, "status", "--short")
            if head.returncode == 0:
                actual_commit = head.stdout.strip() or None
            else:
                git_error = head.stderr.strip() or "git rev-parse failed"
            if status.returncode == 0:
                upstream_status = status.stdout.rstrip().splitlines()
            elif git_error is None:
                git_error = status.stderr.strip() or "git status failed"
        except Exception as error:
            git_error = repr(error)

    components: list[dict[str, object]] = []
    component_sources: dict[str, list[dict[str, object]]] = {}
    for component_id, component in UPSTREAM_COMPONENTS.items():
        sources = []
        for source in component["sources"]:
            source_record = dict(source)
            relative_path = Path(str(source_record["path"]))
            source_record.update(
                file_evidence(
                    UPSTREAM_REFERENCE_ROOT / relative_path,
                    label=str(relative_path),
                )
            )
            sources.append(source_record)
        component_sources[component_id] = sources
        components.append(
            {
                "id": component_id,
                "label": component["label"],
                "route": component["route"],
                "readySelector": component["readySelector"],
                "sources": sources,
            }
        )

    cases: list[dict[str, object]] = []
    for case_id, case in UPSTREAM_CASES.items():
        component_ids = list(case["components"])
        sources = [
            {"component": component_id, **source}
            for component_id in component_ids
            for source in component_sources[component_id]
        ]
        cases.append(
            {
                "id": case_id,
                "focus": case["focus"],
                "components": component_ids,
                "sources": sources,
            }
        )

    source_complete = all(
        bool(source["exists"])
        for component in components
        for source in component["sources"]
    )
    return {
        "schemaVersion": 1,
        "generatedAt": now_iso(),
        "repository": UPSTREAM_REPOSITORY,
        "path": str(UPSTREAM_REFERENCE_ROOT),
        "pathExists": root_exists,
        "expectedCommit": UPSTREAM_REFERENCE_COMMIT,
        "actualCommit": actual_commit,
        "commitMatches": actual_commit == UPSTREAM_REFERENCE_COMMIT,
        "gitStatusShort": upstream_status,
        "gitError": git_error,
        "sourceComplete": source_complete,
        "components": components,
        "cases": cases,
        "capture": {
            "status": "NOT_RUN",
            "requestedViewports": [
                {"name": name, **viewport}
                for name, viewport in UPSTREAM_VIEWPORTS
            ],
            "results": [],
        },
    }


def prepare_root_evidence(settings: Settings) -> dict[str, object]:
    settings.evidence_dir.mkdir(parents=True, exist_ok=True)
    (settings.evidence_dir / "commands.log").write_text("", encoding="utf-8")
    append_command_log(
        settings.evidence_dir,
        "INVOCATION " + shlex.join([sys.executable, *sys.argv]),
    )
    git_state = collect_git_state(settings.evidence_dir)
    materials = collect_materials()
    write_json(settings.evidence_dir / "materials.json", materials)
    upstream_reference = build_upstream_reference(settings.evidence_dir)
    write_json(
        settings.evidence_dir / "upstream-reference.json", upstream_reference
    )

    node = run_logged_command(
        settings.evidence_dir, ["node", "--version"], cwd=ROOT
    )
    pnpm = run_logged_command(
        settings.evidence_dir, ["pnpm", "--version"], cwd=ROOT
    )
    executor = {
        "user": getpass.getuser(),
        "uid": os.getuid() if hasattr(os, "getuid") else None,
        "pid": os.getpid(),
        "hostname": socket.gethostname(),
        "ci": os.environ.get("CI", "").lower() in {"1", "true", "yes"},
        "githubActor": os.environ.get("GITHUB_ACTOR"),
        "githubRunId": os.environ.get("GITHUB_RUN_ID"),
    }
    environment = {
        "schemaVersion": 1,
        "capturedAt": now_iso(),
        "suite": "UX-PARITY",
        "mode": "formal preview",
        "cwd": str(ROOT),
        "evidenceDir": str(settings.evidence_dir),
        "executor": executor,
        "platform": platform.platform(),
        "machine": platform.machine(),
        "python": {
            "executable": sys.executable,
            "version": sys.version,
        },
        "node": node.stdout.strip() or None,
        "pnpm": pnpm.stdout.strip() or None,
        "playwright": package_version("playwright"),
        "baseUrl": settings.base_url,
        "previewManagedByScript": settings.manage_preview,
        "viewport": VIEWPORT,
        "deviceScaleFactor": 1,
        "locale": "zh-CN",
        "timezone": "Asia/Shanghai",
        "git": git_state,
        "materials": materials,
        "upstream": {
            "repository": UPSTREAM_REPOSITORY,
            "path": str(UPSTREAM_REFERENCE_ROOT),
            "expectedCommit": UPSTREAM_REFERENCE_COMMIT,
            "actualCommit": upstream_reference["actualCommit"],
            "commitMatches": upstream_reference["commitMatches"],
        },
        "artifacts": {
            "commands": "commands.log",
            "gitState": "git-state.json",
            "gitStatus": "git-status.txt",
            "materials": "materials.json",
            "upstreamReference": "upstream-reference.json",
        },
    }
    write_json(settings.evidence_dir / "environment.json", environment)
    return upstream_reference


def update_environment(
    settings: Settings, **updates: object
) -> None:
    path = settings.evidence_dir / "environment.json"
    environment = json.loads(path.read_text(encoding="utf-8"))
    environment.update(updates)
    write_json(path, environment)


def git_head() -> str:
    try:
        return subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip()
    except Exception:
        return "unknown-head"


def default_evidence_dir() -> Path:
    run_id = os.environ.get(
        "BLACKBOX_RUN_ID", datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    )
    return (
        ROOT
        / "output"
        / "acceptance"
        / git_head()
        / "UX-PARITY"
        / run_id
        / "local"
    )


def parse_args() -> Settings:
    parser = argparse.ArgumentParser(
        description=(
            "Run DOCX Viewer, DOCX Editor, and XLSX Viewer UX parity checks "
            "against a formal Vite preview."
        )
    )
    parser.add_argument(
        "--base-url",
        default=os.environ.get("UX_PARITY_BASE_URL", ""),
        help=(
            "Use an already-running formal preview, for example "
            "http://127.0.0.1:4173. When omitted, this script starts preview."
        ),
    )
    parser.add_argument(
        "--evidence-dir",
        default=os.environ.get("BLACKBOX_EVIDENCE_DIR", ""),
        help="Evidence directory. Defaults to output/acceptance/<HEAD>/UX-PARITY/<run-id>/local.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("CI_UX_PARITY_PREVIEW_PORT", DEFAULT_PORT)),
        help=f"Preview port when this script starts the server (default: {DEFAULT_PORT}).",
    )
    parser.add_argument(
        "--timeout-ms",
        type=int,
        default=int(os.environ.get("UX_PARITY_TIMEOUT_MS", DEFAULT_TIMEOUT_MS)),
        help=f"Per readiness/action timeout (default: {DEFAULT_TIMEOUT_MS}).",
    )
    args = parser.parse_args()

    if args.port < 1 or args.port > 65535:
        parser.error("--port must be between 1 and 65535")
    if args.timeout_ms < 1_000:
        parser.error("--timeout-ms must be at least 1000")

    manage_preview = not bool(args.base_url.strip())
    base_url = (
        args.base_url.rstrip("/")
        if not manage_preview
        else f"http://127.0.0.1:{args.port}"
    )
    evidence_dir = (
        Path(args.evidence_dir).expanduser().resolve()
        if args.evidence_dir
        else default_evidence_dir().resolve()
    )
    return Settings(
        base_url=base_url,
        evidence_dir=evidence_dir,
        port=args.port,
        timeout_ms=args.timeout_ms,
        manage_preview=manage_preview,
    )


def port_in_use(port: int) -> bool:
    with socket.socket() as sock:
        return sock.connect_ex(("127.0.0.1", port)) == 0


def wait_for_preview(
    base_url: str,
    process: subprocess.Popen[bytes] | None,
    timeout_seconds: float = 30,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if process is not None and process.poll() is not None:
            raise RuntimeError(f"preview exited early with {process.returncode}")
        try:
            with urllib.request.urlopen(base_url, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"formal preview did not become ready at {base_url}")


def stop_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=5)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            process.kill()


def launch_browser(playwright) -> tuple[Browser, str]:
    failures: list[str] = []
    try:
        return playwright.chromium.launch(headless=True), "bundled-chromium"
    except PlaywrightError as error:
        failures.append(str(error))

    for candidate in (
        Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        Path("/usr/bin/google-chrome"),
        Path("/usr/bin/google-chrome-stable"),
    ):
        if not candidate.exists():
            continue
        try:
            return (
                playwright.chromium.launch(
                    headless=True, executable_path=str(candidate)
                ),
                str(candidate),
            )
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


def upstream_capture_artifact(
    path: Path, evidence_dir: Path
) -> dict[str, object]:
    record = file_evidence(path)
    try:
        record["path"] = str(path.relative_to(evidence_dir))
    except ValueError:
        pass
    return record


def capture_upstream_page(
    browser: Browser,
    component_id: str,
    component: dict[str, object],
    viewport_name: str,
    viewport: dict[str, int],
    base_url: str,
    evidence_dir: Path,
    timeout_ms: int,
) -> dict[str, object]:
    capture_dir = evidence_dir / "upstream-reference" / component_id / viewport_name
    capture_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport=viewport,
        device_scale_factor=1,
        locale="en-US",
        timezone_id="Asia/Shanghai",
        accept_downloads=False,
    )
    page = context.new_page()
    page.set_default_timeout(timeout_ms)
    page.set_default_navigation_timeout(timeout_ms * 3)
    evidence = BrowserEvidence(page)
    url = f"{base_url}{component['route']}"
    started_at = now_iso()
    result: dict[str, object]
    try:
        response = page.goto(url, wait_until="networkidle", timeout=timeout_ms * 3)
        if response is None:
            raise RuntimeError(f"upstream navigation returned no response for {url}")
        if response.status >= 400:
            raise RuntimeError(
                f"upstream navigation returned HTTP {response.status} for {url}"
            )
        preview = page.locator(str(component["readySelector"])).first
        preview.wait_for(state="visible", timeout=timeout_ms * 2)
        page.wait_for_timeout(1_500)
        box = preview.bounding_box()
        if box is None or box["width"] < 200 or box["height"] < 200:
            raise RuntimeError(
                f"upstream component preview has no usable layout: {box!r}"
            )

        page_path = capture_dir / "page.png"
        component_path = capture_dir / "component.png"
        page.screenshot(path=str(page_path), full_page=True)
        preview.screenshot(path=str(component_path))
        violations = evidence.violations()
        # The fixed upstream development site emits font-preload warnings of
        # its own. Keep every warning in evidence, but only an upstream console
        # error, page error, failed request or bad response invalidates the
        # reference screenshot.
        blocking_violations = {
            **violations,
            "console": [
                entry
                for entry in violations["console"]
                if entry.get("type") == "error"
            ],
        }
        status = "PASS" if not any(blocking_violations.values()) else "FAIL"
        result = {
            "component": component_id,
            "route": component["route"],
            "url": url,
            "viewport": {"name": viewport_name, **viewport},
            "status": status,
            "startedAt": started_at,
            "finishedAt": now_iso(),
            "httpStatus": response.status,
            "componentBox": box,
            "screenshots": {
                "page": upstream_capture_artifact(page_path, evidence_dir),
                "component": upstream_capture_artifact(
                    component_path, evidence_dir
                ),
            },
            "browserViolations": violations,
            "blockingViolations": blocking_violations,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(capture_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "component": component_id,
            "route": component["route"],
            "url": url,
            "viewport": {"name": viewport_name, **viewport},
            "status": "FAIL",
            "startedAt": started_at,
            "finishedAt": now_iso(),
            "error": str(error),
            "traceback": traceback.format_exc(),
        }
    finally:
        evidence.save(capture_dir)
        context.close()
    write_json(capture_dir / "result.json", result)
    return result


def finalize_upstream_reference(
    settings: Settings,
    reference: dict[str, object],
) -> dict[str, object]:
    write_json(settings.evidence_dir / "upstream-reference.json", reference)
    capture = reference["capture"]
    update_environment(
        settings,
        upstreamCapture={
            "status": capture["status"],
            "reasonCode": capture.get("reasonCode"),
            "reason": capture.get("reason"),
            "artifact": "upstream-reference.json",
        },
    )
    return reference


def block_upstream_capture(
    settings: Settings,
    reference: dict[str, object],
    reason_code: str,
    reason: str,
    **details: object,
) -> dict[str, object]:
    reference["capture"].update(
        {
            "status": "BLOCKED",
            "reasonCode": reason_code,
            "reason": reason,
            "finishedAt": now_iso(),
            **details,
        }
    )
    append_command_log(
        settings.evidence_dir,
        f"UPSTREAM_CAPTURE_BLOCKED reasonCode={reason_code} reason={reason}",
    )
    return finalize_upstream_reference(settings, reference)


def probe_and_capture_upstream_reference(
    settings: Settings,
    reference: dict[str, object],
) -> dict[str, object]:
    capture = reference["capture"]
    capture["startedAt"] = now_iso()
    if not reference["pathExists"]:
        return block_upstream_capture(
            settings,
            reference,
            "reference_path_missing",
            f"fixed upstream path does not exist: {UPSTREAM_REFERENCE_ROOT}",
        )
    if not reference["commitMatches"]:
        return block_upstream_capture(
            settings,
            reference,
            "commit_mismatch",
            (
                "fixed upstream worktree is not at the required commit: "
                f"expected {UPSTREAM_REFERENCE_COMMIT}, got "
                f"{reference['actualCommit']}"
            ),
        )
    if not reference["sourceComplete"]:
        missing_sources = [
            source["path"]
            for component in reference["components"]
            for source in component["sources"]
            if not source["exists"]
        ]
        return block_upstream_capture(
            settings,
            reference,
            "source_missing",
            "one or more fixed upstream source entries are missing",
            missingSources=missing_sources,
        )

    upstream_app = UPSTREAM_REFERENCE_ROOT / "apps" / "v4"
    dependency_paths = [
        UPSTREAM_REFERENCE_ROOT / "node_modules",
        upstream_app / "node_modules",
        UPSTREAM_REFERENCE_ROOT / "node_modules" / ".bin" / "next",
        upstream_app / "node_modules" / ".bin" / "next",
    ]
    dependency_state = [
        {
            "path": str(path),
            "exists": path.exists(),
            "kind": "directory" if path.is_dir() else "file",
        }
        for path in dependency_paths
    ]
    capture["dependencyPaths"] = dependency_state

    pnpm_path = shutil.which("pnpm")
    if pnpm_path is None:
        return block_upstream_capture(
            settings,
            reference,
            "pnpm_unavailable",
            "pnpm is not available, so the fixed upstream app cannot start",
        )

    probe_command = [pnpm_path, "exec", "next", "--version"]
    probe = run_logged_command(
        settings.evidence_dir,
        probe_command,
        cwd=upstream_app,
        timeout=30,
    )
    capture["directStartProbe"] = {
        "command": shlex.join(probe_command),
        "cwd": str(upstream_app),
        "exitCode": probe.returncode,
        "stdout": probe.stdout.rstrip(),
        "stderr": probe.stderr.rstrip(),
    }
    if probe.returncode != 0:
        return block_upstream_capture(
            settings,
            reference,
            "upstream_dependencies_missing",
            (
                "the fixed upstream checkout has no executable Next.js "
                "dependency; direct `pnpm exec next --version` failed and "
                "the test does not install or synthesize dependencies"
            ),
        )

    try:
        upstream_port = int(
            os.environ.get("UX_PARITY_UPSTREAM_PORT", UPSTREAM_DEFAULT_PORT)
        )
    except ValueError:
        return block_upstream_capture(
            settings,
            reference,
            "invalid_upstream_port",
            "UX_PARITY_UPSTREAM_PORT is not an integer",
        )
    if upstream_port < 1 or upstream_port > 65535:
        return block_upstream_capture(
            settings,
            reference,
            "invalid_upstream_port",
            f"upstream port is outside 1..65535: {upstream_port}",
        )
    if port_in_use(upstream_port):
        return block_upstream_capture(
            settings,
            reference,
            "upstream_port_in_use",
            (
                f"port {upstream_port} is already in use; refusing to attach "
                "to an unverified server"
            ),
        )

    base_url = f"http://127.0.0.1:{upstream_port}"
    start_command = [
        pnpm_path,
        "exec",
        "next",
        "dev",
        "--webpack",
        "--hostname",
        "127.0.0.1",
        "--port",
        str(upstream_port),
    ]
    upstream_dir = settings.evidence_dir / "upstream-reference"
    upstream_dir.mkdir(parents=True, exist_ok=True)
    preview_log_path = upstream_dir / "preview.log"
    process: subprocess.Popen[bytes] | None = None
    preview_log = None
    try:
        preview_log = preview_log_path.open("wb")
        upstream_environment = {
            **os.environ,
            "NEXT_PUBLIC_APP_URL": f"{base_url}/ui",
            "NEXT_PUBLIC_BASE_PATH": "/ui",
            "NEXT_PUBLIC_ASSET_PREFIX": "/ui-static",
        }
        append_command_log(
            settings.evidence_dir,
            (
                f"START cwd={upstream_app} command={shlex.join(start_command)} "
                "env=NEXT_PUBLIC_BASE_PATH=/ui,"
                "NEXT_PUBLIC_ASSET_PREFIX=/ui-static"
            ),
        )
        process = subprocess.Popen(
            start_command,
            cwd=upstream_app,
            env=upstream_environment,
            stdout=preview_log,
            stderr=subprocess.STDOUT,
            start_new_session=True,
        )
        capture["server"] = {
            "command": shlex.join(start_command),
            "cwd": str(upstream_app),
            "baseUrl": base_url,
            "pid": process.pid,
            "log": str(preview_log_path.relative_to(settings.evidence_dir)),
        }
        wait_for_preview(
            f"{base_url}/ui/docs/components/docx-viewer",
            process,
            timeout_seconds=max(90, settings.timeout_ms / 1000 * 3),
        )

        with sync_playwright() as playwright:
            browser, browser_source = launch_browser(playwright)
            capture["browser"] = {
                "version": browser.version,
                "source": browser_source,
            }
            results: list[dict[str, object]] = []
            try:
                for component_id, component in UPSTREAM_COMPONENTS.items():
                    for viewport_name, viewport in UPSTREAM_VIEWPORTS:
                        results.append(
                            capture_upstream_page(
                                browser,
                                component_id,
                                component,
                                viewport_name,
                                viewport,
                                base_url,
                                settings.evidence_dir,
                                settings.timeout_ms,
                            )
                        )
            finally:
                browser.close()
        capture["results"] = results
        capture["status"] = (
            "PASS"
            if results and all(result["status"] == "PASS" for result in results)
            else "FAIL"
        )
        capture["finishedAt"] = now_iso()
        return finalize_upstream_reference(settings, reference)
    except Exception as error:
        return block_upstream_capture(
            settings,
            reference,
            "upstream_start_or_capture_failed",
            str(error),
            traceback=traceback.format_exc(),
        )
    finally:
        if process is not None:
            append_command_log(
                settings.evidence_dir,
                f"STOP pid={process.pid} signal=SIGTERM label=upstream-reference",
            )
        stop_process(process)
        if preview_log is not None:
            preview_log.close()


def goto_ready(
    page: Page,
    settings: Settings,
    route: str,
    component_test_id: str,
    component_requires_ready: bool = True,
) -> None:
    page.goto(
        f"{settings.base_url}/#{route}",
        wait_until="networkidle",
        timeout=settings.timeout_ms * 2,
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )
    component = page.get_by_test_id(component_test_id)
    component.wait_for(state="visible", timeout=settings.timeout_ms)
    if component_requires_ready:
        page.wait_for_function(
            """testId => document.querySelector(`[data-testid="${testId}"]`)?.dataset.state === 'ready'""",
            arg=component_test_id,
            timeout=settings.timeout_ms,
        )


def element_layout(locator) -> dict[str, object]:
    return locator.evaluate(
        r"""element => {
          const rect = element.getBoundingClientRect();
          const controls = [...element.querySelectorAll('button,input,select,textarea')]
            .filter((control) => {
              const style = getComputedStyle(control);
              const box = control.getBoundingClientRect();
              return style.display !== 'none' && style.visibility !== 'hidden' &&
                box.width > 0 && box.height > 0;
            })
            .map((control) => {
              const box = control.getBoundingClientRect();
              return {
                tag: control.tagName.toLowerCase(),
                testId: control.getAttribute('data-testid'),
                title: control.getAttribute('title'),
                ariaLabel: control.getAttribute('aria-label'),
                text: (control.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120),
                left: box.left,
                right: box.right,
                width: box.width,
              };
            });
          return {
            left: rect.left,
            right: rect.right,
            width: rect.width,
            clientWidth: element.clientWidth,
            scrollWidth: element.scrollWidth,
            scrollLeft: element.scrollLeft,
            hiddenPixels: Math.max(0, element.scrollWidth - element.clientWidth),
            overflowX: getComputedStyle(element).overflowX,
            controls,
            clippedControls: controls.filter((control) =>
              control.left < rect.left - 1 || control.right > rect.right + 1
            ),
          };
        }"""
    )


def assert_no_horizontal_hiding(name: str, metrics: dict[str, object]) -> None:
    hidden_pixels = int(metrics["hiddenPixels"])
    clipped_controls = metrics["clippedControls"]
    assert hidden_pixels <= 2, (
        f"{name} hides {hidden_pixels}px horizontally: "
        + json.dumps(metrics, ensure_ascii=False)
    )
    assert not clipped_controls, (
        f"{name} has controls outside its visible bounds: "
        + json.dumps(clipped_controls, ensure_ascii=False)
    )


def assert_no_page_overflow(page: Page, name: str) -> dict[str, int]:
    metrics = page.evaluate(
        """() => ({
          viewportWidth: window.innerWidth,
          documentWidth: document.documentElement.scrollWidth,
          bodyWidth: document.body.scrollWidth,
        })"""
    )
    assert metrics["documentWidth"] <= metrics["viewportWidth"] + 2, (
        f"{name} has page-level horizontal overflow: {metrics}"
    )
    assert metrics["bodyWidth"] <= metrics["viewportWidth"] + 2, (
        f"{name} body exceeds viewport: {metrics}"
    )
    return metrics


def read_json_test_id(page: Page, test_id: str) -> dict[str, object]:
    raw = page.get_by_test_id(test_id).text_content() or ""
    parsed = json.loads(raw)
    assert isinstance(parsed, dict), f"{test_id} must contain a JSON object"
    return parsed


def wait_for_model_text(
    page: Page,
    node_index: int,
    expected_text: str,
    timeout_ms: int,
) -> dict[str, object]:
    deadline = time.monotonic() + timeout_ms / 1000
    last: dict[str, object] | None = None
    while time.monotonic() < deadline:
        last = read_json_test_id(page, "editor-model-snapshot")
        nodes = last.get("nodes")
        if (
            isinstance(nodes, list)
            and node_index < len(nodes)
            and isinstance(nodes[node_index], dict)
            and nodes[node_index].get("text") == expected_text
        ):
            return last
        page.wait_for_timeout(40)
    raise AssertionError(
        f"timed out waiting for editor node {node_index} text {expected_text!r}; "
        f"last model={last}"
    )


def wait_for_model_condition(
    page: Page,
    predicate,
    timeout_ms: int,
    description: str,
) -> dict[str, object]:
    deadline = time.monotonic() + timeout_ms / 1000
    last: dict[str, object] | None = None
    while time.monotonic() < deadline:
        last = read_json_test_id(page, "editor-model-snapshot")
        if predicate(last):
            return last
        page.wait_for_timeout(40)
    raise AssertionError(f"timed out waiting for {description}; last model={last}")


def set_dom_text_range(locator, start: int, end: int) -> None:
    locator.evaluate(
        """(element, offsets) => {
          const textNodes = [];
          const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
          let current = walker.nextNode();
          while (current) {
            const parent = current.parentElement;
            if (!parent?.closest('[data-docx-numbering-label="true"]')) textNodes.push(current);
            current = walker.nextNode();
          }
          const resolve = (offset) => {
            let remaining = offset;
            for (const node of textNodes) {
              const length = node.textContent?.length || 0;
              if (remaining <= length) return {node, offset: remaining};
              remaining -= length;
            }
            const node = textNodes[textNodes.length - 1];
            if (!node) return {node: element, offset: 0};
            return {node, offset: node.textContent?.length || 0};
          };
          const anchor = resolve(offsets.start);
          const focus = resolve(offsets.end);
          // Real mouse selection focuses the contenteditable before the
          // browser establishes the expanded range. Doing this afterwards can
          // collapse the range through the editor focus handler.
          element.focus();
          const range = document.createRange();
          range.setStart(anchor.node, anchor.offset);
          range.setEnd(focus.node, focus.offset);
          const selection = window.getSelection();
          selection?.removeAllRanges();
          selection?.addRange(range);
          document.dispatchEvent(new Event('selectionchange'));
        }""",
        {"start": start, "end": end},
    )


def run_ux_docx_001(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_ready(page, settings, "/docx-viewer", "docx-viewer")

    sample = page.get_by_test_id("docx-sample-select")
    option_values = sample.locator("option").evaluate_all(
        "options => options.map((option) => option.value)"
    )
    assert "legal-contract.docx" in option_values, option_values
    if sample.input_value() != "legal-contract.docx":
        sample.select_option("legal-contract.docx")
        page.wait_for_function(
            """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('legal-contract.docx')""",
            timeout=settings.timeout_ms,
        )
        page.wait_for_function(
            """() => document.querySelector('[data-testid="docx-viewer"]')?.dataset.state === 'ready'""",
            timeout=settings.timeout_ms,
        )
    page.get_by_test_id("docx-page").first.wait_for(
        state="visible", timeout=settings.timeout_ms
    )
    page.wait_for_timeout(300)
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)

    toolbar = page.get_by_test_id("docx-viewer-toolbar")
    toolbar.wait_for(state="visible")
    toolbar_metrics = element_layout(toolbar)
    assert_no_horizontal_hiding("DOCX viewer toolbar", toolbar_metrics)

    page_input = page.get_by_test_id("docx-page-current")
    zoom_select = page.get_by_test_id("docx-zoom-select")
    sidebar_toggle = page.get_by_test_id("docx-sidebar-toggle")
    download_button = page.get_by_test_id("docx-download")
    for locator in (page_input, zoom_select, sidebar_toggle, download_button):
        locator.wait_for(state="visible")

    navigation_text = toolbar.locator('[aria-label="Page navigation"]').inner_text()
    total_match = re.search(r"\bof\s+(\d+)\b", navigation_text)
    assert total_match, navigation_text
    total_pages = int(total_match.group(1))
    assert total_pages >= 2, f"multi-page fixture rendered only {total_pages} page(s)"
    assert int(page_input.input_value()) >= 1

    page_surface = page.get_by_test_id("docx-page").first
    before_box = page_surface.bounding_box()
    assert before_box and before_box["width"] > 0, before_box
    zoom_select.select_option("125")
    expect(zoom_select).to_have_value("125", timeout=settings.timeout_ms)
    page.wait_for_function(
        """minimumWidth => {
          const page = document.querySelector('[data-testid="docx-page"]');
          return page && page.getBoundingClientRect().width >= minimumWidth;
        }""",
        arg=before_box["width"] * 1.15,
        timeout=settings.timeout_ms,
    )
    after_box = page_surface.bounding_box()
    assert after_box and after_box["width"] > before_box["width"] * 1.15

    sidebar_toggle.click()
    expect(sidebar_toggle).to_have_attribute(
        "aria-pressed", "true", timeout=settings.timeout_ms
    )
    thumbnail_panel = page.get_by_test_id("docx-thumbnail-panel")
    thumbnail_panel.wait_for(state="visible", timeout=settings.timeout_ms)
    thumbnails = page.get_by_test_id("docx-thumbnail")
    assert thumbnails.count() >= 1
    thumbnail_canvas = thumbnails.first.locator("canvas")
    expect(thumbnail_canvas).to_have_attribute(
        "data-thumbnail-state", "ready", timeout=settings.timeout_ms
    )
    thumbnail_pixels = thumbnail_canvas.evaluate(
        """canvas => {
          const context = canvas.getContext('2d');
          if (!context || canvas.width < 6 || canvas.height < 6) return null;
          const pixels = context.getImageData(2, 2, canvas.width - 4, canvas.height - 4).data;
          let darkPixels = 0;
          for (let index = 0; index < pixels.length; index += 4) {
            if (pixels[index] < 235 || pixels[index + 1] < 235 || pixels[index + 2] < 235) darkPixels += 1;
          }
          return {
            contentBlocks: Number(canvas.dataset.thumbnailContentBlocks || 0),
            darkPixelRatio: darkPixels / (pixels.length / 4),
            encodedLength: canvas.toDataURL('image/png').length,
          };
        }"""
    )
    assert thumbnail_pixels is not None
    assert thumbnail_pixels["contentBlocks"] > 0, thumbnail_pixels
    assert thumbnail_pixels["darkPixelRatio"] > 0.002, thumbnail_pixels
    canvas_box = thumbnail_canvas.bounding_box()
    assert canvas_box and canvas_box["width"] > 0 and canvas_box["height"] > 0

    page_input.fill("2")
    page_input.press("Enter")
    expect(page_input).to_have_value("2", timeout=settings.timeout_ms)
    page.wait_for_function(
        """() => document.querySelector('[data-testid="docx-thumbnail"][aria-current="page"]')?.dataset.thumbnailPageIndex === '1'""",
        timeout=settings.timeout_ms,
    )

    expect(download_button).to_be_enabled(timeout=settings.timeout_ms)
    download_dir = attempt_dir / "download"
    download_dir.mkdir(parents=True, exist_ok=True)
    with page.expect_download(timeout=settings.timeout_ms) as download_info:
        download_button.click()
    download = download_info.value
    safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", download.suggested_filename)
    saved_download = download_dir / (safe_name or "document.docx")
    download.save_as(saved_download)
    download_bytes = saved_download.read_bytes()
    assert len(download_bytes) > 2
    assert download_bytes[:2] == b"PK", "downloaded DOCX is not a ZIP container"

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    page.set_viewport_size({"width": 1280, "height": 720})
    page.wait_for_timeout(150)
    laptop_toolbar = element_layout(toolbar)
    assert_no_horizontal_hiding("DOCX viewer laptop toolbar", laptop_toolbar)
    laptop_page = assert_no_page_overflow(page, "DOCX viewer laptop layout")
    page.screenshot(path=str(attempt_dir / "laptop.png"), full_page=True)
    return {
        "expectedDownloads": 1,
        "toolbar": toolbar_metrics,
        "laptop": {"toolbar": laptop_toolbar, "page": laptop_page},
        "pageNavigation": {
            "current": int(page_input.input_value()),
            "total": total_pages,
        },
        "zoom": {
            "before": before_box["width"],
            "after": after_box["width"],
            "value": zoom_select.input_value(),
        },
        "thumbnails": {
            "count": thumbnails.count(),
            "firstCanvas": canvas_box,
            "firstCanvasPixels": thumbnail_pixels,
        },
        "download": {
            "suggestedFilename": download.suggested_filename,
            "savedPath": str(saved_download),
            "bytes": len(download_bytes),
        },
    }


def run_docx_editor(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_ready(
        page,
        settings,
        "/docx-editor",
        "docx-editor",
        component_requires_ready=False,
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="editor-editable-state"]')?.dataset.state === 'editable'
          && document.querySelectorAll('[data-testid="editor-paragraph"][contenteditable="true"]').length >= 3""",
        timeout=settings.timeout_ms,
    )
    page.get_by_test_id("docx-editor-toolbar").wait_for(
        state="visible", timeout=settings.timeout_ms
    )
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)

    topbar = page.locator(
        '[data-testid="docx-editor-toolbar"] .docx-toolbar-topbar'
    )
    formatting_bar = page.locator(
        '[data-testid="docx-editor-toolbar"] .docx-toolbar'
    )
    assert topbar.count() == 1 and formatting_bar.count() == 1
    topbar_metrics = element_layout(topbar)
    formatting_metrics = element_layout(formatting_bar)
    assert_no_horizontal_hiding("DOCX editor top toolbar", topbar_metrics)
    assert_no_horizontal_hiding("DOCX editor formatting toolbar", formatting_metrics)

    target = page.locator(
        '[data-testid="editor-paragraph"][contenteditable="true"]'
    ).filter(has_text="This paragraph is ready for editing").first
    target.wait_for(state="visible", timeout=settings.timeout_ms)
    node_index_raw = target.get_attribute("data-node-index")
    assert node_index_raw is not None
    node_index = int(node_index_raw)
    original_text = target.inner_text()
    marker = "〔UX-PARITY〕"

    target.click()
    target.evaluate(
        """element => {
          element.focus();
          const range = document.createRange();
          range.selectNodeContents(element);
          range.collapse(false);
          const selection = window.getSelection();
          selection?.removeAllRanges();
          selection?.addRange(range);
        }"""
    )
    page.keyboard.type(marker)
    page.locator("h2").first.click()

    expected_text = original_text + marker
    committed_model = wait_for_model_text(
        page, node_index, expected_text, settings.timeout_ms
    )
    history_state = page.get_by_test_id("editor-history-state")
    expect(history_state).to_have_attribute(
        "data-can-undo", "true", timeout=settings.timeout_ms
    )
    undo = page.get_by_test_id("editor-undo")
    expect(undo).to_be_enabled(timeout=settings.timeout_ms)
    undo.click()
    restored_model = wait_for_model_text(
        page, node_index, original_text, settings.timeout_ms
    )
    expect(target).to_have_text(original_text, timeout=settings.timeout_ms)
    expect(history_state).to_have_attribute(
        "data-can-redo", "true", timeout=settings.timeout_ms
    )

    # A real browser range must drive the formatting command. This catches the
    # old path where toolbar commands only used a synthetic paragraph range.
    range_text = "editing"
    range_start = original_text.index(range_text)
    set_dom_text_range(target, range_start, range_start + len(range_text))
    page.wait_for_function(
        """expected => {
          const raw = document.querySelector('[data-testid="editor-selection-snapshot"]')?.textContent;
          if (!raw) return false;
          const range = JSON.parse(raw).activeTextRange;
          return range?.start?.offset === expected.start && range?.end?.offset === expected.end;
        }""",
        arg={"start": range_start, "end": range_start + len(range_text)},
        timeout=settings.timeout_ms,
    )
    bold = page.get_by_test_id("editor-bold")
    expect(bold).to_be_enabled(timeout=settings.timeout_ms)
    bold.click()

    def selected_text_is_bold(snapshot: dict[str, object]) -> bool:
        nodes = snapshot.get("nodes")
        if not isinstance(nodes, list) or node_index >= len(nodes):
            return False
        node = nodes[node_index]
        if not isinstance(node, dict):
            return False
        runs = node.get("runs")
        return isinstance(runs, list) and any(
            isinstance(run, dict)
            and run.get("text") == range_text
            and isinstance(run.get("style"), dict)
            and run["style"].get("bold") is True
            for run in runs
        )

    formatted_model = wait_for_model_condition(
        page,
        selected_text_is_bold,
        settings.timeout_ms,
        "the real DOM selection to become bold",
    )
    undo.click()
    unformatted_model = wait_for_model_condition(
        page,
        lambda snapshot: not selected_text_is_bold(snapshot),
        settings.timeout_ms,
        "the selected formatting undo",
    )

    # Enter must split the paragraph at the real caret; Backspace at the start
    # of the new paragraph must merge it again without losing text.
    nodes_before_enter = unformatted_model.get("nodes")
    assert isinstance(nodes_before_enter, list)
    split_offset = original_text.index(".") + 1
    set_dom_text_range(target, split_offset, split_offset)
    page.keyboard.press("Enter")

    def paragraph_is_split(snapshot: dict[str, object]) -> bool:
        nodes = snapshot.get("nodes")
        if not isinstance(nodes, list) or len(nodes) != len(nodes_before_enter) + 1:
            return False
        left = nodes[node_index]
        right = nodes[node_index + 1]
        return (
            isinstance(left, dict)
            and isinstance(right, dict)
            and left.get("text") == original_text[:split_offset]
            and right.get("text") == original_text[split_offset:]
        )

    split_model = wait_for_model_condition(
        page, paragraph_is_split, settings.timeout_ms, "Enter to split the paragraph"
    )
    second_paragraph = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index + 1}"]'
    ).first
    second_paragraph.wait_for(state="visible", timeout=settings.timeout_ms)
    set_dom_text_range(second_paragraph, 0, 0)
    page.keyboard.press("Backspace")
    merged_model = wait_for_model_condition(
        page,
        lambda snapshot: (
            isinstance(snapshot.get("nodes"), list)
            and len(snapshot["nodes"]) == len(nodes_before_enter)
            and isinstance(snapshot["nodes"][node_index], dict)
            and snapshot["nodes"][node_index].get("text") == original_text
        ),
        settings.timeout_ms,
        "Backspace to merge the split paragraphs",
    )

    # Read-only is a product state, not only a contenteditable attribute.
    read_only = page.get_by_test_id("editor-readonly")
    read_only.check()
    expect(bold).to_be_disabled(timeout=settings.timeout_ms)
    target = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index}"]'
    ).first
    assert target.get_attribute("contenteditable") != "true"
    read_only_snapshot = read_json_test_id(page, "editor-model-snapshot")
    target.evaluate("element => element.focus()")
    page.keyboard.type("SHOULD-NOT-APPEAR")
    page.wait_for_timeout(100)
    assert read_json_test_id(page, "editor-model-snapshot") == read_only_snapshot
    read_only.uncheck()

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    page.set_viewport_size({"width": 1280, "height": 720})
    page.wait_for_timeout(150)
    laptop_topbar = element_layout(topbar)
    laptop_formatting = element_layout(formatting_bar)
    assert_no_horizontal_hiding("DOCX editor laptop top toolbar", laptop_topbar)
    assert_no_horizontal_hiding(
        "DOCX editor laptop formatting toolbar", laptop_formatting
    )
    laptop_page = assert_no_page_overflow(page, "DOCX editor laptop layout")
    page.screenshot(path=str(attempt_dir / "laptop.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "toolbars": {
            "top": topbar_metrics,
            "formatting": formatting_metrics,
        },
        "laptop": {
            "top": laptop_topbar,
            "formatting": laptop_formatting,
            "page": laptop_page,
        },
        "editing": {
            "nodeIndex": node_index,
            "marker": marker,
            "originalText": original_text,
            "committedText": committed_model["nodes"][node_index]["text"],
            "restoredText": restored_model["nodes"][node_index]["text"],
            "canRedoAfterUndo": history_state.get_attribute("data-can-redo"),
        },
        "selectionFormatting": {
            "text": range_text,
            "start": range_start,
            "end": range_start + len(range_text),
            "formattedRuns": formatted_model["nodes"][node_index]["runs"],
        },
        "paragraphKeys": {
            "splitLeft": split_model["nodes"][node_index]["text"],
            "splitRight": split_model["nodes"][node_index + 1]["text"],
            "merged": merged_model["nodes"][node_index]["text"],
        },
        "readOnly": {"modelUnchanged": True},
    }


def goto_docx_editor_ready(page: Page, settings: Settings) -> None:
    goto_ready(
        page,
        settings,
        "/docx-editor",
        "docx-editor",
        component_requires_ready=False,
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="editor-editable-state"]')?.dataset.state === 'editable'
          && document.querySelectorAll('[data-testid="editor-paragraph"][contenteditable="true"]').length >= 3""",
        timeout=settings.timeout_ms,
    )
    page.get_by_test_id("docx-editor-toolbar").wait_for(
        state="visible", timeout=settings.timeout_ms
    )


def docx_acceptance_paragraph(page: Page, settings: Settings):
    target = page.locator(
        '[data-testid="editor-paragraph"][contenteditable="true"]'
    ).filter(has_text="This paragraph is ready for editing").first
    target.wait_for(state="visible", timeout=settings.timeout_ms)
    node_index_raw = target.get_attribute("data-node-index")
    assert node_index_raw is not None
    return target, int(node_index_raw), target.inner_text()


def docx_selected_text_is_bold(
    snapshot: dict[str, object], node_index: int, selected_text: str
) -> bool:
    nodes = snapshot.get("nodes")
    if not isinstance(nodes, list) or node_index >= len(nodes):
        return False
    node = nodes[node_index]
    if not isinstance(node, dict):
        return False
    runs = node.get("runs")
    return isinstance(runs, list) and any(
        isinstance(run, dict)
        and run.get("text") == selected_text
        and isinstance(run.get("style"), dict)
        and run["style"].get("bold") is True
        for run in runs
    )


def wait_for_docx_selection_offsets(
    page: Page, start: int, end: int, timeout_ms: int
) -> None:
    page.wait_for_function(
        """expected => {
          const raw = document.querySelector('[data-testid="editor-selection-snapshot"]')?.textContent;
          if (!raw) return false;
          const range = JSON.parse(raw).activeTextRange;
          return range?.start?.offset === expected.start && range?.end?.offset === expected.end;
        }""",
        arg={"start": start, "end": end},
        timeout=timeout_ms,
    )


def run_ux_docx_002(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_docx_editor_ready(page, settings)
    target, node_index, original_text = docx_acceptance_paragraph(page, settings)
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)

    selected_text = "editing"
    start = original_text.index(selected_text)
    end = start + len(selected_text)
    set_dom_text_range(target, start, end)
    wait_for_docx_selection_offsets(page, start, end, settings.timeout_ms)

    bold = page.get_by_test_id("editor-bold")
    undo = page.get_by_test_id("editor-undo")
    redo = page.get_by_test_id("editor-redo")
    expect(bold).to_be_enabled(timeout=settings.timeout_ms)
    bold.click()
    formatted = wait_for_model_condition(
        page,
        lambda snapshot: docx_selected_text_is_bold(
            snapshot, node_index, selected_text
        ),
        settings.timeout_ms,
        "the real DOM selection to become bold",
    )
    expect(undo).to_be_enabled(timeout=settings.timeout_ms)

    undo.click()
    unformatted = wait_for_model_condition(
        page,
        lambda snapshot: not docx_selected_text_is_bold(
            snapshot, node_index, selected_text
        ),
        settings.timeout_ms,
        "the selected formatting undo",
    )
    expect(redo).to_be_enabled(timeout=settings.timeout_ms)
    wait_for_docx_selection_offsets(page, start, end, settings.timeout_ms)

    redo.click()
    redone = wait_for_model_condition(
        page,
        lambda snapshot: docx_selected_text_is_bold(
            snapshot, node_index, selected_text
        ),
        settings.timeout_ms,
        "the selected formatting redo",
    )
    wait_for_docx_selection_offsets(page, start, end, settings.timeout_ms)
    assert target.inner_text() == original_text

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "nodeIndex": node_index,
        "selectedText": selected_text,
        "selection": {"start": start, "end": end},
        "formattedRuns": formatted["nodes"][node_index]["runs"],
        "undoRuns": unformatted["nodes"][node_index]["runs"],
        "redoRuns": redone["nodes"][node_index]["runs"],
        "textUnchanged": True,
    }


def run_ux_docx_003(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_docx_editor_ready(page, settings)
    target, node_index, original_text = docx_acceptance_paragraph(page, settings)
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)
    model_before = read_json_test_id(page, "editor-model-snapshot")
    nodes_before = model_before.get("nodes")
    assert isinstance(nodes_before, list)

    split_offset = original_text.index(".") + 1
    set_dom_text_range(target, split_offset, split_offset)
    page.keyboard.press("Enter")

    def paragraph_is_split(snapshot: dict[str, object]) -> bool:
        nodes = snapshot.get("nodes")
        if not isinstance(nodes, list) or len(nodes) != len(nodes_before) + 1:
            return False
        left = nodes[node_index]
        right = nodes[node_index + 1]
        return (
            isinstance(left, dict)
            and isinstance(right, dict)
            and left.get("text") == original_text[:split_offset]
            and right.get("text") == original_text[split_offset:]
        )

    split_model = wait_for_model_condition(
        page,
        paragraph_is_split,
        settings.timeout_ms,
        "Enter to split the paragraph at the real caret",
    )
    second_paragraph = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index + 1}"]'
    ).first
    second_paragraph.wait_for(state="visible", timeout=settings.timeout_ms)
    set_dom_text_range(second_paragraph, 0, 0)
    page.keyboard.press("Backspace")
    merged_model = wait_for_model_condition(
        page,
        lambda snapshot: (
            isinstance(snapshot.get("nodes"), list)
            and len(snapshot["nodes"]) == len(nodes_before)
            and isinstance(snapshot["nodes"][node_index], dict)
            and snapshot["nodes"][node_index].get("text") == original_text
        ),
        settings.timeout_ms,
        "Backspace to merge the split paragraphs",
    )

    read_only = page.get_by_test_id("editor-readonly")
    bold = page.get_by_test_id("editor-bold")
    read_only.check()
    expect(bold).to_be_disabled(timeout=settings.timeout_ms)
    target = page.locator(
        f'[data-testid="editor-paragraph"][data-node-index="{node_index}"]'
    ).first
    assert target.get_attribute("contenteditable") != "true"
    read_only_snapshot = read_json_test_id(page, "editor-model-snapshot")
    target.evaluate("element => element.focus()")
    page.keyboard.type("SHOULD-NOT-APPEAR")
    page.keyboard.press("Enter")
    page.keyboard.press("Backspace")
    page.keyboard.press("Control+b")
    page.wait_for_timeout(150)
    assert read_json_test_id(page, "editor-model-snapshot") == read_only_snapshot

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "split": {
            "left": split_model["nodes"][node_index]["text"],
            "right": split_model["nodes"][node_index + 1]["text"],
        },
        "merged": merged_model["nodes"][node_index]["text"],
        "readOnly": {
            "contentEditable": target.get_attribute("contenteditable"),
            "toolbarDisabled": bold.is_disabled(),
            "modelUnchanged": True,
        },
    }


def docx_first_image(
    snapshot: dict[str, object], minimum_width: int = 0
) -> dict[str, object] | None:
    nodes = snapshot.get("nodes")
    if not isinstance(nodes, list):
        return None
    for node_index, node in enumerate(nodes):
        if not isinstance(node, dict) or node.get("type") != "paragraph":
            continue
        runs = node.get("runs")
        if not isinstance(runs, list):
            continue
        for child_index, run in enumerate(runs):
            if (
                isinstance(run, dict)
                and run.get("type") == "image"
                and int(run.get("widthPx") or 0) >= minimum_width
            ):
                return {
                    **run,
                    "nodeIndex": node_index,
                    "childIndex": child_index,
                }
    return None


def run_ux_docx_004(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_docx_editor_ready(page, settings)
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)

    initial_table = page.locator(
        '[data-docx-table-host="true"]'
    ).filter(has_text="Launch readiness").first
    initial_table.wait_for(state="visible", timeout=settings.timeout_ms)
    row_span_cell = initial_table.locator('td[rowspan="2"]').first
    expect(row_span_cell).to_contain_text(
        "Launch readiness", timeout=settings.timeout_ms
    )
    row_span_value = row_span_cell.get_attribute("rowspan")
    assert row_span_value == "2"
    initial_rows = initial_table.locator(
        "table.docx-table-host-table > tbody > tr"
    )
    assert initial_rows.count() == 2
    assert initial_rows.nth(0).locator(":scope > td").count() == 3
    assert initial_rows.nth(1).locator(":scope > td").count() == 2, (
        "vMergeContinuation rendered a duplicate table cell"
    )
    baseline = read_json_test_id(page, "editor-model-snapshot")
    baseline_table_count = int(baseline.get("tableCount") or 0)
    baseline_image_count = int(baseline.get("imageCount") or 0)
    assert baseline_table_count >= 1

    acceptance = page.locator("details.acceptance-panel")
    if acceptance.get_attribute("open") is None:
        acceptance.locator("summary").click()
    insert_table = page.get_by_test_id("editor-test-insert-table")
    insert_table.click()
    inserted = wait_for_model_condition(
        page,
        lambda snapshot: (
            int(snapshot.get("tableCount") or 0) == baseline_table_count + 1
            and isinstance(snapshot.get("nodes"), list)
            and any(
                isinstance(node, dict)
                and node.get("type") == "table"
                and isinstance(node.get("rows"), list)
                and len(node["rows"]) >= 3
                and all(
                    isinstance(row, dict)
                    and isinstance(row.get("cells"), list)
                    and len(row["cells"]) >= 3
                    for row in node["rows"]
                )
                for node in snapshot["nodes"]
            )
        ),
        settings.timeout_ms,
        "the public table command to insert and extend a table",
    )
    inserted_nodes = inserted["nodes"]
    inserted_table_index = max(
        index
        for index, node in enumerate(inserted_nodes)
        if isinstance(node, dict) and node.get("type") == "table"
    )
    cell_editor = page.locator(
        f'[data-docx-table-cell-paragraph-host="true"]'
        f'[data-docx-table-index="{inserted_table_index}"]'
        '[data-docx-table-row-index="0"]'
        '[data-docx-table-cell-index="0"]'
    ).first
    cell_editor.wait_for(state="visible", timeout=settings.timeout_ms)
    cell_text = "Quarterly owner"
    cell_editor.fill(cell_text)
    page.locator("h2").first.click()
    edited_table = wait_for_model_condition(
        page,
        lambda snapshot: (
            isinstance(snapshot.get("nodes"), list)
            and inserted_table_index < len(snapshot["nodes"])
            and isinstance(snapshot["nodes"][inserted_table_index], dict)
            and snapshot["nodes"][inserted_table_index].get("type") == "table"
            and cell_text
            in str(
                snapshot["nodes"][inserted_table_index]["rows"][0]["cells"][0].get(
                    "text", ""
                )
            )
        ),
        settings.timeout_ms,
        "the inserted table cell edit to reach the public model snapshot",
    )

    page.get_by_test_id("editor-test-insert-image").click()
    image_model = wait_for_model_condition(
        page,
        lambda snapshot: (
            int(snapshot.get("imageCount") or 0) == baseline_image_count + 1
            and docx_first_image(snapshot, 160) is not None
        ),
        settings.timeout_ms,
        "the public image command to create a positioned image",
    )
    inserted_image = docx_first_image(image_model, 160)
    assert inserted_image is not None
    floating = page.get_by_test_id("docx-floating-image").last
    floating.wait_for(state="visible", timeout=settings.timeout_ms)
    before_image_box = floating.bounding_box()
    assert before_image_box and before_image_box["width"] >= 150
    floating.click()
    wrap_toolbar = page.get_by_test_id("docx-image-wrap-toolbar")
    wrap_toolbar.wait_for(state="visible", timeout=settings.timeout_ms)
    resize_handle = floating.locator(
        '.docx-image-resize-handle--se[data-testid="docx-image-resize-handle"]'
    )
    handle_box = resize_handle.bounding_box()
    assert handle_box
    page.mouse.move(
        handle_box["x"] + handle_box["width"] / 2,
        handle_box["y"] + handle_box["height"] / 2,
    )
    page.mouse.down()
    page.mouse.move(
        handle_box["x"] + handle_box["width"] / 2 + 28,
        handle_box["y"] + handle_box["height"] / 2 + 18,
        steps=4,
    )
    page.mouse.up()
    resized_model = wait_for_model_condition(
        page,
        lambda snapshot: (
            (image := docx_first_image(snapshot, 180)) is not None
            and int(image.get("heightPx") or 0) >= 100
        ),
        settings.timeout_ms,
        "the visible image resize handle to update model dimensions",
    )
    resized_image = docx_first_image(resized_model, 180)
    assert resized_image is not None

    page.get_by_test_id("docx-image-wrap-front").click()
    front_model = wait_for_model_condition(
        page,
        lambda snapshot: (
            (image := docx_first_image(snapshot, 180)) is not None
            and isinstance(image.get("floating"), dict)
            and image["floating"].get("wrapType") == "none"
            and image["floating"].get("behindDocument") is False
            and int(image["floating"].get("zIndex") or 0) >= 1
        ),
        settings.timeout_ms,
        "the image wrap mode to change to in front of text",
    )
    page.get_by_test_id("docx-image-wrap-square").click()
    square_model = wait_for_model_condition(
        page,
        lambda snapshot: (
            (image := docx_first_image(snapshot, 180)) is not None
            and isinstance(image.get("floating"), dict)
            and image["floating"].get("wrapType") == "square"
        ),
        settings.timeout_ms,
        "the image wrap mode to return to square",
    )

    image_node_index = floating.get_attribute("data-image-node-index")
    image_child_index = floating.get_attribute("data-image-child-index")
    assert image_node_index is not None and image_child_index is not None
    wrapped_image = page.locator(
        f'[data-docx-paragraph-node-index="{image_node_index}"] '
        f'[data-docx-image-child-index="{image_child_index}"]'
        '[data-docx-image-layout="wrapped"]'
    ).first
    wrapped_image.wait_for(state="visible", timeout=settings.timeout_ms)
    page.wait_for_timeout(100)
    wrapped_box = wrapped_image.bounding_box()
    square_overlay_box = floating.bounding_box()
    assert wrapped_box and square_overlay_box
    for edge in ("x", "y", "width", "height"):
        assert abs(wrapped_box[edge] - square_overlay_box[edge]) <= 2, (
            "DOCX wrapped image controls are detached from the in-flow image: "
            f"edge={edge}, image={wrapped_box}, overlay={square_overlay_box}"
        )

    wrap_geometry = wrapped_image.evaluate(
        r"""image => {
          const imageRect = image.getBoundingClientRect();
          const paragraph = image.closest('[data-docx-paragraph-host="true"]');
          if (!paragraph) return null;
          const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
          const textRects = [];
          while (walker.nextNode()) {
            const node = walker.currentNode;
            if (!(node.textContent || '').trim()) continue;
            const range = document.createRange();
            range.selectNodeContents(node);
            for (const rect of range.getClientRects()) {
              if (rect.width > 0 && rect.height > 0) {
                textRects.push({
                  left: rect.left,
                  right: rect.right,
                  top: rect.top,
                  bottom: rect.bottom,
                });
              }
            }
          }
          const verticallyOverlaps = rect =>
            rect.top < imageRect.bottom - 1 && rect.bottom > imageRect.top + 1;
          const intersects = rect => verticallyOverlaps(rect) &&
            rect.left < imageRect.right - 1 && rect.right > imageRect.left + 1;
          return {
            image: {
              left: imageRect.left,
              right: imageRect.right,
              top: imageRect.top,
              bottom: imageRect.bottom,
            },
            collidingTextRects: textRects.filter(intersects),
            besideTextRects: textRects.filter(rect =>
              verticallyOverlaps(rect) && rect.left >= imageRect.right - 1
            ),
          };
        }"""
    )
    assert wrap_geometry is not None
    assert not wrap_geometry["collidingTextRects"], (
        "DOCX square-wrap image still covers paragraph text: "
        + json.dumps(wrap_geometry, ensure_ascii=False)
    )
    assert wrap_geometry["besideTextRects"], (
        "DOCX square-wrap image did not leave usable text beside the image: "
        + json.dumps(wrap_geometry, ensure_ascii=False)
    )

    title_paragraph = page.locator(
        '[data-docx-paragraph-node-index="0"]'
    ).first
    title_box = title_paragraph.bounding_box()
    assert title_box
    title_overlaps_image = not (
        title_box["x"] + title_box["width"] <= wrapped_box["x"] or
        wrapped_box["x"] + wrapped_box["width"] <= title_box["x"] or
        title_box["y"] + title_box["height"] <= wrapped_box["y"] or
        wrapped_box["y"] + wrapped_box["height"] <= title_box["y"]
    )
    assert not title_overlaps_image, (
        f"DOCX wrapped image covers the document title: title={title_box}, "
        f"image={wrapped_box}"
    )

    page_surface = page.get_by_test_id("docx-page").first
    before_page_box = page_surface.bounding_box()
    assert before_page_box and before_page_box["width"] > 0
    zoom = page.locator(
        '[data-testid="docx-editor-toolbar"] select[title="Zoom"]'
    )
    zoom.select_option("130")
    page.wait_for_function(
        """minimum => document.querySelector('[data-testid="docx-page"]')?.getBoundingClientRect().width >= minimum""",
        arg=before_page_box["width"] * 1.2,
        timeout=settings.timeout_ms,
    )
    after_page_box = page_surface.bounding_box()
    assert after_page_box and after_page_box["width"] > before_page_box["width"] * 1.2

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    page.set_viewport_size({"width": 1280, "height": 720})
    page.wait_for_timeout(150)
    topbar = page.locator(
        '[data-testid="docx-editor-toolbar"] .docx-toolbar-topbar'
    )
    formatting = page.locator(
        '[data-testid="docx-editor-toolbar"] .docx-toolbar'
    )
    laptop_topbar = element_layout(topbar)
    laptop_formatting = element_layout(formatting)
    assert_no_horizontal_hiding("DOCX editor laptop top toolbar", laptop_topbar)
    assert_no_horizontal_hiding(
        "DOCX editor laptop formatting toolbar", laptop_formatting
    )
    laptop_page = assert_no_page_overflow(page, "DOCX editor laptop layout")
    page.screenshot(path=str(attempt_dir / "laptop.png"), full_page=True)

    return {
        "expectedDownloads": 0,
        "rowSpan": row_span_value,
        "insertedTable": {
            "nodeIndex": inserted_table_index,
            "rows": len(edited_table["nodes"][inserted_table_index]["rows"]),
            "columns": len(
                edited_table["nodes"][inserted_table_index]["rows"][0]["cells"]
            ),
            "cellText": edited_table["nodes"][inserted_table_index]["rows"][0][
                "cells"
            ][0]["text"],
        },
        "image": {
            "initial": inserted_image,
            "beforeBox": before_image_box,
            "resized": resized_image,
            "front": docx_first_image(front_model, 180),
            "square": docx_first_image(square_model, 180),
            "wrappedBox": wrapped_box,
            "overlayBox": square_overlay_box,
            "wrapGeometry": wrap_geometry,
        },
        "zoom": {
            "value": zoom.input_value(),
            "before": before_page_box["width"],
            "after": after_page_box["width"],
        },
        "laptop": {
            "top": laptop_topbar,
            "formatting": laptop_formatting,
            "page": laptop_page,
        },
    }


def xlsx_grid_metrics(page: Page) -> dict[str, object]:
    return page.get_by_test_id("xlsx-grid").evaluate(
        """grid => {
          const space = grid.querySelector('[data-testid="xlsx-grid-scroll-space"]');
          const canvas = grid.querySelector('.xlsx-grid__body');
          if (!space || !canvas) return null;
          const gridRect = grid.getBoundingClientRect();
          const spaceRect = space.getBoundingClientRect();
          const canvasRect = canvas.getBoundingClientRect();
          return {
            clientWidth: grid.clientWidth,
            clientHeight: grid.clientHeight,
            scrollWidth: grid.scrollWidth,
            scrollHeight: grid.scrollHeight,
            gridRect: {width: gridRect.width, height: gridRect.height},
            spaceRect: {width: spaceRect.width, height: spaceRect.height},
            bodyCanvas: {
              width: canvasRect.width,
              height: canvasRect.height,
              pixelWidth: canvas.width,
              pixelHeight: canvas.height,
            },
          };
        }"""
    )


def xlsx_canvas_coverage(page: Page) -> dict[str, object]:
    result = page.locator(".xlsx-grid__body").evaluate(
        """canvas => {
          const context = canvas.getContext('2d');
          if (!context || canvas.width < 4 || canvas.height < 4) return null;
          const startX = Math.floor(canvas.width * 0.6);
          const startY = Math.floor(canvas.height * 0.4);
          const width = canvas.width - startX;
          const height = canvas.height - startY;
          const pixels = context.getImageData(startX, startY, width, height).data;
          const colors = new Map();
          for (let index = 0; index < pixels.length; index += 4) {
            const key = `${pixels[index]},${pixels[index + 1]},${pixels[index + 2]},${pixels[index + 3]}`;
            colors.set(key, (colors.get(key) || 0) + 1);
          }
          let dominant = '';
          let dominantCount = 0;
          for (const [color, count] of colors.entries()) {
            if (count > dominantCount) {
              dominant = color;
              dominantCount = count;
            }
          }
          const isDifferent = (x, y) => {
            const offset = (y * width + x) * 4;
            return `${pixels[offset]},${pixels[offset + 1]},${pixels[offset + 2]},${pixels[offset + 3]}` !== dominant;
          };
          const verticalCandidates = [];
          for (let x = 0; x < width; x += 1) {
            let different = 0;
            for (let y = 0; y < height; y += 1) {
              if (isDifferent(x, y)) different += 1;
            }
            if (different / height > 0.7) verticalCandidates.push(x);
          }
          const horizontalCandidates = [];
          for (let y = 0; y < height; y += 1) {
            let different = 0;
            for (let x = 0; x < width; x += 1) {
              if (isDifferent(x, y)) different += 1;
            }
            if (different / width > 0.7) horizontalCandidates.push(y);
          }
          const clusters = (positions) => positions.reduce((count, position, index) =>
            count + (index === 0 || position > positions[index - 1] + 1 ? 1 : 0), 0);
          const total = width * height;
          return {
            sample: {startX, startY, width, height},
            dominantColor: dominant,
            distinctColors: colors.size,
            nonDominantRatio: total > 0 ? (total - dominantCount) / total : 0,
            verticalGridLineClusters: clusters(verticalCandidates),
            horizontalGridLineClusters: clusters(horizontalCandidates),
          };
        }"""
    )
    assert result is not None, "XLSX body canvas has no readable pixels"
    return result


def xlsx_canvas_region_hash(page: Page, top: int, height: int) -> str:
    return page.locator(".xlsx-grid__body").evaluate(
        """(canvas, region) => {
          const ratio = window.devicePixelRatio || 1;
          const y = Math.max(0, Math.floor(region.top * ratio));
          const h = Math.max(1, Math.min(canvas.height - y, Math.floor(region.height * ratio)));
          const copy = document.createElement('canvas');
          copy.width = canvas.width;
          copy.height = h;
          const context = copy.getContext('2d');
          if (!context) return '';
          context.drawImage(canvas, 0, y, canvas.width, h, 0, 0, canvas.width, h);
          const encoded = copy.toDataURL('image/png');
          let hash = 2166136261;
          for (let index = 0; index < encoded.length; index += 1) {
            hash ^= encoded.charCodeAt(index);
            hash = Math.imul(hash, 16777619);
          }
          return (hash >>> 0).toString(16).padStart(8, '0');
        }""",
        {"top": top, "height": height},
    )


def xlsx_first_column_boundary(page: Page) -> int:
    raw = page.get_by_test_id("xlsx-grid").get_attribute(
        "data-first-column-width"
    )
    boundary = round(float(raw or "-1"))
    assert boundary > 20, f"could not detect first XLSX column boundary: {boundary}"
    return boundary


def run_xlsx_viewer(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_ready(page, settings, "/xlsx-viewer", "xlsx-viewer")
    grid = page.get_by_test_id("xlsx-grid")
    grid.wait_for(state="visible", timeout=settings.timeout_ms)
    assert int(grid.get_attribute("data-display-row-count") or "0") >= 200
    assert int(grid.get_attribute("data-display-column-count") or "0") >= 50
    toolbar = page.get_by_test_id("xlsx-toolbar")
    formatting_toolbar = page.get_by_test_id("xlsx-format-toolbar")
    toolbar.wait_for(state="visible", timeout=settings.timeout_ms)
    formatting_toolbar.wait_for(state="visible", timeout=settings.timeout_ms)
    page.wait_for_timeout(300)
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)

    toolbar_metrics = element_layout(toolbar)
    formatting_metrics = element_layout(formatting_toolbar)
    assert_no_horizontal_hiding("XLSX file toolbar", toolbar_metrics)
    assert_no_horizontal_hiding("XLSX formatting toolbar", formatting_metrics)

    grid_metrics = xlsx_grid_metrics(page)
    assert grid_metrics is not None
    assert grid_metrics["scrollWidth"] > grid_metrics["clientWidth"] + 100, grid_metrics
    assert grid_metrics["scrollHeight"] > grid_metrics["clientHeight"] + 100, grid_metrics
    body_canvas = grid_metrics["bodyCanvas"]
    assert body_canvas["width"] >= grid_metrics["clientWidth"] - 50, grid_metrics
    assert body_canvas["height"] >= grid_metrics["clientHeight"] - 26, grid_metrics

    coverage = xlsx_canvas_coverage(page)
    assert coverage["nonDominantRatio"] > 0.01, coverage
    assert coverage["verticalGridLineClusters"] >= 2, coverage
    assert coverage["horizontalGridLineClusters"] >= 4, coverage

    name_box = page.get_by_test_id("xlsx-name-box")
    formula_input = page.get_by_test_id("xlsx-formula-input")
    selection_overlay = page.locator(".xlsx-selection-overlay")

    # Select A1 through the public grid, then navigate to the known A3 = ARPU cell.
    grid.click(position={"x": 60, "y": 36})
    expect(name_box).to_have_value("A1", timeout=settings.timeout_ms)
    page.keyboard.press("ArrowDown")
    page.keyboard.press("ArrowDown")
    expect(name_box).to_have_value("A3", timeout=settings.timeout_ms)
    expect(formula_input).to_have_value("ARPU", timeout=settings.timeout_ms)

    page.keyboard.press("Shift+ArrowRight")
    expect(selection_overlay).to_have_attribute(
        "data-selection-address", "A3:B3", timeout=settings.timeout_ms
    )
    page.keyboard.press("ArrowLeft")
    expect(name_box).to_have_value("A3", timeout=settings.timeout_ms)
    expect(formula_input).to_have_value("ARPU", timeout=settings.timeout_ms)

    page.keyboard.press("Enter")
    cell_editor = page.get_by_test_id("xlsx-cell-editor")
    cell_editor.wait_for(state="visible", timeout=settings.timeout_ms)
    expect(cell_editor).to_have_value("ARPU", timeout=settings.timeout_ms)
    edited_value = "ARPU-UX-PARITY"
    cell_editor.fill(edited_value)
    name_box.click()
    cell_editor.wait_for(state="hidden", timeout=settings.timeout_ms)
    expect(formula_input).to_have_value(edited_value, timeout=settings.timeout_ms)

    undo = page.get_by_test_id("xlsx-undo")
    expect(undo).to_be_enabled(timeout=settings.timeout_ms)
    undo.click()
    expect(formula_input).to_have_value("ARPU", timeout=settings.timeout_ms)
    restored_value = formula_input.input_value()

    resize_before = xlsx_grid_metrics(page)
    assert resize_before is not None
    first_boundary = xlsx_first_column_boundary(page)
    grid_box = grid.bounding_box()
    assert grid_box
    resize_x = grid_box["x"] + 48 + first_boundary
    resize_y = grid_box["y"] + 12
    page.mouse.move(resize_x, resize_y)
    page.mouse.down()
    page.mouse.move(resize_x + 36, resize_y, steps=4)
    page.mouse.up()
    page.wait_for_function(
        """minimum => {
          const space = document.querySelector('[data-testid="xlsx-grid-scroll-space"]');
          return space && space.getBoundingClientRect().width >= minimum;
        }""",
        arg=resize_before["spaceRect"]["width"] + 20,
        timeout=settings.timeout_ms,
    )
    resize_after = xlsx_grid_metrics(page)
    assert resize_after is not None

    search_toggle = page.get_by_test_id("xlsx-search-toggle")
    search_toggle.click()
    search_panel = page.get_by_test_id("xlsx-search-panel")
    search_panel.wait_for(state="visible", timeout=settings.timeout_ms)
    search_input = page.get_by_test_id("xlsx-search-input")
    search_input.fill("ARPU")
    expect(search_panel).to_contain_text("A3", timeout=settings.timeout_ms)
    search_panel.locator("button").filter(has_text="A3").first.click()
    expect(name_box).to_have_value("A3", timeout=settings.timeout_ms)
    search_input.press("Escape")

    read_only = page.get_by_test_id("xlsx-ribbon-read-only")
    read_only.check()
    expect(read_only).to_be_checked(timeout=settings.timeout_ms)
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-viewer"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )
    page.get_by_test_id("xlsx-grid").click(position={"x": 60, "y": 36})
    page.keyboard.press("Enter")
    assert page.get_by_test_id("xlsx-cell-editor").count() == 0, (
        "read-only mode opened a cell editor"
    )
    read_only = page.get_by_test_id("xlsx-ribbon-read-only")
    read_only.uncheck()
    expect(read_only).not_to_be_checked(timeout=settings.timeout_ms)
    page.wait_for_function(
        """() => document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-viewer"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )

    # The sales fixture has a real frozen first row. Its first body row must
    # remain pixel-identical while rows below it change after scrolling.
    sample = page.get_by_test_id("xlsx-sample-select")
    sample.select_option("sales-table.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('sales-table.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-viewer"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )
    grid = page.get_by_test_id("xlsx-grid")
    grid.wait_for(state="visible", timeout=settings.timeout_ms)
    page.wait_for_timeout(300)
    frozen_before = xlsx_canvas_region_hash(page, 0, 23)
    scrolling_before = xlsx_canvas_region_hash(page, 28, 96)
    grid.evaluate(
        """element => {
          element.scrollTop = Math.min(360, element.scrollHeight - element.clientHeight);
          element.dispatchEvent(new Event('scroll'));
        }"""
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="xlsx-grid"]')?.scrollTop > 100""",
        timeout=settings.timeout_ms,
    )
    page.wait_for_timeout(250)
    frozen_after = xlsx_canvas_region_hash(page, 0, 23)
    scrolling_after = xlsx_canvas_region_hash(page, 28, 96)
    assert frozen_before == frozen_after, {
        "frozenBefore": frozen_before,
        "frozenAfter": frozen_after,
    }
    assert scrolling_before != scrolling_after, {
        "scrollingBefore": scrolling_before,
        "scrollingAfter": scrolling_after,
    }

    sample.select_option("charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('charts-images.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && Number(document.querySelector('[data-testid="xlsx-chart-overlay"]')?.dataset.chartCount || 0) > 0
          && document.querySelectorAll('[data-testid="xlsx-image-item"]').length > 0""",
        timeout=settings.timeout_ms,
    )
    chart_item = page.get_by_test_id("xlsx-chart-item").first
    image_item = page.get_by_test_id("xlsx-image-item").first
    chart_item.wait_for(state="visible", timeout=settings.timeout_ms)
    image_item.wait_for(state="visible", timeout=settings.timeout_ms)
    page.locator('[data-testid="xlsx-image-item"] img[data-image-state="ready"]').first.wait_for(
        state="visible", timeout=settings.timeout_ms
    )
    chart_before = chart_item.bounding_box()
    image_before = image_item.bounding_box()
    assert chart_before and image_before
    grid = page.get_by_test_id("xlsx-grid")
    grid.evaluate(
        """element => {
          element.scrollLeft = Math.min(120, element.scrollWidth - element.clientWidth);
          element.scrollTop = Math.min(120, element.scrollHeight - element.clientHeight);
          element.dispatchEvent(new Event('scroll'));
        }"""
    )
    page.wait_for_timeout(250)
    chart_after = chart_item.bounding_box()
    image_after = image_item.bounding_box()
    assert chart_after and image_after
    assert (
        abs(chart_after["x"] - chart_before["x"]) > 40
        or abs(chart_after["y"] - chart_before["y"]) > 40
    ), {"before": chart_before, "after": chart_after}
    assert (
        abs(image_after["x"] - image_before["x"]) > 40
        or abs(image_after["y"] - image_before["y"]) > 40
    ), {"before": image_before, "after": image_after}
    page.screenshot(path=str(attempt_dir / "charts-images.png"), full_page=True)

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    page.set_viewport_size({"width": 1280, "height": 720})
    page.wait_for_timeout(150)
    laptop_toolbar = element_layout(toolbar)
    laptop_formatting = element_layout(formatting_toolbar)
    assert_no_horizontal_hiding("XLSX laptop file toolbar", laptop_toolbar)
    assert_no_horizontal_hiding(
        "XLSX laptop formatting toolbar", laptop_formatting
    )
    laptop_page = assert_no_page_overflow(page, "XLSX laptop layout")
    page.screenshot(path=str(attempt_dir / "laptop.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "toolbars": {
            "file": toolbar_metrics,
            "formatting": formatting_metrics,
        },
        "laptop": {
            "file": laptop_toolbar,
            "formatting": laptop_formatting,
            "page": laptop_page,
        },
        "grid": grid_metrics,
        "canvasCoverage": coverage,
        "selection": {
            "singleCell": "A3",
            "range": "A3:B3",
            "formulaValue": "ARPU",
        },
        "editing": {
            "editedValue": edited_value,
            "restoredValue": restored_value,
        },
        "search": {"query": "ARPU", "selected": "A3"},
        "columnResize": {
            "detectedBoundary": first_boundary,
            "beforeWidth": resize_before["spaceRect"]["width"],
            "afterWidth": resize_after["spaceRect"]["width"],
        },
        "readOnly": {"editorBlocked": True},
        "freeze": {
            "topRowBefore": frozen_before,
            "topRowAfter": frozen_after,
            "scrollingBefore": scrolling_before,
            "scrollingAfter": scrolling_after,
        },
        "drawings": {
            "chartBefore": chart_before,
            "chartAfter": chart_after,
            "imageBefore": image_before,
            "imageAfter": image_after,
        },
    }


def goto_xlsx_ready(page: Page, settings: Settings) -> None:
    goto_ready(page, settings, "/xlsx-viewer", "xlsx-viewer")
    grid = page.get_by_test_id("xlsx-grid")
    grid.wait_for(state="visible", timeout=settings.timeout_ms)
    page.wait_for_function(
        """() => {
          const grid = document.querySelector('[data-testid="xlsx-grid"]');
          return Number(grid?.dataset.displayRowCount || 0) >= 200
            && Number(grid?.dataset.displayColumnCount || 0) >= 50;
        }""",
        timeout=settings.timeout_ms,
    )


def reset_xlsx_scroll(page: Page) -> None:
    page.get_by_test_id("xlsx-grid").evaluate(
        """element => {
          element.scrollLeft = 0;
          element.scrollTop = 0;
          element.dispatchEvent(new Event('scroll'));
        }"""
    )


def xlsx_cell_position(page: Page, row: int, col: int) -> dict[str, float]:
    grid = page.get_by_test_id("xlsx-grid")
    first_width = float(grid.get_attribute("data-first-column-width") or "80")
    first_height = float(grid.get_attribute("data-first-row-height") or "24")
    return {
        "x": 48 + first_width * col + first_width / 2,
        "y": 24 + first_height * row + first_height / 2,
    }


def select_xlsx_cell(page: Page, row: int, col: int) -> None:
    reset_xlsx_scroll(page)
    page.get_by_test_id("xlsx-grid").click(position=xlsx_cell_position(page, row, col))


def run_ux_xlsx_001(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_xlsx_ready(page, settings)
    grid = page.get_by_test_id("xlsx-grid")
    toolbar = page.get_by_test_id("xlsx-toolbar")
    formatting = page.get_by_test_id("xlsx-format-toolbar")
    toolbar.wait_for(state="visible", timeout=settings.timeout_ms)
    formatting.wait_for(state="visible", timeout=settings.timeout_ms)
    page.wait_for_timeout(300)
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)

    toolbar_metrics = element_layout(toolbar)
    formatting_metrics = element_layout(formatting)
    assert_no_horizontal_hiding("XLSX file toolbar", toolbar_metrics)
    assert_no_horizontal_hiding("XLSX formatting toolbar", formatting_metrics)

    metrics = xlsx_grid_metrics(page)
    assert metrics is not None
    assert metrics["scrollWidth"] > metrics["clientWidth"] + 100, metrics
    assert metrics["scrollHeight"] > metrics["clientHeight"] + 100, metrics
    assert int(grid.get_attribute("data-display-row-count") or "0") >= 200
    assert int(grid.get_attribute("data-display-column-count") or "0") >= 50
    coverage = xlsx_canvas_coverage(page)
    assert coverage["nonDominantRatio"] > 0.01, coverage
    assert coverage["distinctColors"] >= 4, coverage
    assert coverage["verticalGridLineClusters"] >= 2, coverage
    assert coverage["horizontalGridLineClusters"] >= 4, coverage

    name_box = page.get_by_test_id("xlsx-name-box")
    formula = page.get_by_test_id("xlsx-formula-input")
    select_xlsx_cell(page, 0, 0)
    expect(name_box).to_have_value("A1", timeout=settings.timeout_ms)
    expect(formula).to_have_value("Driver", timeout=settings.timeout_ms)
    page.keyboard.press("ArrowDown")
    page.keyboard.press("ArrowDown")
    expect(name_box).to_have_value("A3", timeout=settings.timeout_ms)
    expect(formula).to_have_value("ARPU", timeout=settings.timeout_ms)

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    page.set_viewport_size({"width": 1280, "height": 720})
    page.wait_for_timeout(150)
    laptop_toolbar = element_layout(toolbar)
    laptop_formatting = element_layout(formatting)
    assert_no_horizontal_hiding("XLSX laptop file toolbar", laptop_toolbar)
    assert_no_horizontal_hiding(
        "XLSX laptop formatting toolbar", laptop_formatting
    )
    laptop_page = assert_no_page_overflow(page, "XLSX laptop layout")
    page.screenshot(path=str(attempt_dir / "laptop.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "grid": metrics,
        "canvasCoverage": coverage,
        "activeCell": {"address": name_box.input_value(), "value": formula.input_value()},
        "toolbars": {"file": toolbar_metrics, "formatting": formatting_metrics},
        "laptop": {
            "file": laptop_toolbar,
            "formatting": laptop_formatting,
            "page": laptop_page,
        },
    }


def run_ux_xlsx_002(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_xlsx_ready(page, settings)
    grid = page.get_by_test_id("xlsx-grid")
    overlay = page.locator(".xlsx-selection-overlay")
    name_box = page.get_by_test_id("xlsx-name-box")
    formula = page.get_by_test_id("xlsx-formula-input")
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)
    reset_xlsx_scroll(page)

    grid_box = grid.bounding_box()
    assert grid_box
    first_width = xlsx_first_column_boundary(page)
    first_height = float(grid.get_attribute("data-first-row-height") or "24")
    page.mouse.click(
        grid_box["x"] + 48 + first_width / 2,
        grid_box["y"] + 12,
    )
    expect(overlay).to_have_attribute(
        "data-selection-address", "A1:A200", timeout=settings.timeout_ms
    )
    column_selection = overlay.get_attribute("data-selection-address")

    page.mouse.click(
        grid_box["x"] + 24,
        grid_box["y"] + 24 + 2 * first_height + first_height / 2,
    )
    expect(overlay).to_have_attribute(
        "data-selection-address", "A3:AX3", timeout=settings.timeout_ms
    )
    row_selection = overlay.get_attribute("data-selection-address")

    select_xlsx_cell(page, 2, 0)
    expect(name_box).to_have_value("A3", timeout=settings.timeout_ms)
    expect(formula).to_have_value("ARPU", timeout=settings.timeout_ms)
    page.keyboard.press("Shift+ArrowRight")
    expect(overlay).to_have_attribute(
        "data-selection-address", "A3:B3", timeout=settings.timeout_ms
    )
    range_selection = overlay.get_attribute("data-selection-address")
    page.keyboard.press("ArrowLeft")
    expect(name_box).to_have_value("A3", timeout=settings.timeout_ms)

    page.keyboard.press("Enter")
    editor = page.get_by_test_id("xlsx-cell-editor")
    editor.wait_for(state="visible", timeout=settings.timeout_ms)
    expect(editor).to_have_value("ARPU", timeout=settings.timeout_ms)
    editor.fill("CANCELLED-UX-PARITY")
    editor.press("Escape")
    editor.wait_for(state="hidden", timeout=settings.timeout_ms)
    expect(formula).to_have_value("ARPU", timeout=settings.timeout_ms)
    escape_cancelled = formula.input_value() == "ARPU"

    page.keyboard.press("Enter")
    editor = page.get_by_test_id("xlsx-cell-editor")
    editor.wait_for(state="visible", timeout=settings.timeout_ms)
    committed_value = "ARPU-COMMITTED"
    editor.fill(committed_value)
    editor.press("Enter")
    editor.wait_for(state="hidden", timeout=settings.timeout_ms)
    expect(formula).to_have_value(committed_value, timeout=settings.timeout_ms)

    undo = page.get_by_test_id("xlsx-undo")
    expect(undo).to_be_enabled(timeout=settings.timeout_ms)
    undo.click()
    expect(formula).to_have_value("ARPU", timeout=settings.timeout_ms)

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "columnSelection": column_selection,
        "rowSelection": row_selection,
        "rangeSelection": range_selection,
        "escapeCancelled": escape_cancelled,
        "enterCommitted": committed_value,
        "undoRestored": formula.input_value(),
    }


def run_ux_xlsx_003(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_xlsx_ready(page, settings)
    grid = page.get_by_test_id("xlsx-grid")
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)
    reset_xlsx_scroll(page)
    grid_box = grid.bounding_box()
    assert grid_box

    before_column_resize = xlsx_grid_metrics(page)
    assert before_column_resize is not None
    first_boundary = xlsx_first_column_boundary(page)
    column_x = grid_box["x"] + 48 + first_boundary
    column_y = grid_box["y"] + 12
    page.mouse.move(column_x, column_y)
    page.mouse.down()
    page.mouse.move(column_x + 36, column_y, steps=4)
    page.mouse.up()
    page.wait_for_function(
        """minimum => document.querySelector('[data-testid="xlsx-grid-scroll-space"]')?.getBoundingClientRect().width >= minimum""",
        arg=before_column_resize["spaceRect"]["width"] + 20,
        timeout=settings.timeout_ms,
    )
    after_column_resize = xlsx_grid_metrics(page)
    assert after_column_resize is not None

    before_row_resize = after_column_resize
    row_x = grid_box["x"] + 24
    row_y = grid_box["y"] + 24 + 24
    page.mouse.move(row_x, row_y)
    page.mouse.down()
    page.mouse.move(row_x, row_y + 18, steps=4)
    page.mouse.up()
    page.wait_for_function(
        """minimum => document.querySelector('[data-testid="xlsx-grid-scroll-space"]')?.getBoundingClientRect().height >= minimum""",
        arg=before_row_resize["spaceRect"]["height"] + 10,
        timeout=settings.timeout_ms,
    )
    after_row_resize = xlsx_grid_metrics(page)
    assert after_row_resize is not None

    sample = page.get_by_test_id("xlsx-sample-select")
    sample.select_option("sales-table.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('sales-table.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-viewer"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )
    grid = page.get_by_test_id("xlsx-grid")
    page.wait_for_timeout(300)
    frozen_before = xlsx_canvas_region_hash(page, 0, 23)
    scrolling_before = xlsx_canvas_region_hash(page, 28, 96)
    grid.evaluate(
        """element => {
          element.scrollTop = Math.min(360, element.scrollHeight - element.clientHeight);
          element.dispatchEvent(new Event('scroll'));
        }"""
    )
    page.wait_for_function(
        """() => document.querySelector('[data-testid="xlsx-grid"]')?.scrollTop > 100""",
        timeout=settings.timeout_ms,
    )
    page.wait_for_timeout(250)
    frozen_after = xlsx_canvas_region_hash(page, 0, 23)
    scrolling_after = xlsx_canvas_region_hash(page, 28, 96)
    assert frozen_before == frozen_after
    assert scrolling_before != scrolling_after

    sample.select_option("charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('charts-images.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && Number(document.querySelector('[data-testid="xlsx-chart-overlay"]')?.dataset.chartCount || 0) > 0
          && document.querySelectorAll('[data-testid="xlsx-image-item"]').length > 0""",
        timeout=settings.timeout_ms,
    )
    chart = page.get_by_test_id("xlsx-chart-item").first
    image = page.get_by_test_id("xlsx-image-item").first
    chart.wait_for(state="visible", timeout=settings.timeout_ms)
    image.wait_for(state="visible", timeout=settings.timeout_ms)
    page.locator(
        '[data-testid="xlsx-image-item"] img[data-image-state="ready"]'
    ).first.wait_for(state="visible", timeout=settings.timeout_ms)
    chart_before = chart.bounding_box()
    image_before = image.bounding_box()
    assert chart_before and image_before
    grid = page.get_by_test_id("xlsx-grid")
    grid.evaluate(
        """element => {
          element.scrollLeft = Math.min(120, element.scrollWidth - element.clientWidth);
          element.scrollTop = Math.min(120, element.scrollHeight - element.clientHeight);
          element.dispatchEvent(new Event('scroll'));
        }"""
    )
    page.wait_for_timeout(250)
    chart_after = chart.bounding_box()
    image_after = image.bounding_box()
    assert chart_after and image_after
    assert (
        abs(chart_after["x"] - chart_before["x"]) > 40
        or abs(chart_after["y"] - chart_before["y"]) > 40
    )
    assert (
        abs(image_after["x"] - image_before["x"]) > 40
        or abs(image_after["y"] - image_before["y"]) > 40
    )
    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "columnResize": {
            "before": before_column_resize["spaceRect"]["width"],
            "after": after_column_resize["spaceRect"]["width"],
        },
        "rowResize": {
            "before": before_row_resize["spaceRect"]["height"],
            "after": after_row_resize["spaceRect"]["height"],
        },
        "freeze": {
            "topBefore": frozen_before,
            "topAfter": frozen_after,
            "scrollingBefore": scrolling_before,
            "scrollingAfter": scrolling_after,
        },
        "drawings": {
            "chartBefore": chart_before,
            "chartAfter": chart_after,
            "imageBefore": image_before,
            "imageAfter": image_after,
        },
    }


def run_ux_xlsx_004(
    page: Page, attempt_dir: Path, settings: Settings
) -> dict[str, object]:
    goto_xlsx_ready(page, settings)
    grid = page.get_by_test_id("xlsx-grid")
    name_box = page.get_by_test_id("xlsx-name-box")
    formula = page.get_by_test_id("xlsx-formula-input")
    page.screenshot(path=str(attempt_dir / "before.png"), full_page=True)
    reset_xlsx_scroll(page)
    grid.click(position=xlsx_cell_position(page, 2, 1), button="right")
    menu = page.get_by_test_id("xlsx-context-menu")
    menu.wait_for(state="visible", timeout=settings.timeout_ms)
    expect(name_box).to_have_value("B3", timeout=settings.timeout_ms)
    original_value = formula.input_value()
    assert original_value, "right-click target B3 must contain a value"
    menu.get_by_role("menuitem", name="清除内容").click()
    expect(formula).to_have_value("", timeout=settings.timeout_ms)
    undo = page.get_by_test_id("xlsx-undo")
    expect(undo).to_be_enabled(timeout=settings.timeout_ms)
    undo.click()
    expect(formula).to_have_value(original_value, timeout=settings.timeout_ms)
    undo_restored = formula.input_value() == original_value

    search_toggle = page.get_by_test_id("xlsx-search-toggle")
    search_toggle.click()
    search_panel = page.get_by_test_id("xlsx-search-panel")
    search_panel.wait_for(state="visible", timeout=settings.timeout_ms)
    search_input = page.get_by_test_id("xlsx-search-input")
    search_input.fill("Operating Income")
    cross_sheet_result = search_panel.locator(
        "button.xlsx-toolbar__search-result"
    ).filter(has_text="P&L!").first
    cross_sheet_result.wait_for(state="visible", timeout=settings.timeout_ms)
    result_text = cross_sheet_result.inner_text()
    address_match = re.search(r"P&L!([A-Z]+\d+)", result_text)
    assert address_match, result_text
    expected_address = address_match.group(1)
    cross_sheet_result.click()
    selected_tab = page.locator(
        '[data-testid="xlsx-sheet-tab"][data-sheet-name="P&L"]'
    )
    expect(selected_tab).to_have_attribute(
        "aria-selected", "true", timeout=settings.timeout_ms
    )
    expect(name_box).to_have_value(expected_address, timeout=settings.timeout_ms)
    search_input.press("Escape")

    read_only = page.get_by_test_id("xlsx-ribbon-read-only")
    read_only.check()
    expect(read_only).to_be_checked(timeout=settings.timeout_ms)
    grid = page.get_by_test_id("xlsx-grid")
    reset_xlsx_scroll(page)
    grid.click(position=xlsx_cell_position(page, 0, 0), button="right")
    menu = page.get_by_test_id("xlsx-context-menu")
    menu.wait_for(state="visible", timeout=settings.timeout_ms)
    assert menu.get_by_role("menuitem", name="复制").count() == 1
    assert menu.get_by_role("menuitem", name="清除内容").count() == 0
    page.locator(".xlsx-contextmenu-backdrop").click(position={"x": 2, "y": 2})
    grid.click(position=xlsx_cell_position(page, 0, 0))
    page.keyboard.press("Enter")
    assert page.get_by_test_id("xlsx-cell-editor").count() == 0
    read_only.uncheck()
    expect(read_only).not_to_be_checked(timeout=settings.timeout_ms)

    sample = page.get_by_test_id("xlsx-sample-select")
    sample.select_option("sales-table.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('sales-table.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-viewer"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )
    grid = page.get_by_test_id("xlsx-grid")
    reset_xlsx_scroll(page)
    select_xlsx_cell(page, 1, 0)
    expect(formula).to_have_value("North", timeout=settings.timeout_ms)
    unsorted_first = formula.input_value()
    grid.click(position=xlsx_cell_position(page, 0, 0), button="right")
    menu = page.get_by_test_id("xlsx-context-menu")
    menu.wait_for(state="visible", timeout=settings.timeout_ms)
    expect(name_box).to_have_value("A1", timeout=settings.timeout_ms)
    sort_ascending = page.get_by_test_id("xlsx-sort-ascending")
    sort_ascending.wait_for(state="visible", timeout=settings.timeout_ms)
    sort_ascending.click()
    select_xlsx_cell(page, 1, 0)
    expect(formula).to_have_value("APAC", timeout=settings.timeout_ms)
    sorted_first = formula.input_value()
    assert sorted_first != unsorted_first

    sample.select_option("charts-images.xlsx")
    page.wait_for_function(
        """() => document.querySelector('[data-testid="loaded-file"]')?.textContent?.includes('charts-images.xlsx')
          && document.querySelector('[data-testid="page-status"]')?.dataset.state === 'ready'
          && document.querySelector('[data-testid="xlsx-viewer"]')?.dataset.state === 'ready'""",
        timeout=settings.timeout_ms,
    )
    chart_tab = page.locator(
        '[data-testid="xlsx-sheet-tab"][data-sheet-name="Revenue Chart"]'
    )
    chart_tab.wait_for(state="visible", timeout=settings.timeout_ms)
    chart_tab.click()
    expect(chart_tab).to_have_attribute(
        "aria-selected", "true", timeout=settings.timeout_ms
    )
    chart_sheet = page.get_by_test_id("xlsx-chartsheet")
    chart_sheet.wait_for(state="visible", timeout=settings.timeout_ms)
    expect(chart_sheet).to_have_attribute(
        "data-state", "ready", timeout=settings.timeout_ms
    )
    chart = page.get_by_test_id("xlsx-chartsheet-chart").first
    chart.wait_for(state="visible", timeout=settings.timeout_ms)
    chart_svg = chart.locator("svg").first
    chart_svg.wait_for(state="visible", timeout=settings.timeout_ms)
    chart_box = chart_svg.bounding_box()
    assert chart_box and chart_box["width"] > 100 and chart_box["height"] > 100
    expect(chart_svg).to_have_attribute(
        "aria-label", "Revenue by Quarter", timeout=settings.timeout_ms
    )
    expect(chart_svg).to_contain_text(
        "Revenue by Quarter", timeout=settings.timeout_ms
    )
    expect(chart_svg).to_contain_text("Q1", timeout=settings.timeout_ms)
    expect(chart_svg).to_contain_text("Q4", timeout=settings.timeout_ms)
    chart_marks = chart_svg.locator("[data-xlsx-chart-series-index]")
    assert chart_marks.count() >= 4, "chartsheet has no rendered data marks"

    page.screenshot(path=str(attempt_dir / "after.png"), full_page=True)
    return {
        "expectedDownloads": 0,
        "contextMenu": {
            "target": "B3",
            "originalValue": original_value,
            "cleared": True,
            "undoRestored": undo_restored,
        },
        "search": {
            "query": "Operating Income",
            "result": result_text,
            "selectedAddress": expected_address,
            "selectedSheet": "P&L",
        },
        "readOnly": {"editorBlocked": True, "writeCommandsHidden": True},
        "tableSort": {
            "table": "SalesRecords",
            "column": "Region",
            "direction": "ascending",
            "valueBefore": unsorted_first,
            "firstValue": sorted_first,
        },
        "chartsheet": {
            "fixture": "charts-images.xlsx",
            "tab": chart_tab.get_attribute("data-sheet-name"),
            "chartCount": page.get_by_test_id("xlsx-chartsheet-chart").count(),
            "dataMarkCount": chart_marks.count(),
            "chartBox": chart_box,
        },
    }


CASES = (
    ("UX-DOCX-001", "/docx-viewer", run_ux_docx_001),
    ("UX-DOCX-002", "/docx-editor", run_ux_docx_002),
    ("UX-DOCX-003", "/docx-editor", run_ux_docx_003),
    ("UX-DOCX-004", "/docx-editor", run_ux_docx_004),
    ("UX-XLSX-001", "/xlsx-viewer", run_ux_xlsx_001),
    ("UX-XLSX-002", "/xlsx-viewer", run_ux_xlsx_002),
    ("UX-XLSX-003", "/xlsx-viewer", run_ux_xlsx_003),
    ("UX-XLSX-004", "/xlsx-viewer", run_ux_xlsx_004),
)


def run_attempt(
    browser: Browser,
    case_id: str,
    route: str,
    runner,
    attempt: int,
    settings: Settings,
) -> dict[str, object]:
    attempt_dir = settings.evidence_dir / case_id / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport=VIEWPORT,
        device_scale_factor=1,
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
        accept_downloads=True,
    )
    page = context.new_page()
    page.set_default_timeout(settings.timeout_ms)
    page.set_default_navigation_timeout(settings.timeout_ms * 2)
    evidence = BrowserEvidence(page)
    started_at = datetime.now().astimezone().isoformat()
    result: dict[str, object]
    try:
        details = runner(page, attempt_dir, settings)
        expected_downloads = int(details.get("expectedDownloads", 0))
        actual_downloads = len(evidence.events["downloads"])
        assert actual_downloads == expected_downloads, {
            "expectedDownloads": expected_downloads,
            "actualDownloads": evidence.events["downloads"],
        }
        evidence.assert_clean()
        result = {
            "id": case_id,
            "route": route,
            "attempt": attempt,
            "status": "PASS",
            "startedAt": started_at,
            "finishedAt": datetime.now().astimezone().isoformat(),
            "details": details,
            "events": evidence.events,
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": case_id,
            "route": route,
            "attempt": attempt,
            "status": "FAIL",
            "startedAt": started_at,
            "finishedAt": datetime.now().astimezone().isoformat(),
            "error": str(error),
            "traceback": traceback.format_exc(),
            "events": evidence.events,
        }
    finally:
        evidence.save(attempt_dir)
        context.close()
    write_json(attempt_dir / "result.json", result)
    return result


def case_status(attempts: list[dict[str, object]]) -> str:
    if attempts[0]["status"] == "PASS":
        return "PASS"
    if len(attempts) == 2 and attempts[1]["status"] == "PASS":
        return "FLAKY"
    return "FAIL"


def run_suite(settings: Settings) -> int:
    settings.evidence_dir.mkdir(parents=True, exist_ok=True)
    upstream_reference = prepare_root_evidence(settings)
    upstream_reference = probe_and_capture_upstream_reference(
        settings, upstream_reference
    )
    upstream_status = str(upstream_reference["capture"]["status"])
    preview: subprocess.Popen[bytes] | None = None
    preview_log = None
    try:
        if settings.manage_preview:
            if port_in_use(settings.port):
                raise RuntimeError(
                    f"port {settings.port} is already in use; pass --base-url to use "
                    "an explicitly managed formal preview"
                )
            preview_log = (settings.evidence_dir / "preview.log").open("wb")
            append_command_log(
                settings.evidence_dir,
                (
                    f"START cwd={ROOT} command=pnpm --filter demo preview "
                    f"--host 127.0.0.1 --port {settings.port}"
                ),
            )
            preview = subprocess.Popen(
                [
                    "pnpm",
                    "--filter",
                    "demo",
                    "preview",
                    "--host",
                    "127.0.0.1",
                    "--port",
                    str(settings.port),
                ],
                cwd=ROOT,
                stdout=preview_log,
                stderr=subprocess.STDOUT,
                start_new_session=True,
            )
        wait_for_preview(settings.base_url, preview)

        with sync_playwright() as playwright:
            browser, browser_source = launch_browser(playwright)
            browser_version = browser.version
            results: list[dict[str, object]] = []
            try:
                for case_id, route, runner in CASES:
                    attempts = [
                        run_attempt(
                            browser, case_id, route, runner, 1, settings
                        )
                    ]
                    if attempts[0]["status"] == "FAIL":
                        attempts.append(
                            run_attempt(
                                browser, case_id, route, runner, 2, settings
                            )
                        )
                    results.append(
                        {
                            "id": case_id,
                            "route": route,
                            "status": case_status(attempts),
                            "attempts": attempts,
                        }
                    )
            finally:
                browser.close()

        local_stable = all(result["status"] == "PASS" for result in results)
        stable = local_stable and upstream_status == "PASS"
        suite_result = (
            "PASS"
            if stable
            else "BLOCKED" if upstream_status == "BLOCKED" else "FAIL"
        )
        update_environment(
            settings,
            localBrowser={"version": browser_version, "source": browser_source},
            localCases={
                "status": "PASS" if local_stable else "FAIL",
                "count": len(results),
            },
        )
        summary = {
            "suite": "UX-PARITY",
            "mode": "formal preview",
            "result": suite_result,
            "commit": git_head(),
            "baseUrl": settings.base_url,
            "viewport": VIEWPORT,
            "browser": browser_version,
            "browserSource": browser_source,
            "upstreamReference": {
                "repository": UPSTREAM_REPOSITORY,
                "expectedCommit": UPSTREAM_REFERENCE_COMMIT,
                "actualCommit": upstream_reference["actualCommit"],
                "commitMatches": upstream_reference["commitMatches"],
                "captureStatus": upstream_status,
                "artifact": "upstream-reference.json",
            },
            "results": results,
        }
        write_json(settings.evidence_dir / "summary.json", summary)
        return 0 if stable else 2 if suite_result == "BLOCKED" else 1
    finally:
        if preview is not None:
            append_command_log(
                settings.evidence_dir,
                f"STOP pid={preview.pid} signal=SIGTERM label=local-formal-preview",
            )
        stop_process(preview)
        if preview_log is not None:
            preview_log.close()


def main() -> int:
    settings = parse_args()
    try:
        return run_suite(settings)
    except Exception as error:
        settings.evidence_dir.mkdir(parents=True, exist_ok=True)
        write_json(
            settings.evidence_dir / "summary.json",
            {
                "suite": "UX-PARITY",
                "mode": "formal preview",
                "result": "BLOCKED",
                "commit": git_head(),
                "baseUrl": settings.base_url,
                "error": str(error),
                "traceback": traceback.format_exc(),
            },
        )
        print(f"UX-PARITY blocked: {error}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
