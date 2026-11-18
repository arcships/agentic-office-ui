#!/usr/bin/env python3

"""为桌面 PowerPoint 保存的正式素材生成浏览器事件与能力基准。"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from urllib.request import urlopen


ROOT = Path(__file__).resolve().parents[2]
FIXTURES = ROOT / "tests/fixtures/pptx/playback"
MANIFEST = FIXTURES / "manifest.json"
PORT = 5207
APP_URL = f"http://127.0.0.1:{PORT}"
POWERPOINT_VERSION = "16.110.3 (16.110.26070318)"


SPECS = {
    "timing-click-with-after.pptx": {
        "slide": 0,
        "args": ["--steps", "1"],
        "focus": "第 1 页：普通单击、与上一动画同时、上一动画之后和状态重建。",
    },
    "timing-delay-repeat-reverse.pptx": {
        "slide": 6,
        "args": ["--repeat-check"],
        "focus": "第 7 页：延迟、有限重复、自动往返、缓动和恢复。",
    },
    "entrance-exit-fade.pptx": {
        "slide": 10,
        "args": ["--overlap-check"],
        "focus": "第 11 页：淡入、淡出和最终隐藏状态。",
    },
    "wipe-scale-rotate.pptx": {
        "slide": 6,
        "args": ["--repeat-check"],
        "focus": "第 7 页：右向擦除、缩放和旋转的并行属性轨道。",
    },
    "motion-straight.pptx": {
        "slide": 6,
        "args": ["--repeat-check"],
        "focus": "第 7 页：直线路径到页面坐标的换算。",
    },
    "text-by-paragraph.pptx": {
        "slide": 11,
        "args": ["--paragraph-check"],
        "focus": "第 12 页：三段文字按三次单击依次显示。",
    },
    "trigger-shape-click.pptx": {
        "slide": 0,
        "args": ["--shape-trigger-check"],
        "focus": "第 1 页：指定对象点击只启动对应触发组。",
    },
    "transition-cut-fade-push-wipe.pptx": {
        "slide": 1,
        "args": ["--steps", "1", "--transition-check", "--pause-check"],
        "focus": "第 2 至 4 页和第 8 页：淡化、推进、擦除和直接切换。",
    },
    "media-audio-video.pptx": {
        "slide": 5,
        "args": ["--media-check", "--resource-check"],
        "focus": "第 6 页：嵌入音视频、暂停继续、裁剪、循环、音量和释放。",
    },
    "media-bookmark.pptx": {
        "slide": 5,
        "args": ["--media-check"],
        "focus": "第 6 页：媒体越过 middle 书签后触发指定效果。",
    },
    "hidden-and-internal-link.pptx": {
        "slide": 9,
        "args": ["--action-check"],
        "focus": "第 9 至 10 页：隐藏页、内部跳转和外部链接请求。",
    },
    "overlap-same-property.pptx": {
        "slide": 10,
        "args": ["--overlap-check"],
        "focus": "第 11 页：同对象透明度同时写入时后声明效果胜出。",
    },
    "morph-explicit-name.pptx": {
        "slide": 3,
        "args": ["--auto-advance-check", "--approximation", "safe"],
        "focus": "第 4 至 5 页：!!hero 强身份 Morph 和自动换页。",
    },
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def wait_for_server() -> None:
    for _ in range(100):
        try:
            with urlopen(APP_URL, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.1)
    raise RuntimeError("PPTX 基准捕获服务启动超时。")


def stop_server(server: subprocess.Popen) -> None:
    """结束整个开发服务器进程组，避免 vite 子进程残留到下一次捕获。"""
    if server.poll() is not None:
        return
    try:
        group = os.getpgid(server.pid)
    except ProcessLookupError:
        return
    try:
        os.killpg(group, signal.SIGTERM)
    except ProcessLookupError:
        return
    try:
        server.wait(timeout=5)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(group, signal.SIGKILL)
        except ProcessLookupError:
            pass


def stable_snapshot(file_name: str, spec: dict, raw: dict) -> dict:
    assertions = {
        key: raw.get(key)
        for key in [
            "clickBoundary", "effectStarts", "effectEnds", "stepChanges", "previousBoundary",
            "transitionLayersDuring", "transitionLayersAfter", "autoAdvanceSlideIndex",
            "autoAdvanceStatus", "mediaReleasedOnSlide", "mediaPolicy", "mediaTracks", "mediaUrlRevoked",
            "remainingChildrenAfterDispose", "repeatFinalScale", "indefiniteFinalStatus",
            "hiddenDirectRejected", "hiddenDirectSlideIndex", "hiddenSkippedSlideIndex",
            "externalActionHandled", "internalActionHandled", "internalActionSlideIndex",
            "actionRequests", "windowOpenCalls", "overlapInitialState", "overlapFinalState",
            "paragraphStates", "shapeTriggerHandled", "shapeTriggerKey",
        ]
        if raw.get(key) is not None
    }
    return {
        "schemaVersion": 1,
        "fixture": file_name,
        "powerPointVersion": POWERPOINT_VERSION,
        "focus": spec["focus"],
        "focusSlideIndex": spec["slide"],
        "document": {
            "slides": raw["playbackSlides"],
            "objects": raw["objects"],
            "nodes": raw["nodes"],
            "media": raw["media"],
            "hiddenSlides": raw["hiddenSlides"],
            "actions": raw["actions"],
        },
        "capability": {
            "discovered": raw["discovered"],
            "strict": raw["strict"],
            "approximate": raw["approximate"],
            "static": raw["static"],
            "unparsed": raw["unparsed"],
            "features": raw["capabilityFeatures"],
        },
        "focusedSlide": raw["focusedSlide"],
        "schedule": raw["schedule"],
        "eventSequence": [
            event for event in raw.get("eventSequence", [])
            if event["type"] not in {"statechange", "capability"}
        ],
        "assertions": assertions,
    }


def write_readme(directory: Path, file_name: str, spec: dict, digest: str) -> None:
    directory.joinpath("README.md").write_text(
        "\n".join([
            f"# {file_name}",
            "",
            spec["focus"],
            "",
            f"- 来源：仓库生成的组合素材，经桌面 Microsoft PowerPoint {POWERPOINT_VERSION} 另存为；",
            "- 许可：项目自有测试素材；",
            f"- SHA-256：`{digest}`；",
            "- `expected-events.json`：Chromium 中解析、对象、触发组、事件和能力报告基准；",
            "",
            "复测命令由 `scripts/fixtures/capture-pptx-formal-baselines.py` 统一执行；",
            "`--check` 只重新捕获并比对，不写入基准。",
            "",
        ]),
        encoding="utf-8",
    )


def capture_fixture(file_name: str, spec: dict) -> dict:
    """在 Chromium 中重新捕获一份稳定快照。"""
    fixture = FIXTURES / file_name
    if not fixture.is_file():
        raise RuntimeError(f"缺少正式素材：{file_name}")
    directory = FIXTURES / fixture.stem
    directory.mkdir(exist_ok=True)
    raw_output = directory / ".observed-events.json"
    sample_url = f"{APP_URL}/@fs{fixture}"
    command = [
        sys.executable, "tests/blackbox/pptx_playback_model.py",
        "--app-url", APP_URL,
        "--sample-url", sample_url,
        "--initial-slide", str(spec["slide"]),
        "--browser", "chromium",
        "--quiet",
        "--output", str(raw_output),
        *spec["args"],
    ]
    try:
        subprocess.run(command, cwd=ROOT, check=True)
        raw = json.loads(raw_output.read_text(encoding="utf-8"))
    finally:
        raw_output.unlink(missing_ok=True)
    return stable_snapshot(file_name, spec, raw)


def baseline_diff(file_name: str, spec: dict, snapshot: dict) -> list[str]:
    """比对现有基准与新快照，返回差异行；空列表表示一致。"""
    target = FIXTURES / Path(file_name).stem / "expected-events.json"
    expected = target.read_text(encoding="utf-8").splitlines()
    observed = (json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n").splitlines()
    return [
        line
        for line in difflib.unified_diff(expected, observed, fromfile="expected-events.json", tofile="observed-events.json", lineterm="")
        if not line.startswith(("---", "+++"))
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description="重新捕获并校验正式 PPTX 浏览器事件基准。")
    parser.add_argument("--check", action="store_true", help="只重新捕获并比对现有基准，不写入任何文件。")
    parser.add_argument("--fixtures", default="", help="逗号分隔的素材文件名，默认全部。")
    args = parser.parse_args()

    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    entries = {entry["file"]: entry for entry in manifest["entries"]}
    if set(entries) != set(SPECS):
        raise RuntimeError("正式素材清单与基准映射不一致。")
    selected = [name.strip() for name in args.fixtures.split(",") if name.strip()] or list(SPECS)
    unknown = [name for name in selected if name not in SPECS]
    if unknown:
        raise RuntimeError(f"未找到正式素材：{', '.join(unknown)}")

    server = subprocess.Popen(
        ["pnpm", "--filter=pptx-playback-lab", "dev", "--host", "127.0.0.1", "--port", str(PORT)],
        cwd=ROOT,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )
    failures: list[tuple[str, list[str]]] = []
    try:
        wait_for_server()
        for file_name in selected:
            spec = SPECS[file_name]
            snapshot = capture_fixture(file_name, spec)
            if args.check:
                diff = baseline_diff(file_name, spec, snapshot)
                if diff:
                    failures.append((file_name, diff))
                continue
            fixture = FIXTURES / file_name
            directory = FIXTURES / fixture.stem
            directory.joinpath("expected-events.json").write_text(
                json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            digest = sha256(fixture)
            write_readme(directory, file_name, spec, digest)
            entries[file_name].update({
                "status": "ready",
                "source": f"scripts/fixtures/generate-pptx-playback-fixtures.py -> Microsoft PowerPoint {POWERPOINT_VERSION} 另存为",
                "license": "项目自有测试素材",
                "sha256": digest,
                "powerpointVersion": POWERPOINT_VERSION,
                "baseline": f"{fixture.stem}/expected-events.json",
                "focusSlideIndex": spec["slide"],
            })
    finally:
        stop_server(server)

    if args.check:
        if failures:
            for file_name, diff in failures:
                print(f"FAIL: {file_name} 的浏览器基准已变化（前 80 行）：")
                print("\n".join(diff[:80]))
            print(f"FAIL: {len(failures)} 份基准与浏览器重新捕获的结果不一致。")
            sys.exit(1)
        print(f"PASS: {len(selected)} 份正式 PPTX 事件基准与 Chromium 重新捕获一致。")
        return

    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"PASS: 已生成 {len(selected)} 份正式 PPTX 事件基准。")


if __name__ == "__main__":
    main()
