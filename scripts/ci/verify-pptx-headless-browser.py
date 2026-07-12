#!/usr/bin/env python3

import atexit
import os
import subprocess
import time
import urllib.error
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright


url = os.environ.get("PPTX_HEADLESS_URL", "http://127.0.0.1:4178/#/pptx-headless")
root = Path(__file__).resolve().parents[2]
fixture = Path(os.environ.get(
    "PPTX_HEADLESS_FIXTURE",
    str(root / "tests/fixtures/pptx/playback-controlled.pptx"),
))
second_fixture = Path(os.environ.get(
    "PPTX_HEADLESS_SECOND_FIXTURE",
    str(root / "tests/fixtures/pptx/playback/timing-click-with-after.pptx"),
))

if not fixture.is_file():
    raise FileNotFoundError(f"找不到 PPTX 验证文件：{fixture}")
if not second_fixture.is_file():
    raise FileNotFoundError(f"找不到第二个 PPTX 验证文件：{second_fixture}")


def server_ready() -> bool:
    try:
        with urllib.request.urlopen(url.split("/#", 1)[0], timeout=1) as response:
            return response.status < 500
    except (OSError, urllib.error.URLError):
        return False


server = None
if not server_ready():
    server = subprocess.Popen(
        ["pnpm", "--filter=demo", "dev", "--host", "127.0.0.1", "--port", "4178"],
        cwd=root,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    atexit.register(server.terminate)
    for _ in range(100):
        if server_ready():
            break
        if server.poll() is not None:
            raise RuntimeError("PPTX 最小组合验证服务启动失败。")
        time.sleep(0.1)
    else:
        raise TimeoutError("PPTX 最小组合验证服务启动超时。")

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1440, "height": 900})
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.goto(url, wait_until="networkidle")

    page.get_by_test_id("pptx-headless-file").set_input_files(str(fixture))
    stage = page.get_by_test_id("pptx-stage")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.documentState === 'ready'",
        timeout=30000,
    )
    assert stage.get_attribute("data-slide-index") == "0"
    assert page.locator("[data-pptx-object-key]").count() > 0

    initial_boundary = int(stage.get_attribute("data-click-boundary") or "0")
    page.get_by_test_id("pptx-headless-next").click()
    page.wait_for_function(
        "([boundary]) => Number(document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.clickBoundary ?? 0) > boundary",
        arg=[initial_boundary],
        timeout=10000,
    )
    assert stage.get_attribute("data-slide-index") == "0", "第一次下一步不应直接换页"

    page.get_by_test_id("pptx-headless-reset").click()
    page.wait_for_function(
        "() => Number(document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.clickBoundary ?? -1) === 0",
        timeout=10000,
    )
    stage.click(position={"x": 8, "y": 8})
    page.wait_for_function(
        "() => Number(document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.clickBoundary ?? 0) >= 1",
        timeout=10000,
    )
    assert stage.get_attribute("data-slide-index") == "0", "第一次舞台点击不应直接换页"

    page.get_by_test_id("pptx-headless-remount").click()
    page.get_by_text("舞台已卸载。").wait_for()
    page.get_by_test_id("pptx-headless-remount").click()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.documentState === 'ready'",
        timeout=30000,
    )
    assert page.get_by_test_id("pptx-stage").get_attribute("data-slide-index") == "0"
    assert page.locator("[data-pptx-object-key]").count() > 0
    assert not errors, errors

    page.get_by_test_id("pptx-headless-file").set_input_files(str(fixture))
    page.get_by_test_id("pptx-headless-file").set_input_files(str(second_fixture))
    page.wait_for_function(
        "([name]) => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.fileName === name",
        arg=[second_fixture.name],
        timeout=30000,
    )
    assert stage.get_attribute("data-slide-index") == "0"

    repeat_fixture = root / "tests/fixtures/pptx/playback/timing-delay-repeat-reverse.pptx"
    page.get_by_test_id("pptx-headless-file").set_input_files(str(repeat_fixture))
    page.wait_for_function(
        "([name]) => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.fileName === name",
        arg=[repeat_fixture.name],
        timeout=30000,
    )
    page.get_by_test_id("pptx-headless-play").click()
    page.wait_for_function(
        "() => ['running', 'waiting', 'ended'].includes(document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.playbackStatus ?? '')",
        timeout=10000,
    )
    page.get_by_test_id("pptx-headless-pause").click()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.playbackStatus === 'paused'",
        timeout=10000,
    )
    page.get_by_test_id("pptx-headless-resume").click()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.playbackStatus !== 'paused'",
        timeout=10000,
    )

    hidden_fixture = root / "tests/fixtures/pptx/playback/hidden-and-internal-link.pptx"
    page.get_by_test_id("pptx-headless-file").set_input_files(str(hidden_fixture))
    page.wait_for_function(
        "([name]) => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.fileName === name",
        arg=[hidden_fixture.name],
        timeout=30000,
    )
    page.get_by_test_id("pptx-headless-slide-8").click()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.slideIndex === '7'",
        timeout=10000,
    )
    for _ in range(12):
        page.get_by_test_id("pptx-headless-next").click()
        page.wait_for_timeout(120)
        if stage.get_attribute("data-slide-index") != "7":
            break
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.slideIndex === '9'",
        timeout=10000,
    )

    page.get_by_test_id("pptx-headless-file").set_input_files(str(fixture))
    page.wait_for_function(
        "([name]) => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.fileName === name",
        arg=[fixture.name],
        timeout=30000,
    )
    primary_boundary = int(stage.get_attribute("data-click-boundary") or "0")
    page.get_by_test_id("pptx-headless-second-toggle").click()
    second_stage = page.get_by_test_id("pptx-headless-second-stage")
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-headless-second-stage\"]')?.dataset.documentState === 'ready'",
        timeout=30000,
    )
    page.get_by_test_id("pptx-headless-second-next").click()
    page.wait_for_function(
        "() => Number(document.querySelector('[data-testid=\"pptx-headless-second-stage\"]')?.dataset.clickBoundary ?? 0) >= 1",
        timeout=10000,
    )
    assert int(stage.get_attribute("data-click-boundary") or "0") == primary_boundary
    assert second_stage.get_attribute("data-slide-index") == "0"

    media_fixture = root / "tests/fixtures/pptx/playback/media-audio-video.pptx"
    page.get_by_test_id("pptx-headless-second-toggle").click()
    page.get_by_test_id("pptx-headless-file").set_input_files(str(media_fixture))
    page.wait_for_function(
        "([name]) => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.fileName === name",
        arg=[media_fixture.name],
        timeout=30000,
    )
    page.get_by_test_id("pptx-headless-slide-6").click()
    page.wait_for_function(
        "() => document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.slideIndex === '5'",
        timeout=10000,
    )
    assert stage.locator("video, audio").count() > 0
    page.get_by_test_id("pptx-headless-play").click()
    page.wait_for_function(
        "() => ['running', 'waiting', 'blocked', 'ended'].includes(document.querySelector('[data-testid=\"pptx-stage\"]')?.dataset.playbackStatus ?? '')",
        timeout=10000,
    )
    assert not errors, errors
    browser.close()

print("PASS: PPTX 最小组合完成逐步动画、重新挂载、快速换文件、暂停恢复、隐藏页、双播放器和媒体验证。")
