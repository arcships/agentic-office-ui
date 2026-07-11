from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import signal
import socket
import struct
import subprocess
from tempfile import TemporaryDirectory
import time
import traceback
import urllib.request
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
import zlib

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_evidence import BrowserEvidence


ROOT = Path(__file__).resolve().parents[2]
OUTPUT = Path(
    os.environ.get(
        "BLACKBOX_EVIDENCE_DIR",
        ROOT / "output" / "acceptance" / "bb-stress-budget",
    )
).resolve()
PORT = int(os.environ.get("CI_STRESS_PREVIEW_PORT", "4182"))
BASE_URL = f"http://127.0.0.1:{PORT}"


DOCX_CONTENT_TYPES = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
DOCX_ROOT_RELS = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOCX_DOCUMENT = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>budget fixture</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'''

XLSX_CONTENT_TYPES = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
XLSX_ROOT_RELS = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
XLSX_WORKBOOK = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'''
XLSX_WORKBOOK_RELS = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'''
XLSX_SHEET = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row></sheetData></worksheet>'''


def zip_info(name: str) -> ZipInfo:
    info = ZipInfo(name, date_time=(2020, 1, 1, 0, 0, 0))
    info.compress_type = ZIP_DEFLATED
    return info


def png_chunk(kind: bytes, data: bytes) -> bytes:
    checksum = zlib.crc32(kind)
    checksum = zlib.crc32(data, checksum) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", checksum)


def declared_png(width: int, height: int, *, complete: bool = True) -> bytes:
    """A tiny PNG header whose dimensions can exercise pre-decode limits."""
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    value = b"\x89PNG\r\n\x1a\n" + png_chunk(b"IHDR", ihdr)
    return value + png_chunk(b"IEND", b"") if complete else value


def write_zip(path: Path, entries: list[tuple[str, bytes]]) -> dict[str, object]:
    with ZipFile(path, "w") as archive:
        for name, data in entries:
            archive.writestr(zip_info(name), data)
    with ZipFile(path) as archive:
        details = [
            {
                "name": item.filename,
                "compressedBytes": item.compress_size,
                "uncompressedBytes": item.file_size,
                "ratio": item.file_size / max(1, item.compress_size),
            }
            for item in archive.infolist()
        ]
    data = path.read_bytes()
    return {
        "fileName": path.name,
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "entries": details,
        "entryCount": len(details),
        "uncompressedBytes": sum(int(item["uncompressedBytes"]) for item in details),
        "xmlBytes": sum(
            int(item["uncompressedBytes"])
            for item in details
            if str(item["name"]).endswith((".xml", ".rels"))
            or item["name"] == "[Content_Types].xml"
        ),
    }


def docx_entries(document: bytes = DOCX_DOCUMENT, extras: list[tuple[str, bytes]] | None = None):
    return [
        ("[Content_Types].xml", DOCX_CONTENT_TYPES),
        ("_rels/.rels", DOCX_ROOT_RELS),
        ("word/document.xml", document),
        *(extras or []),
    ]


def xlsx_entries(extras: list[tuple[str, bytes]] | None = None):
    return [
        ("[Content_Types].xml", XLSX_CONTENT_TYPES),
        ("_rels/.rels", XLSX_ROOT_RELS),
        ("xl/workbook.xml", XLSX_WORKBOOK),
        ("xl/_rels/workbook.xml.rels", XLSX_WORKBOOK_RELS),
        ("xl/worksheets/sheet1.xml", XLSX_SHEET),
        *(extras or []),
    ]


def build_fixtures(directory: Path) -> tuple[dict[str, dict[str, object]], list[dict[str, object]]]:
    fixtures: dict[str, dict[str, object]] = {}
    deterministic_noise = b"".join(
        hashlib.sha256(f"budget-noise-{index}".encode()).digest()
        for index in range(64)
    )

    def add(key: str, suffix: str, entries: list[tuple[str, bytes]]) -> dict[str, object]:
        path = directory / f"{key}.{suffix}"
        manifest = write_zip(path, entries)
        fixtures[key] = {"path": path, "format": suffix, "manifest": manifest}
        return manifest

    normal_docx = add("normal-docx", "docx", docx_entries())
    add("normal-xlsx", "xlsx", xlsx_entries())
    large_input = add(
        "input-over-1k",
        "docx",
        docx_entries(
            b'''<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'''
            + b"".join(b"<w:p><w:r><w:t>input budget paragraph</w:t></w:r></w:p>" for _ in range(80))
            + b"<w:sectPr/></w:body></w:document>"
            , extras=[("custom/input-noise.bin", deterministic_noise)]
        ),
    )
    assert int(large_input["bytes"]) > 1024
    single_entry = add(
        "single-entry",
        "docx",
        docx_entries(DOCX_DOCUMENT.replace(b"budget fixture", b"S" * 1400)),
    )
    assert max(int(item["uncompressedBytes"]) for item in single_entry["entries"]) > 512
    add("entry-count", "docx", docx_entries())
    ratio_manifest = add(
        "compression-ratio",
        "xlsx",
        xlsx_entries([("xl/media/repeated.bin", b"R" * 4096)]),
    )
    assert max(float(item["ratio"]) for item in ratio_manifest["entries"]) > 2

    unsafe_entries = {
        "traversal": [("../escape.xml", b"<root/>")],
        "absolute": [("/absolute.xml", b"<root/>")],
        "backslash": [("word\\escape.xml", b"<root/>")],
        "normalized-duplicate": [
            ("word/caf\u00e9.xml", b"<root/>") ,
            ("word/cafe\u0301.xml", b"<root/>") ,
        ],
    }
    for name, extras in unsafe_entries.items():
        add(f"unsafe-{name}", "docx", docx_entries(extras=extras))

    deep = b'<?xml version="1.0"?><a><b><c><d><e><f/></e></d></c></b></a>'
    malformed = b'<?xml version="1.0"?><w:document><w:body><w:p></w:body>'
    entity = b'<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x "x">]><a>&x;</a>'
    long_attribute = b'<?xml version="1.0"?><a value="' + b"A" * 400 + b'"/>'
    add("xml-depth", "docx", docx_entries(deep))
    add("xml-malformed", "docx", docx_entries(malformed))
    add("xml-entity", "docx", docx_entries(entity))
    add("xml-long-attribute", "docx", docx_entries(long_attribute))

    page_paragraphs = b"".join(
        b"<w:p><w:r><w:t>page budget " + str(index).encode() + b" " + b"content " * 100 + b"</w:t></w:r></w:p>"
        for index in range(60)
    )
    add(
        "docx-many-pages",
        "docx",
        docx_entries(
            b'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            + page_paragraphs
            + b"<w:sectPr/></w:body></w:document>"
        ),
    )
    add(
        "docx-wrong-mime",
        "docx",
        [
            ("[Content_Types].xml", XLSX_CONTENT_TYPES),
            ("_rels/.rels", DOCX_ROOT_RELS),
            ("word/document.xml", DOCX_DOCUMENT),
        ],
    )

    for count in (4, 5, 6):
        extras = [(f"custom/item-{index}.bin", bytes([index])) for index in range(count - 3)]
        add(f"entry-boundary-{count}", "docx", docx_entries(extras=extras))

    combined_manifest = add(
        "combined-uncompressed",
        "xlsx",
        xlsx_entries([
            ("xl/custom/part-1.bin", b"A" * 900),
            ("xl/custom/part-2.bin", b"B" * 900),
            ("xl/custom/part-3.bin", b"C" * 900),
        ]),
    )
    max_entry = max(int(item["uncompressedBytes"]) for item in combined_manifest["entries"])
    total = int(combined_manifest["uncompressedBytes"])

    oversized_image = declared_png(100_000, 100_000)
    malformed_image = declared_png(100, 100, complete=False)
    high_resolution_image = declared_png(6_000, 6_000)
    aggregate_images = [
        ("word/media/image-1.png", declared_png(5_000, 5_000)),
        ("word/media/image-2.png", declared_png(5_000, 5_000)),
    ]
    add("image-width-docx", "docx", docx_entries(extras=[("word/media/bomb.png", oversized_image)]))
    add("image-width-xlsx", "xlsx", xlsx_entries([("xl/media/bomb.png", oversized_image)]))
    add("image-pixels-docx", "docx", docx_entries(extras=[("word/media/bomb.png", declared_png(8_000, 8_000))]))
    add("image-malformed-docx", "docx", docx_entries(extras=[("word/media/broken.png", malformed_image)]))
    add("image-malformed-xlsx", "xlsx", xlsx_entries([("xl/media/broken.png", malformed_image)]))
    add("image-high-resolution", "docx", docx_entries(extras=[("word/media/high-resolution.png", high_resolution_image)]))
    add("image-total-docx", "docx", docx_entries(extras=aggregate_images))
    add(
        "image-total-xlsx",
        "xlsx",
        xlsx_entries([
            ("xl/media/image-1.png", declared_png(5_000, 5_000)),
            ("xl/media/image-2.png", declared_png(5_000, 5_000)),
        ]),
    )

    cases = [
        {"id": "STRESS-001", "fixture": "input-over-1k", "limits": {"maxInputBytes": 1024}, "error": {"code": "LIMIT_EXCEEDED", "phase": "input", "limit": "maxInputBytes", "allowed": 1024}},
        {"id": "STRESS-002", "fixture": "single-entry", "limits": {"maxSingleEntryBytes": 512}, "error": {"code": "LIMIT_EXCEEDED", "phase": "archive", "limit": "maxSingleEntryBytes", "allowed": 512}},
        {"id": "STRESS-003-ENTRIES", "fixture": "entry-count", "limits": {"maxArchiveEntries": 2}, "error": {"code": "LIMIT_EXCEEDED", "phase": "archive", "limit": "maxArchiveEntries", "actual": 3, "allowed": 2}},
        {"id": "STRESS-003-RATIO", "fixture": "compression-ratio", "limits": {"maxCompressionRatio": 2}, "error": {"code": "LIMIT_EXCEEDED", "phase": "archive", "limit": "maxCompressionRatio", "allowed": 2}},
        {"id": "STRESS-004-DOCX-WIDTH", "fixture": "image-width-docx", "limits": {"maxImageWidth": 32_768}, "error": {"code": "IMAGE_LIMIT_EXCEEDED", "phase": "image", "limit": "maxImageWidth", "actual": 100_000, "allowed": 32_768}},
        {"id": "STRESS-004-XLSX-WIDTH", "fixture": "image-width-xlsx", "limits": {"maxImageWidth": 32_768}, "error": {"code": "IMAGE_LIMIT_EXCEEDED", "phase": "image", "limit": "maxImageWidth", "actual": 100_000, "allowed": 32_768}},
        {"id": "STRESS-004-DOCX-PIXELS", "fixture": "image-pixels-docx", "limits": {"maxImageWidth": 10_000, "maxImageHeight": 10_000, "maxSingleImagePixels": 40_000_000}, "error": {"code": "IMAGE_LIMIT_EXCEEDED", "phase": "image", "limit": "maxSingleImagePixels", "actual": 64_000_000, "allowed": 40_000_000}},
        {"id": "STRESS-004-DOCX-MALFORMED", "fixture": "image-malformed-docx", "limits": {}, "error": {"code": "INVALID_IMAGE", "phase": "image"}},
        {"id": "STRESS-004-XLSX-MALFORMED", "fixture": "image-malformed-xlsx", "limits": {}, "error": {"code": "INVALID_IMAGE", "phase": "image"}},
        {"id": "STRESS-004-HIGH-RESOLUTION", "fixture": "image-high-resolution", "limits": {"maxImageWidth": 6_000, "maxImageHeight": 6_000, "maxSingleImagePixels": 36_000_000, "maxTotalImagePixels": 36_000_000}, "state": "ready"},
        *[
            {"id": f"STRESS-011-{name.upper()}", "fixture": f"unsafe-{name}", "limits": {}, "error": {"phase": "archive", "codes": ["PARSE_FAILED", "INVALID_WORKBOOK"]}}
            for name in unsafe_entries
        ],
        {"id": "STRESS-012-DEPTH", "fixture": "xml-depth", "limits": {"maxXmlDepth": 4}, "error": {"code": "LIMIT_EXCEEDED", "phase": "xml", "limit": "maxXmlDepth", "allowed": 4}},
        {"id": "STRESS-012-MALFORMED", "fixture": "xml-malformed", "limits": {}, "error": {"phase": "xml", "codes": ["PARSE_FAILED", "INVALID_WORKBOOK"]}},
        {"id": "STRESS-012-ENTITY", "fixture": "xml-entity", "limits": {}, "error": {"phase": "xml", "codes": ["PARSE_FAILED", "INVALID_WORKBOOK"]}},
        {"id": "STRESS-012-ATTRIBUTE", "fixture": "xml-long-attribute", "limits": {"maxXmlAttributeBytes": 128}, "error": {"code": "LIMIT_EXCEEDED", "phase": "xml", "limit": "maxXmlAttributeBytes", "allowed": 128}},
        {"id": "STRESS-013-DOCX-PAGES", "fixture": "docx-many-pages", "limits": {"maxDocxPages": 1}, "recoveryLimits": {"maxDocxPages": 10}, "error": {"code": "LIMIT_EXCEEDED", "phase": "layout", "limit": "maxDocxPages", "actual": 2, "allowed": 1}},
        {"id": "STRESS-013-LIMIT-MINUS-1", "fixture": "entry-boundary-4", "limits": {"maxArchiveEntries": 5}, "state": "ready"},
        {"id": "STRESS-013-LIMIT", "fixture": "entry-boundary-5", "limits": {"maxArchiveEntries": 5}, "state": "ready"},
        {"id": "STRESS-013-LIMIT-PLUS-1", "fixture": "entry-boundary-6", "limits": {"maxArchiveEntries": 5}, "error": {"code": "LIMIT_EXCEEDED", "phase": "archive", "limit": "maxArchiveEntries", "actual": 6, "allowed": 5}},
        {"id": "STRESS-014-UNCOMPRESSED", "fixture": "combined-uncompressed", "limits": {"maxSingleEntryBytes": max_entry, "maxUncompressedBytes": total - 1}, "error": {"code": "LIMIT_EXCEEDED", "phase": "archive", "limit": "maxUncompressedBytes", "allowed": total - 1}},
        {"id": "STRESS-014-DOCX-TOTAL-PIXELS", "fixture": "image-total-docx", "limits": {"maxSingleImagePixels": 30_000_000, "maxTotalImagePixels": 40_000_000}, "error": {"code": "IMAGE_LIMIT_EXCEEDED", "phase": "image", "limit": "maxTotalImagePixels", "actual": 50_000_000, "allowed": 40_000_000}},
        {"id": "STRESS-014-XLSX-TOTAL-PIXELS", "fixture": "image-total-xlsx", "limits": {"maxSingleImagePixels": 30_000_000, "maxTotalImagePixels": 40_000_000}, "error": {"code": "IMAGE_LIMIT_EXCEEDED", "phase": "image", "limit": "maxTotalImagePixels", "actual": 50_000_000, "allowed": 40_000_000}},
        {"id": "STRESS-MIME-DOCX", "fixture": "docx-wrong-mime", "limits": {}, "error": {"phase": "archive", "codes": ["PARSE_FAILED"]}},
    ]
    requested_cases = {
        item.strip()
        for item in os.environ.get("CI_STRESS_CASES", "").split(",")
        if item.strip()
    }
    if requested_cases:
        unknown_cases = requested_cases.difference(str(case["id"]) for case in cases)
        if unknown_cases:
            raise ValueError(f"unknown CI_STRESS_CASES: {sorted(unknown_cases)}")
        cases = [case for case in cases if str(case["id"]) in requested_cases]
    manifests = [
        {"id": key, "format": value["format"], **value["manifest"]}
        for key, value in fixtures.items()
    ]
    assert int(normal_docx["entryCount"]) == 3
    return fixtures, cases, manifests


def port_in_use() -> bool:
    with socket.socket() as sock:
        return sock.connect_ex(("127.0.0.1", PORT)) == 0


def wait_for_preview(process: subprocess.Popen[bytes]) -> None:
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"preview exited early with {process.returncode}")
        try:
            with urllib.request.urlopen(BASE_URL, timeout=1) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.2)
    raise TimeoutError(f"preview did not become ready at {BASE_URL}")


def launch_browser(playwright):
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
            return playwright.chromium.launch(headless=True, executable_path=str(candidate)), str(candidate)
        except PlaywrightError as error:
            failures.append(str(error))
    raise RuntimeError("no Chromium executable could launch: " + " | ".join(failures))


def stop_process(process: subprocess.Popen[bytes]) -> None:
    if process.poll() is not None:
        return
    try:
        os.killpg(process.pid, signal.SIGTERM)
        process.wait(timeout=5)
    except Exception:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except Exception:
            process.kill()


def parse_number(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return float(value)


def run_attempt(browser, case: dict[str, object], fixtures: dict[str, dict[str, object]], attempt: int) -> dict[str, object]:
    attempt_dir = OUTPUT / str(case["id"]) / f"attempt-{attempt}"
    attempt_dir.mkdir(parents=True, exist_ok=True)
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
    )
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()
    evidence = BrowserEvidence(page)
    fixture = fixtures[str(case["fixture"])]
    expected_state = str(case.get("state", "error"))
    try:
        page.goto(f"{BASE_URL}/#/runtime-limits", wait_until="networkidle")
        page.wait_for_function("() => document.querySelector('[data-testid=page-status]')?.dataset.state === 'ready'")
        page.get_by_test_id("runtime-limits-format").select_option(str(fixture["format"]))
        page.get_by_test_id("runtime-limits-json").fill(json.dumps(case["limits"]))
        page.get_by_test_id("runtime-limits-apply").click()
        active = json.loads(page.get_by_test_id("runtime-limits-active").inner_text())
        for key, value in dict(case["limits"]).items():
            assert active[key] == value, {"key": key, "expected": value, "actual": active.get(key)}
        page.get_by_test_id("runtime-limits-file-input").set_input_files(str(fixture["path"]))
        page.wait_for_function(
            "state => document.querySelector('[data-testid=runtime-limits-status]')?.dataset.state === state",
            arg=expected_state,
            timeout=30_000,
        )

        error_details = None
        diagnostics_before_recovery = page.get_by_test_id("runtime-limits-diagnostics").inner_text()
        if expected_state == "error":
            error = page.get_by_test_id("runtime-limits-error")
            error_details = {
                "code": error.get_attribute("data-error-code"),
                "phase": error.get_attribute("data-error-phase"),
                "limit": error.get_attribute("data-error-limit"),
                "actual": parse_number(error.get_attribute("data-error-actual")),
                "allowed": parse_number(error.get_attribute("data-error-allowed")),
            }
            expected = dict(case["error"])
            if "code" in expected:
                assert error_details["code"] == expected["code"], error_details
            else:
                assert error_details["code"] in expected["codes"], error_details
            assert error_details["phase"] == expected.get("phase"), error_details
            if "limit" in expected:
                assert error_details["limit"] == expected["limit"], error_details
            if "actual" in expected:
                assert error_details["actual"] == float(expected["actual"]), error_details
            if "allowed" in expected:
                assert error_details["allowed"] == float(expected["allowed"]), error_details
                assert error_details["actual"] is not None and error_details["actual"] > error_details["allowed"], error_details
            page.screenshot(path=str(attempt_dir / "error.png"), full_page=True)

            normal = fixtures[f"normal-{fixture['format']}"]
            page.get_by_test_id("runtime-limits-json").fill(
                json.dumps(case.get("recoveryLimits", {}))
            )
            page.get_by_test_id("runtime-limits-apply").click()
            page.get_by_test_id("runtime-limits-file-input").set_input_files(str(normal["path"]))
            page.wait_for_function(
                "() => document.querySelector('[data-testid=runtime-limits-status]')?.dataset.state === 'ready'",
                timeout=30_000,
            )
            page.screenshot(path=str(attempt_dir / "recovered.png"), full_page=True)
        else:
            assert page.get_by_test_id("runtime-limits-error").count() == 0
            page.screenshot(path=str(attempt_dir / "ready.png"), full_page=True)

        evidence.assert_clean()
        assert not evidence.events["downloads"], evidence.events["downloads"]
        result: dict[str, object] = {
            "id": case["id"],
            "attempt": attempt,
            "status": "PASS",
            "format": fixture["format"],
            "fixture": fixture["manifest"],
            "limits": case["limits"],
            "expectedState": expected_state,
            "error": error_details,
            "diagnosticsBeforeRecovery": diagnostics_before_recovery,
            "finalState": page.get_by_test_id("runtime-limits-status").get_attribute("data-state"),
        }
    except Exception as error:
        try:
            page.screenshot(path=str(attempt_dir / "failure.png"), full_page=True)
        except Exception:
            pass
        result = {
            "id": case["id"],
            "attempt": attempt,
            "status": "FAIL",
            "error": repr(error),
            "traceback": traceback.format_exc(),
            "events": evidence.events,
        }
    finally:
        evidence.save(attempt_dir)
        context.tracing.stop(path=attempt_dir / "trace.zip")
        context.close()
    (attempt_dir / "result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def main() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    if port_in_use():
        raise RuntimeError(f"port {PORT} is already in use")
    preview_log = (OUTPUT / "preview.log").open("wb")
    preview = subprocess.Popen(
        ["pnpm", "--filter", "demo", "preview", "--host", "127.0.0.1", "--port", str(PORT)],
        cwd=ROOT,
        stdout=preview_log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    try:
        wait_for_preview(preview)
        with TemporaryDirectory(prefix="agentic-office-budget-") as temp:
            fixtures, cases, manifests = build_fixtures(Path(temp))
            (OUTPUT / "fixtures.json").write_text(
                json.dumps(manifests, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            with sync_playwright() as playwright:
                browser, browser_source = launch_browser(playwright)
                browser_version = browser.version
                results = []
                for case in cases:
                    attempts = [run_attempt(browser, case, fixtures, 1)]
                    if attempts[0]["status"] == "FAIL":
                        attempts.append(run_attempt(browser, case, fixtures, 2))
                    status = "PASS" if attempts[0]["status"] == "PASS" else (
                        "FLAKY" if attempts[-1]["status"] == "PASS" else "FAIL"
                    )
                    results.append({"id": case["id"], "status": status, "attempts": attempts})
                browser.close()
        overall = "PASS" if all(item["status"] == "PASS" for item in results) else "FAIL"
        summary = {
            "suite": "BB-STRESS",
            "subset": [str(case["id"]) for case in cases],
            "mode": "formal preview",
            "result": overall,
            "browser": browser_version,
            "browserSource": browser_source,
            "results": results,
        }
        (OUTPUT / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return 0 if overall == "PASS" else 1
    finally:
        stop_process(preview)
        preview_log.close()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        (OUTPUT / "summary.json").write_text(
            json.dumps(
                {"suite": "BB-STRESS", "result": "BLOCKED", "error": repr(error), "traceback": traceback.format_exc()},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"BLOCKED: {error}")
        raise SystemExit(2)
