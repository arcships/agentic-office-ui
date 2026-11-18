"""Verify the generated office fixtures and their document-level features.

The manifest only pins file names, bytes and hashes. This script opens every
fixture with a real parser and checks the behaviours the black-box suites rely
on: PDF page counts, `/Rotate` metadata, image-only pages, encryption and
searchable text (including CJK). It is executed by the PDF CI suite, so a
regenerated fixture that silently loses rotation or gains a text layer fails
the build.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from PIL import Image
from pypdf import PdfReader
from pypdf.errors import FileNotDecryptedError

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "apps" / "demo" / "public" / "samples"

valid_docx = [
    "demo.docx",
    "legal-contract.docx",
    "invoice-table.docx",
    "report-with-image.docx",
    "chinese-mixed.docx",
]
valid_pdf = [
    "sample.pdf",
    "scanned-invoice.pdf",
    "rotated-pages.pdf",
    "large-contract.pdf",
    "cjk-report.pdf",
]
valid_xlsx = [
    "financial-model.xlsx",
    "sales-table.xlsx",
    "charts-images.xlsx",
    "large-grid.xlsx",
]
valid_png = ["invoice.png", "contract-page.png"]
valid_json = ["field-citations.json", "ocr-layout.json", "manifest.json"]
corrupted = ["corrupted.docx", "corrupted.pdf", "corrupted.xlsx"]

# Page counts, per-page /Rotate values and searchable markers the PDF black-box
# suites assert in the browser. Keep this table in sync with tests/blackbox.
pdf_expectations: dict[str, dict[str, object]] = {
    "sample.pdf": {
        "pages": 4,
        "rotations": [0, 0, 0, 0],
        "page_markers": [f"ALPHA-{page}" for page in range(1, 5)],
    },
    "scanned-invoice.pdf": {
        "pages": 3,
        "rotations": [0, 0, 0],
        "image_only": True,
    },
    "rotated-pages.pdf": {
        "pages": 2,
        "rotations": [0, 90],
        "page_markers": ["Portrait Page", "Rotated Page"],
    },
    "large-contract.pdf": {
        "pages": 240,
        "minimum_pages": 200,
        "rotations": [0],
    },
    "cjk-report.pdf": {
        "pages": 2,
        "rotations": [0, 0],
        "page_markers": ["CJK-ALPHA-1", "CJK-ALPHA-2", "合同"],
    },
}
encrypted_pdf = {"file": "encrypted.pdf", "pages": 2, "password": "arcships", "wrong_password": "not-the-password"}


def pdf_rotation(page) -> int:
    return int(page.get("/Rotate") or 0)


def pdf_text(page) -> str:
    return page.extract_text() or ""


report: dict[str, dict[str, object]] = {
    "docx": {},
    "pdf": {},
    "xlsx": {},
    "png": {},
    "json": {},
    "corrupted": {},
}
failures: list[str] = []


def expect(condition: bool, message: str) -> None:
    if not condition:
        failures.append(message)


for name in valid_docx:
    path = SAMPLES / name
    doc = Document(path)
    report["docx"][name] = {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "bytes": path.stat().st_size,
    }
    expect(len(doc.paragraphs) > 0, f"{name}: expected at least one paragraph")

for name in valid_pdf:
    path = SAMPLES / name
    expected = pdf_expectations[name]
    reader = PdfReader(str(path))
    expect(not reader.is_encrypted, f"{name}: fixture must not be encrypted")
    pages = len(reader.pages)
    texts = [pdf_text(page) for page in reader.pages]
    rotations = [pdf_rotation(page) for page in reader.pages]
    report["pdf"][name] = {
        "pages": pages,
        "rotations": rotations,
        "textLengths": [len(text) for text in texts],
        "bytes": path.stat().st_size,
    }
    expect(pages == int(expected["pages"]), f"{name}: expected {expected['pages']} pages, got {pages}")
    if "minimum_pages" in expected:
        expect(
            pages >= int(expected["minimum_pages"]),
            f"{name}: expected at least {expected['minimum_pages']} pages, got {pages}",
        )
    if "rotations" in expected:
        fixture_rotations = expected["rotations"]
        assert isinstance(fixture_rotations, list)
        if len(fixture_rotations) == 1:
            mismatched = [rotation for rotation in rotations if rotation != fixture_rotations[0]]
            expect(
                not mismatched,
                f"{name}: every page must be /Rotate {fixture_rotations[0]}, got {rotations}",
            )
        else:
            expect(
                rotations == fixture_rotations,
                f"{name}: expected page rotations {fixture_rotations}, got {rotations}",
            )
    if expected.get("image_only"):
        expect(
            all(not text.strip() for text in texts),
            f"{name}: image-only fixture must not expose a native text layer, got {texts}",
        )
        expect(
            any("/XObject" in (page.get("/Resources") or {}) for page in reader.pages),
            f"{name}: image-only fixture must draw at least one image XObject",
        )
    for marker in expected.get("page_markers", []):
        expect(
            any(marker in text for text in texts),
            f"{name}: extracted text is missing marker {marker!r}",
        )

encrypted_path = SAMPLES / encrypted_pdf["file"]
encrypted_reader = PdfReader(str(encrypted_path))
expect(encrypted_reader.is_encrypted, f"{encrypted_pdf['file']}: fixture must declare /Encrypt")
if encrypted_reader.is_encrypted:
    expect(
        encrypted_reader.decrypt(encrypted_pdf["wrong_password"]) == 0,
        f"{encrypted_pdf['file']}: a wrong password must not unlock the fixture",
    )
    locked = PdfReader(str(encrypted_path))
    expect(locked.is_encrypted, f"{encrypted_pdf['file']}: a fresh reader must still be locked")
    try:
        len(locked.pages)
        locked_error = ""
    except FileNotDecryptedError as error:
        locked_error = str(error)
    expect(locked_error != "", f"{encrypted_pdf['file']}: locked pages must raise FileNotDecryptedError")
    decrypted = PdfReader(str(encrypted_path))
    expect(
        decrypted.decrypt(encrypted_pdf["password"]) == 1,
        f"{encrypted_pdf['file']}: documented password must unlock the fixture",
    )
    decrypted_texts = [pdf_text(page) for page in decrypted.pages]
    report["pdf"][encrypted_pdf["file"]] = {
        "pages": len(decrypted.pages),
        "encrypted": True,
        "textLengths": [len(text) for text in decrypted_texts],
        "bytes": encrypted_path.stat().st_size,
    }
    expect(
        len(decrypted.pages) == int(encrypted_pdf["pages"]),
        f"{encrypted_pdf['file']}: expected {encrypted_pdf['pages']} pages after decryption",
    )
    expect(
        all("Encrypted" in text or "Payment" in text for text in decrypted_texts),
        f"{encrypted_pdf['file']}: decrypted pages lost their text layer",
    )

for name in valid_xlsx:
    path = SAMPLES / name
    wb = load_workbook(path, data_only=False, read_only=False)
    report["xlsx"][name] = {
        "sheets": wb.sheetnames,
        "bytes": path.stat().st_size,
    }
    expect(len(wb.sheetnames) > 0, f"{name}: workbook must contain at least one sheet")
    wb.close()

for name in valid_png:
    path = SAMPLES / name
    with Image.open(path) as img:
        report["png"][name] = {
            "size": img.size,
            "mode": img.mode,
            "bytes": path.stat().st_size,
        }
        expect(img.size[0] > 0 and img.size[1] > 0, f"{name}: image must have a positive size")

for name in valid_json:
    path = SAMPLES / name
    data = json.loads(path.read_text(encoding="utf-8"))
    report["json"][name] = {
        "type": type(data).__name__,
        "bytes": path.stat().st_size,
    }

for name in corrupted:
    path = SAMPLES / name
    is_zip = zipfile.is_zipfile(path)
    report["corrupted"][name] = {
        "is_zip": is_zip,
        "bytes": path.stat().st_size,
    }
    expect(not is_zip, f"{name}: negative corrupted fixture must not be a zip")

report["result"] = "FAIL" if failures else "PASS"
print(json.dumps(report, ensure_ascii=False, indent=2))
if failures:
    raise SystemExit("material verification failed:\n" + "\n".join(f"- {failure}" for failure in failures))
