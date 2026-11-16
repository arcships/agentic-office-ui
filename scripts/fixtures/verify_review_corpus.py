#!/usr/bin/env python3
"""Verify that test-data/corpus contains the features its manifest claims.

Run: uv run --with pypdf --with openpyxl --with python-docx --with python-pptx \
         python scripts/fixtures/verify_review_corpus.py
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]
CORPUS = ROOT / "test-data" / "corpus"
FAILURES: list[str] = []


def check(name: str, condition: bool, detail: str) -> None:
    if not condition:
        FAILURES.append(f"{name}: {detail}")


def docx_features(name: str) -> dict:
    with zipfile.ZipFile(CORPUS / name) as z:
        doc = z.read("word/document.xml").decode("utf8", "ignore")
        comments = z.read("word/comments.xml").decode("utf8", "ignore") if "word/comments.xml" in z.namelist() else ""
        return {
            "vmerge": len(re.findall(r"<w:vMerge", doc)),
            "gridspan": len(re.findall(r"<w:gridSpan", doc)),
            "comments": len(re.findall(r"<w:comment\b", comments)),
            "comment_ranges": len(re.findall(r"<w:commentRangeStart", doc)),
            "footnote_refs": len(re.findall(r"<w:footnoteReference", doc)),
            "sections": len(re.findall(r"<w:sectPr", doc)),
            "columns": re.findall(r"<w:cols[^>]*w:num=\"(\d+)\"", doc),
            "header_refs": len(re.findall(r"<w:headerReference", doc)),
            "nested_table": doc.count("<w:tbl>") > 1,
            "multi_paragraph_cell": len(re.findall(r"<w:tc>[\s\S]*?<w:p\b[\s\S]*?<w:p\b", doc)) > 0,
        }


def main() -> int:
    manifest = json.loads((CORPUS / "manifest.json").read_text())
    entries = {entry["file"]: entry for entry in manifest["entries"]}
    assert len(entries) >= 12, "corpus is unexpectedly small"

    # byte sizes match the manifest
    for name, entry in entries.items():
        path = CORPUS / name
        check(name, path.exists(), "missing file")
        if path.exists():
            check(name, path.stat().st_size == entry["bytes"], "byte size drifted from manifest")

    merged = docx_features("docx-merged-table.docx")
    check("docx-merged-table", merged["vmerge"] >= 2, f"expected vMerge, got {merged['vmerge']}")
    check("docx-merged-table", merged["gridspan"] >= 1, "missing gridSpan")
    check("docx-merged-table", merged["nested_table"], "missing nested table")

    comments = docx_features("docx-comments.docx")
    check("docx-comments", comments["comments"] >= 3, f"expected >=3 comments, got {comments['comments']}")
    check("docx-comments", comments["comment_ranges"] >= 2, "expected at least two comment ranges")
    check("docx-comments", comments["footnote_refs"] >= 1, "missing footnote reference")
    check("docx-comments", comments["header_refs"] >= 1, "missing header reference (header comment)")

    mixed = docx_features("docx-mixed-orientation.docx")
    check("docx-mixed-orientation", mixed["sections"] >= 2, "expected two sections")

    columns = docx_features("docx-two-columns.docx")
    check("docx-two-columns", "2" in columns["columns"], "expected w:cols num=2")

    charts = load_workbook(CORPUS / "xlsx-charts.xlsx")
    ws = charts[charts.sheetnames[0]]
    check("xlsx-charts", len(ws._charts) >= 5, f"expected >=5 charts, got {len(ws._charts)}")
    with zipfile.ZipFile(CORPUS / "xlsx-charts.xlsx") as z:
        check("xlsx-charts", any(n.startswith("xl/charts/") for n in z.namelist()), "missing chart parts")

    cf = load_workbook(CORPUS / "xlsx-conditional-formats.xlsx")
    ws = cf[cf.sheetnames[0]]
    rules = sum(len(r.rules) for r in ws.conditional_formatting)
    check("xlsx-conditional-formats", rules >= 3, f"expected >=3 conditional rules, got {rules}")

    dv = load_workbook(CORPUS / "xlsx-validation.xlsx")
    ws = dv[dv.sheetnames[0]]
    check("xlsx-validation", len(ws.data_validations.dataValidation) >= 2, "expected two validations")

    frozen = load_workbook(CORPUS / "xlsx-frozen-hidden.xlsx")
    ws = frozen[frozen.sheetnames[0]]
    check("xlsx-frozen-hidden", ws.freeze_panes == "B3", f"freeze is {ws.freeze_panes}")
    check("xlsx-frozen-hidden", bool(ws.row_dimensions[3].hidden), "row 3 not hidden")
    check("xlsx-frozen-hidden", bool(ws.column_dimensions["D"].hidden), "column D not hidden")
    check("xlsx-frozen-hidden", str(frozen.epoch).startswith("1904"), f"epoch is {frozen.epoch}")

    rotated = PdfReader(str(CORPUS / "pdf-rotated-90.pdf"))
    check("pdf-rotated-90", int(rotated.pages[0].get("/Rotate", 0)) == 90, "page is not rotated")

    image_only = PdfReader(str(CORPUS / "pdf-image-only.pdf"))
    text = (image_only.pages[0].extract_text() or "").strip()
    check("pdf-image-only", text == "", f"image-only fixture has {len(text)} chars of text")

    encrypted = PdfReader(str(CORPUS / "pdf-encrypted.pdf"))
    check("pdf-encrypted", encrypted.is_encrypted, "PDF is not encrypted")

    many = PdfReader(str(CORPUS / "pdf-300-pages.pdf"))
    check("pdf-300-pages", len(many.pages) == 300, f"expected 300 pages, got {len(many.pages)}")

    deck = Presentation(str(CORPUS / "pptx-deck-basic.pptx"))
    check("pptx-deck-basic", len(deck.slides) == 5, f"expected 5 slides, got {len(deck.slides)}")
    check(
        "pptx-deck-basic",
        any(s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip() for s in deck.slides),
        "no slide has notes",
    )

    if FAILURES:
        print("corpus verification FAILED")
        for failure in FAILURES:
            print(f"  - {failure}")
        return 1
    print(f"corpus verification passed ({len(entries)} fixtures)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
