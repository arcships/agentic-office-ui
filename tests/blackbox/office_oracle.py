"""LibreOffice oracle check for the review corpus.

Converts DOCX/XLSX/PPTX fixtures with headless LibreOffice and asserts the PDF
page counts recorded in `office_oracle_baseline.json`. This is a fast,
Playwright-free sanity check that the fixtures are valid and that the oracle
result has not drifted; it also prints the known local viewer page count so the
pagination gap stays visible.

    uv run python tests/blackbox/office_oracle.py

Requires `soffice` and `pdfinfo` (LibreOffice + poppler). If either is missing
the script prints SKIP and exits 0.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BASELINE = Path(__file__).with_name("office_oracle_baseline.json")


def find_tool(name: str) -> str | None:
    return shutil.which(name)


def pdf_pages(pdf: Path, pdfinfo: str) -> int | None:
    result = subprocess.run([pdfinfo, str(pdf)], capture_output=True, text=True, timeout=30)
    for line in result.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":", 1)[1].strip())
    return None


def convert(soffice: str, profile: str, source: Path, outdir: Path) -> Path | None:
    result = subprocess.run(
        [
            soffice,
            f"-env:UserInstallation=file://{profile}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(source),
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    pdf = outdir / f"{source.stem}.pdf"
    return pdf if pdf.exists() else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", help="write the JSON report to this path")
    args = parser.parse_args()

    soffice = find_tool("soffice")
    pdfinfo = find_tool("pdfinfo")
    if not soffice or not pdfinfo:
        print("SKIP: soffice/pdfinfo not available")
        return 0

    baseline = json.loads(BASELINE.read_text(encoding="utf-8"))
    failures: list[str] = []
    report: dict[str, dict] = {}

    with tempfile.TemporaryDirectory(prefix="office-oracle-") as tmp:
        profile = str(Path(tmp) / "lo-profile")
        for name, entry in baseline["fixtures"].items():
            source = ROOT / entry["file"]
            if not source.exists():
                failures.append(f"{name}: missing fixture {entry['file']}")
                continue
            pdf = convert(soffice, profile, source, Path(tmp))
            if pdf is None:
                failures.append(f"{name}: LibreOffice could not convert the file")
                continue
            pages = pdf_pages(pdf, pdfinfo)
            report[name] = {
                "file": entry["file"],
                "oracle_pages": pages,
                "expected_pages": entry["pages"],
                "local_pages": entry.get("local_pages"),
            }
            if pages != entry["pages"]:
                failures.append(f"{name}: oracle pages {pages} != baseline {entry['pages']}")
            if entry.get("local_pages") is not None and pages is not None:
                report[name]["gap"] = pages - entry["local_pages"]

    payload = {
        "suite": "office-oracle",
        "oracle": baseline.get("oracle"),
        "cases": report,
        "failures": failures,
        "verdict": "FAIL" if failures else "PASS",
    }
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if failures:
        print("oracle FAILED")
        for failure in failures:
            print(f"  - {failure}")
        return 1
    print("oracle PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
