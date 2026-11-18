from __future__ import annotations

import json
import os
import posixpath
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from openpyxl import __version__ as openpyxl_version
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[2]
EXPECTED_OPENPYXL = "3.1.5"
CHART_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)


def resolve_part(rels_path: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    if rels_path == "_rels/.rels":
        owner = ""
    else:
        prefix, _, file_name = rels_path.rpartition("/_rels/")
        owner = f"{prefix}/{file_name[:-5]}" if prefix else file_name[:-5]
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


def validate_relationship_targets(archive: zipfile.ZipFile) -> int:
    names = set(archive.namelist())
    checked = 0
    for name in sorted(item for item in names if item.endswith(".rels")):
        root = ElementTree.fromstring(archive.read(name))
        for relationship in root:
            if relationship.get("TargetMode") == "External":
                continue
            target = relationship.get("Target") or ""
            if not target or target.startswith("#"):
                continue
            resolved = resolve_part(name, target)
            assert resolved in names, f"{name} target {target} is missing ({resolved})"
            checked += 1
    return checked


def main() -> int:
    assert openpyxl_version == EXPECTED_OPENPYXL, (
        f"expected pinned openpyxl {EXPECTED_OPENPYXL}, got {openpyxl_version}"
    )

    output_dir = Path(tempfile.mkdtemp(prefix="xlsx-export-openpyxl-"))
    output_path = output_dir / "roundtrip.xlsx"
    subprocess.run(
        [
            os.environ.get("NODE", "node"),
            str(ROOT / "scripts" / "ci" / "export-xlsx-roundtrip.mjs"),
            str(output_path),
        ],
        cwd=ROOT,
        check=True,
    )

    with zipfile.ZipFile(output_path) as archive:
        names = set(archive.namelist())
        assert "xl/charts/chart1.xml" in names, "worksheet chart part is missing"
        assert "xl/charts/chart2.xml" in names, "chartsheet chart part is missing"
        assert "xl/media/image1.png" in names, "drawing media part is missing"

        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert sheet_xml.count("<drawing") == 1, (
            "a worksheet must contain exactly one <drawing> element"
        )

        content_types = archive.read("[Content_Types].xml").decode("utf-8")
        assert "/xl/charts/chart1.xml" in content_types
        assert CHART_CONTENT_TYPE in content_types

        relationship_count = validate_relationship_targets(archive)
        assert relationship_count > 0, "the package must contain relationships"

        drawing_rels = [
            name
            for name in names
            if name.startswith("xl/drawings/_rels/") and name.endswith(".rels")
        ]
        assert drawing_rels, "the package must contain drawing relationships"
        chart_relationships = 0
        for name in drawing_rels:
            root = ElementTree.fromstring(archive.read(name))
            chart_relationships += sum(
                1 for relationship in root if (relationship.get("Type") or "").endswith("/chart")
            )
        assert chart_relationships >= 2, "both worksheet and chartsheet charts must be linked"

    workbook = load_workbook(output_path, data_only=False)
    try:
        assert workbook.sheetnames == ["Dashboard", "Revenue Chart"]
        assert workbook["Dashboard"]["A1"].value == "OPENPYXL-ROUNDTRIP"
        charts = getattr(workbook["Dashboard"], "_charts", [])
        assert charts, "openpyxl must reopen the worksheet chart"
        chart_title = charts[0].title
        title_text = "".join(
            run.t or ""
            for paragraph in (chart_title.tx.rich.p if chart_title.tx and chart_title.tx.rich else [])
            for run in (paragraph.r or [])
        )
        assert title_text == "Openpyxl roundtrip chart", title_text
    finally:
        workbook.close()

    summary = {
        "openpyxl": openpyxl_version,
        "output": str(output_path),
        "bytes": output_path.stat().st_size,
        "relationships": relationship_count,
        "chartRelationships": chart_relationships,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
