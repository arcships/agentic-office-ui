from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import __version__ as openpyxl_version
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[2]
EXPECTED_OPENPYXL = "3.1.5"


def main() -> int:
    assert openpyxl_version == EXPECTED_OPENPYXL, (
        f"expected pinned openpyxl {EXPECTED_OPENPYXL}, got {openpyxl_version}"
    )

    output_dir = Path(tempfile.mkdtemp(prefix="xlsx-controller-export-"))
    output_path = output_dir / "controller-export.xlsx"
    result = subprocess.run(
        [
            os.environ.get("NODE", "node"),
            str(ROOT / "scripts" / "ci" / "export-xlsx-controller.mjs"),
            str(output_path),
        ],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise AssertionError(
            "controller export failed:\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )
    driver = json.loads(result.stdout.strip().splitlines()[-1])

    workbook = load_workbook(output_path, data_only=False)
    try:
        dashboard = workbook["Dashboard"]
        assert dashboard["A1"].value == "CONTROLLER-EXPORT", dashboard["A1"].value
        assert dashboard["A1"].font.bold is True, dashboard["A1"].font
        assert str(dashboard["A1"].fill.fgColor.rgb).endswith("1C3A5E"), dashboard["A1"].fill
        assert dashboard["E6"].value == "CONTROLLER-EXPORT", dashboard["E6"].value
        assert dashboard["E6"].font.bold is True, dashboard["E6"].font
        assert dashboard["E6"].number_format == "0.00%", dashboard["E6"].number_format

        merged = {str(range_) for range_ in dashboard.merged_cells.ranges}
        assert "A14:C15" in merged, f"the new merge is missing: {sorted(merged)}"
        assert "A7:C8" in merged, f"the fixture merge was lost: {sorted(merged)}"

        # Read-only fixture features must survive the editable export.
        assert dashboard.freeze_panes == "C3", dashboard.freeze_panes
        assert len(dashboard.data_validations.dataValidation) == 1
        assert dashboard.data_validations.dataValidation[0].formula1 == '"1,2,3"'
        assert workbook.sheetnames == ["Dashboard", "Revenue Chart"]
    finally:
        workbook.close()

    summary = {
        "openpyxl": openpyxl_version,
        "driver": driver,
        "bytes": output_path.stat().st_size,
        "editedCell": "E6",
        "editedHeader": "A1",
        "numberFormat": "0.00%",
        "merged": ["A7:C8", "A14:C15"],
        "freezePanes": "C3",
        "dataValidations": 1,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
