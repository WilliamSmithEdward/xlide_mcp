"""The server exposes pyOfficeEditor's measured Excel range operations."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


def test_excel_error_checks_find_number_stored_as_text(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_cells", file_path=str(workbook), sheet="Sheet1", start_cell="A1",
        data=[["123"]],
    )
    result = call("xlide_check_cells", file_path=str(workbook), sheet="Sheet1")
    assert {("A1", "numberStoredAsText")} <= {
        (entry["cell"], entry["rule"]) for entry in result["checks"]
    }


def test_sort_and_remove_duplicates_keep_headers_and_first_rows(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_cells", file_path=str(workbook), sheet="Sheet1", start_cell="A1",
        data=[["Name", "Score"], ["B", 2], ["A", 1], ["B", 2], ["C", 3]],
    )
    sorted_rows = call(
        "xlide_sort_rows", file_path=str(workbook), sheet="Sheet1",
        cell_range="A1:B5", header=True, keys=[{"column": "A"}],
    )
    assert sorted_rows["saved"] is True
    assert sorted_rows["recalculated"] is False
    read = call(
        "xlide_read_cells", file_path=str(workbook), sheet="Sheet1", cell_range="A1:B5"
    )
    assert read["values"] == [
        ["Name", "Score"], ["A", 1], ["B", 2], ["B", 2], ["C", 3]
    ]

    removed = call(
        "xlide_remove_duplicates", file_path=str(workbook), sheet="Sheet1",
        cell_range="A1:B5", columns=["A", "B"], header=True,
    )
    assert removed["rows_removed"] == 1
    read = call(
        "xlide_read_cells", file_path=str(workbook), sheet="Sheet1", cell_range="A1:B5"
    )
    assert read["values"] == [
        ["Name", "Score"], ["A", 1], ["B", 2], ["C", 3], [None, None]
    ]


def test_copy_requires_overwrite_acknowledgment_and_moves_formulas(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_cells", file_path=str(workbook), sheet="Sheet1", start_cell="A1",
        data=[[5, "=A1*2"]],
    )
    before = workbook.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_copy_cells", file_path=str(workbook), sheet="Sheet1",
            source_range="A1:B1", destination="A2",
        )
    assert "allow_overwrite=true" in refusal.value.message
    assert workbook.read_bytes() == before

    copied = call(
        "xlide_copy_cells", file_path=str(workbook), sheet="Sheet1",
        source_range="A1:B1", destination="A2", allow_overwrite=True,
    )
    assert copied["destination_range"] == "A2:B2"
    assert copied["recalculated"] is False
    read = call(
        "xlide_read_cells", file_path=str(workbook), sheet="Sheet1",
        cell_range="A2:B2", include="both",
    )
    assert read["values"][0][0] == 5
    assert read["formulas"][0][1] == "=A2*2"
