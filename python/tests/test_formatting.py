"""How a range looks, and the guards on changing it.

The property worth protecting here is that only what was asked for changes. A
tool that rebuilt a cell's whole format to set one flag would quietly flatten
every number format under a header row the first time somebody made it bold, and
nothing would say so.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


def formats(path: Path, sheet: str, reference: str) -> Any:
    """The stored format of one cell, read back through the library itself."""
    from xlide_mcp import cells

    with cells.open_workbook(path) as book:
        return cells.sheet_named(book, sheet)[reference].format


def test_a_header_row_can_be_made_bold(call: Callable[..., Any], workbook: Path) -> None:
    call(
        "xlide_write_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[["Region", "Total"]],
    )
    answer = call(
        "xlide_format_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        cell_range="A1:B1",
        bold=True,
    )
    assert answer["applied"] == ["font"]
    assert answer["saved"] is True
    assert formats(workbook, "Sheet1", "A1").font.bold is True


def test_only_what_was_asked_for_changes(call: Callable[..., Any], workbook: Path) -> None:
    """The regression this tool is shaped to avoid. Setting bold must not take
    the number format with it."""
    call(
        "xlide_write_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[1234.5]],
    )
    call(
        "xlide_format_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        cell_range="A1",
        number_format="#,##0.00",
    )
    call(
        "xlide_format_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        cell_range="A1",
        bold=True,
    )

    stored = formats(workbook, "Sheet1", "A1")
    assert stored.font.bold is True
    assert stored.number_format == "#,##0.00"


def test_several_changes_land_in_one_call(call: Callable[..., Any], workbook: Path) -> None:
    answer = call(
        "xlide_format_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        cell_range="A1:B2",
        bold=True,
        fill_color="FFFF00",
        border_style="thin",
        horizontal="center",
        number_format="0%",
    )
    assert answer["applied"] == ["font", "fill", "border", "alignment", "number format"]

    stored = formats(workbook, "Sheet1", "A1")
    assert stored.font.bold is True
    assert stored.alignment.horizontal == "center"
    assert stored.number_format == "0%"


def test_a_colour_is_taken_with_or_without_its_hash(
    call: Callable[..., Any], workbook: Path
) -> None:
    """A calling model writes '#FF0000' about as often as 'FF0000', and refusing
    one of them is a round trip spent on punctuation."""
    for colour in ("#FF0000", "FF0000", "ffff0000"):
        call(
            "xlide_format_cells",
            file_path=str(workbook),
            sheet="Sheet1",
            cell_range="A1",
            fill_color=colour,
        )


def test_a_colour_that_is_not_one_says_so(call: Callable[..., Any], workbook: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_format_cells",
            file_path=str(workbook),
            sheet="Sheet1",
            cell_range="A1",
            fill_color="cornflower",
        )
    assert "RRGGBB" in refusal.value.message


def test_an_unknown_border_style_lists_the_real_ones(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_format_cells",
            file_path=str(workbook),
            sheet="Sheet1",
            cell_range="A1",
            border_style="squiggly",
        )
    assert "thin" in refusal.value.message


def test_a_call_that_changes_nothing_is_refused(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Saying nothing happened is better than saving a file for no reason and
    reporting success."""
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_format_cells",
            file_path=str(workbook),
            sheet="Sheet1",
            cell_range="A1",
        )
    assert "Nothing to change" in refusal.value.message


def test_cells_can_be_merged_and_split_again(
    call: Callable[..., Any], workbook: Path
) -> None:
    from xlide_mcp import cells

    call(
        "xlide_format_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        cell_range="A1:C1",
        merge="merge",
    )
    with cells.open_workbook(workbook) as book:
        assert [str(r) for r in cells.sheet_named(book, "Sheet1").merged_ranges] == ["A1:C1"]

    call(
        "xlide_format_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        cell_range="A1:C1",
        merge="unmerge",
    )
    with cells.open_workbook(workbook) as book:
        assert list(cells.sheet_named(book, "Sheet1").merged_ranges) == []


def test_an_unknown_sheet_names_the_ones_that_exist(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_format_cells",
            file_path=str(workbook),
            sheet="Ghost",
            cell_range="A1",
            bold=True,
        )
    assert "Sheet1" in refusal.value.message


def test_a_binary_workbook_is_refused_with_the_reason(
    call: Callable[..., Any], workspace: Path
) -> None:
    """A .xlsb keeps its grid outside the OOXML package, so the package reader
    cannot format it. The refusal says which formats do work."""
    import pyopenvba

    binary = workspace / "Binary.xlsb"
    with pyopenvba.ExcelFile.create_new(binary) as book:
        book.save()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_format_cells",
            file_path=str(binary),
            sheet="Sheet1",
            cell_range="A1",
            bold=True,
        )
    assert ".xlsm" in refusal.value.message
