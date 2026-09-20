"""Sheets, and the rows and columns in them.

The interesting assertions here are about what moves with a change rather than
the change itself. Renaming a sheet has to take the formulas that named it;
inserting a row has to take every reference below it. Both are upstream's work,
and both are exactly what a caller would not check by hand.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


def formula(path: Path, sheet: str, reference: str) -> str | None:
    from xlide_mcp import cells

    return cells.read(path, sheet, reference).formulas[0][0]


def test_a_sheet_can_be_added_named_and_moved(
    call: Callable[..., Any], workbook: Path
) -> None:
    added = call(
        "xlide_manage_sheet", file_path=str(workbook), action="add", sheet="Summary"
    )
    assert added["created"] is True
    assert added["sheets"] == ["Sheet1", "Summary"]

    moved = call(
        "xlide_manage_sheet",
        file_path=str(workbook),
        action="move",
        sheet="Summary",
        index=0,
    )
    assert moved["sheets"] == ["Summary", "Sheet1"]


def test_renaming_a_sheet_takes_the_formulas_with_it(
    call: Callable[..., Any], workbook: Path
) -> None:
    """The reason renaming goes through the library rather than a string edit.
    A formula naming the old sheet has to follow, or the workbook opens broken."""
    call("xlide_manage_sheet", file_path=str(workbook), action="add", sheet="Data")
    call(
        "xlide_write_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[["=SUM(Data!A1:A5)"]],
    )

    call(
        "xlide_manage_sheet",
        file_path=str(workbook),
        action="rename",
        sheet="Data",
        new_name="Q1 Data",
    )
    # A name with a space has to come back quoted, or Excel cannot parse it.
    assert formula(workbook, "Sheet1", "A1") == "=SUM('Q1 Data'!A1:A5)"


def test_a_sheet_can_be_hidden_and_shown(call: Callable[..., Any], workbook: Path) -> None:
    call("xlide_manage_sheet", file_path=str(workbook), action="add", sheet="Working")
    call("xlide_manage_sheet", file_path=str(workbook), action="hide", sheet="Working")

    listed = call("xlide_list_sheets", file_path=str(workbook))["sheets"]
    assert next(s for s in listed if s["name"] == "Working")["hidden"] is True

    call("xlide_manage_sheet", file_path=str(workbook), action="show", sheet="Working")
    listed = call("xlide_list_sheets", file_path=str(workbook))["sheets"]
    assert next(s for s in listed if s["name"] == "Working")["hidden"] is False


def test_hiding_the_last_visible_sheet_is_refused(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Excel refuses to open a workbook with every sheet hidden, so the refusal
    belongs here, where it can still be explained, rather than there."""
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_sheet", file_path=str(workbook), action="hide", sheet="Sheet1")
    assert "only visible sheet" in refusal.value.message


def test_a_sheet_can_be_protected_and_released(
    call: Callable[..., Any], workbook: Path
) -> None:
    protected = call(
        "xlide_manage_sheet", file_path=str(workbook), action="protect", sheet="Sheet1"
    )
    assert protected["protected"] is True

    released = call(
        "xlide_manage_sheet", file_path=str(workbook), action="unprotect", sheet="Sheet1"
    )
    assert released["protected"] is False


def test_inserting_rows_moves_the_references_below(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[10], [20], ["=SUM(A1:A2)"]],
    )
    call(
        "xlide_manage_rows_columns",
        file_path=str(workbook),
        sheet="Sheet1",
        action="insert",
        which="rows",
        first=1,
        count=2,
    )
    assert formula(workbook, "Sheet1", "A5") == "=SUM(A3:A4)"


def test_deleting_columns_breaks_what_pointed_into_them(
    call: Callable[..., Any], workbook: Path
) -> None:
    """#REF! is the right answer, not a failure. Excel does exactly this, and a
    tool that silently repointed the formula somewhere else would be worse."""
    call(
        "xlide_write_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[1, 2, "=SUM(A1:B1)"]],
    )
    call(
        "xlide_manage_rows_columns",
        file_path=str(workbook),
        sheet="Sheet1",
        action="delete",
        which="columns",
        first=1,
        count=2,
    )
    assert "#REF!" in (formula(workbook, "Sheet1", "A1") or "")


def test_rows_can_be_sized_hidden_and_grouped(
    call: Callable[..., Any], workbook: Path
) -> None:
    from xlide_mcp import cells

    for action, extra in (
        ("resize", {"size": 30}),
        ("hide", {}),
        ("group", {}),
    ):
        call(
            "xlide_manage_rows_columns",
            file_path=str(workbook),
            sheet="Sheet1",
            action=action,
            which="rows",
            first=2,
            count=3,
            **extra,
        )

    with cells.open_workbook(workbook) as book:
        sheet = cells.sheet_named(book, "Sheet1")
        assert sheet.row_height(2) == 30
        assert sheet.row_hidden(2) is True
        assert sheet.row_outline_level(2) >= 1


def test_an_unknown_action_lists_the_real_ones(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_sheet", file_path=str(workbook), action="obliterate", sheet="Sheet1")
    assert "rename" in refusal.value.message


def test_rows_or_columns_has_to_be_one_of_them(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_rows_columns",
            file_path=str(workbook),
            sheet="Sheet1",
            action="insert",
            which="diagonals",
            first=1,
        )
    assert "'rows' or 'columns'" in refusal.value.message
