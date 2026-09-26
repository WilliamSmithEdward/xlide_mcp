"""The OOXML grid: conversions, formula shifting, and a read/write round trip.

The write path is the one worth testing hard. It splices into the workbook's own
sheet XML, so the thing that must hold is that everything it did not touch comes
back byte for byte.
"""

from __future__ import annotations

import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

from xlide_mcp.xlsx import (
    Workbook,
    XlsxError,
    column_to_index,
    formula_for_display,
    formula_for_file,
    index_to_column,
    parse_range,
    shift_formula,
)


@pytest.mark.parametrize(
    ("letters", "index"),
    [("A", 1), ("Z", 26), ("AA", 27), ("AZ", 52), ("BA", 53), ("XFD", 16384)],
)
def test_column_conversions_round_trip(letters: str, index: int) -> None:
    assert column_to_index(letters) == index
    assert index_to_column(index) == letters


def test_column_index_refuses_past_the_last_column() -> None:
    assert column_to_index("XFE") == 0


def test_parse_range_normalizes_corners_and_anchors() -> None:
    assert str(parse_range("$D$50:$A$1")) == "A1:D50"
    assert str(parse_range("B3")) == "B3"
    assert str(parse_range("Sheet1!A1:B2")) == "A1:B2"


def test_parse_range_refuses_nonsense() -> None:
    with pytest.raises(XlsxError):
        parse_range("not a range")


def test_future_functions_get_the_prefix_the_file_stores() -> None:
    assert formula_for_file("XLOOKUP(A1,B:B,C:C)") == "_xlfn.XLOOKUP(A1,B:B,C:C)"
    # FILTER only works on a worksheet, so it carries the longer prefix.
    assert formula_for_file("FILTER(A:A,B:B)") == "_xlfn._xlws.FILTER(A:A,B:B)"
    # SUM predates the future-function table and is stored as typed.
    assert formula_for_file("SUM(A1:A9)") == "SUM(A1:A9)"
    assert formula_for_display("_xlfn.XLOOKUP(A1,B:B,C:C)") == "XLOOKUP(A1,B:B,C:C)"


def test_a_function_name_inside_a_string_is_left_alone() -> None:
    assert formula_for_file('CONCAT("XLOOKUP(")') == '_xlfn.CONCAT("XLOOKUP(")'


def test_shift_formula_moves_relative_references_only() -> None:
    assert shift_formula("A1+$B$2", 2, 1) == "B3+$B$2"
    assert shift_formula("A1+B$2", 2, 0) == "A3+B$2"
    assert shift_formula("A1", 0, 0) == "A1"


def test_shift_formula_does_not_mistake_a_function_for_a_reference() -> None:
    """LOG10( parses as column LOG row 10 without the guard, and a shared formula
    filled down would silently become LOG11(."""
    assert shift_formula("LOG10(A1)", 1, 0) == "LOG10(A2)"
    assert shift_formula("SUM(A1:A9)", 1, 0) == "SUM(A2:A10)"


def test_shift_formula_leaves_string_literals_alone() -> None:
    assert shift_formula('IF(A1="B2","B2",A1)', 1, 0) == 'IF(A2="B2","B2",A2)'


def test_shift_formula_off_the_sheet_becomes_ref_error() -> None:
    assert shift_formula("A1", -5, 0) == "#REF!"


def test_read_and_write_cells(call: Callable[..., Any], plain_workbook: Path) -> None:
    written = call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[["Item", "Qty", "Each"], ["Widget", 3, 2.5], ["Gadget", 1, 10]],
    )
    assert written["cells_written"] == 9
    assert written["range"] == "A1:C3"
    assert written["recalculated"] is False
    assert written["cells_overwritten"] == 0

    read = call(
        "xlide_read_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        cell_range="A1:C3",
    )
    assert read["values"][0] == ["Item", "Qty", "Each"]
    assert read["values"][1] == ["Widget", 3, 2.5]


def test_writing_a_formula_stores_it_and_reads_it_back_as_typed(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[1], [2], ["=SUM(A1:A2)"], ["=XLOOKUP(1,A1:A2,A1:A2)"]],
    )
    read = call(
        "xlide_read_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        cell_range="A1:A4",
        include="both",
    )
    assert read["formulas"][2][0] == "=SUM(A1:A2)"
    # Stored with the prefix, handed back the way it was typed.
    assert read["formulas"][3][0] == "=XLOOKUP(1,A1:A2,A1:A2)"
    # No result has been calculated, because nothing ran Excel.
    assert read["values"][2][0] is None
    assert "note" in read


def test_a_write_asks_excel_to_recalculate_on_open(plain_workbook: Path) -> None:
    book = Workbook(plain_workbook)
    book.write("Sheet1", "A1", [["=1+1"]])
    book.save()
    with zipfile.ZipFile(plain_workbook) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in workbook_xml


def test_a_write_leaves_every_other_part_byte_for_byte(plain_workbook: Path) -> None:
    with zipfile.ZipFile(plain_workbook) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}

    book = Workbook(plain_workbook)
    book.write("Sheet1", "A1", [["hello"]])
    book.save()

    with zipfile.ZipFile(plain_workbook) as archive:
        after = {name: archive.read(name) for name in archive.namelist()}

    changed = {name for name in before if before.get(name) != after.get(name)}
    # Only the sheet written and the workbook part that carries the recalculation
    # flag are allowed to move. The Power Query part in particular must not.
    assert changed <= {"xl/worksheets/sheet1.xml", "xl/workbook.xml"}, changed
    assert set(before) - set(after) <= {"xl/calcChain.xml"}


def test_overwriting_content_is_reported(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[["first"], ["=1+1"]],
    )
    before = plain_workbook.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_cells", file_path=str(plain_workbook), sheet="Sheet1",
            start_cell="A1", data=[["second"], ["third"]],
        )
    assert "allow_overwrite=true" in refusal.value.message
    assert plain_workbook.read_bytes() == before
    again = call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[["second"], ["third"]],
        allow_overwrite=True,
    )
    assert again["cells_overwritten"] == 2
    assert again["formulas_replaced"] == 1
    assert "warning" in again


def test_large_write_reports_every_overwritten_cell(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    data = [["old"] * 101 for _ in range(200)]
    call("xlide_write_cells", file_path=str(plain_workbook), sheet="Sheet1",
         start_cell="A1", data=data)
    again = call("xlide_write_cells", file_path=str(plain_workbook), sheet="Sheet1",
                 start_cell="A1", data=[["new"] * 101 for _ in range(200)],
                 allow_overwrite=True)
    assert again["cells_overwritten"] == 20_200


def test_unknown_sheet_names_the_ones_that_exist(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_cells",
            file_path=str(plain_workbook),
            sheet="Nope",
            cell_range="A1",
        )
    assert "Sheet1" in refusal.value.message


def test_a_binary_workbook_is_not_read_from_the_package(
    call: Callable[..., Any], workspace: Path
) -> None:
    """A .xlsb keeps its grid in binary records rather than OOXML. Where Excel is
    not there to answer, the refusal says why and what would change it; where it
    is, Excel answers and the result says so. Either way the package reader does
    not pretend."""
    import pyopenvba

    from xlide_mcp import grid

    binary = workspace / "Binary.xlsb"
    with pyopenvba.ExcelFile.create_new(binary) as book:
        book.save()

    if not grid.excel_available():
        with pytest.raises(ToolFailure) as refusal:
            call("xlide_list_sheets", file_path=str(binary))
        assert ".xlsb" in refusal.value.message
        assert "VBA project in the file is still fully readable" in refusal.value.message
        return

    result = call("xlide_list_sheets", file_path=str(binary), timeout=240)
    assert result["source"] == "excel"
    assert result["sheets"]


def test_unreadable_pivots_are_unknown_rather_than_absent() -> None:
    from xlide_mcp.cells import _sheet_info

    class BrokenPivotSheet:
        name = "Summary"
        used_range = None
        visible = "visible"

        @property
        def pivot_tables(self) -> list[Any]:
            raise ValueError("pivot cache is damaged")

    summary = _sheet_info(BrokenPivotSheet()).summary()
    assert summary["pivot_tables"] is None
    assert "pivot cache is damaged" in summary["pivot_tables_error"]


def test_unreadable_sheet_visibility_is_unknown_rather_than_visible() -> None:
    from xlide_mcp.cells import _sheet_info

    class BrokenVisibility:
        name = "Summary"
        used_range = None

        @property
        def pivot_tables(self) -> list[Any]:
            return []

        @property
        def visible(self) -> str:
            raise ValueError("sheet state is damaged")

    summary = _sheet_info(BrokenVisibility()).summary()
    assert summary["hidden"] is None
    assert "sheet state is damaged" in summary["hidden_error"]

    class UnknownVisibility(BrokenVisibility):
        visible = "mystery"

    unknown = _sheet_info(UnknownVisibility()).summary()
    assert unknown["hidden"] is None
    assert "unknown state" in unknown["hidden_error"]


def test_unreadable_cell_presentation_is_not_returned_as_blank() -> None:
    from xlide_mcp.cells import CellsError, _runs, _text

    class Cell:
        a1 = "B2"
        reference = "B2"

        @property
        def text(self) -> str:
            raise ValueError("number format is damaged")

    class Sheet:
        def get_rich_text(self, _reference: str) -> None:
            raise ValueError("rich text record is damaged")

    with pytest.raises(CellsError) as refusal:
        _text(Cell())
    assert "B2" in str(refusal.value)
    assert "include='values'" in str(refusal.value)

    with pytest.raises(CellsError) as refusal:
        _runs(Sheet(), Cell())
    assert "B2" in str(refusal.value)
    assert "rich text record is damaged" in str(refusal.value)


def test_a_partial_binary_cell_grid_is_refused(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], workspace: Path
) -> None:
    import pyopenvba

    from xlide_mcp import grid

    binary = workspace / "Binary.xlsb"
    with pyopenvba.ExcelFile.create_new(binary) as book:
        book.save()

    monkeypatch.setattr(grid, "excel_available", lambda: True)
    monkeypatch.setattr(grid, "read_cells", lambda *_args: [[1]])
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_cells", file_path=str(binary), sheet="Sheet1",
            cell_range="A1:B2",
        )
    assert "No partial cell values were returned" in refusal.value.message
    assert "requires 2 rows and 2 columns" in refusal.value.message


def test_a_huge_binary_range_is_refused_before_excel_starts(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], workspace: Path
) -> None:
    import pyopenvba

    from xlide_mcp import grid

    binary = workspace / "Binary.xlsb"
    with pyopenvba.ExcelFile.create_new(binary) as book:
        book.save()

    monkeypatch.setattr(grid, "excel_available", lambda: True)

    def should_not_start(*_args: Any) -> Any:
        raise AssertionError("Excel must not start for an oversized request")

    monkeypatch.setattr(grid, "read_cells", should_not_start)
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_cells", file_path=str(binary), sheet="Sheet1",
            cell_range="A1:Z10000",
        )
    assert "over the 20,000" in refusal.value.message
    assert "Read it in blocks" in refusal.value.message


def test_a_huge_range_is_refused_rather_than_returned(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_cells",
            file_path=str(plain_workbook),
            sheet="Sheet1",
            cell_range="A1:Z10000",
        )
    assert "Read it in blocks" in refusal.value.message
