"""Tables, names, validation, conditional rules and links.

These are the parts of a workbook that say what it means rather than what it
holds, and the assertions worth making are about the translation between how a
caller would phrase something and how Excel stores it. A dropdown of three
colours is the string `"Red,Green,Blue"` with its quotes; the same argument
naming cells is a reference and must not be quoted. Getting that backwards
produces a workbook that opens and then behaves wrongly, which is the failure
mode nobody notices.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


@pytest.fixture
def sheet_with_data(call: Callable[..., Any], workbook: Path) -> Path:
    call(
        "xlide_write_cells",
        file_path=str(workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[["Region", "Amount"], ["North", 120], ["South", 80]],
    )
    return workbook


# ----------------------------------------------------------------------- tables


def test_a_table_takes_its_columns_from_the_header_row(
    call: Callable[..., Any], sheet_with_data: Path
) -> None:
    added = call(
        "xlide_manage_table",
        file_path=str(sheet_with_data),
        action="add",
        sheet="Sheet1",
        table_name="Sales",
        cell_range="A1:B3",
    )
    assert added["columns"] == ["Region", "Amount"]

    listed = call("xlide_manage_table", file_path=str(sheet_with_data), action="list")
    assert [t["name"] for t in listed["tables"]] == ["Sales"]

    call(
        "xlide_manage_table",
        file_path=str(sheet_with_data),
        action="remove",
        table_name="Sales",
    )
    assert call("xlide_manage_table", file_path=str(sheet_with_data))["count"] == 0


def test_adding_a_table_without_a_range_says_what_is_missing(
    call: Callable[..., Any], sheet_with_data: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_table",
            file_path=str(sheet_with_data),
            action="add",
            sheet="Sheet1",
            table_name="Sales",
        )
    assert "header row" in refusal.value.message


# ------------------------------------------------------------------ defined names


def test_a_defined_name_round_trips(call: Callable[..., Any], workbook: Path) -> None:
    call(
        "xlide_manage_name",
        file_path=str(workbook),
        action="add",
        name="TaxRate",
        refers_to="Sheet1!$B$2",
    )
    listed = call("xlide_manage_name", file_path=str(workbook))["names"]
    entry = next(n for n in listed if n["name"] == "TaxRate")
    assert entry["refers_to"] == "Sheet1!$B$2"
    assert entry["scope"] == "(workbook)"

    call("xlide_manage_name", file_path=str(workbook), action="remove", name="TaxRate")
    assert all(n["name"] != "TaxRate" for n in call(
        "xlide_manage_name", file_path=str(workbook)
    )["names"])


# -------------------------------------------------------------------- validation


def test_a_typed_out_dropdown_is_quoted_and_a_reference_is_not(
    call: Callable[..., Any], workbook: Path
) -> None:
    """The one translation in this tool. Excel stores an inline list as a single
    quoted string and a cell-driven list as a bare reference, and sending either
    one in the other's shape gives a dropdown that is empty or literal."""
    typed = call(
        "xlide_manage_validation",
        file_path=str(workbook),
        sheet="Sheet1",
        action="add",
        cell_range="C1:C9",
        kind="list",
        formula1="Red,Green,Blue",
    )
    assert typed["formula1"] == '"Red,Green,Blue"'

    referenced = call(
        "xlide_manage_validation",
        file_path=str(workbook),
        sheet="Sheet1",
        action="add",
        cell_range="D1:D9",
        kind="list",
        formula1="$H$1:$H$9",
    )
    assert referenced["formula1"] == "$H$1:$H$9"


def test_validation_can_be_listed_and_cleared(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_manage_validation",
        file_path=str(workbook),
        sheet="Sheet1",
        action="add",
        cell_range="C1:C9",
        kind="list",
        formula1="Yes,No",
    )
    assert call(
        "xlide_manage_validation", file_path=str(workbook), sheet="Sheet1"
    )["count"] == 1

    call(
        "xlide_manage_validation",
        file_path=str(workbook),
        sheet="Sheet1",
        action="clear",
    )
    assert call(
        "xlide_manage_validation", file_path=str(workbook), sheet="Sheet1"
    )["count"] == 0


def test_validation_needs_something_to_allow(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_validation",
            file_path=str(workbook),
            sheet="Sheet1",
            action="add",
            cell_range="C1:C9",
            kind="list",
        )
    assert "formula1 is required" in refusal.value.message


# -------------------------------------------------------- conditional formatting


def test_a_rule_paints_cells_that_compare_true(
    call: Callable[..., Any], sheet_with_data: Path
) -> None:
    call(
        "xlide_manage_conditional_format",
        file_path=str(sheet_with_data),
        sheet="Sheet1",
        action="add",
        cell_range="B2:B3",
        rule="cell_is",
        operator="greaterThan",
        value="100",
        fill_color="FFC7CE",
    )
    listed = call(
        "xlide_manage_conditional_format", file_path=str(sheet_with_data), sheet="Sheet1"
    )
    assert listed["count"] == 1


def test_a_rule_with_no_paint_is_refused(
    call: Callable[..., Any], sheet_with_data: Path
) -> None:
    """A rule that highlights nothing is a rule nobody can see, and Excel will
    happily store one. Saying so is more use than saving it."""
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_conditional_format",
            file_path=str(sheet_with_data),
            sheet="Sheet1",
            action="add",
            cell_range="B2:B3",
            rule="cell_is",
            operator="greaterThan",
            value="100",
        )
    assert "highlights nothing" in refusal.value.message


def test_a_colour_scale_needs_no_paint_of_its_own(
    call: Callable[..., Any], sheet_with_data: Path
) -> None:
    """The exception to the rule above: a gradient is its own paint."""
    call(
        "xlide_manage_conditional_format",
        file_path=str(sheet_with_data),
        sheet="Sheet1",
        action="add",
        cell_range="B2:B3",
        rule="color_scale",
    )


# -------------------------------------------------------------------- hyperlinks


def test_a_link_out_and_a_link_inside_both_work(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_manage_hyperlink",
        file_path=str(workbook),
        sheet="Sheet1",
        action="add",
        cell_range="A1",
        target="https://example.com",
        display="Docs",
    )
    call(
        "xlide_manage_hyperlink",
        file_path=str(workbook),
        sheet="Sheet1",
        action="add",
        cell_range="A2",
        location="Sheet1!A10",
    )
    listed = call("xlide_manage_hyperlink", file_path=str(workbook), sheet="Sheet1")
    assert listed["count"] == 2

    call(
        "xlide_manage_hyperlink",
        file_path=str(workbook),
        sheet="Sheet1",
        action="remove",
        cell_range="A1",
    )
    assert call(
        "xlide_manage_hyperlink", file_path=str(workbook), sheet="Sheet1"
    )["count"] == 1


def test_a_link_to_nowhere_says_which_argument_is_missing(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_hyperlink",
            file_path=str(workbook),
            sheet="Sheet1",
            action="add",
            cell_range="A1",
        )
    assert "target" in refusal.value.message and "location" in refusal.value.message


# -------------------------------------------------------------------- page setup


def test_page_setup_reads_without_changing_anything(
    call: Callable[..., Any], workbook: Path
) -> None:
    answer = call("xlide_page_setup", file_path=str(workbook), sheet="Sheet1")
    assert answer["sheet"] == "Sheet1"
    assert "margins" in answer
    assert isinstance(answer["print_area"], list)
