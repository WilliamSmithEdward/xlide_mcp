"""Notes, threaded comments and autofilters: pyOfficeEditor 0.3's additions.

A filter is the one to watch. Excel does not apply one when it opens a
workbook, it shows the rows as the file marks them, so a filter written with its
criteria and no hidden rows opens with every row showing. The rows are worked
out here, and these tests read them back from the rows themselves.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

SALES = [
    ["Region", "Amount"],
    ["West", 120],
    ["East", 80],
    ["West", 45],
    ["North", 300],
    ["East", 10],
]


@pytest.fixture
def sales(call: Callable[..., Any], plain_workbook: Path) -> Path:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=SALES,
    )
    return plain_workbook


def _hidden_rows(path: Path) -> set[int]:
    import pyofficeeditor.excel as excel

    with excel.Workbook.open(path) as book:
        sheet = book.sheets[0]
        return {row for row in range(1, len(SALES) + 1) if sheet.row_hidden(row)}


# ----------------------------------------------------------------- comments


def test_a_note_is_written_read_and_removed(call: Callable[..., Any], sales: Path) -> None:
    set_note = call(
        "xlide_manage_comment",
        file_path=str(sales),
        sheet="Sheet1",
        action="set",
        cell="b2",
        text="Includes the June correction.",
        author="Analyst",
    )
    assert set_note["comments"] == [
        {
            "cell": "B2",
            "kind": "note",
            "text": "Includes the June correction.",
            "author": "Analyst",
            "visible": False,
        }
    ]
    listed = call("xlide_manage_comment", file_path=str(sales), sheet="Sheet1")
    assert listed["count"] == 1

    call("xlide_manage_comment", file_path=str(sales), sheet="Sheet1", action="remove", cell="B2")
    assert call("xlide_manage_comment", file_path=str(sales), sheet="Sheet1")["count"] == 0


def test_a_thread_takes_replies_and_resolves(call: Callable[..., Any], sales: Path) -> None:
    common = {"file_path": str(sales), "sheet": "Sheet1", "cell": "B5"}
    call("xlide_manage_comment", action="set", kind="thread", text="Is 300 right?", **common)
    call("xlide_manage_comment", action="reply", text="Yes, one large order.", **common)
    resolved = call("xlide_manage_comment", action="resolve", **common)
    (thread,) = resolved["comments"]
    assert thread["kind"] == "thread"
    assert thread["resolved"] is True
    assert [reply["text"] for reply in thread["replies"]] == ["Yes, one large order."]


def test_an_absolute_reference_is_the_same_cell(call: Callable[..., Any], sales: Path) -> None:
    """$B$4 was written to B4 and then looked for as $B$4, so the result came back
    with no comments, as if the write had not happened."""
    written = call(
        "xlide_manage_comment",
        file_path=str(sales),
        sheet="Sheet1",
        action="set",
        cell="$B$4",
        text="Checked",
        author="A",
    )
    assert written["cell"] == "B4"
    assert [entry["text"] for entry in written["comments"]] == ["Checked"]
    listed = call("xlide_manage_comment", file_path=str(sales), sheet="Sheet1", cell="$b$4")
    assert listed["count"] == 1


def test_removing_nothing_says_so(call: Callable[..., Any], sales: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_comment", file_path=str(sales), sheet="Sheet1", action="remove",
             cell="C9")
    assert "has no note or comment" in refusal.value.message


# ---------------------------------------------------------------- autofilters


def test_a_filter_hides_the_rows_it_filters_out(call: Callable[..., Any], sales: Path) -> None:
    result = call(
        "xlide_manage_filter",
        file_path=str(sales),
        sheet="Sheet1",
        action="set",
        cell_range="A1:B6",
        column="Region",
        criteria1="=West",
    )
    assert result["rows"] == {"hidden": 3, "shown": 2}
    assert _hidden_rows(sales) == {3, 5, 6}
    (listed,) = result["filters"]
    assert listed["range"] == "A1:B6"
    assert listed["columns"] == [{"column": 1, "values": ["West"]}]


def test_a_second_column_keeps_the_first(call: Callable[..., Any], sales: Path) -> None:
    common = {"file_path": str(sales), "sheet": "Sheet1", "action": "set"}
    call("xlide_manage_filter", cell_range="A1:B6", column="Region", criteria1="=West", **common)
    both = call("xlide_manage_filter", column="2", criteria1=">=100", **common)
    assert _hidden_rows(sales) == {3, 4, 5, 6}, "only West at 100 or more is left"
    assert [entry["column"] for entry in both["filters"][0]["columns"]] == [1, 2]


def test_top_and_above_average_are_worked_out(call: Callable[..., Any], sales: Path) -> None:
    common = {"file_path": str(sales), "sheet": "Sheet1", "action": "set", "column": "Amount"}
    call("xlide_manage_filter", cell_range="A1:B6", top=2, **common)
    assert _hidden_rows(sales) == {3, 4, 6}, "300 and 120 are the top two"
    call("xlide_manage_filter", dynamic="aboveAverage", **common)
    assert _hidden_rows(sales) == {3, 4, 6}, "the average is 111, so the same two"


def test_clearing_shows_every_row(call: Callable[..., Any], sales: Path) -> None:
    common = {"file_path": str(sales), "sheet": "Sheet1"}
    call("xlide_manage_filter", action="set", cell_range="A1:B6", column="Region",
         criteria1="=East", **common)
    call("xlide_manage_filter", action="clear", **common)
    assert _hidden_rows(sales) == set()
    assert call("xlide_manage_filter", **common)["filters"] == []


def test_a_table_is_filtered_by_its_column_name(call: Callable[..., Any], sales: Path) -> None:
    call("xlide_manage_table", file_path=str(sales), action="add", sheet="Sheet1",
         table_name="Sales", cell_range="A1:B6")
    result = call(
        "xlide_manage_filter",
        file_path=str(sales),
        sheet="Sheet1",
        action="set",
        table="sales",
        column="Amount",
        criteria1="<50",
    )
    assert _hidden_rows(sales) == {2, 3, 5}
    assert result["filters"][0]["table"] == "Sales"


def test_a_filter_needs_one_kind_of_criterion(call: Callable[..., Any], sales: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_filter", file_path=str(sales), sheet="Sheet1", action="set",
             cell_range="A1:B6", column="Amount", criteria1=">1", top=3)
    assert "exactly one of criteria1, top or dynamic" in refusal.value.message

    with pytest.raises(ToolFailure) as unknown:
        call("xlide_manage_filter", file_path=str(sales), sheet="Sheet1", action="set",
             cell_range="A1:B6", column="Price", criteria1=">1")
    assert "No column headed 'Price'" in unknown.value.message
