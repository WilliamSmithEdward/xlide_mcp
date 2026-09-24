"""Working formulas out without Excel: pyOfficeEditor's formula engine.

A value read from the file is what Excel last calculated. Before pyOfficeEditor
0.3 nothing here could do better, so a formula written through this server had
no value until Excel next opened the workbook. The engine changes what can be
answered, not what is claimed: a calculated value says it was, and a cell the
engine cannot work out keeps Excel's cached value and is named with the reason.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


@pytest.fixture
def ledger(call: Callable[..., Any], plain_workbook: Path) -> Path:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[2], [3], ["=SUM(A1:A2)"], ['=WEBSERVICE("http://example.com")'], ["=A4*2"]],
    )
    return plain_workbook


def test_a_written_formula_has_no_value_until_it_is_calculated(
    call: Callable[..., Any], ledger: Path
) -> None:
    plain = call("xlide_read_cells", file_path=str(ledger), sheet="Sheet1", cell_range="A3")
    assert plain["recalculated"] is False
    assert "calculate=true" in plain["note"]

    before = ledger.read_bytes()
    worked = call(
        "xlide_read_cells", file_path=str(ledger), sheet="Sheet1", cell_range="A3", calculate=True
    )
    assert worked["values"] == [[5]]
    assert worked["recalculated"] is True
    assert worked["calculated_by"] == "pyOfficeEditor formula engine"
    assert ledger.read_bytes() == before, "calculating reads; it never writes"


def test_what_the_engine_cannot_work_out_is_named_not_passed_off(
    call: Callable[..., Any], ledger: Path
) -> None:
    worked = call(
        "xlide_read_cells",
        file_path=str(ledger),
        sheet="Sheet1",
        cell_range="A1:A5",
        calculate=True,
    )
    assert worked["recalculated"] is False
    assert worked["kept_cached"]["A4"] == "the function WEBSERVICE"
    assert "could not work out" in worked["kept_cached"]["A5"]
    assert worked["values"][2] == [5]


def test_the_text_a_cell_shows_follows_its_number_format(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="B1",
        data=[[0.125]],
    )
    call(
        "xlide_format_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        cell_range="B1",
        number_format="0.0%",
    )
    shown = call(
        "xlide_read_cells", file_path=str(plain_workbook), sheet="Sheet1", cell_range="B1",
        include="text",
    )
    assert shown["text"] == [["12.5%"]]


def test_a_formula_can_be_tried_before_it_is_written(
    call: Callable[..., Any], ledger: Path
) -> None:
    answer = call(
        "xlide_evaluate_formula", file_path=str(ledger), sheet="Sheet1", formula="=SUM(A1:A2)*10"
    )
    assert answer["value"] == 50
    assert answer["formula"] == "=SUM(A1:A2)*10"

    error = call("xlide_evaluate_formula", file_path=str(ledger), sheet="Sheet1", formula="A1/0")
    assert error["value"] == "#DIV/0!"


def test_a_formula_the_engine_cannot_read_is_refused(
    call: Callable[..., Any], ledger: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_evaluate_formula", file_path=str(ledger), sheet="Sheet1", formula="=SUM(")
    assert "not a formula the engine can read" in refusal.value.message
