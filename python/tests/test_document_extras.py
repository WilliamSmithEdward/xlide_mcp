"""Charts, text in several fonts, and what list_sheets says about pivots and chart tabs."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
from conftest import ToolFailure

MONTHS = [["Month", "Sales", "Costs"], ["Jan", 10, 7], ["Feb", 14, 9], ["Mar", 12, 8]]


@pytest.fixture
def figures(call: Callable[..., Any], plain_workbook: Path) -> Path:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=MONTHS,
    )
    return plain_workbook


def test_a_chart_is_added_and_listed_with_its_series(
    call: Callable[..., Any], figures: Path
) -> None:
    added = call(
        "xlide_add_chart",
        file_path=str(figures),
        sheet="Sheet1",
        data_range="A1:C4",
        chart_type="lineMarkers",
        cell="E2",
        title="Sales and costs",
        chart_name="Trend",
    )
    chart = added["chart"]
    assert chart["name"] == "Trend"
    assert chart["type"] == ["line"]
    assert [series["name"] for series in chart["series"]] == ["Sales", "Costs"]
    assert chart["series"][0]["formula"].startswith("=SERIES(")

    listed = call("xlide_list_shapes", file_path=str(figures))
    trend = next(s for s in listed["sheets"][0]["shapes"] if s["name"] == "Trend")
    assert trend["kind"] == "chart"
    assert trend["chart"]["title"] == "Sales and costs"
    assert trend["cells"].startswith("E2")


def test_an_unknown_chart_type_names_the_ones_there_are(
    call: Callable[..., Any], figures: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_add_chart",
            file_path=str(figures),
            sheet="Sheet1",
            data_range="A1:C4",
            chart_type="radar",
            cell="E2",
        )
    assert "column, bar, line, lineMarkers" in refusal.value.message


def test_text_in_several_fonts_is_written_and_read_back(
    call: Callable[..., Any], figures: Path
) -> None:
    runs = [
        {"text": "Total ", "bold": True, "color": "C00000"},
        {"text": "(estimated)", "italic": True},
    ]
    call(
        "xlide_write_cells",
        file_path=str(figures),
        sheet="Sheet1",
        start_cell="A6",
        data=[[{"rich_text": runs}, "plain"]],
    )
    read = call(
        "xlide_read_cells",
        file_path=str(figures),
        sheet="Sheet1",
        cell_range="A6:B6",
        include="rich_text",
    )
    first, second = read["rich_text"][0]
    assert second is None, "plain text has no runs"
    assert [run["text"] for run in first] == ["Total ", "(estimated)"]
    assert first[0]["bold"] is True
    assert first[0]["color"].endswith("C00000")
    assert first[1]["italic"] is True
    values = call("xlide_read_cells", file_path=str(figures), sheet="Sheet1", cell_range="A6")
    assert values["values"] == [["Total (estimated)"]]


def test_a_malformed_run_is_refused_before_anything_is_written(
    call: Callable[..., Any], figures: Path
) -> None:
    before = figures.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_cells",
            file_path=str(figures),
            sheet="Sheet1",
            start_cell="A6",
            data=[[{"rich_text": [{"text": "x", "colour": "FF0000"}]}]],
        )
    assert "colour" in refusal.value.message
    assert figures.read_bytes() == before


def test_a_size_that_is_not_a_number_is_refused_by_name(
    call: Callable[..., Any], figures: Path
) -> None:
    """float("big") escaped as a bare "Error executing tool"."""
    before = figures.read_bytes()
    for size in ("big", -3):
        with pytest.raises(ToolFailure) as refusal:
            call(
                "xlide_write_cells",
                file_path=str(figures),
                sheet="Sheet1",
                start_cell="A6",
                data=[[{"rich_text": [{"text": "x", "size": size}]}]],
            )
        assert "size is in points" in refusal.value.message
    assert figures.read_bytes() == before


def test_a_pivot_table_is_described_by_where_it_is_and_what_it_reads() -> None:
    from xlide_mcp.cells import _pivot

    by_range = SimpleNamespace(
        name="PivotTable1", location="H3:J9", source_name=None, source_sheet="Data",
        source_range="A1:C100",
    )
    by_name = SimpleNamespace(
        name="PivotTable2", location="L3:N9", source_name="SalesTable", source_sheet=None,
        source_range=None,
    )
    assert _pivot(by_range) == {"name": "PivotTable1", "range": "H3:J9", "source": "Data!A1:C100"}
    assert _pivot(by_name)["source"] == "SalesTable"
