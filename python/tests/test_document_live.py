"""Excel's own word on what the document-surface tools write, and on what they read.

Every check here is one a reader of the package cannot make about itself: that
Excel opens the file without a repair and sees what was written. And two things
the library reads but cannot build, a pivot table and a chart sheet, are built
by Excel for the listing to find.

Run with:  pytest -m live
"""

from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(sys.platform != "win32", reason="needs Windows with Excel installed"),
]

SALES = [
    ["Region", "Amount"],
    ["West", 120],
    ["East", 80],
    ["West", 45],
    ["North", 300],
    ["East", 10],
]

READ_BACK = """
Public Function ReadBack() As String
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    ReadBack = CStr(ws.AutoFilterMode) & "|" & _
               CStr(ws.Rows(3).Hidden) & "|" & CStr(ws.Rows(2).Hidden) & "|" & _
               ws.Range("C2").Comment.Text & "|" & _
               CStr(ws.ChartObjects.Count) & "|" & _
               CStr(ws.Range("D1").Characters(1, 6).Font.Bold) & "|" & _
               CStr(ws.Range("D1").Characters(7, 2).Font.Bold)
End Function
"""


def test_excel_sees_the_filter_note_chart_and_runs_this_server_wrote(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    path = str(plain_workbook)
    sheet = {"file_path": path, "sheet": "Sheet1"}
    call("xlide_write_cells", start_cell="A1", data=SALES, **sheet)
    call(
        "xlide_manage_filter", action="set", cell_range="A1:B6", column="Region",
        criteria1="=West", **sheet,
    )
    call("xlide_manage_comment", action="set", cell="C2", text="Checked", author="QA", **sheet)
    call(
        "xlide_write_cells", start_cell="D1",
        data=[[{"rich_text": [{"text": "Total ", "bold": True}, {"text": "42"}]}]], **sheet,
    )
    call("xlide_add_chart", data_range="A1:B6", chart_type="column", cell="F2", **sheet)

    from pyvbaharness import ExcelSession

    with ExcelSession() as excel:
        excel.open_document(path, read_only=True)
        result = excel.run_vba(READ_BACK, proc="ReadBack", timeout=240)

    assert result.outcome == "passed", result.error
    mode, east_hidden, west_hidden, note, charts, bold, plain = str(result.value).split("|")
    assert mode == "True"
    assert (east_hidden, west_hidden) == ("True", "False"), "Excel shows what the filter kept"
    assert note.endswith("Checked")
    assert charts == "1"
    assert (bold, plain) == ("True", "False")


BUILD_PIVOT_AND_CHART_SHEET = """
Public Function Build(ByVal target As String) As String
    Dim wb As Workbook, data As Worksheet, report As Worksheet
    Set wb = Application.Workbooks.Add
    Set data = wb.Worksheets(1)
    data.Name = "Data"
    data.Range("A1:B4").Value = [{"Region","Amount";"West",1;"East",2;"West",3}]
    Set report = wb.Worksheets.Add(After:=data)
    report.Name = "Report"
    wb.PivotCaches.Create(1, data.Range("A1:B4")).CreatePivotTable _
        report.Range("B3"), "RegionTotals"
    wb.Charts.Add(After:=report).Name = "Trend"
    Application.DisplayAlerts = False
    wb.SaveAs target, 51
    wb.Close False
    Application.DisplayAlerts = True
    Build = target
End Function
"""


def test_a_pivot_table_and_a_chart_sheet_excel_made_are_listed(
    call: Callable[..., Any], workspace: Path
) -> None:
    target = workspace / "Report.xlsx"
    from pyvbaharness import ExcelSession

    with ExcelSession() as excel:
        result = excel.run_vba(
            BUILD_PIVOT_AND_CHART_SHEET, proc="Build", args=(str(target),), timeout=240
        )
    assert result.outcome == "passed", result.error

    listed = call("xlide_list_sheets", file_path=str(target))
    report = next(sheet for sheet in listed["sheets"] if sheet["name"] == "Report")
    (pivot,) = report["pivot_tables"]
    assert pivot["name"] == "RegionTotals"
    assert "Data" in pivot["source"]
    assert [tab["name"] for tab in listed["chart_sheets"]] == ["Trend"]
