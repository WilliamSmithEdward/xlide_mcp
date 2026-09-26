"""Keep the committed shapes fixture honest, and rebuild it when it has to change.

`fixtures/shapes.xlsm` is the one binary in the repository, and a binary fixture
is a claim about what Excel writes that nothing checks. This runs Excel, builds
the same workbook again, and asserts the reader sees the same thing in the fresh
one as in the committed one. If Excel ever changes how it stores a form control,
this is what says so.

Run with:  pytest -m live
Rebuild:   pytest -m live -k rebuild --rebuild-fixture
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pytest

from xlide_mcp.shapes import read_sheet_shapes

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(sys.platform != "win32", reason="needs Windows with Excel installed"),
]


def _excel_session() -> object:
    from pyvbaharness import ExcelSession, HarnessConfig

    return ExcelSession(HarnessConfig(lock_wait_s=60.0))

FIXTURE = Path(__file__).parent / "fixtures" / "shapes.xlsm"

# The workbook the fixture is. Excel authors it, because a Forms-toolbar button
# lives in four parts that only Excel writes in agreement.
BUILD_SOURCE = """
Public Function BuildIt(ByVal target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim btn As Object
    Dim shp As Object
    Dim cb As Object

    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Controls"

    Set btn = ws.Buttons.Add(ws.Range("B2").Left, ws.Range("B2").Top, 90, 24)
    btn.Name = "RunButton"
    btn.Caption = "Run it"
    btn.OnAction = "Module1.DoTheThing"

    Set shp = ws.Shapes.AddShape(1, ws.Range("E2").Left, ws.Range("E2").Top, 120, 40)
    shp.Name = "GoShape"
    shp.TextFrame.Characters.Text = "Go"
    shp.OnAction = "Module1.DoTheThing"

    Set cb = ws.CheckBoxes.Add(ws.Range("B6").Left, ws.Range("B6").Top, 100, 18)
    cb.Name = "Ready"
    cb.Caption = "Ready?"
    cb.LinkedCell = "$D$6"

    wb.VBProject.VBComponents.Add(1).Name = "Module1"
    wb.VBProject.VBComponents("Module1").CodeModule.AddFromString _
        "Public Sub DoTheThing()" & vbCrLf & "End Sub"

    Application.DisplayAlerts = False
    wb.SaveAs target, 52
    wb.Close False
    Application.DisplayAlerts = True
    BuildIt = target
End Function
"""


def build_with_excel(target: Path) -> None:
    with _excel_session() as excel:
        result = excel.run_vba(
            BUILD_SOURCE, proc="BuildIt", args=(str(target),), timeout=240
        )
    assert result.outcome == "passed", f"Excel could not build the fixture: {result.error}"
    assert target.is_file()


def readable(path: Path) -> dict[str, dict[str, object]]:
    """Everything the reader claims about a workbook, keyed by shape name."""
    out: dict[str, dict[str, object]] = {}
    for sheet, shapes in read_sheet_shapes(path).items():
        for shape in shapes:
            out[f"{sheet}!{shape.name}"] = shape.summary()
    return out


def test_a_freshly_built_workbook_reads_like_the_committed_fixture(
    tmp_path: Path,
) -> None:
    """The check that makes a committed binary trustworthy. It compares what the
    reader sees, not the bytes: Excel varies ids and creation GUIDs between
    saves, and none of that is what the fixture is for."""
    fresh = tmp_path / "fresh.xlsm"
    build_with_excel(fresh)

    assert readable(fresh) == readable(FIXTURE), (
        "Excel now writes this workbook differently than the committed fixture "
        "records. Rebuild it with --rebuild-fixture and read the diff carefully."
    )


def test_rebuild_the_fixture(tmp_path: Path, request: pytest.FixtureRequest) -> None:
    """Rebuilds the committed fixture. Skipped unless --rebuild-fixture is passed,
    because overwriting a fixture is not something a test run should do by
    accident."""
    if not request.config.getoption("--rebuild-fixture"):
        pytest.skip("pass --rebuild-fixture to overwrite the committed fixture")
    fresh = tmp_path / "fresh.xlsm"
    build_with_excel(fresh)
    shutil.copy2(fresh, FIXTURE)
    assert readable(FIXTURE), "the rebuilt fixture holds no shapes"


READ_BACK = """
Public Function ReadBack() As String
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Controls")
    ReadBack = ws.Buttons("RunButton").OnAction & "|" & _
               ws.Shapes("GoShape").OnAction & "|" & _
               CStr(ws.Shapes.Count)
End Function
"""


def test_excel_accepts_a_workbook_whose_shape_macro_this_server_changed(
    tmp_path: Path,
) -> None:
    """The gate for the write path. Excel opens the edited package without a
    repair prompt and reports the macro this server put there, on both the modern
    shape and the Forms button whose macro lives in two parts."""
    import shutil

    from xlide_mcp.shapes import set_shape_macro

    workbook = tmp_path / "Shapes.xlsm"
    shutil.copy2(FIXTURE, workbook)
    set_shape_macro(workbook, "Controls", "RunButton", "Module1.Renamed")
    set_shape_macro(workbook, "Controls", "GoShape", "Module1.Renamed")

    with _excel_session() as excel:
        excel.open_document(str(workbook), read_only=True)
        result = excel.run_vba(READ_BACK, proc="ReadBack", timeout=240)

    assert result.outcome == "passed", result.error
    button, drawing, count = str(result.value).split("|")
    # Excel resolves the stored [N]! prefix to the workbook's own name.
    assert button.endswith("Module1.Renamed"), button
    assert drawing == "Module1.Renamed"
    assert count == "3", "every shape survived the edit"


READ_ADDED = """
Public Function ReadAdded() As String
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Controls")
    ReadAdded = ws.Buttons("Refresh").OnAction & "|" & _
                ws.Buttons("Refresh").Caption & "|" & _
                ws.CheckBoxes("Include").LinkedCell & "|" & _
                ws.Shapes("Include").TopLeftCell.Address(False, False) & "|" & _
                CStr(ws.Shapes.Count)
End Function
"""


def test_excel_opens_a_workbook_this_server_added_and_removed_controls_in(
    tmp_path: Path,
) -> None:
    """The gate for adding and removing. A control is four parts, and a
    disagreement among them is a workbook Excel repairs on open, which a
    reader of the package alone cannot see."""
    from xlide_mcp.shapes import add_shape, remove_shape

    workbook = tmp_path / "Shapes.xlsm"
    shutil.copy2(FIXTURE, workbook)
    add_shape(
        workbook, "Controls", name="Refresh", kind="button", cell="E8",
        text="Refresh data", macro="Module1.DoTheThing",
    )
    add_shape(
        workbook, "Controls", name="Include", kind="checkBox", cell="B10",
        text="Include totals", linked_cell="$D$10",
    )
    remove_shape(workbook, "Controls", "GoShape")

    with _excel_session() as excel:
        excel.open_document(str(workbook), read_only=True)
        result = excel.run_vba(READ_ADDED, proc="ReadAdded", timeout=240)

    assert result.outcome == "passed", result.error
    macro, caption, linked, corner, count = str(result.value).split("|")
    assert macro.endswith("Module1.DoTheThing"), macro
    assert caption == "Refresh data"
    assert linked == "$D$10"
    assert corner == "B10", "placed on the cell it was given"
    assert count == "4", "RunButton, Ready and the two added; GoShape gone"
