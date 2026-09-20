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
    from pyvbaharness import ExcelSession

    with ExcelSession() as excel:
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
