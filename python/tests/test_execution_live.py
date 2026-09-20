"""The execution layer against real Office. Skipped unless asked for.

Run with:  pytest -m live

These start desktop applications, so they are slow, they are Windows-only, and a
PowerPoint one puts a window on screen. They are the only tests that can tell you
the run path works, because everything they exercise is behaviour of Office rather
than of this code.
"""

from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(sys.platform != "win32", reason="needs Windows with Office installed"),
]

# PyVbaLog is supplied by the harness at run time and does not exist statically,
# which is why it appears here and not in the fixture the analysis tests use.
LOGGING_MODULE = """Option Explicit

Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function

Public Sub Boom()
    Err.Raise 513, "Tests", "deliberate failure"
End Sub
"""

TESTS_MODULE = """Option Explicit

Public Sub TestArithmetic()
    PyVbaAssertEqual 4, 2 + 2
End Sub

Public Sub TestDeliberateFailure()
    PyVbaAssertEqual 5, 2 + 2, "two and two are four"
End Sub
"""


@pytest.fixture
def runnable(call: Callable[..., Any], workbook: Path) -> Path:
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Runner",
        source=LOGGING_MODULE.replace("\n", "\r\n"),
    )
    return workbook


def test_run_macro_returns_its_value(call: Callable[..., Any], runnable: Path) -> None:
    result = call(
        "xlide_run_macro",
        file_path=str(runnable),
        procedure="Runner.AddNums",
        args=[20, 22],
        timeout=120,
    )
    assert result["outcome"] == "passed"
    assert result["value"] == 42


def test_a_vba_error_comes_back_as_data(call: Callable[..., Any], runnable: Path) -> None:
    """No dialog opens, and the error arrives as a value the caller can act on."""
    result = call(
        "xlide_run_macro", file_path=str(runnable), procedure="Runner.Boom", timeout=120
    )
    assert result["outcome"] == "vba-error"
    assert result["passed"] is False
    assert result["error"]["number"] == 513
    assert "deliberate failure" in result["error"]["description"]
    # A procedure already in the document was never instrumented, so there is no
    # failing line. The result has to say so rather than leave a silent null.
    assert result["error"]["line"] is None
    assert "xlide_run_vba" in result["error"]["line_note"]


def test_injected_source_reports_the_failing_line_and_stack(
    call: Callable[..., Any],
) -> None:
    """The other half of the same fact: injected source is instrumented, so it
    carries the line and the call stack that run_macro cannot."""
    result = call(
        "xlide_run_vba",
        source=(
            "Public Sub Main()\r\n"
            "    Dim x As Long\r\n"
            "    x = 1\r\n"
            '    Err.Raise 513, "Tests", "deliberate failure"\r\n'
            "End Sub\r\n"
        ),
        procedure="Main",
        timeout=120,
    )
    assert result["outcome"] == "vba-error"
    assert result["error"]["number"] == 513
    assert result["error"]["line"] == 4
    assert result["error"]["stack"], "a stack is the point of trapping the error in VBA"
    assert "line_note" not in result["error"]


def test_run_vba_without_a_file(call: Callable[..., Any]) -> None:
    result = call(
        "xlide_run_vba",
        source='Public Sub Main()\r\n    PyVbaLog "ran"\r\nEnd Sub\r\n',
        procedure="Main",
        timeout=120,
    )
    assert result["outcome"] == "passed"
    assert result["output"] == ["ran"]


def test_a_runaway_run_is_terminated(call: Callable[..., Any]) -> None:
    result = call(
        "xlide_run_vba",
        source="Public Sub Main()\r\n    Do\r\n    Loop\r\nEnd Sub\r\n",
        procedure="Main",
        timeout=10,
    )
    assert result["outcome"] == "timeout"
    assert "deadline" in result["note"]


def test_run_tests_reports_each_case(call: Callable[..., Any], workbook: Path) -> None:
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Suite",
        source=TESTS_MODULE.replace("\n", "\r\n"),
    )
    report = call(
        "xlide_run_tests", file_path=str(workbook), module_name="Suite", timeout=120
    )
    by_name = {case["name"]: case for case in report["cases"]}
    assert by_name["TestArithmetic"]["passed"] is True
    assert by_name["TestDeliberateFailure"]["passed"] is False
    assert report["failed"] == 1
    assert report["verdict"] == "failures"


def test_run_tests_refuses_a_file_with_none(call: Callable[..., Any], workbook: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_run_tests", file_path=str(workbook))
    assert "No test procedures found" in refusal.value.message


def test_compile_check_accepts_clean_code(call: Callable[..., Any], runnable: Path) -> None:
    result = call("xlide_compile_check", file_path=str(runnable), timeout=120)
    assert result["outcome"] in {"accepted", "infrastructure-failure"}
    if result["outcome"] == "infrastructure-failure":
        pytest.skip("the compile check could not complete; the result is unknown, not a failure")
    assert result["accepted"] is True


def test_a_workbook_this_server_wrote_opens_in_excel(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    """The gate that matters for the write path: Excel opens the spliced package
    and reads back what was written, with no repair prompt."""
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[2], [3], ["=SUM(A1:A2)"]],
    )
    result = call(
        "xlide_run_vba",
        source=(
            "Public Function ReadBack() As String\r\n"
            "    ReadBack = CStr(ActiveWorkbook.Worksheets(1).Range(\"A3\").Value) & \"|\" & "
            "ActiveWorkbook.Worksheets(1).Range(\"A3\").Formula\r\n"
            "End Function\r\n"
        ),
        procedure="ReadBack",
        file_path=str(plain_workbook),
        timeout=180,
    )
    assert result["outcome"] == "passed", result
    value, formula = str(result["value"]).split("|", 1)
    assert formula == "=SUM(A1:A2)", "Excel reads the formula back as it was typed"
    assert value == "5", "Excel recalculated on open, because the write asked it to"
