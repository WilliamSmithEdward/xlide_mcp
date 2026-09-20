"""Running VBA in each of the four applications, not just in Excel.

The execution layer was only ever exercised against Excel, and the three others
each behave differently in a way that had to be found rather than assumed.

Word behaves like Excel. Access has no read-only automation mode at all, so a run
there writes to the database as it goes and the caller has to say so. PowerPoint
refuses to run hidden, so its window appears on screen, and it is single-instance:
it cannot start while someone has PowerPoint open, and the harness refuses rather
than taking over a process it did not create.

Run with:  pytest -m live
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

MODULE = (
    "Option Explicit\r\n"
    "\r\n"
    "Public Function Twice(ByVal n As Long) As Long\r\n"
    "    Twice = n * 2\r\n"
    "End Function\r\n"
)

RUNNABLE = [
    pytest.param(("excel", "Book.xlsm"), id="excel"),
    pytest.param(("word", "Doc.docm"), id="word"),
    pytest.param(("powerpoint", "Deck.pptm"), id="powerpoint"),
]


@pytest.fixture
def runnable(request: pytest.FixtureRequest, workspace: Path) -> tuple[str, Path]:
    import pyopenvba

    host, file_name = request.param
    maker = {
        "excel": pyopenvba.ExcelFile,
        "word": pyopenvba.WordFile,
        "powerpoint": pyopenvba.PowerPointFile,
    }[host]
    path = workspace / file_name
    with maker.create_new(path) as handle:
        handle.vba_project().add_module("Helpers", MODULE)
        handle.save()
    return host, path


@pytest.mark.parametrize("runnable", RUNNABLE, indirect=True)
def test_a_macro_runs_and_returns_its_value(
    call: Callable[..., Any], runnable: tuple[str, Path]
) -> None:
    _, path = runnable
    result = call(
        "xlide_run_macro",
        file_path=str(path),
        procedure="Helpers.Twice",
        args=[21],
        timeout=300,
    )
    assert result["outcome"] == "passed"
    assert result["value"] == 42


@pytest.mark.parametrize("runnable", RUNNABLE, indirect=True)
def test_injected_source_runs_in_every_host(
    call: Callable[..., Any], runnable: tuple[str, Path]
) -> None:
    host, path = runnable
    result = call(
        "xlide_run_vba",
        source='Public Sub Main()\r\n    PyVbaLog "ran in ' + host + '"\r\nEnd Sub\r\n',
        procedure="Main",
        file_path=str(path),
        timeout=300,
    )
    assert result["outcome"] == "passed"
    assert result["output"] == [f"ran in {host}"]


# ------------------------------------------------------------------- Access


@pytest.fixture
def database(workspace: Path) -> Path:
    import pyopenvba

    path = workspace / "App.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        db.vba_project().add_module("Helpers", "Option Compare Database\r\n" + MODULE)
        db.save()
    return path


def test_access_refuses_a_read_only_run_and_says_why(
    call: Callable[..., Any], database: Path
) -> None:
    """Access has no read-only automation mode. Passing the refusal through from
    the library would name a parameter the caller never saw, so the decision is
    put to them in their own terms instead."""
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_run_macro",
            file_path=str(database),
            procedure="Helpers.Twice",
            args=[21],
            timeout=300,
        )
    assert "read_only=false" in refusal.value.message
    assert "cannot be undone" in refusal.value.message


def test_access_runs_when_the_caller_allows_the_write(
    call: Callable[..., Any], database: Path
) -> None:
    result = call(
        "xlide_run_macro",
        file_path=str(database),
        procedure="Helpers.Twice",
        args=[21],
        read_only=False,
        timeout=300,
    )
    assert result["outcome"] == "passed"
    assert result["value"] == 42


def test_a_compile_check_on_access_is_refused_rather_than_faked(
    call: Callable[..., Any], database: Path
) -> None:
    """A read-only check that writes to the file is not a read-only check."""
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_compile_check", file_path=str(database))
    assert "xlide_analyze" in refusal.value.message


def test_a_compile_check_passes_on_the_hosts_that_have_one(
    call: Callable[..., Any], workspace: Path
) -> None:
    import pyopenvba

    path = workspace / "Doc.docm"
    with pyopenvba.WordFile.create_new(path) as document:
        document.vba_project().add_module("Helpers", MODULE)
        document.save()

    result = call("xlide_compile_check", file_path=str(path), timeout=300)
    if result["outcome"] == "infrastructure-failure":
        pytest.skip("the check could not complete; the result is unknown, not a failure")
    assert result["accepted"] is True


# ----------------------------------------- the formats the package cannot read


def test_a_binary_workbooks_grid_goes_through_excel(
    call: Callable[..., Any], workspace: Path
) -> None:
    """A .xlsb keeps its grid in BIFF12 records rather than OOXML. Excel reads
    and writes its own formats, so it answers, and because it opened the workbook
    the values are ones it has just calculated rather than a cached result."""
    import pyopenvba

    path = workspace / "Binary.xlsb"
    with pyopenvba.ExcelFile.create_new(path) as book:
        book.vba_project().add_module("Helpers", MODULE)
        book.save()

    sheets = call("xlide_list_sheets", file_path=str(path), timeout=300)
    assert sheets["source"] == "excel"
    sheet = sheets["sheets"][0]["name"]

    written = call(
        "xlide_write_cells",
        file_path=str(path),
        sheet=sheet,
        start_cell="A1",
        data=[[2], [3], ["=SUM(A1:A2)"]],
        timeout=300,
    )
    assert written["source"] == "excel"
    assert written["saved"] is True

    read = call(
        "xlide_read_cells", file_path=str(path), sheet=sheet, cell_range="A1:A3", timeout=300
    )
    assert read["source"] == "excel"
    assert read["recalculated"] is True
    # Excel calculated the SUM, which the package path never could have.
    assert read["values"] == [[2.0], [3.0], [5.0]]


def test_a_write_through_excel_leaves_no_modules_behind(
    call: Callable[..., Any], workspace: Path
) -> None:
    """Driving Excel from injected VBA puts the harness's own support modules in
    the project, and a save then makes them part of the user's file for good.
    Measured before the fix: a cell write left three of them there."""
    import pyopenvba

    path = workspace / "Binary.xlsb"
    with pyopenvba.ExcelFile.create_new(path) as book:
        book.vba_project().add_module("Helpers", MODULE)
        book.save()
    before = {m["name"] for m in call("xlide_list_modules", file_path=str(path))["modules"]}

    call(
        "xlide_write_cells",
        file_path=str(path),
        sheet="Sheet1",
        start_cell="A1",
        data=[[1]],
        timeout=300,
    )

    after = {m["name"] for m in call("xlide_list_modules", file_path=str(path))["modules"]}
    assert after == before, f"the write added {sorted(after - before)}"
    assert not any(name.startswith("PyVba") for name in after)
