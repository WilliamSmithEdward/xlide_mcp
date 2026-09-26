"""Worksheet cells in the formats the package reader cannot open.

`.xlsx`, `.xlsm` and `.xlam` keep their grid in OOXML, which `xlsx.py` reads and
writes on any platform with no Office installed. `.xlsb` keeps it in BIFF12
binary records and `.xls` in BIFF8 inside a compound file. Implementing either is
a library rather than a feature, and the one program that already reads and
writes both correctly is Excel.

So on Windows with Excel, those two formats go through Excel. It is slower - a
session costs a few seconds to start - and it is used only where the package
reader cannot answer, never in preference to it.

One difference matters enough to travel with every result. A value read out of
the package is what Excel last calculated and stored; a value read through Excel
is what Excel has just calculated, because opening the workbook recalculates it.
The second is better data and it is not the same data, so `source` and
`recalculated` say which one came back rather than leaving a caller to assume.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import ToolError
from .hosts import HostInfo

# Formats Excel will save back in place. The harness refuses the legacy ones on
# save because they drop things silently, and a save through VBA has the same
# hazard, so a write to one is refused rather than attempted.
WRITABLE_THROUGH_EXCEL = frozenset({".xlsb", ".xlsm", ".xlam", ".xlsx"})

# Reading every sheet's name and used range in one call, rather than one COM
# round trip per question.
_SHEET_SURVEY = """
Public Function SurveySheets() As String
    Dim ws As Object, out As String
    For Each ws In ActiveWorkbook.Worksheets
        out = out & ws.Name & Chr(9) & ws.UsedRange.Address(False, False) & _
              Chr(9) & CStr(ws.Visible = -1) & Chr(10)
    Next ws
    SurveySheets = out
End Function
"""

_COUNT_OCCUPIED = """
Public Function CountOccupied(ByVal sheetName As String, ByVal firstCell As String, _
                              ByVal rowCount As Long, ByVal columnCount As Long) As Long
    Dim area As Range
    Set area = ActiveWorkbook.Worksheets(sheetName).Range(firstCell).Resize(rowCount, columnCount)
    CountOccupied = Application.WorksheetFunction.CountA(area)
End Function
"""


@dataclass(frozen=True)
class Sheet:
    name: str
    used_range: str
    hidden: bool


def excel_available() -> bool:
    """Whether this machine can answer for a format the package reader cannot."""
    if sys.platform != "win32":
        return False
    try:
        import pyvbaharness  # noqa: F401
    except ImportError:
        return False
    return True


def refuse(info: HostInfo, operation: str) -> ToolError:
    """Why a binary format cannot be reached here, and what would change it."""
    reason = (
        f"{operation} reads the OOXML package, and {info.extension} does not use one: "
        f"a .xlsb keeps its grid in binary records and a .xls inside a compound file."
    )
    if sys.platform != "win32":
        return ToolError(
            f"{reason} Excel reads its own formats, so on Windows with Excel installed this "
            "works; on this machine it cannot. The VBA project in the file is still fully "
            "readable."
        )
    return ToolError(
        f"{reason} Excel can do it, and this server will use it once pyvbaharness is "
        "installed. Ask the user before installing anything, then: "
        "pip install 'xlide-mcp[live]'. The VBA project in the file is still fully readable."
    )


def _session() -> Any:
    from .tools.execution import _session as make_session

    return make_session(HostInfo(host="excel", extension="", readable=True))


def survey_sheets(path: Path, timeout: float) -> list[Sheet]:
    """Every worksheet's name, used range and visibility, through Excel."""
    with _session() as excel:
        _open(excel, path, read_only=True)
        result = excel.run_vba(_SHEET_SURVEY, proc="SurveySheets", timeout=timeout)
    _require(result, f"reading the sheets of {path.name}")
    return _parse_sheet_survey(str(result.value or ""))


def _parse_sheet_survey(raw: str) -> list[Sheet]:
    """Decode Excel's sheet survey without quietly dropping malformed rows."""
    sheets: list[Sheet] = []
    for line_number, line in enumerate(raw.split("\n"), start=1):
        line = line.rstrip("\r")
        if not line:
            continue
        parts = line.split("\t")
        if len(parts) != 3 or not parts[0] or not parts[1]:
            raise ToolError(
                f"Excel returned malformed worksheet details on line {line_number}. "
                "No sheet list was returned; retry xlide_list_sheets."
            )
        visible = parts[2].strip().lower()
        if visible not in {"true", "false", "-1", "0"}:
            raise ToolError(
                f"Excel returned an unknown visibility state for {parts[0]!r}: "
                f"{parts[2]!r}. No sheet list was returned."
            )
        sheets.append(
            Sheet(
                name=parts[0],
                used_range=parts[1],
                hidden=visible in {"false", "0"},
            )
        )
    return sheets


def read_cells(path: Path, sheet: str, cell_range: str, timeout: float) -> list[list[Any]]:
    """A block of cells as Excel has just calculated them."""
    with _session() as excel:
        _open(excel, path, read_only=True)
        try:
            return excel.read_range(sheet, cell_range, timeout=timeout)
        except Exception as exc:
            raise ToolError(f"Excel could not read {cell_range} on {sheet!r}: {exc}") from exc


def write_cells(
    path: Path, sheet: str, start_cell: str, data: list[list[Any]], timeout: float,
    *, allow_overwrite: bool = False,
) -> dict[str, Any]:
    """Write a block through Excel and save the workbook in its own format.

    A legacy format is refused rather than saved: saving one means choosing a
    format, and the wrong choice drops what that format cannot hold, silently,
    with alerts suppressed.
    """
    if path.suffix.lower() not in WRITABLE_THROUGH_EXCEL:
        raise ToolError(
            f"Writing cells to {path.suffix} is not offered. Excel can open it, but saving it "
            "back means choosing a format, and the wrong choice drops what the format cannot "
            "hold with the alerts suppressed. Ask the user to save it as .xlsb or .xlsm first."
        )
    rows = len(data)
    columns = max((len(row) for row in data), default=0)
    if not rows or not columns:
        raise ToolError("data has no cells to write.")
    # Run this in a read-only instance that is discarded. The harness injects
    # support modules for run_vba, which must never be saved with the workbook.
    with _session() as excel:
        _open(excel, path, read_only=True)
        count = excel.run_vba(
            _COUNT_OCCUPIED, proc="CountOccupied",
            args=(sheet, start_cell, rows, columns), timeout=timeout,
        )
    _require(count, f"checking cells in {path.name}")
    occupied = int(count.value or 0)
    if occupied and not allow_overwrite:
        raise ToolError(
            f"Writing at {start_cell} would overwrite {occupied} cells. Nothing was written. "
            "Read the range, ask the user, then call again with allow_overwrite=true."
        )
    before = _module_names(path)
    with _session() as excel:
        _open(excel, path, read_only=False)
        try:
            excel.write_range(sheet, start_cell, data, timeout=timeout)
        except Exception as exc:
            raise ToolError(f"Excel could not write to {sheet!r}: {exc}") from exc
        # The session's own save, not a save driven from injected VBA. Injecting
        # anything puts the harness's support modules into the project, and a
        # save then makes them part of the user's workbook for good: measured,
        # a cell write left PyVbaUserCode, PyVbaHarnessRunner and PyVbaHarnessCall
        # behind in the file.
        try:
            excel.save_as(str(path), timeout=timeout)
        except Exception as exc:
            raise ToolError(f"Excel could not save {path.name}: {exc}") from exc

    left_behind = sorted(_module_names(path) - before)
    if left_behind:
        raise ToolError(
            f"The write left {', '.join(left_behind)} in the VBA project, which was not "
            "there before and is not the user's. The workbook has been saved with them; "
            "remove them with xlide_delete_module and report this."
        )
    return {"saved": True, "saved_as": str(path), "cells_overwritten": occupied}


def _module_names(path: Path) -> set[str]:
    """The project's modules, for checking a write added none of its own."""
    try:
        import pyopenvba

        with pyopenvba.ExcelFile(path) as book:
            return {module.name for module in book.vba_project().modules}
    except Exception:
        return set()


def _open(session: Any, path: Path, *, read_only: bool) -> None:
    try:
        session.open_workbook(str(path), read_only=read_only)
    except Exception as exc:
        raise ToolError(f"Excel could not open {path.name}: {exc}") from exc


def _require(result: Any, what: str) -> None:
    """Turn a run that did not pass into the reason it did not."""
    outcome = getattr(result, "outcome", "runner-error")
    if outcome == "passed":
        return
    error = getattr(result, "error", None)
    detail = getattr(error, "description", "") if error is not None else ""
    if outcome == "timeout":
        raise ToolError(
            f"Excel did not finish {what} before the deadline, and was terminated. "
            "A larger timeout may help; a workbook that recalculates on open can be slow."
        )
    if outcome == "modal-blocked":
        raise ToolError(
            f"Excel put up a dialog while {what}, so it was stopped. The workbook may be "
            "asking for a password, or warning about links."
        )
    raise ToolError(f"Excel failed while {what}: {detail or outcome}")
