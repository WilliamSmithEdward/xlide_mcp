"""The worksheet grid, through pyOfficeEditor.

This server used to read and write the grid itself, in `xlsx.py`, ported from
XLIDE because no library covered it. pyOfficeEditor covers it now, and holds byte
fidelity as a correctness property: editing one cell leaves every other byte of
the package alone. That is a stronger guarantee than a splice, and it is the kind
of format knowledge AGENTS.md says belongs upstream rather than here.

So this module is an adapter, not a reader. What it owns is the translation
between what pyOfficeEditor returns and what this server's tools have always
answered, because the conformance corpus pins those answers and a swap that
changed one would be a swap that broke a port:

* A whole number comes back as an integer, not a float. `values[0][0]` is `1`.
* A formula comes back with its leading `=`. pyOfficeEditor stores what the file
  stores, which has no `=`; the formula bar shows one, and so does this.
* A date is an ISO string and an error is its code, both JSON-safe, because a
  tool result is JSON and `datetime.date` is not.

`xlsx.py` keeps only the package surface the drawing layer still needs.
"""

from __future__ import annotations

import contextlib
import datetime as dt
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from . import locks
from .errors import ToolError

# One read's ceiling. A grid larger than this is not one an agent reasons about;
# it is one it asks narrower questions of.
MAX_CELLS_PER_READ = 20_000


class CellsError(ToolError):
    """A grid this server cannot read, with the reason a caller can act on."""


@dataclass(frozen=True)
class SheetInfo:
    name: str
    used_range: str
    hidden: bool


@dataclass(frozen=True)
class NamedRange:
    name: str
    refers_to: str


@dataclass(frozen=True)
class ReadResult:
    sheet: str
    reference: str
    rows: int
    columns: int
    values: list[list[Any]]
    formulas: list[list[str | None]]

    @property
    def formula_count(self) -> int:
        return sum(1 for row in self.formulas for formula in row if formula)


@dataclass(frozen=True)
class WriteResult:
    sheet: str
    reference: str
    cells_written: int
    cells_overwritten: int
    formulas_replaced: int


def open_workbook(path: Path) -> Any:
    """The workbook, or a refusal naming what this reader does not open."""
    import pyofficeeditor.excel as excel
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        return excel.Workbook.open(path)
    except PyOfficeEditorError as exc:
        raise CellsError(f"{path.name}: {exc}") from exc
    except PermissionError as exc:
        raise CellsError(locks.lock_message(path, "Excel", reading=True)) from exc
    except OSError as exc:
        raise CellsError(f"{path.name} could not be opened: {exc}") from exc


@contextlib.contextmanager
def editing(path: Path) -> Iterator[Any]:
    """Open a workbook, let the caller change it, save it.

    Every write tool in this server goes through here rather than opening and
    saving for itself, so that a save is never the step somebody forgot and a
    failure half way through never leaves a saved file behind.
    """
    with open_workbook(path) as book:
        yield book
        save(book, path)


def save(book: Any, path: Path) -> None:
    """Save a workbook, or refuse naming what holds it."""
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        book.save()
    except PyOfficeEditorError as exc:
        raise CellsError(f"{path.name} could not be saved: {exc}") from exc
    except PermissionError as exc:
        raise CellsError(locks.lock_message(path, "Excel")) from exc


def sheet_named(book: Any, name: str) -> Any:
    """A sheet by name, matched without case, or a refusal listing the real ones."""
    wanted = (name or "").strip().casefold()
    for sheet in book.sheets:
        if sheet.name.casefold() == wanted:
            return sheet
    listed = ", ".join(book.sheet_names) or "(none)"
    raise CellsError(f"No sheet named {name!r}. Sheets in this workbook: {listed}.")


def sheets(path: Path) -> list[SheetInfo]:
    with open_workbook(path) as book:
        return [_sheet_info(sheet) for sheet in book.sheets]


def named_ranges(path: Path) -> list[NamedRange]:
    with open_workbook(path) as book:
        return [
            NamedRange(name=entry.name, refers_to=entry.refers_to)
            for entry in book.defined_names
        ]


def read(path: Path, sheet_name: str, reference: str) -> ReadResult:
    with open_workbook(path) as book:
        sheet = sheet_named(book, sheet_name)
        area = _range(sheet, reference)
        # Measured from the reference, before a single cell is built. Asking for
        # a quarter of a million cells should cost a refusal, not the memory to
        # assemble them and then throw them away.
        span = area.reference
        height = span.end.row - span.start.row + 1
        width = span.end.column - span.start.column + 1
        if height * width > MAX_CELLS_PER_READ:
            raise CellsError(
                f"{span} is {height * width:,} cells, over the {MAX_CELLS_PER_READ:,} "
                "a single read returns. Read it in blocks."
            )
        rows = list(area.rows())
        return ReadResult(
            sheet=sheet.name,
            reference=str(area.reference),
            rows=height,
            columns=width,
            values=[[_value(cell.value) for cell in row] for row in rows],
            formulas=[[_formula(cell.formula) for cell in row] for row in rows],
        )


def write(path: Path, sheet_name: str, start_cell: str, data: list[list[Any]]) -> WriteResult:
    """Put values and formulas into a block, and say what was displaced.

    A value written over a formula removes the formula, which is what typing into
    the cell does. What the block held first is counted before anything moves, so
    the answer can tell the user they overwrote their own data rather than leaving
    them to find out.
    """
    if not data or not any(isinstance(row, list) for row in data):
        raise CellsError("data must be a list of rows, each row a list of cell values.")

    with open_workbook(path) as book:
        sheet = sheet_named(book, sheet_name)
        width = max(len(row) for row in data)
        first = _cell_ref(sheet, start_cell)
        block = _block(first, len(data), width)

        overwritten = formulas_replaced = 0
        if len(data) * width <= MAX_CELLS_PER_READ:
            for row in _range(sheet, block).rows():
                for cell in row:
                    if cell.formula is not None:
                        formulas_replaced += 1
                    elif cell.value is not None:
                        overwritten += 1

        written = 0
        for row_offset, row_values in enumerate(data):
            for column_offset, raw in enumerate(row_values):
                cell = sheet.cell(first.row + row_offset, first.column + column_offset)
                _put(cell, raw)
                written += 1
        save(book, path)

    return WriteResult(
        sheet=sheet.name,
        reference=block,
        cells_written=written,
        cells_overwritten=overwritten + formulas_replaced,
        formulas_replaced=formulas_replaced,
    )


# ------------------------------------------------------------------ the parts


def _hidden(sheet: Any) -> bool:
    """Whether a sheet is hidden, however this pyOfficeEditor spells it.

    `Worksheet.visible` arrives in 0.2. On 0.1.1 the state is only in the
    workbook part, as `state="hidden"` on the sheet entry, so it is read from
    there. Feature-detected rather than version-detected: the moment the
    attribute exists this stops touching XML, with no pin to remember.

    A sheet is hidden whether it is hidden or very hidden. An agent asking which
    sheets are hidden means both, and the difference is only how hard Excel
    makes it to unhide.
    """
    visible = getattr(sheet, "visible", None)
    if visible is not None:
        return str(getattr(visible, "value", visible)).lower() != "visible"
    try:
        entry = _sheet_entry(sheet)
    except Exception:
        return False
    return entry in {"hidden", "veryhidden"}


def _sheet_entry(sheet: Any) -> str:
    """The `state` attribute on this sheet's entry in the workbook part."""
    import re

    book = sheet.workbook
    xml = book.package.xml(book.workbook_part).text
    for match in re.finditer(r"<sheet\b[^>]*>", xml):
        markup = match.group(0)
        name = re.search(r'\bname="([^"]*)"', markup)
        if name is None or name.group(1) != sheet.name:
            continue
        state = re.search(r'\bstate="([^"]*)"', markup)
        return state.group(1).strip().lower() if state else "visible"
    return "visible"


def _sheet_info(sheet: Any) -> SheetInfo:
    used = sheet.used_range
    return SheetInfo(
        name=sheet.name,
        used_range=str(used) if used is not None else "",
        hidden=_hidden(sheet),
    )


def _range(sheet: Any, reference: str) -> Any:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        return sheet.range(reference)
    except PyOfficeEditorError as exc:
        raise CellsError(f"{reference!r} is not a range this sheet can read: {exc}") from exc
    except ValueError as exc:
        raise CellsError(f"{reference!r} is not an A1-style range: {exc}") from exc


def _cell_ref(sheet: Any, reference: str) -> Any:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        return sheet[reference].reference
    except (PyOfficeEditorError, ValueError, KeyError) as exc:
        raise CellsError(f"{reference!r} is not an A1-style cell reference: {exc}") from exc


def _block(first: Any, height: int, width: int) -> str:
    from pyofficeeditor.excel import MAX_COLUMN, MAX_ROW, column_letter

    last_row = first.row + height - 1
    last_column = first.column + width - 1
    if last_row > MAX_ROW or last_column > MAX_COLUMN:
        raise CellsError(
            f"The block starting at {first} runs past the end of the sheet."
        )
    start = f"{column_letter(first.column)}{first.row}"
    end = f"{column_letter(last_column)}{last_row}"
    return start if start == end else f"{start}:{end}"


def _value(raw: Any) -> Any:
    """What pyOfficeEditor read, as something a JSON result can carry.

    The integer rule is not cosmetic. A worksheet stores every number as a
    double, so a count of 3 reads back as 3.0, and a result that answered 3.0
    where the file shows 3 would have an agent reporting a number the user
    cannot find in their workbook.
    """
    if raw is None or isinstance(raw, (str, bool)):
        return raw
    if isinstance(raw, (dt.datetime, dt.date)):
        return raw.isoformat()
    if isinstance(raw, float):
        return int(raw) if raw.is_integer() and abs(raw) < 2**53 else raw
    if isinstance(raw, int):
        return raw
    # A CellError, which is a value in its own right: the cell holds #DIV/0!,
    # it does not hold the text of it.
    return str(raw)


def _formula(raw: str | None) -> str | None:
    """The formula as the formula bar shows it.

    The file stores `SUM(A1:A2)`; every human and every agent writes
    `=SUM(A1:A2)`, and a round trip that dropped the `=` would make a read that
    fed straight back into a write turn a formula into text.
    """
    if raw is None:
        return None
    return raw if raw.startswith("=") else "=" + raw


def _put(cell: Any, raw: Any) -> None:
    """One cell, written the way typing into it would write it."""
    if isinstance(raw, str) and raw.startswith("="):
        cell.formula = raw[1:]
        return
    # Writing a value clears any formula that was there, which is what typing
    # over a formula does.
    if cell.formula is not None:
        cell.formula = None
    cell.value = raw
