"""Worksheet cells: what the VBA is usually about.

A cell value read from the file is the result Excel last calculated and stored,
not a result computed here. That distinction is why every write answers with
`recalculated: false`: a formula written by this server, and every formula that
depends on a cell it writes, keeps its old cached result until Excel next opens
the workbook. Reporting a number Excel has not calculated as though it had is the
failure these tools are shaped to prevent.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from ..config import Settings
from ..errors import ToolError
from ..hosts import host_info
from ..paths import require_writable, resolve_path
from ..xlsx import MAX_CELLS_PER_READ, CellRange, Workbook, parse_cell_ref
from ._common import read_only, writes


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_sheets",
        title="List worksheets",
        annotations=read_only("List worksheets"),
        description=(
            "Lists the worksheets in an Excel file with their used ranges, whether each is "
            "hidden, and the workbook's named ranges. Call this before reading cells, so the "
            "range you ask for is one that holds data. Works on .xlsx, .xlsm and .xlam; a "
            ".xlsb or .xls keeps its grid in a binary format this server does not read."
        ),
    )
    def list_sheets(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
    ) -> dict[str, Any]:
        path = _excel_path(file_path, settings)
        book = Workbook(path)
        return {
            "path": str(path),
            "sheets": [
                {
                    "name": sheet.name,
                    "used_range": sheet.used_range or "(empty)",
                    "hidden": sheet.hidden,
                }
                for sheet in book.sheets()
            ],
            "named_ranges": [
                {"name": n.name, "refers_to": n.refers_to} for n in book.named_ranges()
            ],
        }

    @server.tool(
        name="xlide_read_cells",
        title="Read cells",
        annotations=read_only("Read cells"),
        description=(
            "Reads a range of cells from a worksheet and returns a grid of values, formulas, "
            "or both. Values are what Excel last calculated and stored, so a formula whose "
            "inputs changed outside Excel shows its old result. Ask for formulas when you need "
            "to understand what a sheet computes, and values when you need what it currently "
            "shows. At most 20,000 cells per call."
        ),
    )
    def read_cells(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        cell_range: Annotated[
            str, Field(description="A1-style range, such as A1:D50, or a single cell.")
        ],
        include: Annotated[
            str,
            Field(
                default="values",
                description="'values', 'formulas' or 'both'.",
            ),
        ] = "values",
    ) -> dict[str, Any]:
        path = _excel_path(file_path, settings)
        wanted = (include or "values").strip().lower()
        if wanted not in {"values", "formulas", "both"}:
            raise ToolError("include must be 'values', 'formulas' or 'both'.")

        book = Workbook(path)
        area, grid = book.read(sheet, cell_range)
        result: dict[str, Any] = {
            "path": str(path),
            "sheet": book.canonical_sheet_name(sheet),
            "range": str(area),
            "rows": area.last_row - area.first_row + 1,
            "columns": area.last_column - area.first_column + 1,
        }
        if wanted in {"values", "both"}:
            result["values"] = [[cell.value for cell in row] for row in grid]
        if wanted in {"formulas", "both"}:
            result["formulas"] = [[cell.formula for cell in row] for row in grid]
        formula_count = sum(1 for row in grid for cell in row if cell.formula)
        if formula_count and wanted != "formulas":
            result["note"] = (
                f"{formula_count} of these cells hold formulas. The values are what Excel last "
                "calculated and stored, not a recalculation."
            )
        return result

    @server.tool(
        name="xlide_list_shapes",
        title="List shapes and buttons",
        annotations=read_only("List shapes and buttons"),
        description=(
            "Lists what sits on a worksheet's drawing layer: buttons, form controls, "
            "AutoShapes, text boxes, pictures, charts and groups, each with the cells it "
            "covers and, where it has one, the macro a click runs. Call it when asked how a "
            "workbook is started, or before renaming a Sub: a button's OnAction names a "
            "procedure and nothing rewrites it. ActiveX controls are listed but have no "
            "macro; their code is event procedures in the sheet's module."
        ),
    )
    def list_shapes(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[
            str,
            Field(default="", description="One worksheet. Empty lists every sheet's shapes."),
        ] = "",
    ) -> dict[str, Any]:
        path = _excel_path(file_path, settings)
        from ..shapes import read_sheet_shapes

        by_sheet = read_sheet_shapes(path, sheet.strip() or None)
        sheets = [
            {"sheet": name, "shapes": [shape.summary() for shape in shapes]}
            for name, shapes in by_sheet.items()
        ]
        total = sum(len(entry["shapes"]) for entry in sheets)
        with_macros = [
            {"sheet": entry["sheet"], "shape": shape["name"], "macro": shape["macro"]}
            for entry in sheets
            for shape in entry["shapes"]
            if shape.get("macro")
        ]
        result: dict[str, Any] = {
            "path": str(path),
            "shape_count": total,
            "sheets": sheets,
        }
        if with_macros:
            # The link an agent is usually actually after, lifted out of the tree.
            result["macros_run_by_shapes"] = with_macros
        return result

    @server.tool(
        name="xlide_set_shape_macro",
        title="Point a shape at a macro",
        annotations=writes("Point a shape at a macro", destructive=True),
        description=(
            "Changes which macro an existing shape or button runs when clicked, and saves "
            "the workbook. An empty macro clears the link. Use it after writing a Sub, so a "
            "button actually calls it, and after renaming one, because nothing rewrites an "
            "OnAction. Give the procedure as Proc or Module.Proc; it must already exist in "
            "the project, so write it first. This changes an existing shape only: adding or "
            "deleting one is not offered, because a form control lives in four parts that "
            "have to agree and a wrong one produces a workbook Excel repairs on open."
        ),
    )
    def set_shape_macro(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet the shape is on.")],
        shape_name: Annotated[
            str, Field(description="Shape name, as xlide_list_shapes reports it.")
        ],
        macro: Annotated[
            str,
            Field(
                description=(
                    "The procedure to run, as Proc or Module.Proc. Empty clears the link."
                )
            ),
        ],
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_set_shape_macro")
        path = _excel_path(file_path, settings)
        from ..shapes import set_shape_macro as write_macro

        result = write_macro(path, sheet, shape_name, macro)
        result["path"] = str(path)
        result["saved"] = True
        result["note"] = (
            "The link is stored. Excel runs the procedure on the next click, and refuses "
            "one that is not in the project, so check it exists with xlide_list_procedures."
        )
        return result

    @server.tool(
        name="xlide_write_cells",
        title="Write cells",
        annotations=writes("Write cells", destructive=True),
        description=(
            "Writes a rectangular block of values and formulas into a worksheet, starting at "
            "one cell, and saves the file. Each row of data is a row of the sheet. A string "
            "starting with '=' is written as a formula, as you would type it into Excel 365 "
            "with no _xlfn prefixes; anything else is a value. A value written over a formula "
            "removes that formula, which is what typing into the cell does. Only the rows you "
            "touch are rewritten, so charts, styles, pivot caches and the VBA project are "
            "untouched. Ask the user before overwriting cells that hold data. Nothing "
            "recalculates until Excel next opens the workbook."
        ),
    )
    def write_cells(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        start_cell: Annotated[
            str, Field(description="Top-left cell of the block, such as B2.")
        ],
        data: Annotated[
            list[list[Any]],
            Field(
                description=(
                    "Rows of cell values. Numbers, strings, booleans, null for empty, and "
                    "strings starting with '=' for formulas."
                )
            ),
        ],
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_write_cells")
        path = _excel_path(file_path, settings)
        if not data:
            raise ToolError("data is empty; nothing to write.")

        book = Workbook(path)
        sheet_name = book.canonical_sheet_name(sheet)
        # What the block held before, so the answer can say what was displaced.
        # The user gets told they overwrote data by the agent, not by the workbook.
        first_row, first_column = parse_cell_ref(start_cell)
        target = CellRange(
            first_row,
            first_column,
            first_row + len(data) - 1,
            first_column + max((len(row) for row in data), default=1) - 1,
        )
        replaced_values = replaced_formulas = 0
        if target.cell_count <= MAX_CELLS_PER_READ:
            _, before_grid = book.read(sheet_name, str(target))
            for row in before_grid:
                for cell in row:
                    if cell.formula is not None:
                        replaced_formulas += 1
                    elif cell.value is not None:
                        replaced_values += 1

        written, count = book.write(sheet_name, start_cell, data)
        book.save()
        result: dict[str, Any] = {
            "path": str(path),
            "sheet": sheet_name,
            "range": str(written),
            "cells_written": count,
            "cells_overwritten": replaced_values + replaced_formulas,
            "formulas_replaced": replaced_formulas,
            "saved": True,
            "recalculated": False,
            "note": (
                "Values and formulas are stored. Excel recalculates the workbook the next time "
                "it opens it; until then, do not report a computed result as current."
            ),
        }
        if replaced_values or replaced_formulas:
            result["warning"] = (
                f"{replaced_values + replaced_formulas} cells already held content and were "
                f"overwritten, {replaced_formulas} of them formulas. Tell the user."
            )
        return result


def _excel_path(raw: str, settings: Settings) -> Path:
    path = resolve_path(raw, settings)
    info = host_info(path)
    if info.host != "excel":
        raise ToolError(
            f"{path.name} is a {info.title} file. Worksheet cells exist in Excel files only."
        )
    if not info.supports_sheets:
        raise ToolError(
            f"Worksheet cells are read from the OOXML package, which {info.extension} is not. "
            "Supported: .xlsx, .xlsm and .xlam. The VBA project in this file is still readable."
        )
    return path
