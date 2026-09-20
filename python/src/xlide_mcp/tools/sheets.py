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

from .. import cells, grid
from ..config import Settings, clamp_timeout
from ..errors import ToolError
from ..hosts import host_info
from ..paths import require_writable, resolve_path
from ._common import bound, read_only, writes


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
        timeout: Annotated[
            float,
            Field(default=0, ge=0, description="Seconds allowed if Excel has to answer."),
        ] = 0,
    ) -> dict[str, Any]:
        path, info = _any_excel(file_path, settings)
        if not info.supports_sheets:
            return _through_excel_sheets(path, info, clamp_timeout(timeout or None, settings))
        return {
            "path": str(path),
            "source": "file",
            "sheets": [
                {
                    "name": sheet.name,
                    "used_range": sheet.used_range or "(empty)",
                    "hidden": sheet.hidden,
                }
                for sheet in cells.sheets(path)
            ],
            "named_ranges": [
                {"name": n.name, "refers_to": n.refers_to} for n in cells.named_ranges(path)
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
        timeout: Annotated[
            float,
            Field(default=0, ge=0, description="Seconds allowed if Excel has to answer."),
        ] = 0,
    ) -> dict[str, Any]:
        path, info = _any_excel(file_path, settings)
        wanted = (include or "values").strip().lower()
        if wanted not in {"values", "formulas", "both"}:
            raise ToolError("include must be 'values', 'formulas' or 'both'.")
        if not info.supports_sheets:
            return _through_excel_read(
                path, info, sheet, cell_range, wanted,
                clamp_timeout(timeout or None, settings),
            )

        block = cells.read(path, sheet, cell_range)
        result: dict[str, Any] = {
            "path": str(path),
            "sheet": block.sheet,
            "source": "file",
            "recalculated": False,
            "range": block.reference,
            "rows": block.rows,
            "columns": block.columns,
        }
        if wanted in {"values", "both"}:
            result["values"] = block.values
        if wanted in {"formulas", "both"}:
            result["formulas"] = block.formulas
        formula_count = block.formula_count
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
        path = excel_with_sheets(file_path, settings)
        from ..shapes import read_sheet_shapes

        by_sheet = read_sheet_shapes(path, sheet.strip() or None)
        sheets = []
        notes: list[str] = []
        for name, shapes in by_sheet.items():
            shown, note = bound(
                [shape.summary() for shape in shapes],
                "shapes",
                f"Ask for one sheet with sheet={name!r}.",
            )
            sheets.append({"sheet": name, "shapes": shown})
            if note:
                notes.append(f"{name}: {note}")
        total = sum(len(shapes) for shapes in by_sheet.values())
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
        if notes:
            result["note"] = " ".join(notes)
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
        path = excel_with_sheets(file_path, settings)
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
        timeout: Annotated[
            float,
            Field(default=0, ge=0, description="Seconds allowed if Excel has to do the write."),
        ] = 0,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_write_cells")
        path, info = _any_excel(file_path, settings)
        if not data:
            raise ToolError("data is empty; nothing to write.")
        if not info.supports_sheets:
            return _through_excel_write(
                path, info, sheet, start_cell, data,
                clamp_timeout(timeout or None, settings),
            )

        written = cells.write(path, sheet, start_cell, data)
        result: dict[str, Any] = {
            "path": str(path),
            "sheet": written.sheet,
            "range": written.reference,
            "cells_written": written.cells_written,
            "source": "file",
            "cells_overwritten": written.cells_overwritten,
            "formulas_replaced": written.formulas_replaced,
            "saved": True,
            "recalculated": False,
            "note": (
                "Values and formulas are stored. Excel recalculates the workbook the next time "
                "it opens it; until then, do not report a computed result as current."
            ),
        }
        if written.cells_overwritten:
            result["warning"] = (
                f"{written.cells_overwritten} cells already held content and were "
                f"overwritten, {written.formulas_replaced} of them formulas. Tell the user."
            )
        return result


def excel_with_sheets(raw: str, settings: Settings) -> Path:
    """An Excel file whose grid the package reader can open.

    Public because every tool that edits the document surface needs the same
    gate, and a second copy of it would be a second set of refusal messages.
    """
    path, info = _any_excel(raw, settings)
    if not info.supports_sheets:
        raise ToolError(
            f"Worksheet cells are read from the OOXML package, which {info.extension} is not. "
            "Supported: .xlsx, .xlsm and .xlam. The VBA project in this file is still readable."
        )
    return path


def _any_excel(raw: str, settings: Settings) -> tuple[Path, Any]:
    """An Excel file of any format, with how its grid has to be reached."""
    path = resolve_path(raw, settings)
    info = host_info(path)
    if info.host != "excel":
        raise ToolError(
            f"{path.name} is a {info.title} file. Worksheet cells exist in Excel files only."
        )
    return path, info


# --------------------------------------------------- the formats Excel opens


def _through_excel_sheets(path: Path, info: Any, timeout: float) -> dict[str, Any]:
    if not grid.excel_available():
        raise grid.refuse(info, "Listing worksheets")
    sheets = grid.survey_sheets(path, timeout)
    return {
        "path": str(path),
        "source": "excel",
        "sheets": [
            {"name": s.name, "used_range": s.used_range, "hidden": s.hidden} for s in sheets
        ],
        "named_ranges": [],
        "note": (
            f"{info.extension} keeps its grid in a binary format this server does not read, "
            "so Excel answered. Named ranges are not surveyed on this path."
        ),
    }


def _through_excel_read(
    path: Path, info: Any, sheet: str, cell_range: str, include: str, timeout: float
) -> dict[str, Any]:
    if not grid.excel_available():
        raise grid.refuse(info, "Reading cells")
    if include != "values":
        raise ToolError(
            f"Only values can be read from {info.extension}. Excel answers this one, and it "
            "returns what each cell evaluates to rather than the formula behind it."
        )
    rows = grid.read_cells(path, sheet, cell_range, timeout)
    return {
        "path": str(path),
        "sheet": sheet,
        "source": "excel",
        "recalculated": True,
        "range": cell_range,
        "rows": len(rows),
        "columns": max((len(r) for r in rows), default=0),
        "values": rows,
        "note": (
            f"{info.extension} keeps its grid in a binary format this server does not read, "
            "so Excel opened the workbook and these values are the ones it has just "
            "calculated, not a cached result."
        ),
    }


def _through_excel_write(
    path: Path, info: Any, sheet: str, start_cell: str, data: list[list[Any]], timeout: float
) -> dict[str, Any]:
    if not grid.excel_available():
        raise grid.refuse(info, "Writing cells")
    result = grid.write_cells(path, sheet, start_cell, data, timeout)
    rows = len(data)
    columns = max((len(row) for row in data), default=0)
    return {
        "path": str(path),
        "sheet": sheet,
        "source": "excel",
        "cells_written": rows * columns,
        "saved": bool(result.get("saved")),
        "recalculated": True,
        "note": (
            f"{info.extension} keeps its grid in a binary format this server does not write, "
            "so Excel made the change and saved the workbook in its own format. Unlike a "
            "write to the package, this one did recalculate."
        ),
    }
