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
            "hidden, and the pivot tables on each, the chart sheets, and the workbook's named "
            "ranges. Call this before reading cells, so the range you ask for is one that "
            "holds data. Works on .xlsx, .xlsm and .xlam; a .xlsb or .xls keeps its grid in a "
            "binary format, and Excel answers for those on Windows."
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
        result: dict[str, Any] = {
            "path": str(path),
            "source": "file",
            "sheets": [sheet.summary() for sheet in cells.sheets(path)],
            "named_ranges": [
                {"name": n.name, "refers_to": n.refers_to} for n in cells.named_ranges(path)
            ],
        }
        charts = cells.chart_sheets(path)
        if charts:
            result["chart_sheets"] = charts
        return result

    @server.tool(
        name="xlide_read_cells",
        title="Read cells",
        annotations=read_only("Read cells"),
        description=(
            "Reads a range of cells from a worksheet and returns a grid of values, formulas, "
            "both, or the text each cell shows under its number format. Values are what Excel "
            "last calculated and stored, so a formula whose inputs changed outside Excel shows "
            "its old result; calculate=true works every formula out first with pyOfficeEditor's "
            "formula engine, in memory, and names any cell it could not work out. Ask for "
            "formulas to understand what a sheet computes, and values for what it shows. At "
            "most 20,000 cells per call."
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
                description=(
                    "'values', 'formulas', 'both', 'text' for what each cell shows, such as "
                    "12.5% or 1/15/2024, or 'rich_text' for the font runs of text written in "
                    "more than one font."
                ),
            ),
        ] = "values",
        calculate: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Work every formula out before reading, with the formula engine. The file "
                    "is not changed. Use it after writing inputs, to see the results now."
                ),
            ),
        ] = False,
        timeout: Annotated[
            float,
            Field(default=0, ge=0, description="Seconds allowed if Excel has to answer."),
        ] = 0,
    ) -> dict[str, Any]:
        path, info = _any_excel(file_path, settings)
        wanted = (include or "values").strip().lower()
        if wanted not in {"values", "formulas", "both", "text", "rich_text"}:
            raise ToolError(
                "include must be 'values', 'formulas', 'both', 'text' or 'rich_text'."
            )
        if not info.supports_sheets:
            return _through_excel_read(
                path, info, sheet, cell_range, wanted,
                clamp_timeout(timeout or None, settings),
            )

        block = cells.read(
            path,
            sheet,
            cell_range,
            calculate=calculate,
            text=wanted == "text",
            rich=wanted == "rich_text",
        )
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
        if wanted == "text":
            result["text"] = block.texts
        if wanted == "rich_text":
            result["rich_text"] = block.rich
        formula_count = block.formula_count
        if block.calculated:
            kept = block.kept_cached or {}
            result["recalculated"] = not kept
            result["calculated_by"] = "pyOfficeEditor formula engine"
            if kept:
                result["kept_cached"] = kept
            result["note"] = (
                "Worked out in memory by pyOfficeEditor's formula engine, which gives what "
                "Excel cached for the 10,958 formulas it is tested against; the file still "
                "holds what Excel last calculated. "
                + (
                    f"{len(kept)} cells here kept Excel's cached value, listed in kept_cached "
                    "with the reason."
                    if kept
                    else "Every formula here was worked out."
                )
            )
        elif formula_count and wanted != "formulas":
            result["note"] = (
                f"{formula_count} of these cells hold formulas. The values are what Excel last "
                "calculated and stored, not a recalculation. calculate=true works them out."
            )
        return result

    @server.tool(
        name="xlide_evaluate_formula",
        title="Evaluate a formula",
        annotations=read_only("Evaluate a formula"),
        description=(
            "Works out what a formula would give in a cell of a workbook, without writing it "
            "there: to check a formula before xlide_write_cells, or to ask the workbook a "
            "question, such as =SUMIFS(Sales[Amount],Sales[Region],\"West\"). Uses "
            "pyOfficeEditor's formula engine, 493 of Excel's functions, with the workbook's "
            "cells as they stand. A formula written for one cell answers with one value, as "
            "Excel would give it there. Changes nothing. .xlsx, .xlsm and .xlam."
        ),
    )
    def evaluate_formula(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet the formula is evaluated on.")],
        formula: Annotated[str, Field(description="The formula, with or without its '='.")],
        at: Annotated[
            str,
            Field(
                default="A1",
                description="The cell it is imagined in, which relative references count from.",
            ),
        ] = "A1",
    ) -> dict[str, Any]:
        path = excel_with_sheets(file_path, settings)
        value = cells.evaluate(path, sheet, formula, at)
        return {
            "path": str(path),
            "sheet": sheet,
            "formula": formula if formula.strip().startswith("=") else "=" + formula.strip(),
            "at": (at or "A1").strip().upper(),
            "value": value,
            "calculated_by": "pyOfficeEditor formula engine",
        }

    @server.tool(
        name="xlide_list_shapes",
        title="List shapes and buttons",
        annotations=read_only("List shapes and buttons"),
        description=(
            "Lists what sits on a worksheet's drawing layer: buttons, form controls, "
            "AutoShapes, text boxes, pictures, charts and groups, each with the cells it "
            "covers, its position in points and, where it has one, the macro a click runs. "
            "A form control also carries what it holds: a check box's state, a list's "
            "chosen item, a spinner's value and bounds, the cell it is linked to. Call it "
            "when asked how a workbook is started, or before renaming a Sub: a button's "
            "OnAction names a procedure and nothing rewrites it. ActiveX controls are "
            "listed but have no macro; their code is event procedures in the sheet's module."
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
        from ..shapes import control_states, read_sheet_shapes

        by_sheet = read_sheet_shapes(path, sheet.strip() or None)
        states = control_states(path)
        sheets = []
        notes: list[str] = []
        for name, shapes in by_sheet.items():
            known = states.get(name, {})
            summaries = []
            for shape in shapes:
                entry = shape.summary()
                entry.update(known.get(shape.name.casefold(), {}))
                summaries.append(entry)
            shown, note = bound(
                summaries,
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
            "the project, so write it first. To add a button with its macro, or to remove "
            "one, use xlide_manage_shape."
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
        name="xlide_manage_shape",
        title="Add or remove a shape or button",
        annotations=writes("Add or remove a shape or button", destructive=True),
        description=(
            "Adds a Forms control, an AutoShape, a text box, a line or a picture to a "
            "worksheet, or removes one, and saves the workbook. The usual case is a button "
            "that runs a macro: kind='button', a caption in text, and the procedure in macro "
            "as Proc or Module.Proc, written first. Place it with cell, its top-left cell, "
            "or with left and top in points; width and height are points, with a size to "
            "start from when left at 0. Check boxes, option buttons, lists and drop-downs "
            "take linked_cell, and lists and drop-downs list_range. A Forms control lives "
            "in four parts that have to agree, and they are written and removed together. "
            "Removing a shape that runs a macro leaves the macro; removing a chart is not "
            "offered. Excel workbooks only; .xlsx, .xlsm and .xlam."
        ),
    )
    def manage_shape(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[str, Field(description="'add' or 'remove'.")],
        shape_name: Annotated[
            str, Field(description="The shape's name: new for add, as listed for remove.")
        ],
        kind: Annotated[
            str,
            Field(
                default="button",
                description=(
                    "For add: button, checkBox, optionButton, dropDown, listBox, scrollBar, "
                    "spinner, label, groupBox, shape, textBox, line or picture."
                ),
            ),
        ] = "button",
        cell: Annotated[
            str, Field(default="", description="For add: the top-left cell, such as B2.")
        ] = "",
        left: Annotated[
            float, Field(default=-1.0, description="For add without cell: points from the left.")
        ] = -1.0,
        top: Annotated[
            float, Field(default=-1.0, description="For add without cell: points from the top.")
        ] = -1.0,
        width: Annotated[float, Field(default=0.0, ge=0, description="Points. 0 picks one.")] = 0.0,
        height: Annotated[
            float, Field(default=0.0, ge=0, description="Points. 0 picks one.")
        ] = 0.0,
        text: Annotated[
            str,
            Field(
                default="",
                description="A control's caption, a shape's text, or a picture's alt text.",
            ),
        ] = "",
        macro: Annotated[
            str, Field(default="", description="The procedure a click runs: Proc or Module.Proc.")
        ] = "",
        linked_cell: Annotated[
            str, Field(default="", description="The cell a control writes its value to: $D$6.")
        ] = "",
        list_range: Annotated[
            str,
            Field(default="", description="The cells a list or drop-down offers: $H$1:$H$9."),
        ] = "",
        geometry: Annotated[
            str,
            Field(
                default="",
                description="For kind='shape': the preset, such as rect, roundRect or ellipse.",
            ),
        ] = "",
        image_path: Annotated[
            str, Field(default="", description="For kind='picture': a PNG, JPEG or GIF file.")
        ] = "",
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_manage_shape")
        path = excel_with_sheets(file_path, settings)
        wanted = (action or "").strip().lower()
        if wanted not in {"add", "remove"}:
            raise ToolError("action must be 'add' or 'remove'.")
        if not shape_name.strip():
            raise ToolError("shape_name is required.")
        from .. import shapes as drawing

        if wanted == "remove":
            detail = drawing.remove_shape(path, sheet, shape_name.strip())
            result: dict[str, Any] = {"path": str(path), "action": wanted, "saved": True, **detail}
            if detail.get("had_macro"):
                result["note"] = (
                    f"It ran {detail['had_macro']}, which is still in the project; nothing "
                    "else calls it from this sheet unless another shape does."
                )
            return result

        image = resolve_path(image_path, settings) if image_path.strip() else None
        detail = drawing.add_shape(
            path,
            sheet,
            name=shape_name.strip(),
            kind=kind.strip(),
            cell=cell,
            left=left if left >= 0 else None,
            top=top if top >= 0 else None,
            width=width,
            height=height,
            text=text,
            macro=macro,
            linked_cell=linked_cell,
            list_range=list_range,
            geometry=geometry,
            image=image,
        )
        result = {"path": str(path), "action": wanted, "saved": True, **detail}
        if macro.strip():
            result["note"] = (
                f"A click runs {macro.strip()}. Excel refuses one that is not in the project, "
                "so check it exists with xlide_list_procedures."
            )
        return result

    @server.tool(
        name="xlide_add_chart",
        title="Add a chart",
        annotations=writes("Add a chart", destructive=False, idempotent=False),
        description=(
            "Adds a chart of a block of cells to a worksheet and saves the workbook, written "
            "as Excel's Insert Chart writes one of the same type. The block's first row names "
            "the series and its first column holds the categories, as Excel reads a selection; "
            "a block taller than it is wide makes a series of each column. The data can be on "
            "another sheet, as Data!A1:C13. Place it with cell or left and top in points. "
            "xlide_list_shapes reports it afterwards with each series' formula."
        ),
    )
    def add_chart(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet the chart goes on.")],
        data_range: Annotated[
            str,
            Field(description="The block to chart, header row and category column included."),
        ],
        chart_type: Annotated[
            str,
            Field(
                default="column",
                description="column, bar, line, lineMarkers, pie, doughnut, scatter or area.",
            ),
        ] = "column",
        cell: Annotated[
            str, Field(default="", description="The chart's top-left cell, such as H2.")
        ] = "",
        left: Annotated[
            float, Field(default=-1.0, description="Without cell: points from the left.")
        ] = -1.0,
        top: Annotated[
            float, Field(default=-1.0, description="Without cell: points from the top.")
        ] = -1.0,
        width: Annotated[
            float, Field(default=0.0, ge=0, description="Points. 0 gives Excel's 360.")
        ] = 0.0,
        height: Annotated[
            float, Field(default=0.0, ge=0, description="Points. 0 gives Excel's 216.")
        ] = 0.0,
        title: Annotated[
            str, Field(default="", description="The title. Empty keeps Excel's automatic one.")
        ] = "",
        chart_name: Annotated[
            str, Field(default="", description="Its name. Empty gives Chart 1, Chart 2 and so on.")
        ] = "",
        series_in: Annotated[
            str,
            Field(
                default="",
                description="'columns' or 'rows' to say which way the series run. Empty decides.",
            ),
        ] = "",
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_add_chart")
        path = excel_with_sheets(file_path, settings)
        from .. import shapes as drawing

        detail = drawing.add_chart(
            path,
            sheet,
            data=data_range,
            chart_type=chart_type,
            cell=cell,
            left=left if left >= 0 else None,
            top=top if top >= 0 else None,
            width=width,
            height=height,
            title=title,
            name=chart_name,
            series_in=series_in,
        )
        return {"path": str(path), "saved": True, **detail}

    @server.tool(
        name="xlide_write_cells",
        title="Write cells",
        annotations=writes("Write cells", destructive=True),
        description=(
            "Writes a rectangular block of values and formulas into a worksheet, starting at "
            "one cell, and saves the file. Each row of data is a row of the sheet. A string "
            "starting with '=' is written as a formula, as you would type it into Excel 365 "
            "with no _xlfn prefixes; anything else is a value. A cell given as "
            "{\"rich_text\": [{\"text\": \"Total \", \"bold\": true}, {\"text\": \"42\"}]} "
            "is text in more than one font. A value written over a formula removes that "
            "formula, which is what typing into the cell does. Only the rows you touch are "
            "rewritten, so charts, styles, pivot caches and the VBA project are untouched. Ask "
            "the user before overwriting cells that hold data. The file keeps no calculated "
            "result until Excel next opens it; xlide_read_cells with calculate=true works "
            "the results out now."
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
                    "Rows of cell values. Numbers, strings, booleans, null for empty, strings "
                    "starting with '=' for formulas, and {\"rich_text\": [runs]} for text in "
                    "several fonts, each run a text with any of bold, italic, strike, "
                    "underline, size, font, color (RRGGBB) and script."
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
                "it opens it; until then the file holds no current result. xlide_read_cells "
                "with calculate=true works the results out now, and says which it could not."
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
