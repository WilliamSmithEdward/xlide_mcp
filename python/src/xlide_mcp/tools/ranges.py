"""Excel's Sort, Remove Duplicates, and Copy/Paste over worksheet ranges."""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import BaseModel, Field

from .. import cells
from ..config import Settings
from ..errors import ToolError
from ..paths import require_writable
from ._common import page, read_only, writes
from .sheets import excel_with_sheets


class SortColumn(BaseModel):
    column: str = Field(description="Column letter for a range/filter, or header name for a table.")
    descending: bool = Field(default=False, description="Largest values first.")


PASTE_KINDS = (
    "all", "formulas", "values", "formats", "comments", "validation",
    "all_except_borders", "column_widths", "formulas_and_number_formats",
    "values_and_number_formats", "all_merging_conditional_formats", "link",
)
PASTE_OPERATIONS = ("add", "subtract", "multiply", "divide")


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_check_cells",
        title="Find Excel cell error indicators",
        annotations=read_only("Find Excel cell error indicators"),
        description=(
            "Lists cells Excel would mark with a green error-checking triangle, such as "
            "numbers stored as text, inconsistent formulas, or formulas with cached errors. "
            "Returns the rule and cell for each finding. Excel hides ignored findings unless "
            "include_ignored=true. Formula-dependent checks use the file's cached results, "
            "which may be stale until Excel recalculates and saves it. Works on .xlsx, .xlsm "
            "and .xlam; page through long lists with offset and next_offset."
        ),
    )
    def check_cells(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        rules: Annotated[
            list[str] | None,
            Field(default=None, description="Excel error-rule names; omit for its enabled rules."),
        ] = None,
        include_ignored: Annotated[
            bool, Field(default=False, description="Include findings Excel was told to ignore.")
        ] = False,
        offset: Annotated[
            int, Field(default=0, ge=0, description="Findings to skip.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="Most findings to return.")
        ] = 300,
    ) -> dict[str, Any]:
        from pyofficeeditor.exceptions import PyOfficeEditorError

        path = excel_with_sheets(file_path, settings)
        with cells.open_workbook(path) as book:
            owner = cells.sheet_named(book, sheet)
            try:
                found = owner.error_checks(rules or None, include_ignored=include_ignored)
            except (PyOfficeEditorError, ValueError) as exc:
                raise ToolError(f"Could not check cells: {exc}") from exc
            name = owner.name
        entries = [
            {"cell": str(check.reference), "rule": check.rule, "ignored": check.ignored}
            for check in found
        ]
        shown, next_offset = page(entries, "diagnostics", offset, max_results)
        result: dict[str, Any] = {
            "path": str(path), "sheet": name, "count": len(entries), "checks": shown,
            "offset": offset, "next_offset": next_offset,
        }
        if next_offset is not None:
            result["note"] = (
                f"{len(entries)} checks in all; call again with offset={next_offset} "
                "for the next page."
            )
        return result

    @server.tool(
        name="xlide_sort_rows",
        title="Sort worksheet rows",
        annotations=writes("Sort worksheet rows", destructive=True),
        description=(
            "Sorts a range, a table's data rows, or a sheet autofilter's data rows, and "
            "saves the workbook. Keys are evaluated in order; each names a column letter "
            "for a range or filter, or a header name for a table. A range can keep its first "
            "row as a header. Hidden rows keep their places. Formulas, formatting, notes and "
            "links move with their rows. This changes existing data order and has no undo in "
            "the file: ask the user before calling it. Works on .xlsx, .xlsm and .xlam."
        ),
    )
    def sort_rows(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        keys: Annotated[
            list[SortColumn], Field(min_length=1, max_length=64,
                description="Ordered sort keys; each names a column and direction.")
        ],
        target: Annotated[
            str, Field(default="range", description="range, table or filter.")
        ] = "range",
        cell_range: Annotated[
            str, Field(default="", description="For target='range': block such as A1:D20.")
        ] = "",
        table_name: Annotated[
            str, Field(default="", description="For target='table': table name.")
        ] = "",
        header: Annotated[
            bool, Field(default=False, description="For a range: leave its first row in place.")
        ] = False,
        match_case: Annotated[
            bool, Field(default=False, description="Use case-sensitive text order.")
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_sort_rows")
        wanted = target.strip().lower()
        if wanted not in {"range", "table", "filter"}:
            raise ToolError("target must be 'range', 'table' or 'filter'.")
        if wanted == "range" and not cell_range.strip():
            raise ToolError("cell_range is required for target='range'.")
        if wanted == "table" and not table_name.strip():
            raise ToolError("table_name is required for target='table'.")
        from pyofficeeditor.excel import SortKey
        from pyofficeeditor.exceptions import PyOfficeEditorError

        by = [SortKey(key.column.strip(), key.descending) for key in keys]
        path = excel_with_sheets(file_path, settings)
        with cells.editing(path) as book:
            owner = cells.sheet_named(book, sheet)
            try:
                if wanted == "range":
                    owner.sort(cell_range.strip(), by, header=header, match_case=match_case)
                elif wanted == "table":
                    owner.sort_table(table_name.strip(), by, match_case=match_case)
                else:
                    owner.sort_auto_filter(by, match_case=match_case)
            except (PyOfficeEditorError, ValueError, KeyError) as exc:
                raise ToolError(f"Could not sort {wanted}: {exc}") from exc
            name = owner.name
        return {
            "path": str(path), "sheet": name, "target": wanted, "saved": True,
            "recalculated": False, "keys": [key.model_dump() for key in keys],
            "note": "Excel recalculates formulas the next time it opens the workbook.",
        }

    @server.tool(
        name="xlide_remove_duplicates",
        title="Remove duplicate worksheet rows",
        annotations=writes("Remove duplicate worksheet rows", destructive=True),
        description=(
            "Removes later rows with the same values in the chosen columns, keeping the "
            "first row of each set, and saves the workbook. A range inside a table acts on "
            "the table and shrinks it. An empty columns list compares every column. This "
            "deletes data and cannot be undone in the file: ask the user before calling it. "
            "Works on .xlsx, .xlsm and .xlam."
        ),
    )
    def remove_duplicates(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        cell_range: Annotated[str, Field(description="Range such as A1:D20.")],
        columns: Annotated[
            list[str] | None,
            Field(default=None, description="Column letters to compare; omit for all."),
        ] = None,
        header: Annotated[
            bool, Field(default=False, description="For a plain range: keep its first row.")
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_remove_duplicates")
        from pyofficeeditor.exceptions import PyOfficeEditorError

        path = excel_with_sheets(file_path, settings)
        with cells.editing(path) as book:
            owner = cells.sheet_named(book, sheet)
            try:
                removed = owner.remove_duplicates(
                    cell_range.strip(), columns or None, header=header
                )
            except (PyOfficeEditorError, ValueError, KeyError) as exc:
                raise ToolError(f"Could not remove duplicates: {exc}") from exc
            name = owner.name
        return {
            "path": str(path), "sheet": name, "range": cell_range.strip(),
            "rows_removed": removed, "saved": True, "recalculated": False,
            "note": "Excel recalculates formulas the next time it opens the workbook.",
        }

    @server.tool(
        name="xlide_copy_cells",
        title="Copy or paste cells",
        annotations=writes("Copy or paste cells", destructive=True),
        description=(
            "Copies a range to another cell or block, on the same sheet or another sheet, "
            "and saves the workbook. Paste Special can select values, formulas, formats, "
            "comments, validation, links or other parts; it can also transpose, skip blanks "
            "or combine values arithmetically. The destination may hold data, which the paste "
            "can replace. Set allow_overwrite=true only after the user agrees to replace it. "
            "Formulas and their relative references move as Excel's Copy does. Works on "
            ".xlsx, .xlsm and .xlam."
        ),
    )
    def copy_cells(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Source worksheet name.")],
        source_range: Annotated[str, Field(description="Source range such as A1:D20.")],
        destination: Annotated[str, Field(description="Destination cell or repeatable block.")],
        destination_sheet: Annotated[
            str, Field(default="", description="Destination sheet; empty means the source sheet.")
        ] = "",
        paste: Annotated[
            str,
            Field(
                default="all",
                description=(
                    "all, values, formulas, formats, comments, validation, link, "
                    "or another Excel Paste Special kind."
                ),
            ),
        ] = "all",
        operation: Annotated[
            str,
            Field(default="", description="add, subtract, multiply or divide; empty replaces."),
        ] = "",
        skip_blanks: Annotated[
            bool,
            Field(
                default=False,
                description="Leave destination cells unchanged for blank source cells.",
            ),
        ] = False,
        transpose: Annotated[
            bool, Field(default=False, description="Turn source rows into destination columns.")
        ] = False,
        allow_overwrite: Annotated[
            bool,
            Field(
                default=False,
                description="Required: the paste may replace data. Set after user agreement.",
            ),
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_copy_cells")
        if not allow_overwrite:
            raise ToolError(
                "Copying can replace destination data. Ask the user, then pass "
                "allow_overwrite=true to proceed. No cells were changed."
            )
        chosen = paste.strip().lower()
        if chosen not in PASTE_KINDS:
            raise ToolError(f"paste must be one of: {', '.join(PASTE_KINDS)}.")
        operation_name = operation.strip().lower()
        if operation_name and operation_name not in PASTE_OPERATIONS:
            raise ToolError(f"operation must be one of: {', '.join(PASTE_OPERATIONS)}.")
        from pyofficeeditor.exceptions import PyOfficeEditorError

        path = excel_with_sheets(file_path, settings)
        with cells.editing(path) as book:
            source = cells.sheet_named(book, sheet)
            target = cells.sheet_named(book, destination_sheet or sheet)
            try:
                pasted = source.copy_range(
                    source_range.strip(), destination.strip(), to=target, paste=chosen,
                    operation=operation_name or None, skip_blanks=skip_blanks,
                    transpose=transpose,
                )
            except (PyOfficeEditorError, ValueError, KeyError) as exc:
                raise ToolError(f"Could not copy cells: {exc}") from exc
            source_name, target_name = source.name, target.name
        return {
            "path": str(path), "sheet": source_name, "source_range": source_range.strip(),
            "destination_sheet": target_name, "destination_range": pasted.a1,
            "paste": chosen, "saved": True, "recalculated": False,
            "note": "Excel recalculates formulas the next time it opens the workbook.",
        }
