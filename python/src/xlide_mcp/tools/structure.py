"""The shape of a workbook: its sheets, and the rows and columns in them.

Both tools take an `action` rather than splitting into a dozen one-verb tools.
The operations here share their arguments almost entirely, and a calling model
choosing between `add` and `remove` on one tool makes fewer mistakes than one
choosing between twelve tool names that differ by a word.

Inserting and deleting rows is the operation with consequences. Every reference
in the workbook moves with them, and a reference that pointed into deleted cells
becomes `#REF!` exactly as it does in Excel. That is upstream's work, not this
server's, but it is why the deleting actions are marked destructive and say what
they did.
"""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import cells
from ..config import Settings
from ..errors import ToolError
from ..paths import require_writable
from ._common import writes
from .sheets import excel_with_sheets

SHEET_ACTIONS = ("add", "remove", "rename", "move", "hide", "show", "protect", "unprotect")
LINE_ACTIONS = ("insert", "delete", "resize", "hide", "show", "group", "ungroup")


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_manage_sheet",
        title="Add, remove or change a worksheet",
        annotations=writes("Add, remove or change a worksheet", destructive=True),
        description=(
            "Adds, removes, renames, moves, hides, shows, protects or unprotects a worksheet. "
            "Renaming rewrites the formulas and defined names that referred to the old name, "
            "so nothing breaks. Removing a sheet takes its cells, tables and everything on it, "
            "and formulas elsewhere that pointed at it become #REF!, which has no undo: ask "
            "the user first. Protecting a sheet stops editing in Excel; it is not a security "
            "boundary and the password is trivially recovered. Works on .xlsx, .xlsm, .xlam."
        ),
    )
    def manage_sheet(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        action: Annotated[
            str,
            Field(
                description=(
                    "add, remove, rename, move, hide, show, protect or unprotect."
                )
            ),
        ],
        sheet: Annotated[
            str,
            Field(
                default="",
                description="The sheet to act on. For add, the name to give the new one.",
            ),
        ] = "",
        new_name: Annotated[
            str, Field(default="", description="For rename: the new name.")
        ] = "",
        index: Annotated[
            int,
            Field(
                default=-1,
                description="For add and move: 0-based position. -1 means the end.",
            ),
        ] = -1,
        very_hidden: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "For hide: hide it so that Excel's own Unhide dialog does not list it. "
                    "Only the VBA editor can bring it back."
                ),
            ),
        ] = False,
        password: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For protect: an optional password. Excel's sheet password is obfuscation, "
                    "not encryption, so do not use one the user relies on elsewhere."
                ),
            ),
        ] = "",
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_manage_sheet")
        path = excel_with_sheets(file_path, settings)
        wanted = _action(action, SHEET_ACTIONS)

        with cells.editing(path) as book:
            detail = _do_sheet(book, wanted, sheet, new_name, index, very_hidden, password)
            names = list(book.sheet_names)

        return {
            "path": str(path),
            "action": wanted,
            "saved": True,
            **detail,
            "sheets": names,
        }

    @server.tool(
        name="xlide_manage_rows_columns",
        title="Insert, delete or size rows and columns",
        annotations=writes("Insert, delete or size rows and columns", destructive=True),
        description=(
            "Inserts, deletes, resizes, hides, shows, groups or ungroups whole rows or "
            "columns. Inserting and deleting move every reference in the workbook with them: "
            "a formula pointing below an inserted row follows it, and one pointing into "
            "deleted cells becomes #REF!, exactly as Excel does it. Deleting has no undo, so "
            "ask the user first. Widths are in characters and heights in points, which is what "
            "Excel's own dialogs use. Works on .xlsx, .xlsm and .xlam."
        ),
    )
    def manage_rows_columns(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[
            str,
            Field(description="insert, delete, resize, hide, show, group or ungroup."),
        ],
        which: Annotated[str, Field(description="'rows' or 'columns'.")],
        first: Annotated[
            int, Field(ge=1, description="First row or column, 1-based. Column A is 1.")
        ],
        count: Annotated[
            int,
            Field(default=1, ge=1, description="How many, counting from first."),
        ] = 1,
        size: Annotated[
            float,
            Field(
                default=-1,
                description=(
                    "For resize: column width in characters, or row height in points. "
                    "-1 restores the sheet default."
                ),
            ),
        ] = -1,
        collapsed: Annotated[
            bool,
            Field(default=False, description="For group: start the new group collapsed."),
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_manage_rows_columns")
        path = excel_with_sheets(file_path, settings)
        wanted = _action(action, LINE_ACTIONS)
        rows = _rows_or_columns(which)
        last = first + count - 1

        with cells.editing(path) as book:
            sheet_object = cells.sheet_named(book, sheet)
            _do_lines(sheet_object, wanted, rows, first, count, last, size, collapsed)
            name = sheet_object.name

        return {
            "path": str(path),
            "sheet": name,
            "action": wanted,
            "which": "rows" if rows else "columns",
            "first": first,
            "count": count,
            "saved": True,
            "recalculated": False,
            "note": (
                "Excel recalculates the workbook the next time it opens it. Until then a "
                "formula that moved keeps its old cached result."
            ),
        }


# ------------------------------------------------------------------ the parts


def _do_sheet(
    book: Any,
    action: str,
    sheet: str,
    new_name: str,
    index: int,
    very_hidden: bool,
    password: str,
) -> dict[str, Any]:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        if action == "add":
            if not sheet.strip():
                raise ToolError("sheet is the name to give the new worksheet.")
            added = book.add_sheet(sheet.strip(), None if index < 0 else index)
            return {"sheet": added.name, "created": True}

        target = cells.sheet_named(book, sheet)
        name = target.name

        if action == "remove":
            book.remove_sheet(name)
            return {"sheet": name, "removed": True}
        if action == "rename":
            if not new_name.strip():
                raise ToolError("new_name is required for action='rename'.")
            book.rename_sheet(name, new_name.strip())
            return {"renamed_from": name, "sheet": new_name.strip()}
        if action == "move":
            if index < 0:
                raise ToolError("index is required for action='move', 0 being first.")
            book.move_sheet(name, index)
            return {"sheet": name, "moved_to": index}
        if action == "hide":
            if not book.has_another_visible_sheet(name):
                raise ToolError(
                    f"{name} is the only visible sheet. A workbook with every sheet hidden "
                    "is one Excel refuses to open, so this is refused here instead."
                )
            target.visible = "veryHidden" if very_hidden else "hidden"
            return {"sheet": name, "visible": False, "very_hidden": very_hidden}
        if action == "show":
            target.visible = "visible"
            return {"sheet": name, "visible": True}
        if action == "protect":
            target.protect(password=password or None)
            return {"sheet": name, "protected": True, "password_set": bool(password)}
        target.unprotect()
        return {"sheet": name, "protected": False}
    except ToolError:
        raise
    except PyOfficeEditorError as exc:
        raise ToolError(f"{action} failed: {exc}") from exc


def _do_lines(
    sheet: Any,
    action: str,
    rows: bool,
    first: int,
    count: int,
    last: int,
    size: float,
    collapsed: bool,
) -> None:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        if action == "insert":
            (sheet.insert_rows if rows else sheet.insert_columns)(first, count)
            return
        if action == "delete":
            (sheet.delete_rows if rows else sheet.delete_columns)(first, count)
            return
        if action == "resize":
            wanted = None if size < 0 else size
            setter = sheet.set_row_height if rows else sheet.set_column_width
            for number in range(first, last + 1):
                setter(number, wanted)
            return
        if action in {"hide", "show"}:
            setter = sheet.set_row_hidden if rows else sheet.set_column_hidden
            for number in range(first, last + 1):
                setter(number, action == "hide")
            return
        if action == "group":
            grouper = sheet.group_rows if rows else sheet.group_columns
            grouper(first, last, collapsed=collapsed)
            return
        (sheet.ungroup_rows if rows else sheet.ungroup_columns)(first, last)
    except PyOfficeEditorError as exc:
        raise ToolError(f"{action} failed on {'rows' if rows else 'columns'}: {exc}") from exc


def _action(raw: str, allowed: tuple[str, ...]) -> str:
    wanted = (raw or "").strip().lower()
    if wanted not in allowed:
        raise ToolError(f"action must be one of: {', '.join(allowed)}. Got {raw!r}.")
    return wanted


def _rows_or_columns(raw: str) -> bool:
    wanted = (raw or "").strip().lower()
    if wanted in {"rows", "row"}:
        return True
    if wanted in {"columns", "column", "cols", "col"}:
        return False
    raise ToolError(f"which must be 'rows' or 'columns'. Got {raw!r}.")
