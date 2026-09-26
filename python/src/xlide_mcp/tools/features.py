"""The rest of the document surface: tables, names, validation, rules, links.

Four of these five are things a workbook uses to say what it means rather than
what it holds. A dropdown says which values are allowed, a conditional rule says
which are worth looking at, a table names a block so formulas can refer to it by
meaning, and a defined name does the same for a range. An agent asked to explain
a workbook has to read them, and an agent asked to build one has to write them.

Each tool takes an `action` and lists before it changes, so the same call answers
"what is here" and "make this so".
"""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import cells
from ..config import Settings
from ..errors import ToolError
from ..paths import require_writable, resolve_path
from ._common import page, read_only, writes
from .sheets import excel_with_sheets

VALIDATION_KINDS = ("whole", "decimal", "list", "date", "time", "textLength", "custom")
OPERATORS = (
    "between", "notBetween", "equal", "notEqual",
    "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual",
)


def register(server: MCPServer, settings: Settings) -> None:
    # ------------------------------------------------------------------ tables
    @server.tool(
        name="xlide_manage_table",
        title="Excel tables",
        annotations=writes("Excel tables", destructive=True),
        description=(
            "Lists, adds or removes Excel tables, the ListObjects that Ctrl+T creates. A table "
            "names a block so formulas can say Sales[Amount] instead of an address that breaks "
            "when rows move, and it is what a Power Query load writes into. action='list' "
            "reads them and changes nothing. Adding one takes the column names from the "
            "header row of the range you give. Removing one leaves the cells and takes the "
            "table, so structured references to it break."
        ),
    )
    def manage_table(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        action: Annotated[str, Field(default="list", description="list, add or remove.")] = "list",
        sheet: Annotated[
            str, Field(default="", description="Worksheet. Empty lists every sheet's tables.")
        ] = "",
        table_name: Annotated[
            str, Field(default="", description="For add and remove: the table's name.")
        ] = "",
        cell_range: Annotated[
            str,
            Field(
                default="",
                description="For add: the range including its header row, such as A1:D20.",
            ),
        ] = "",
        totals_row: Annotated[
            bool, Field(default=False, description="For add: give the table a totals row.")
        ] = False,
        offset: Annotated[int, Field(default=0, ge=0, description="For list: tables to skip.")] = 0,
        max_results: Annotated[
            int, Field(default=500, ge=1, le=500, description="For list: most tables to return.")
        ] = 500,
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "remove"))
        if wanted == "list":
            path = excel_with_sheets(file_path, settings)
            with cells.open_workbook(path) as book:
                found = [
                    _table_summary(table, owner.name)
                    for owner in book.sheets
                    if not sheet.strip() or owner.name.casefold() == sheet.strip().casefold()
                    for table in owner.tables
                ]
            shown, next_offset = page(found, "tables", offset, max_results)
            result: dict[str, Any] = {
                "path": str(path), "count": len(found), "offset": offset,
                "next_offset": next_offset, "tables": shown,
            }
            return result

        require_writable(settings, "xlide_manage_table")
        path = excel_with_sheets(file_path, settings)
        if not table_name.strip():
            raise ToolError(f"table_name is required for action={wanted!r}.")

        with cells.editing(path) as book:
            target = cells.sheet_named(book, sheet) if sheet.strip() else None
            detail = _do_table(book, target, wanted, table_name.strip(), cell_range, totals_row)
        return {"path": str(path), "action": wanted, "saved": True, **detail}

    # ------------------------------------------------------------ defined names
    @server.tool(
        name="xlide_manage_name",
        title="Defined names",
        annotations=writes("Defined names", destructive=True),
        description=(
            "Lists, adds or removes a workbook's defined names. A defined name is what lets a "
            "formula say TaxRate rather than Config!$B$7, and VBA reads them too, so removing "
            "one can break code as well as formulas. action='list' changes nothing. For a "
            "range, qualify refers_to with its sheet and use absolute addresses to avoid "
            "following the active sheet or moving with a copied formula, for example "
            "Data!$A$1:$A$50."
        ),
    )
    def manage_name(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        action: Annotated[str, Field(default="list", description="list, add or remove.")] = "list",
        name: Annotated[
            str, Field(default="", description="For add and remove: the defined name.")
        ] = "",
        refers_to: Annotated[
            str,
            Field(default="", description="For add: what it points at, such as Data!$A$1:$A$50."),
        ] = "",
        scope: Annotated[
            str,
            Field(
                default="",
                description="A sheet name to scope it to that sheet. Empty means the workbook.",
            ),
        ] = "",
        offset: Annotated[int, Field(default=0, ge=0, description="For list: names to skip.")] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="For list: most names to return.")
        ] = 300,
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "remove"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                found = [
                    {
                        "name": entry.name,
                        "refers_to": entry.refers_to,
                        "scope": entry.scope or "(workbook)",
                        "hidden": bool(entry.hidden),
                    }
                    for entry in book.defined_names
                ]
            shown, next_offset = page(found, "relationships", offset, max_results)
            result: dict[str, Any] = {
                "path": str(path), "count": len(found), "offset": offset,
                "next_offset": next_offset, "names": shown,
            }
            return result

        require_writable(settings, "xlide_manage_name")
        if not name.strip():
            raise ToolError(f"name is required for action={wanted!r}.")

        from pyofficeeditor.exceptions import PyOfficeEditorError

        with cells.editing(path) as book:
            try:
                if wanted == "add":
                    if not refers_to.strip():
                        raise ToolError("refers_to is required for action='add'.")
                    book.add_defined_name(
                        name.strip(), refers_to.strip(), scope=scope.strip() or None
                    )
                    detail = {"name": name.strip(), "refers_to": refers_to.strip()}
                else:
                    book.remove_defined_name(name.strip(), scope=scope.strip() or None)
                    detail = {"removed": name.strip()}
            except ToolError:
                raise
            except (PyOfficeEditorError, KeyError) as exc:
                raise ToolError(f"{wanted} failed for {name!r}: {exc}") from exc
        return {"path": str(path), "action": wanted, "saved": True, **detail}

    # --------------------------------------------------------------- validation
    @server.tool(
        name="xlide_manage_validation",
        title="Data validation",
        annotations=writes("Data validation", destructive=True),
        description=(
            "Lists, adds or clears data validation: what a cell will accept, and the dropdown "
            "it shows. kind='list' with formula1 as a comma-separated set of values gives a "
            "dropdown; kind='list' with a range reference gives one driven by cells. The other "
            "kinds take an operator and one or two formulas, which may be literals or "
            "references. Validation stops typing in Excel, not writing through this server, "
            "and Excel does not re-check cells that already held a value."
        ),
    )
    def manage_validation(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[
            str, Field(default="list", description="list, add or clear.")
        ] = "list",
        cell_range: Annotated[
            str,
            Field(
                default="",
                description="For add and clear: the range. Empty clears the whole sheet.",
            ),
        ] = "",
        kind: Annotated[
            str,
            Field(
                default="list",
                description="whole, decimal, list, date, time, textLength or custom.",
            ),
        ] = "list",
        formula1: Annotated[
            str,
            Field(
                default="",
                description=(
                    "The allowed values. For kind='list', either 'Red,Green,Blue' or a range "
                    "like $H$1:$H$9. For the others, the bound, such as 0."
                ),
            ),
        ] = "",
        formula2: Annotated[
            str,
            Field(default="", description="The second bound, for between and notBetween."),
        ] = "",
        operator: Annotated[
            str,
            Field(default="between", description="between, greaterThan, lessThan and so on."),
        ] = "between",
        allow_blank: Annotated[
            bool, Field(default=True, description="Let the cell be left empty.")
        ] = True,
        error_message: Annotated[
            str, Field(default="", description="What Excel says when the entry is refused.")
        ] = "",
        offset: Annotated[
            int, Field(default=0, ge=0, description="For list: validations to skip.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="For list: most to return.")
        ] = 300,
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "clear"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                target = cells.sheet_named(book, sheet)
                found = [_validation_summary(rule) for rule in target.data_validations]
                name = target.name
            shown, next_offset = page(found, "relationships", offset, max_results)
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found),
                "offset": offset, "next_offset": next_offset, "validations": shown,
            }
            return result

        require_writable(settings, "xlide_manage_validation")
        with cells.editing(path) as book:
            target = cells.sheet_named(book, sheet)
            if wanted == "clear":
                removed = target.clear_data_validations(cell_range.strip() or None)
                detail: dict[str, Any] = {"cleared": removed}
            else:
                detail = _add_validation(
                    target, cell_range, kind, formula1, formula2,
                    operator, allow_blank, error_message,
                )
            name = target.name
        return {"path": str(path), "sheet": name, "action": wanted, "saved": True, **detail}

    # ---------------------------------------------------- conditional formatting
    @server.tool(
        name="xlide_manage_conditional_format",
        title="Conditional formatting",
        annotations=writes("Conditional formatting", destructive=True),
        description=(
            "Lists, adds or clears conditional formatting: the rules that colour cells by what "
            "is in them. rule='cell_is' with an operator and a value paints cells that compare "
            "true; rule='expression' takes a formula written for the top-left cell of the "
            "range and applied relatively, the way Excel's 'Use a formula' box works; "
            "rule='color_scale' and 'data_bar' are the gradient and in-cell bar. The paint "
            "itself is the fill and font arguments, which is what Excel calls a differential "
            "format."
        ),
    )
    def manage_conditional_format(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[
            str, Field(default="list", description="list, add or clear.")
        ] = "list",
        cell_range: Annotated[
            str,
            Field(
                default="",
                description="For add and clear: the range. Empty clears the whole sheet.",
            ),
        ] = "",
        rule: Annotated[
            str,
            Field(
                default="cell_is",
                description="cell_is, expression, color_scale or data_bar.",
            ),
        ] = "cell_is",
        operator: Annotated[
            str,
            Field(default="greaterThan", description="For cell_is: greaterThan, between, etc."),
        ] = "greaterThan",
        value: Annotated[
            str, Field(default="", description="For cell_is: the value compared against.")
        ] = "",
        other_value: Annotated[
            str, Field(default="", description="For cell_is with between: the second value.")
        ] = "",
        formula: Annotated[
            str,
            Field(
                default="",
                description="For expression: a formula for the range's top-left cell.",
            ),
        ] = "",
        fill_color: Annotated[
            str, Field(default="", description="Fill to paint matching cells, as hex.")
        ] = "",
        font_color: Annotated[
            str, Field(default="", description="Text colour for matching cells, as hex.")
        ] = "",
        bold: Annotated[
            bool | None, Field(default=None, description="Make matching cells bold.")
        ] = None,
        offset: Annotated[
            int, Field(default=0, ge=0, description="For list: format ranges to skip.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="For list: most to return.")
        ] = 300,
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "clear"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                target = cells.sheet_named(book, sheet)
                found = [
                    {
                        "range": entry.sqref,
                        "rules": [
                            {
                                "kind": str(getattr(r, "kind", "") or ""),
                                "operator": str(getattr(r, "operator", "") or ""),
                                "priority": getattr(r, "priority", None),
                            }
                            for r in getattr(entry, "rules", []) or []
                        ],
                    }
                    for entry in target.conditional_formats
                ]
                name = target.name
            shown, next_offset = page(found, "relationships", offset, max_results)
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found),
                "offset": offset, "next_offset": next_offset, "formats": shown,
            }
            return result

        require_writable(settings, "xlide_manage_conditional_format")
        with cells.editing(path) as book:
            target = cells.sheet_named(book, sheet)
            if wanted == "clear":
                removed = target.clear_conditional_formats(cell_range.strip() or None)
                detail: dict[str, Any] = {"cleared": removed}
            else:
                detail = _add_conditional(
                    target, cell_range, rule, operator, value, other_value,
                    formula, fill_color, font_color, bold,
                )
            name = target.name
        return {"path": str(path), "sheet": name, "action": wanted, "saved": True, **detail}

    # --------------------------------------------------------------- hyperlinks
    @server.tool(
        name="xlide_manage_hyperlink",
        title="Hyperlinks",
        annotations=writes("Hyperlinks", destructive=True),
        description=(
            "Lists, adds or removes hyperlinks on a worksheet. A link either goes out to a "
            "target, which is a URL or a file path, or inside the workbook to a location such "
            "as 'Summary!A1'. Pass one or the other. The cell's displayed text is separate "
            "from the link and is left alone unless you pass display."
        ),
    )
    def manage_hyperlink(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[
            str, Field(default="list", description="list, add or remove.")
        ] = "list",
        cell_range: Annotated[
            str, Field(default="", description="For add and remove: the cell or range.")
        ] = "",
        target: Annotated[
            str, Field(default="", description="For add: a URL or file path to open.")
        ] = "",
        location: Annotated[
            str,
            Field(default="", description="For add: a place in this workbook, like Summary!A1."),
        ] = "",
        display: Annotated[
            str, Field(default="", description="For add: text to put in the cell.")
        ] = "",
        tooltip: Annotated[
            str, Field(default="", description="For add: the hover text.")
        ] = "",
        offset: Annotated[int, Field(default=0, ge=0, description="For list: links to skip.")] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="For list: most links to return.")
        ] = 300,
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "remove"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                sheet_object = cells.sheet_named(book, sheet)
                found = [
                    {
                        "range": link.ref.a1,
                        "target": getattr(link, "target", "") or "",
                        "location": getattr(link, "location", "") or "",
                        "display": getattr(link, "display", "") or "",
                    }
                    for link in sheet_object.hyperlinks
                ]
                name = sheet_object.name
            shown, next_offset = page(found, "relationships", offset, max_results)
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found),
                "offset": offset, "next_offset": next_offset, "hyperlinks": shown,
            }
            return result

        require_writable(settings, "xlide_manage_hyperlink")
        if not cell_range.strip():
            raise ToolError(f"cell_range is required for action={wanted!r}.")

        from pyofficeeditor.exceptions import PyOfficeEditorError

        with cells.editing(path) as book:
            sheet_object = cells.sheet_named(book, sheet)
            try:
                if wanted == "add":
                    if not target.strip() and not location.strip():
                        raise ToolError(
                            "Pass target for a link out to a URL or file, or location for one "
                            "inside this workbook, such as 'Summary!A1'."
                        )
                    sheet_object.add_hyperlink(
                        cell_range.strip(),
                        target.strip() or None,
                        location=location.strip() or None,
                        display=display.strip() or None,
                        tooltip=tooltip.strip() or None,
                    )
                    detail: dict[str, Any] = {"range": cell_range.strip(), "added": True}
                else:
                    detail = {
                        "range": cell_range.strip(),
                        "removed": sheet_object.remove_hyperlink(cell_range.strip()),
                    }
            except ToolError:
                raise
            except (PyOfficeEditorError, ValueError) as exc:
                raise ToolError(f"{wanted} failed on {cell_range}: {exc}") from exc
            name = sheet_object.name
        return {"path": str(path), "sheet": name, "action": wanted, "saved": True, **detail}

    # --------------------------------------------------------------- page setup
    @server.tool(
        name="xlide_page_setup",
        title="How a sheet prints",
        annotations=read_only("How a sheet prints"),
        description=(
            "Reads how a worksheet is set up to print: orientation, paper size, margins, the "
            "print area, the rows or columns repeated on every page, whether it is scaled to "
            "fit, and the header and footer. Call it when asked why a sheet prints the way it "
            "does, or before changing a layout somebody set up deliberately."
        ),
    )
    def page_setup(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
    ) -> dict[str, Any]:
        resolve_path(file_path, settings)
        path = excel_with_sheets(file_path, settings)
        with cells.open_workbook(path) as book:
            target = cells.sheet_named(book, sheet)
            setup = target.page_setup
            margins = target.page_margins
            options = target.print_options
            answer = {
                "path": str(path),
                "sheet": target.name,
                "orientation": str(getattr(setup, "orientation", "") or ""),
                "paper_size": getattr(setup, "paper_size", None),
                "scale": getattr(setup, "scale", None),
                "fit_to_page": bool(target.fit_to_page),
                "print_area": [str(area) for area in target.print_area],
                "print_titles": target.print_titles or "",
                "margins": {
                    side: getattr(margins, side, None)
                    for side in ("left", "right", "top", "bottom", "header", "footer")
                },
                "gridlines_printed": bool(getattr(options, "grid_lines", False)),
                "headings_printed": bool(getattr(options, "headings", False)),
            }
            header_footer = target.header_footer
            answer["header_footer"] = {
                part: str(getattr(getattr(header_footer, part, None), "text", "") or "")
                for part in ("odd_header", "odd_footer")
                if getattr(header_footer, part, None) is not None
            }
        return answer

    # ----------------------------------------------------------------- comments
    @server.tool(
        name="xlide_manage_comment",
        title="Notes and comments",
        annotations=writes("Notes and comments", destructive=True),
        description=(
            "Lists, writes, answers or removes the comments on a worksheet's cells. Excel has "
            "two kinds: a note, the yellow box that shows on hover, and a threaded comment, "
            "the conversation on the Review tab with replies and a resolved state. "
            "action='list' reads both and changes nothing. 'set' writes a note, replacing one "
            "already on the cell, or with kind='thread' starts a conversation; 'reply' and "
            "'resolve' act on a thread; 'remove' takes either kind off the cell. Use a note to "
            "explain a cell to whoever reads the workbook next."
        ),
    )
    def manage_comment(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[
            str, Field(default="list", description="list, set, reply, resolve or remove.")
        ] = "list",
        cell: Annotated[
            str,
            Field(default="", description="The cell, such as B4. For list, empty lists the sheet."),
        ] = "",
        text: Annotated[
            str, Field(default="", description="For set and reply: what the comment says.")
        ] = "",
        kind: Annotated[
            str, Field(default="note", description="For set: 'note' or 'thread'.")
        ] = "note",
        author: Annotated[
            str,
            Field(
                default="Agent",
                description="Who wrote it, as the workbook shows the name.",
            ),
        ] = "Agent",
        visible: Annotated[
            bool,
            Field(default=False, description="For a note: shown all the time, not only on hover."),
        ] = False,
        resolved: Annotated[
            bool, Field(default=True, description="For resolve: false opens a thread again.")
        ] = True,
        offset: Annotated[
            int, Field(default=0, ge=0, description="For list: comments to skip.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="For list: most comments to return.")
        ] = 300,
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "set", "reply", "resolve", "remove"))
        path = excel_with_sheets(file_path, settings)
        if wanted == "list":
            with cells.open_workbook(path) as book:
                target = cells.sheet_named(book, sheet)
                found = _comments(target, cell.strip())
                name = target.name
            shown, next_offset = page(found, "comments", offset, max_results)
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found),
                "offset": offset, "next_offset": next_offset, "comments": shown,
            }
            return result

        require_writable(settings, "xlide_manage_comment")
        if not cell.strip():
            raise ToolError(f"cell is required for action={wanted!r}.")
        if wanted in {"set", "reply"} and not text.strip():
            raise ToolError(f"text is required for action={wanted!r}.")
        wanted_kind = _one_of(kind, ("note", "thread"), "kind")

        from pyofficeeditor.exceptions import PyOfficeEditorError

        with cells.editing(path) as book:
            target = cells.sheet_named(book, sheet)
            reference = _plain_cell(cell)
            try:
                if wanted == "set" and wanted_kind == "note":
                    target.set_comment(reference, text, author=author.strip(), visible=visible)
                elif wanted == "set":
                    target.add_threaded_comment(reference, text, author=author.strip() or "Agent")
                elif wanted == "reply":
                    target.add_threaded_reply(reference, text, author=author.strip() or "Agent")
                elif wanted == "resolve":
                    target.resolve_threaded_comment(reference, resolved)
                elif not target.remove_comment(reference):
                    raise ToolError(f"{target.name}!{reference} has no note or comment to remove.")
            except ToolError:
                raise
            except (PyOfficeEditorError, KeyError, ValueError) as exc:
                raise ToolError(f"{wanted} failed on {target.name}!{reference}: {exc}") from exc
            now = _comments(target, reference)
            name = target.name
        return {
            "path": str(path),
            "sheet": name,
            "action": wanted,
            "cell": reference,
            "saved": True,
            "comments": now,
        }

    # ------------------------------------------------------------- autofilters
    @server.tool(
        name="xlide_manage_filter",
        title="Autofilters",
        annotations=writes("Autofilters", destructive=True),
        description=(
            "Lists, sets, reapplies or clears the autofilter on a worksheet or on an Excel "
            "table, and hides the rows it filters out. Excel does not apply a filter when it "
            "opens a workbook: it shows the rows as the file marks them, so the rows are "
            "worked out here, held to what Excel keeps. action='list' reads the filters and "
            "changes nothing. 'set' filters one column and keeps what the others already "
            "filter by; call it once per column. Give criteria as VBA's Range.AutoFilter "
            "spells Criteria1 and Criteria2: '>=10', '=North', '=*east*', '=' for blanks, "
            "'<>' for anything; or top for the top N items (negative for the bottom), or "
            "dynamic for aboveAverage, thisMonth, Q1 and the like. 'reapply' filters again "
            "after the data changed; 'clear' removes the filter and shows every row."
        ),
    )
    def manage_filter(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        action: Annotated[
            str, Field(default="list", description="list, set, reapply or clear.")
        ] = "list",
        cell_range: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For set on a sheet: the block with its header row, such as A1:F200. "
                    "Empty keeps the range of the filter already there."
                ),
            ),
        ] = "",
        table: Annotated[
            str, Field(default="", description="Filter this Excel table rather than a range.")
        ] = "",
        column: Annotated[
            str,
            Field(
                default="",
                description="For set: the column, by header text or by number from 1 in the range.",
            ),
        ] = "",
        criteria1: Annotated[
            str, Field(default="", description="For set: Criteria1, such as '>=10' or '=West'.")
        ] = "",
        criteria2: Annotated[
            str, Field(default="", description="For set: Criteria2, such as '<100'.")
        ] = "",
        match_all: Annotated[
            bool,
            Field(
                default=True,
                description="With two criteria: true keeps rows meeting both, false either.",
            ),
        ] = True,
        top: Annotated[
            int,
            Field(
                default=0,
                description="For set: keep the top N items, or with a negative N the bottom N.",
            ),
        ] = 0,
        percent: Annotated[
            bool, Field(default=False, description="top counts percent rather than items.")
        ] = False,
        dynamic: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For set: aboveAverage, belowAverage, today, yesterday, tomorrow, thisWeek, "
                    "lastWeek, nextWeek, thisMonth, lastMonth, nextMonth, thisQuarter, "
                    "lastQuarter, nextQuarter, thisYear, lastYear, nextYear, yearToDate, Q1 to "
                    "Q4 or M1 to M12."
                ),
            ),
        ] = "",
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "set", "reapply", "clear"))
        path = excel_with_sheets(file_path, settings)
        if wanted == "list":
            with cells.open_workbook(path) as book:
                target = cells.sheet_named(book, sheet)
                found = _filters(target)
                name = target.name
            return {"path": str(path), "sheet": name, "filters": found}

        require_writable(settings, "xlide_manage_filter")
        from pyofficeeditor.exceptions import PyOfficeEditorError

        table_name = table.strip()
        with cells.editing(path) as book:
            target = cells.sheet_named(book, sheet)
            owner = _table_named(target, table_name) if table_name else None
            try:
                if wanted == "clear":
                    if owner is not None:
                        target.clear_table_filter(owner.name)
                    else:
                        target.clear_auto_filter()
                    outcome = None
                elif wanted == "reapply":
                    outcome = (
                        target.apply_table_filter(owner.name)
                        if owner is not None
                        else target.apply_auto_filter()
                    )
                else:
                    outcome = _set_filter(
                        target, owner, cell_range.strip(), column, criteria1, criteria2,
                        match_all, top, percent, dynamic,
                    )
            except ToolError:
                raise
            except (PyOfficeEditorError, KeyError, ValueError) as exc:
                raise ToolError(f"The filter could not be {wanted}: {exc}") from exc
            described = _filters(target)
            name = target.name
        result: dict[str, Any] = {
            "path": str(path),
            "sheet": name,
            "action": wanted,
            "saved": True,
            "filters": described,
        }
        if outcome is not None:
            result["rows"] = _outcome(outcome)
        return result


# ------------------------------------------------------------------ the parts


def _set_filter(
    sheet: Any,
    owner: Any,
    cell_range: str,
    column: str,
    criteria1: str,
    criteria2: str,
    match_all: bool,
    top: int,
    percent: bool,
    dynamic: str,
) -> Any:
    """Filter one column, keeping the criteria already on the others."""
    import pyofficeeditor.excel as excel

    chosen = [bool(criteria1.strip()), bool(top), bool(dynamic.strip())]
    if sum(chosen) != 1:
        raise ToolError("Give exactly one of criteria1, top or dynamic.")
    if criteria2.strip() and not criteria1.strip():
        raise ToolError("criteria2 goes with criteria1.")
    if not column.strip():
        raise ToolError("column is required for action='set'.")

    if owner is not None:
        existing = owner.auto_filter
        headers = [entry.name for entry in owner.columns]
        block = None
    else:
        existing = sheet.auto_filter
        reference = cell_range or (existing.ref if existing is not None else "")
        if not reference:
            raise ToolError(
                "This sheet has no filter yet, so give cell_range: the block with its header "
                "row, such as A1:F200."
            )
        block = excel.RangeRef.parse(str(reference)).normalized
        if existing is not None:
            current = excel.RangeRef.parse(existing.ref).normalized
            if str(current) != str(block):
                existing = None  # a filter over another block is replaced, not merged
        headers = [
            str(sheet.get_value(excel.CellRef(block.top, index)) or "")
            for index in range(block.left, block.right + 1)
        ]

    offset = _column_offset(column.strip(), headers)
    if criteria1.strip():
        criterion = excel.criteria(
            criteria1.strip(), criteria2.strip() or None, require_all=match_all
        )
    elif top:
        limit = 100 if percent else 500
        if not 1 <= abs(top) <= limit:
            raise ToolError(
                f"top must be between 1 and {limit}, or -1 and -{limit} for the bottom."
            )
        criterion = excel.Top10Filter(count=abs(top), percent=percent, top=top > 0, threshold=None)
    else:
        kind = dynamic.strip()
        from typing import get_args

        from pyofficeeditor.excel._filters import DynamicType

        allowed = get_args(DynamicType)
        match = next((name for name in allowed if name.casefold() == kind.casefold()), None)
        if match is None:
            raise ToolError(f"dynamic must be one of: {', '.join(allowed)}. Got {dynamic!r}.")
        criterion = excel.DynamicFilter(kind=match, value=None, maximum=None)

    kept = [c for c in (existing.columns if existing is not None else ()) if c.column != offset]
    columns = [*kept, excel.FilterColumn(column=offset, criterion=criterion, hide_dropdown=False)]
    columns.sort(key=lambda entry: entry.column)
    if owner is not None:
        return sheet.set_table_filter(owner.name, columns)
    return sheet.set_auto_filter(block, columns)


def _column_offset(column: str, headers: list[str]) -> int:
    """A filter column's offset from the left edge, from a number or a header."""
    if column.isdigit():
        number = int(column)
        if not 1 <= number <= len(headers):
            raise ToolError(f"column {number} is outside the filter's {len(headers)} columns.")
        return number - 1
    for index, header in enumerate(headers):
        if header.strip().casefold() == column.casefold():
            return index
    listed = ", ".join(repr(h) for h in headers if h) or "(no headers)"
    raise ToolError(f"No column headed {column!r}. The headers are: {listed}.")


def _table_named(sheet: Any, name: str) -> Any:
    for table in sheet.tables:
        if table.name.casefold() == name.casefold():
            return table
    listed = ", ".join(table.name for table in sheet.tables) or "(none)"
    raise ToolError(f"No table named {name!r} on {sheet.name}. Tables there: {listed}.")


def _filters(sheet: Any) -> list[dict[str, Any]]:
    """The sheet's filter and each table's, with what every column filters by."""
    found: list[dict[str, Any]] = []
    sheet_filter = sheet.auto_filter
    if sheet_filter is not None:
        found.append(_filter_summary(sheet_filter, None))
    for table in sheet.tables:
        table_filter = table.auto_filter
        if table_filter is not None and table_filter.columns:
            found.append(_filter_summary(table_filter, table))
    return found


def _filter_summary(auto_filter: Any, table: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {"range": str(auto_filter.ref)}
    if table is not None:
        entry["table"] = table.name
    entry["columns"] = [
        {"column": column.column + 1, **_criterion(column.criterion)}
        for column in auto_filter.columns
        if column.criterion is not None
    ]
    return entry


def _criterion(criterion: Any) -> dict[str, Any]:
    """What a column filters by, said the way Range.AutoFilter would take it."""
    import pyofficeeditor.excel as excel

    symbols = {
        "equal": "=", "notEqual": "<>", "greaterThan": ">", "greaterThanOrEqual": ">=",
        "lessThan": "<", "lessThanOrEqual": "<=",
    }
    if isinstance(criterion, excel.ValueFilter):
        entry: dict[str, Any] = {"values": list(criterion.values)}
        if criterion.blank:
            entry["blanks"] = True
        if criterion.date_groups:
            entry["date_groups"] = len(criterion.date_groups)
        return entry
    if isinstance(criterion, excel.CustomFilter):
        return {
            "criteria": [
                f"{symbols.get(item.operator, item.operator)}{item.value}"
                for item in criterion.comparisons
            ],
            "match_all": bool(criterion.require_all),
        }
    if isinstance(criterion, excel.Top10Filter):
        return {
            "top" if criterion.top else "bottom": criterion.count,
            "percent": bool(criterion.percent),
        }
    if isinstance(criterion, excel.DynamicFilter):
        return {"dynamic": criterion.kind}
    return {
        "kept_as_is": (
            "a colour, icon or other filter this server cannot evaluate; it is kept, and the "
            "rows it hides are left as they are"
        )
    }


def _outcome(outcome: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {"hidden": len(outcome.hidden), "shown": len(outcome.shown)}
    if outcome.undecided:
        entry["left_as_they_were"] = len(outcome.undecided)
    if outcome.unevaluated:
        entry["not_evaluated"] = len(outcome.unevaluated)
    return entry


def _plain_cell(cell: str) -> str:
    """A cell as a comment is keyed: $B$4 and b4 are both B4."""
    return cell.strip().upper().replace("$", "")


def _comments(sheet: Any, cell: str) -> list[dict[str, Any]]:
    """The notes and threads on a sheet, or on one cell, as plain JSON."""
    wanted = _plain_cell(cell)
    found: list[dict[str, Any]] = []
    for note in sheet.comments:
        where = str(getattr(note.ref, "a1", note.ref))
        if wanted and where != wanted:
            continue
        found.append(
            {
                "cell": where,
                "kind": "note",
                "text": note.text,
                "author": note.author,
                "visible": bool(note.visible),
            }
        )
    for thread in sheet.threaded_comments:
        where = str(getattr(thread.ref, "a1", thread.ref))
        if wanted and where != wanted:
            continue
        found.append(
            {
                "cell": where,
                "kind": "thread",
                "text": thread.text,
                "author": thread.author,
                "when": str(thread.when or ""),
                "resolved": bool(thread.resolved),
                "replies": [
                    {"text": reply.text, "author": reply.author, "when": str(reply.when or "")}
                    for reply in thread.replies
                ],
            }
        )
    # A thread keeps a placeholder note beside it for older Excel, which is not
    # a second comment on the cell.
    threads = {entry["cell"] for entry in found if entry["kind"] == "thread"}
    return [e for e in found if not (e["kind"] == "note" and e["cell"] in threads)]


def _action(raw: str, allowed: tuple[str, ...]) -> str:
    wanted = (raw or "").strip().lower()
    for candidate in allowed:
        if candidate.lower() == wanted:
            return candidate
    raise ToolError(f"action must be one of: {', '.join(allowed)}. Got {raw!r}.")


def _table_summary(table: Any, sheet_name: str) -> dict[str, Any]:
    return {
        "name": table.name,
        "sheet": sheet_name,
        "range": str(table.ref),
        "columns": list(table.column_names),
        "has_totals_row": bool(table.has_totals_row),
    }


def _do_table(
    book: Any, sheet: Any, action: str, name: str, reference: str, totals_row: bool
) -> dict[str, Any]:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        if action == "add":
            if sheet is None:
                raise ToolError("sheet is required for action='add'.")
            if not reference.strip():
                raise ToolError(
                    "cell_range is required for action='add', including the header row."
                )
            table = sheet.add_table(name, reference.strip(), totals_row=totals_row)
            return {
                "table": table.name,
                "sheet": sheet.name,
                "range": str(table.ref),
                "columns": list(table.column_names),
            }
        owner = sheet
        if owner is None:
            owner = next(
                (s for s in book.sheets if any(t.name == name for t in s.tables)), None
            )
            if owner is None:
                listed = ", ".join(book.table_names) or "(none)"
                raise ToolError(f"No table named {name!r}. Tables: {listed}.")
        owner.remove_table(name)
        return {"removed": name, "sheet": owner.name}
    except ToolError:
        raise
    except (PyOfficeEditorError, KeyError, ValueError) as exc:
        raise ToolError(f"{action} failed for table {name!r}: {exc}") from exc


def _validation_summary(rule: Any) -> dict[str, Any]:
    return {
        "ranges": [str(r) for r in getattr(rule, "ranges", []) or []],
        "kind": str(getattr(rule, "kind", "") or ""),
        "operator": str(getattr(rule, "operator", "") or ""),
        "formula1": str(getattr(rule, "formula1", "") or ""),
        "formula2": str(getattr(rule, "formula2", "") or ""),
        "allow_blank": bool(getattr(rule, "allow_blank", False)),
    }


def _add_validation(
    sheet: Any,
    reference: str,
    kind: str,
    formula1: str,
    formula2: str,
    operator: str,
    allow_blank: bool,
    error_message: str,
) -> dict[str, Any]:
    from pyofficeeditor.excel import DataValidation
    from pyofficeeditor.exceptions import PyOfficeEditorError

    if not reference.strip():
        raise ToolError("cell_range is required for action='add'.")
    wanted = _one_of(kind, VALIDATION_KINDS, "kind")
    if not formula1.strip():
        raise ToolError("formula1 is required: it is what the cell is allowed to hold.")

    first = formula1.strip()
    if wanted == "list" and not _looks_like_a_reference(first):
        # Excel spells an inline list as one quoted, comma-separated string.
        first = '"' + first.replace('"', "") + '"'

    rule = DataValidation(
        kind=wanted,
        operator=_one_of(operator, OPERATORS, "operator"),
        formula1=first,
        formula2=formula2.strip() or None,
        allow_blank=allow_blank,
        error_message=error_message.strip() or None,
        show_error=bool(error_message.strip()),
    )
    try:
        sheet.add_data_validation(reference.strip(), rule)
    except (PyOfficeEditorError, ValueError) as exc:
        raise ToolError(f"Validation was refused for {reference}: {exc}") from exc
    return {"range": reference.strip(), "kind": wanted, "formula1": first}


def _add_conditional(
    sheet: Any,
    reference: str,
    rule: str,
    operator: str,
    value: str,
    other_value: str,
    formula: str,
    fill_color: str,
    font_color: str,
    bold: bool | None,
) -> dict[str, Any]:
    from pyofficeeditor.excel import Dxf, bar, cell_is, expression, gradient
    from pyofficeeditor.exceptions import PyOfficeEditorError

    if not reference.strip():
        raise ToolError("cell_range is required for action='add'.")
    wanted = _one_of(rule, ("cell_is", "expression", "color_scale", "data_bar"), "rule")

    if wanted == "cell_is":
        if not value.strip():
            raise ToolError("value is required for rule='cell_is'.")
        built = cell_is(
            _one_of(operator, OPERATORS, "operator"),
            _number_or_text(value),
            _number_or_text(other_value) if other_value.strip() else None,
        )
    elif wanted == "expression":
        if not formula.strip():
            raise ToolError("formula is required for rule='expression'.")
        built = expression(formula.strip())
    elif wanted == "color_scale":
        built = gradient()
    else:
        built = bar()

    paint = None
    if fill_color or font_color or bold is not None:
        paint = Dxf.of(
            fill=_color(fill_color, "fill_color") if fill_color else None,
            color=_color(font_color, "font_color") if font_color else None,
            bold=bold,
        )
    elif wanted in {"cell_is", "expression"}:
        raise ToolError(
            "A rule with no paint highlights nothing. Pass fill_color, font_color or bold."
        )

    try:
        sheet.add_conditional_format(reference.strip(), built, dxf=paint)
    except (PyOfficeEditorError, ValueError) as exc:
        raise ToolError(f"The rule was refused for {reference}: {exc}") from exc
    return {"range": reference.strip(), "rule": wanted}


def _looks_like_a_reference(text: str) -> bool:
    """Whether a list source is cells rather than typed-out values."""
    return text.startswith("=") or "!" in text or "$" in text or ":" in text


def _number_or_text(raw: str) -> Any:
    """A comparison value as the number it is, when it is one."""
    text = raw.strip()
    try:
        number = float(text)
    except ValueError:
        return text
    return int(number) if number.is_integer() else number


def _color(raw: str, argument: str) -> str:
    text = raw.strip().lstrip("#").upper()
    if len(text) not in {6, 8} or any(c not in "0123456789ABCDEF" for c in text):
        raise ToolError(
            f"{argument} must be RRGGBB or AARRGGBB hex, such as 'FF0000'. Got {raw!r}."
        )
    return text


def _one_of(raw: str, allowed: tuple[str, ...], argument: str) -> str:
    wanted = (raw or "").strip()
    for candidate in allowed:
        if candidate.casefold() == wanted.casefold():
            return candidate
    raise ToolError(f"{argument} must be one of: {', '.join(allowed)}. Got {raw!r}.")
