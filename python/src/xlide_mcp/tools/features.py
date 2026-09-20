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
from ._common import bound, read_only, writes
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
            shown, note = bound(found, "tables", "Ask for one sheet with sheet=.")
            result: dict[str, Any] = {"path": str(path), "count": len(found), "tables": shown}
            if note:
                result["note"] = note
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
            "formula say TaxRate rather than Config!$B$7, and VBA reads them too, so renaming "
            "or removing one can break code as well as formulas. action='list' changes "
            "nothing. refers_to is a formula, so it needs its sheet and its dollar signs: "
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
            shown, note = bound(found, "relationships", "")
            result: dict[str, Any] = {"path": str(path), "count": len(found), "names": shown}
            if note:
                result["note"] = note
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
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "clear"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                target = cells.sheet_named(book, sheet)
                found = [_validation_summary(rule) for rule in target.data_validations]
                name = target.name
            shown, note = bound(found, "relationships", "")
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found), "validations": shown,
            }
            if note:
                result["note"] = note
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
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "clear"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                target = cells.sheet_named(book, sheet)
                found = [
                    {
                        "range": str(getattr(entry, "reference", "") or ""),
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
            shown, note = bound(found, "relationships", "")
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found), "formats": shown,
            }
            if note:
                result["note"] = note
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
    ) -> dict[str, Any]:
        wanted = _action(action, ("list", "add", "remove"))
        path = excel_with_sheets(file_path, settings)

        if wanted == "list":
            with cells.open_workbook(path) as book:
                sheet_object = cells.sheet_named(book, sheet)
                found = [
                    {
                        "range": str(getattr(link, "reference", "") or ""),
                        "target": getattr(link, "target", "") or "",
                        "location": getattr(link, "location", "") or "",
                        "display": getattr(link, "display", "") or "",
                    }
                    for link in sheet_object.hyperlinks
                ]
                name = sheet_object.name
            shown, note = bound(found, "relationships", "")
            result: dict[str, Any] = {
                "path": str(path), "sheet": name, "count": len(found), "hyperlinks": shown,
            }
            if note:
                result["note"] = note
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


# ------------------------------------------------------------------ the parts


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
