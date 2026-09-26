"""How a range looks: fonts, fills, borders, alignment, number format, merging.

The arguments are flat scalars rather than nested objects on purpose. A calling
model fills in `bold=True, fill_color="FFFF00"` reliably; it fills in a nested
`{"font": {"bold": true}, "fill": {"pattern": "solid", ...}}` less reliably, and
every argument it gets wrong is a round trip the user waits through.

Only what is passed is changed. Each cell keeps the rest of its own format, so
making a header row bold does not flatten the number formats underneath it.
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

# What Excel's own UI offers, which is also what pyOfficeEditor accepts.
BORDER_STYLES = (
    "thin", "medium", "thick", "double", "dotted", "dashed", "hair",
    "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
    "mediumDashDotDot", "slantDashDot", "none",
)
HORIZONTAL = (
    "general", "left", "center", "right", "fill", "justify",
    "centerContinuous", "distributed",
)
VERTICAL = ("top", "center", "bottom", "justify", "distributed")


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_format_cells",
        title="Format cells",
        annotations=writes("Format cells", destructive=True),
        description=(
            "Changes how a range looks: bold and font, fill colour, borders, alignment, "
            "number format, and merging. Only what you pass is changed, so each cell keeps "
            "the rest of its own format and making a header row bold does not flatten the "
            "number formats under it. Colours are RRGGBB or AARRGGBB hex without a leading "
            "hash. A number format is an Excel format code such as '#,##0.00' or 'yyyy-mm-dd'; "
            "that code is also what decides whether a number is shown as a date. Merging "
            "clears values and formulas outside the top-left cell; allow_overwrite is required "
            "when that would discard content. Works on "
            ".xlsx, .xlsm and .xlam."
        ),
    )
    def format_cells(
        file_path: Annotated[str, Field(description="Absolute path to the Excel file.")],
        sheet: Annotated[str, Field(description="Worksheet name, matched without case.")],
        cell_range: Annotated[
            str, Field(description="A1-style range, such as A1:D1, or a single cell.")
        ],
        bold: Annotated[bool | None, Field(default=None, description="Bold on or off.")] = None,
        italic: Annotated[
            bool | None, Field(default=None, description="Italic on or off.")
        ] = None,
        underline: Annotated[
            bool | None, Field(default=None, description="Single underline on or off.")
        ] = None,
        strike: Annotated[
            bool | None, Field(default=None, description="Strikethrough on or off.")
        ] = None,
        font_name: Annotated[
            str, Field(default="", description="Typeface, such as 'Calibri'.")
        ] = "",
        font_size: Annotated[
            float, Field(default=0, ge=0, description="Points. 0 leaves it alone.")
        ] = 0,
        font_color: Annotated[
            str, Field(default="", description="Text colour as hex, such as 'FF0000'.")
        ] = "",
        fill_color: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Solid background colour as hex. 'none' clears the fill back to no fill."
                ),
            ),
        ] = "",
        border_style: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Border on all four sides: thin, medium, thick, double, dotted, dashed, "
                    "hair, or none to remove."
                ),
            ),
        ] = "",
        border_color: Annotated[
            str, Field(default="", description="Border colour as hex. Defaults to automatic.")
        ] = "",
        horizontal: Annotated[
            str,
            Field(default="", description="left, center, right, fill, justify or general."),
        ] = "",
        vertical: Annotated[
            str, Field(default="", description="top, center, bottom, justify or distributed.")
        ] = "",
        wrap_text: Annotated[
            bool | None, Field(default=None, description="Wrap text in the cell.")
        ] = None,
        number_format: Annotated[
            str,
            Field(
                default="",
                description="Excel format code, such as '#,##0.00', '0%' or 'yyyy-mm-dd'.",
            ),
        ] = "",
        merge: Annotated[
            str,
            Field(
                default="",
                description=(
                    "'merge' joins the range into one cell, keeping only the top-left value. "
                    "'unmerge' splits it back. Empty leaves merging alone."
                ),
            ),
        ] = "",
        allow_overwrite: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "For merge: allow clearing values or formulas in cells other than the "
                    "top-left cell. Ask the user first."
                ),
            ),
        ] = False,
        style: Annotated[
            str,
            Field(
                default="",
                description=(
                    "A named cell style, applied first: one of Excel's own, such as Good, Bad, "
                    "Neutral, Title, Heading 1, Total, Input, Output, Note or 20% - Accent1, or "
                    "one the workbook defines. It sets only the parts the style includes."
                ),
            ),
        ] = "",
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_format_cells")
        path = excel_with_sheets(file_path, settings)

        font = _font_changes(bold, italic, underline, strike, font_name, font_size, font_color)
        alignment = _alignment_changes(horizontal, vertical, wrap_text)
        wanted_merge = (merge or "").strip().lower()
        if wanted_merge not in {"", "merge", "unmerge"}:
            raise ToolError("merge must be 'merge', 'unmerge' or empty.")
        if not any(
            [font, alignment, fill_color, border_style, number_format, wanted_merge, style.strip()]
        ):
            raise ToolError(
                "Nothing to change. Pass at least one of style, bold, italic, underline, "
                "strike, font_name, font_size, font_color, fill_color, border_style, "
                "horizontal, vertical, wrap_text, number_format or merge."
            )

        applied: list[str] = []
        merge_loss: dict[str, int] = {}
        with cells.editing(path) as book:
            sheet_object = cells.sheet_named(book, sheet)
            area = _area(sheet_object, cell_range)
            if wanted_merge == "merge":
                merge_loss = _merge_loss(area)
                if merge_loss["cells_cleared"] and not allow_overwrite:
                    formula_word = (
                        "formula" if merge_loss["formulas_cleared"] == 1 else "formulas"
                    )
                    raise ToolError(
                        f"Merging {area.a1} would clear {merge_loss['cells_cleared']} cells "
                        f"outside the top-left cell, including "
                        f"{merge_loss['formulas_cleared']} {formula_word}. Nothing was saved. "
                        "Read the range, ask the user, then call again with "
                        "allow_overwrite=true."
                    )

            # First, so what else this call sets lands on top of the style, as it
            # would if the user picked the style and then made the text bold.
            if style.strip():
                _apply_style(area, style.strip())
                applied.append(f"style {style.strip()}")
            if font:
                area.apply_font(**font)
                applied.append("font")
            if fill_color:
                _apply_fill(area, fill_color)
                applied.append("fill")
            if border_style:
                _apply_border(area, border_style, border_color)
                applied.append("border")
            if alignment:
                area.apply_alignment(**alignment)
                applied.append("alignment")
            if number_format:
                area.apply_number_format(number_format)
                applied.append("number format")
            if wanted_merge:
                _apply_merge(sheet_object, area, wanted_merge)
                applied.append(wanted_merge)

            name = sheet_object.name
            reference = str(area.reference)

        return {
            "path": str(path),
            "sheet": name,
            "range": reference,
            "applied": applied,
            **merge_loss,
            "saved": True,
            "recalculated": False,
            "note": (
                "Formatting is stored. A number format decides how a stored number is shown, "
                "not what it is: the value under it does not change."
            ),
        }


# ------------------------------------------------------------------ the parts


def _merge_loss(area: Any) -> dict[str, int]:
    """Values and formulas pyOfficeEditor will clear outside the merge anchor."""
    anchor = area.reference.start
    occupied = formulas = 0
    for cell in area:
        if cell.row == anchor.row and cell.column == anchor.column:
            continue
        if cell.formula is not None:
            formulas += 1
            occupied += 1
        elif cell.value is not None:
            occupied += 1
    return {"cells_cleared": occupied, "formulas_cleared": formulas}


def _font_changes(
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    strike: bool | None,
    name: str,
    size: float,
    color: str,
) -> dict[str, Any]:
    changes: dict[str, Any] = {}
    if bold is not None:
        changes["bold"] = bold
    if italic is not None:
        changes["italic"] = italic
    if strike is not None:
        changes["strike"] = strike
    if underline is not None:
        # pyOfficeEditor takes Excel's own vocabulary here, where the absence of
        # an underline is spelled "none" rather than False.
        changes["underline"] = "single" if underline else "none"
    if name.strip():
        changes["name"] = name.strip()
    if size:
        changes["size"] = size
    if color.strip():
        changes["color"] = _color(color, "font_color")
    return changes


def _alignment_changes(horizontal: str, vertical: str, wrap: bool | None) -> dict[str, Any]:
    changes: dict[str, Any] = {}
    if horizontal.strip():
        changes["horizontal"] = _one_of(horizontal, HORIZONTAL, "horizontal")
    if vertical.strip():
        changes["vertical"] = _one_of(vertical, VERTICAL, "vertical")
    if wrap is not None:
        changes["wrap_text"] = wrap
    return changes


def _area(sheet: Any, reference: str) -> Any:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        return sheet.range(reference)
    except (PyOfficeEditorError, ValueError) as exc:
        raise ToolError(f"{reference!r} is not an A1-style range: {exc}") from exc


def _apply_fill(area: Any, color: str) -> None:
    from pyofficeeditor.excel import Fill

    if color.strip().lower() == "none":
        area.apply_fill(Fill(pattern="none"))
        return
    area.apply_fill(_color(color, "fill_color"))


def _apply_border(area: Any, style: str, color: str) -> None:
    from pyofficeeditor.excel import Border

    wanted = _one_of(style, BORDER_STYLES, "border_style")
    if wanted == "none":
        area.apply_border(Border())
        return
    area.apply_border(Border.all_sides(wanted, _color(color, "border_color") if color else None))


def _apply_style(area: Any, name: str) -> None:
    """A named style, Excel's own defined in the workbook the first time it is used."""
    from pyofficeeditor.exceptions import PyOfficeEditorError

    try:
        area.apply_style(name)
    except (PyOfficeEditorError, KeyError, ValueError) as exc:
        raise ToolError(
            f"{name!r} is not a cell style this workbook has or Excel defines: {exc}. Excel's "
            "own are named as its Cell Styles gallery shows them, such as Good, Heading 1 or "
            "40% - Accent2."
        ) from exc


def _apply_merge(sheet: Any, area: Any, action: str) -> None:
    from pyofficeeditor.exceptions import PyOfficeEditorError

    reference = area.reference
    try:
        if action == "merge":
            sheet.merge(reference)
        else:
            sheet.unmerge(reference)
    except PyOfficeEditorError as exc:
        raise ToolError(f"{action} failed on {reference}: {exc}") from exc


def _color(raw: str, argument: str) -> str:
    """A hex colour, with the leading hash a caller often sends taken off."""
    text = raw.strip().lstrip("#").upper()
    if len(text) not in {6, 8} or any(c not in "0123456789ABCDEF" for c in text):
        raise ToolError(
            f"{argument} must be RRGGBB or AARRGGBB hex, such as 'FF0000'. Got {raw!r}."
        )
    return text


def _one_of(raw: str, allowed: tuple[str, ...], argument: str) -> str:
    wanted = raw.strip()
    for candidate in allowed:
        if candidate.casefold() == wanted.casefold():
            return candidate
    raise ToolError(f"{argument} must be one of: {', '.join(allowed)}. Got {raw!r}.")
