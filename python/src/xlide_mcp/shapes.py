"""The drawing layer of a worksheet: shapes, buttons, and the macros they run.

An agent asked why a workbook does something has to be able to see the button
that starts it. A very common shape of real automation is a Forms-toolbar button
on a sheet whose OnAction names a Sub, and nothing else in this server surfaces
that link: the module is readable, the button is not, and the connection between
them is invisible.

Excel keeps shapes in two eras, and both are read here.

* Modern shapes live in the sheet's drawing part, `xl/drawings/drawingN.xml`,
  each inside a cell anchor that positions it. A shape's OnAction is the `macro`
  attribute on its own element.
* Form controls are older. Excel keeps each in up to four places that have to
  agree: a VML shape in `xl/drawings/vmlDrawingN.vml` (which cell notes also live
  in), an entry in the sheet's `<controls>`, a `ctrlProp` part, and a hidden
  DrawingML twin. Excel 2007 writes no `<controls>` entry at all, so the VML is
  the only place a control is guaranteed to appear, and it is read last so a
  richer entry from the other parts wins.

The layout was taken from XLIDE's `xlsxShapes.ts`, which measured it against
files Excel 16 saved.

Writing is pyOfficeEditor's. Since 0.3 it adds, removes and repoints shapes,
and keeps a form control's four parts in agreement while it does, which is the
format knowledge this server should never have held. The reader stays for now,
because pyOfficeEditor's shapes do not yet carry four things list_shapes has
always answered and the conformance corpus pins: the cells an anchor covers,
alt text, the hidden flag, and ActiveX controls. What pyOfficeEditor reads that
this does not - a control's state and every shape's position in points - is
merged in by name. When its shapes carry the four, this reader and xlsx.py go
together (WilliamSmithEdward/pyOfficeEditor#4).

One write stays here with the reader: the macro on a control Excel 2007 saved,
which lives only in the VML, where pyOfficeEditor does not look.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .errors import ToolError
from .xlsx import (
    Workbook,
    XlsxError,
    decode_xml,
    encode_attr,
    encode_xml,
    index_to_column,
    next_tag,
)

# The OnAction Excel stores is qualified to the workbook; the formula bar is not.
_WORKBOOK_PREFIX = re.compile(r"^\[\d+\]!")

_RELATIONSHIP_DRAWING = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_RELATIONSHIP_VML = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
)

# A form control's objectType, as the ctrlProp part and the VML both spell it.
_CONTROL_KINDS = {
    "button": "button",
    "checkbox": "checkBox",
    "radio": "optionButton",
    "drop": "dropDown",
    "list": "listBox",
    "scroll": "scrollBar",
    "spin": "spinner",
    "label": "label",
    "gbox": "groupBox",
    "edit": "editBox",
}

_DRAWING_ELEMENTS = frozenset(
    {"xdr:sp", "xdr:grpSp", "xdr:graphicFrame", "xdr:cxnSp", "xdr:pic", "xdr:contentPart"}
)


@dataclass
class Shape:
    name: str
    kind: str
    cells: str = ""
    macro: str = ""
    text: str = ""
    linked_cell: str = ""
    input_range: str = ""
    alt_text: str = ""
    hidden: bool = False
    children: list[Shape] = field(default_factory=list)

    def summary(self) -> dict[str, Any]:
        entry: dict[str, Any] = {"name": self.name, "kind": self.kind}
        for key, value in (
            ("cells", self.cells),
            ("macro", self.macro),
            ("text", self.text),
            ("linked_cell", self.linked_cell),
            ("input_range", self.input_range),
            ("alt_text", self.alt_text),
        ):
            if value:
                entry[key] = value
        if self.hidden:
            entry["hidden"] = True
        if self.children:
            entry["shapes"] = [child.summary() for child in self.children]
        return entry


def read_sheet_shapes(path: Path, sheet_name: str | None = None) -> dict[str, list[Shape]]:
    """Every shape on one sheet, or on all of them, keyed by sheet name."""
    book = Workbook(path)
    sheets = book.sheets()
    if sheet_name:
        wanted = book.canonical_sheet_name(sheet_name)
        sheets = [s for s in sheets if s.name == wanted]
    out: dict[str, list[Shape]] = {}
    for sheet in sheets:
        out[sheet.name] = _shapes_for(book, sheet.part)
    return out


def _shapes_for(book: Workbook, sheet_part: str) -> list[Shape]:
    sheet_xml = book.part_text(sheet_part)
    relationships = book.part_relationships(sheet_part)

    shapes: list[Shape] = []
    # The drawing part first: its shapes carry their own names and macros.
    drawing_part = _related_part(sheet_xml, relationships, "drawing", _RELATIONSHIP_DRAWING)
    if drawing_part:
        drawing_xml = book.optional_part_text(drawing_part)
        if drawing_xml is not None:
            shapes.extend(_read_drawing(drawing_xml))

    # Then the form controls, which can add to a shape the drawing already named.
    controls = _read_controls(book, sheet_part, sheet_xml, relationships)
    by_name = {shape.name.casefold(): shape for shape in shapes}
    for control in controls:
        existing = by_name.get(control.name.casefold())
        if existing is None:
            shapes.append(control)
            by_name[control.name.casefold()] = control
            continue
        # The DrawingML twin carries the name and the anchor; the control parts
        # carry what it does. Neither alone is the whole shape.
        existing.kind = control.kind
        existing.macro = existing.macro or control.macro
        existing.text = control.text or existing.text
        existing.linked_cell = existing.linked_cell or control.linked_cell
        existing.input_range = existing.input_range or control.input_range
        # Excel marks the twin hidden for every form control, visible or not, so
        # the twin's flag says nothing about the control. The VML style does.
        existing.hidden = control.hidden
    return shapes


# ------------------------------------------------------------------ drawings


def _read_drawing(xml: str) -> list[Shape]:
    shapes: list[Shape] = []
    for anchor_start, anchor_end in _elements(xml, {"xdr:twoCellAnchor", "xdr:oneCellAnchor",
                                                    "xdr:absoluteAnchor", "mc:AlternateContent"}):
        body = xml[anchor_start:anchor_end]
        cells = _anchor_range(body)
        # An anchor holds one shape. Anything deeper belongs to that shape.
        for element_name, start, end in _drawing_children(body, 0):
            shape = _drawing_shape(element_name, body[start:end], cells)
            if shape is not None:
                shapes.append(shape)
            break
    return shapes


def _drawing_shape(element_name: str, body: str, cells: str) -> Shape | None:
    properties = _first_tag(body, "xdr:cNvPr")
    if properties is None:
        return None
    name = properties.get("name", "")
    shape = Shape(name=name, kind=_drawing_kind(element_name, body), cells=cells)

    start = _first_tag(body, element_name)
    if start is not None:
        shape.macro = _display_macro(start.get("macro", ""))
    shape.alt_text = properties.get("descr", "")
    shape.hidden = properties.get("hidden") == "1"
    shape.text = _drawing_text(body)
    if element_name == "xdr:grpSp":
        # Walk from inside the group's own opening tag, so the group is not
        # rediscovered as its own first child.
        opening = next_tag(body, 0)
        after = opening.end if opening is not None else 0
        for child_name, child_start, child_end in _drawing_children(body, after):
            child = _drawing_shape(child_name, body[child_start:child_end], "")
            if child is not None:
                shape.children.append(child)
    return shape


def _drawing_kind(element_name: str, body: str) -> str:
    if element_name == "xdr:sp":
        tag = _first_tag(body, "xdr:cNvSpPr")
        return "textBox" if tag is not None and tag.get("txBox") == "1" else "shape"
    if element_name == "xdr:cxnSp":
        return "connector"
    if element_name == "xdr:pic":
        return "picture"
    if element_name == "xdr:grpSp":
        return "group"
    if element_name == "xdr:graphicFrame":
        return "chart" if 'drawingml/2006/chart"' in body else "other"
    return "other"


def _drawing_text(body: str) -> str:
    """A DrawingML text body, one line per paragraph.

    The runs are read out of the paragraph's own slice, and the offsets
    `_element_contents` hands back are relative to what it was given. Indexing the
    whole body with them lands a few characters into the middle of an attribute,
    which is how a shape reading "Go" came back as ":c".
    """
    lines: list[str] = []
    for paragraph_start, paragraph_end in _elements(body, {"a:p"}):
        paragraph = body[paragraph_start:paragraph_end]
        runs = [
            decode_xml(paragraph[start:end])
            for start, end in _element_contents(paragraph, "a:t")
        ]
        lines.append("".join(runs))
    return "\n".join(lines).strip()


def _anchor_range(body: str) -> str:
    """The cells an anchor covers, from its top-left cell to its bottom-right."""
    start = _marker(body, "xdr:from")
    if start is None:
        return ""
    end = _marker(body, "xdr:to")
    first = f"{index_to_column(start[0] + 1)}{start[1] + 1}"
    if end is None:
        return first
    # A shape ending exactly on a cell's top-left corner does not reach into it.
    column = end[0] - 1 if end[2] == 0 and end[0] > start[0] else end[0]
    row = end[1] - 1 if end[3] == 0 and end[1] > start[1] else end[1]
    last = f"{index_to_column(max(column, start[0]) + 1)}{max(row, start[1]) + 1}"
    return first if first == last else f"{first}:{last}"


_MARKER_FIELD = re.compile(r"<xdr:(col|row|colOff|rowOff)>(-?\d+)</xdr:\1>")


def _marker(body: str, name: str) -> tuple[int, int, int, int] | None:
    spans = list(_element_contents(body, name))
    if not spans:
        return None
    inner = body[spans[0][0] : spans[0][1]]
    values = {m.group(1): int(m.group(2)) for m in _MARKER_FIELD.finditer(inner)}
    return (
        values.get("col", 0),
        values.get("row", 0),
        values.get("colOff", 0),
        values.get("rowOff", 0),
    )


# ------------------------------------------------------------- form controls


def _read_controls(
    book: Workbook, sheet_part: str, sheet_xml: str, relationships: dict[str, dict[str, str]]
) -> list[Shape]:
    by_id: dict[int, Shape] = {}

    # The <controls> entries: name, macro, alt text, anchor, and the ctrlProp or
    # ActiveX part behind each.
    for start, end in _elements(sheet_xml, {"control"}):
        tag = _first_tag(sheet_xml[start:end], "control")
        if tag is None:
            continue
        shape_id = _int(tag.get("shapeId"))
        if shape_id in by_id:
            # An ActiveX control repeats its entry in an mc:Fallback, after the
            # one that counts.
            continue
        body = sheet_xml[start:end]
        relationship = relationships.get(tag.get("r:id", "") or tag.get("id", ""), {})
        target = relationship.get("part", "")
        is_active_x = bool(target) and not relationship.get("type", "").endswith("/ctrlProp")

        shape = Shape(
            name=tag.get("name", ""),
            kind="activeX" if is_active_x else "formControl",
            cells=_anchor_range(body),
        )
        control_properties = _first_tag(body, "controlPr")
        if control_properties is not None:
            shape.macro = _display_macro(control_properties.get("macro", ""))
            shape.alt_text = control_properties.get("altText", "")

        if target and not is_active_x:
            properties_xml = book.optional_part_text(target)
            if properties_xml is not None:
                properties = _first_tag(properties_xml, "formControlPr")
                if properties is not None:
                    shape.kind = _CONTROL_KINDS.get(
                        properties.get("objectType", "").lower(), shape.kind
                    )
                    shape.linked_cell = properties.get("fmlaLink", "")
                    shape.input_range = properties.get("fmlaRange", "")
        by_id[shape_id] = shape

    vml_part = _related_part(sheet_xml, relationships, "legacyDrawing", _RELATIONSHIP_VML)
    if vml_part:
        vml = book.optional_part_text(vml_part)
        if vml is not None:
            _merge_vml(vml, by_id)
    return [shape for shape in by_id.values() if shape.name or shape.macro]


_SPID = re.compile(r'\b(?:o:spid|id)="_x0000_s(\d+)"')
_OBJECT_TYPE = re.compile(r'<x:ClientData\b[^>]*\bObjectType="([^"]*)"')


def _merge_vml(vml: str, by_id: dict[int, Shape]) -> None:
    """Every form control has a VML shape; Excel 2007 keeps nothing else.

    Cell notes live here too and are not shapes, which is why the object type is
    checked before anything else is read.
    """
    for start, end in _elements(vml, {"v:shape"}):
        inner = vml[start:end]
        object_type_match = _OBJECT_TYPE.search(inner)
        if object_type_match is None or object_type_match.group(1) == "Note":
            continue
        object_type = object_type_match.group(1)
        head = inner[: inner.find(">") + 1]
        spid = _SPID.search(head)
        if spid is None:
            continue
        shape_id = int(spid.group(1))
        kind = _CONTROL_KINDS.get(object_type.lower(), "formControl")
        shape = by_id.get(shape_id)
        if shape is None:
            shape = Shape(name=f"{object_type} {shape_id % 1024}", kind=kind)
            by_id[shape_id] = shape
        if shape.kind == "formControl":
            shape.kind = kind
        shape.macro = shape.macro or _display_macro(_client_data(inner, "FmlaMacro"))
        shape.linked_cell = shape.linked_cell or _client_data(inner, "FmlaLink")
        shape.input_range = shape.input_range or _client_data(inner, "FmlaRange")
        if not shape.text and shape.kind not in {"dropDown", "listBox", "scrollBar", "spinner"}:
            shape.text = _vml_text(inner)
        if "visibility:hidden" in head.replace(" ", ""):
            shape.hidden = True
        if not shape.cells:
            shape.cells = _vml_range(inner)


def _client_data(inner: str, name: str) -> str:
    match = re.search(rf"<x:{name}>(.*?)</x:{name}>", inner, re.DOTALL)
    return decode_xml(match.group(1)).strip() if match else ""


def _vml_text(inner: str) -> str:
    match = re.search(r"<v:textbox\b.*?</v:textbox>", inner, re.DOTALL)
    if match is None:
        return ""
    stripped = re.sub(r"<[^>]+>", "", match.group(0))
    return decode_xml(stripped).strip()


def _vml_range(inner: str) -> str:
    """A VML control's anchor: eight numbers, the first and fifth being columns."""
    anchor = _client_data(inner, "Anchor")
    parts = [p.strip() for p in anchor.split(",") if p.strip()]
    if len(parts) < 8:
        return ""
    try:
        numbers = [int(p) for p in parts]
    except ValueError:
        return ""
    first = f"{index_to_column(numbers[0] + 1)}{numbers[2] + 1}"
    last = f"{index_to_column(numbers[4] + 1)}{numbers[6] + 1}"
    return first if first == last else f"{first}:{last}"


# ------------------------------------------------------------------- helpers


def _related_part(
    sheet_xml: str, relationships: dict[str, dict[str, str]], element: str, kind: str
) -> str:
    """The part a sheet's <drawing> or <legacyDrawing> element points at."""
    tag = _first_tag(sheet_xml, element)
    if tag is not None:
        identifier = tag.get("r:id") or tag.get("id") or ""
        related = relationships.get(identifier)
        if related:
            return related["part"]
    # Some writers omit the element but keep the relationship.
    for related in relationships.values():
        if related.get("type") == kind:
            return related["part"]
    return ""


def _first_tag(xml: str, name: str) -> dict[str, str] | None:
    position = 0
    while (tag := next_tag(xml, position)) is not None:
        position = tag.end
        if tag.name == name:
            return tag.attrs
    return None


def _elements(xml: str, names: set[str]):
    """Spans of every top-level element with one of these names, non-overlapping."""
    position = 0
    while (tag := next_tag(xml, position)) is not None:
        position = tag.end
        if tag.name not in names or tag.closing:
            continue
        if tag.self_closing:
            yield tag.start, tag.end
            continue
        end = _matching_end(xml, tag.name, tag.end)
        if end < 0:
            break
        yield tag.start, end
        position = end


def _drawing_children(body: str, start: int):
    """Drawing elements at one level, from `start`.

    Each match advances past its own closing tag, so what comes back is siblings:
    a shape inside a group is never yielded beside the group that holds it.
    """
    position = start
    while (tag := next_tag(body, position)) is not None:
        position = tag.end
        if tag.name not in _DRAWING_ELEMENTS or tag.closing:
            continue
        end = tag.end if tag.self_closing else _matching_end(body, tag.name, tag.end)
        if end < 0:
            return
        yield tag.name, tag.start, end
        position = end


def _element_contents(xml: str, name: str):
    """The inner span of every element with this name."""
    position = 0
    while (tag := next_tag(xml, position)) is not None:
        position = tag.end
        if tag.name != name or tag.closing:
            continue
        if tag.self_closing:
            yield tag.end, tag.end
            continue
        close = xml.find(f"</{name}>", tag.end)
        if close < 0:
            break
        yield tag.end, close
        position = close


def _matching_end(xml: str, name: str, after: int) -> int:
    """The offset past this element's closing tag, counting nested same-name ones."""
    depth = 1
    position = after
    while (tag := next_tag(xml, position)) is not None:
        position = tag.end
        if tag.name == name and not tag.self_closing:
            depth += 1
        elif tag.name == f"/{name}":
            depth -= 1
            if depth == 0:
                return tag.end
    return -1


def _display_macro(stored: str) -> str:
    """An OnAction as the formula bar shows it, not as the file stores it."""
    return _WORKBOOK_PREFIX.sub("", stored or "").strip()


def _int(value: str | None) -> int:
    try:
        return int(value or 0)
    except ValueError:
        return 0


# --------------------------------------------------------------------- writing
#
# Every write goes through pyOfficeEditor, which keeps a form control's four
# parts in agreement - the drawing, the sheet's <controls> entry, the ctrlProp
# part and the VML shape - and takes all four away together. This module used to
# write the macro link itself, part by part, and still does for the one kind of
# control pyOfficeEditor does not reach, at the end of this file.

# The kinds list_shapes reports, as pyOfficeEditor's add calls spell them.
CONTROL_KINDS = {
    "button": "Button",
    "checkBox": "CheckBox",
    "optionButton": "Radio",
    "dropDown": "Drop",
    "listBox": "List",
    "scrollBar": "Scroll",
    "spinner": "Spin",
    "label": "Label",
    "groupBox": "GBox",
}
DRAWING_KINDS = frozenset({"shape", "textBox", "line"})
ADDABLE_KINDS = (*CONTROL_KINDS, *sorted(DRAWING_KINDS), "picture")

# A size to start from, in points, when none is given. Nothing about these is
# measured: they are big enough to see and click, and the user resizes from there.
DEFAULT_SIZE = {
    "button": (72.0, 24.0),
    "checkBox": (96.0, 18.0),
    "optionButton": (96.0, 18.0),
    "dropDown": (96.0, 18.0),
    "listBox": (96.0, 72.0),
    "scrollBar": (18.0, 72.0),
    "spinner": (18.0, 36.0),
    "label": (72.0, 18.0),
    "groupBox": (144.0, 96.0),
    "shape": (96.0, 48.0),
    "textBox": (144.0, 36.0),
    "line": (96.0, 0.0),
}

# Excel's answers for a check box or an option button.
_CHECK_STATES = {1: True, -4146: False, 2: "mixed"}


def set_shape_macro(
    path: Path, sheet_name: str, shape_name: str, macro: str
) -> dict[str, Any]:
    """Point an existing shape at a different macro, or at none.

    A form control stores its macro twice, in its VML shape and in the sheet's
    `<controls>` entry, and Excel reads both; pyOfficeEditor writes both. A
    workbook index already on the link, the `[1]` in `[1]!Module1.Go`, is kept.
    A control Excel 2007 saved lives only in the VML drawing, where pyOfficeEditor
    does not look, and its macro is written there directly, as it always was.
    """
    from pyofficeeditor.exceptions import PyOfficeEditorError

    from . import cells

    wanted = (macro or "").strip()
    try:
        with cells.editing(path) as book:
            sheet = cells.sheet_named(book, sheet_name)
            shape = _find_shape(sheet, shape_name)
            previous = _display_macro(shape.macro)
            try:
                sheet.set_shape_macro(shape.name, wanted)
            except (PyOfficeEditorError, KeyError, ValueError) as exc:
                raise ToolError(f"The macro on {shape.name!r} could not be set: {exc}") from exc
            written = (
                ["form control (sheet entry)", "form control (VML)"]
                if shape.control is not None
                else ["drawing"]
            )
            canonical = sheet.name
    except _NotOnSheet as missing:
        legacy = _set_legacy_macro(path, missing.sheet, shape_name, wanted)
        if legacy is None:
            raise
        return legacy
    return {
        "sheet": canonical,
        "shape": shape.name,
        "macro": wanted,
        "previous_macro": previous,
        "parts_written": written,
        "cleared": not wanted,
    }


def add_shape(
    path: Path,
    sheet_name: str,
    *,
    name: str,
    kind: str,
    cell: str = "",
    left: float | None = None,
    top: float | None = None,
    width: float = 0.0,
    height: float = 0.0,
    text: str = "",
    macro: str = "",
    linked_cell: str = "",
    list_range: str = "",
    geometry: str = "",
    image: Path | None = None,
) -> dict[str, Any]:
    """Put a form control, an AutoShape, a text box, a line or a picture on a sheet."""
    from pyofficeeditor.exceptions import PyOfficeEditorError

    from . import cells

    if kind not in ADDABLE_KINDS:
        raise ToolError(f"kind {kind!r} cannot be added. Use one of: {', '.join(ADDABLE_KINDS)}.")
    with cells.editing(path) as book:
        sheet = cells.sheet_named(book, sheet_name)
        clash = next(
            (s for s in _walk(sheet.shapes) if s.name.casefold() == name.casefold()), None
        )
        if clash is not None:
            raise ToolError(f"{sheet.name} already has a shape named {clash.name!r}.")
        if cell.strip():
            x, y = _cell_origin(sheet, cell)
        elif left is not None and top is not None:
            x, y = left, top
        else:
            raise ToolError("Give cell, the top-left cell, or both left and top in points.")
        default_width, default_height = DEFAULT_SIZE.get(kind, (96.0, 48.0))
        try:
            if kind in CONTROL_KINDS:
                added = sheet.add_form_control(
                    name,
                    left=x,
                    top=y,
                    width=width or default_width,
                    height=height or default_height,
                    kind=CONTROL_KINDS[kind],
                    text=text,
                    macro=macro.strip(),
                    linked_cell=linked_cell.strip(),
                    list_range=list_range.strip(),
                )
            elif kind in DRAWING_KINDS:
                added = sheet.add_shape(
                    name,
                    left=x,
                    top=y,
                    width=width or default_width,
                    height=height or default_height,
                    kind=kind,
                    geometry=geometry.strip(),
                    text=text,
                    macro=macro.strip(),
                )
            else:
                if image is None:
                    raise ToolError("A picture needs image_path: a PNG, JPEG or GIF file.")
                added = sheet.add_picture(
                    name,
                    image,
                    left=x,
                    top=y,
                    width=width or None,
                    height=height or None,
                    description=text,
                )
        except ToolError:
            raise
        except (PyOfficeEditorError, KeyError, ValueError, OSError) as exc:
            raise ToolError(f"The {kind} could not be added: {exc}") from exc
        canonical = sheet.name
    return {"sheet": canonical, "added": _position(added) | {"name": added.name, "kind": kind}}


def remove_shape(path: Path, sheet_name: str, shape_name: str) -> dict[str, Any]:
    """Take a shape off a sheet, with every part that belongs to it.

    Refused wherever pyOfficeEditor would take less than that, or cannot reach
    the shape: a chart, whose part it leaves behind (WilliamSmithEdward/
    pyOfficeEditor#3); a group holding a chart, a control or a picture, whose
    parts it leaves behind too (#5 there); a member of a group, which it does not
    reach; and a control kept only in VML, which it does not reach either.
    """
    from pyofficeeditor.exceptions import PyOfficeEditorError

    from . import cells

    try:
        with cells.editing(path) as book:
            sheet = cells.sheet_named(book, sheet_name)
            shape = _find_shape(sheet, shape_name)
            if shape.kind == "chart":
                raise ToolError(
                    "Removing a chart is not offered yet: pyOfficeEditor takes the chart's "
                    "anchor off the drawing and leaves the chart part behind. Ask the user to "
                    "delete it in Excel."
                )
            group = next(
                (
                    top
                    for top in sheet.shapes
                    if any(member.name == shape.name for member in _walk(top.children))
                ),
                None,
            )
            if group is not None:
                raise ToolError(
                    f"{shape.name!r} is inside the group {group.name!r}, and pyOfficeEditor "
                    "removes a shape that stands on its own, not a member of a group. Nothing "
                    "was removed. Ask the user to ungroup it in Excel first."
                )
            stranded = [
                member.name
                for member in _walk(shape.children)
                if member.kind == "chart" or member.control is not None or member.image
            ]
            if stranded:
                raise ToolError(
                    f"{shape.name!r} is a group holding {', '.join(map(repr, stranded))}, and "
                    "pyOfficeEditor takes a group off the drawing without the chart, control "
                    "and picture parts of what is in it, which would be left in the file. "
                    "Nothing was removed. Ask the user to delete it in Excel."
                )
            try:
                sheet.remove_shape(shape.name)
            except (PyOfficeEditorError, KeyError, ValueError) as exc:
                raise ToolError(f"{shape.name!r} could not be removed: {exc}") from exc
            canonical = sheet.name
    except _NotOnSheet as missing:
        listed = read_sheet_shapes(path, missing.sheet).get(missing.sheet, [])
        if any(entry.name == shape_name for entry in listed):
            raise ToolError(
                f"{shape_name!r} is a form control kept only in the sheet's VML drawing, as "
                "Excel 2007 saves one, and pyOfficeEditor removes a control through the "
                "drawing a later Excel writes. Nothing was removed. Ask the user to delete it "
                "in Excel."
            ) from missing
        raise
    return {
        "sheet": canonical,
        "removed": shape.name,
        "had_macro": _display_macro(shape.macro),
    }


def control_states(path: Path) -> dict[str, dict[str, dict[str, Any]]]:
    """Where each shape sits in points, and what each form control holds, by sheet.

    What pyOfficeEditor reads that the drawing-layer reader above does not: a
    check box's state, a list's selected item, a spinner's value and bounds. It
    is merged into list_shapes by name.
    """
    from . import cells

    states: dict[str, dict[str, dict[str, Any]]] = {}
    with cells.open_workbook(path) as book:
        for sheet in book.sheets:
            by_name: dict[str, dict[str, Any]] = {}
            for shape in _walk(sheet.shapes):
                entry: dict[str, Any] = {"position": _position(shape)}
                entry.update(_control_state(shape.control))
                by_name[shape.name.casefold()] = entry
            for chart in sheet.charts:
                by_name.setdefault(chart.name.casefold(), {})["chart"] = chart_summary(chart)
            states[sheet.name] = by_name
    return states


def chart_summary(chart: Any) -> dict[str, Any]:
    """A chart's type, title and series, each series as Excel's formula bar shows it."""
    entry: dict[str, Any] = {"type": list(chart.kinds)}
    if chart.title:
        entry["title"] = chart.title
    entry["series"] = [
        {
            key: value
            for key, value in (
                ("name", series.name),
                ("values", series.values),
                ("categories", series.categories),
                ("formula", _series_formula(series)),
            )
            if value
        }
        for series in chart.series
    ]
    return entry


def _series_formula(series: Any) -> str:
    try:
        return str(series.formula or "")
    except Exception:
        return ""


CHART_TYPES = ("column", "bar", "line", "lineMarkers", "pie", "doughnut", "scatter", "area")


def add_chart(
    path: Path,
    sheet_name: str,
    *,
    data: str,
    chart_type: str,
    cell: str = "",
    left: float | None = None,
    top: float | None = None,
    width: float = 0.0,
    height: float = 0.0,
    title: str = "",
    name: str = "",
    series_in: str = "",
) -> dict[str, Any]:
    """A chart of a block of cells, written as Excel's Insert Chart writes one."""
    from pyofficeeditor.exceptions import PyOfficeEditorError

    from . import cells

    kind = next((t for t in CHART_TYPES if t.casefold() == chart_type.strip().casefold()), None)
    if kind is None:
        raise ToolError(f"chart_type must be one of: {', '.join(CHART_TYPES)}.")
    arrangement = series_in.strip().lower() or None
    if arrangement not in {None, "columns", "rows"}:
        raise ToolError("series_in must be 'columns', 'rows' or empty.")
    with cells.editing(path) as book:
        sheet = cells.sheet_named(book, sheet_name)
        if cell.strip():
            x, y = _cell_origin(sheet, cell)
        elif left is not None and top is not None:
            x, y = left, top
        else:
            raise ToolError("Give cell, the top-left cell, or both left and top in points.")
        try:
            chart = sheet.add_chart(
                kind,
                data.strip(),
                left=x,
                top=y,
                width=width or 360.0,
                height=height or 216.0,
                title=title.strip() or None,
                name=name.strip() or None,
                series_in=arrangement,
            )
        except (PyOfficeEditorError, KeyError, ValueError) as exc:
            raise ToolError(f"The chart could not be added: {exc}") from exc
        canonical = sheet.name
    return {"sheet": canonical, "chart": {"name": chart.name, **chart_summary(chart)}}


def _control_state(control: Any) -> dict[str, Any]:
    if control is None:
        return {}
    kind = str(control.kind)
    state: dict[str, Any] = {}
    if kind in {"CheckBox", "Radio"}:
        state["checked"] = _CHECK_STATES.get(control.value, bool(control.value))
    elif kind in {"Drop", "List"}:
        # Excel counts the chosen item from 1, and 0 is nothing chosen.
        state["selected_item"] = control.value
    elif kind in {"Spin", "Scroll"}:
        state["value"] = control.value
        if control.minimum is not None:
            state["minimum"] = control.minimum
        if control.maximum is not None:
            state["maximum"] = control.maximum
    return state


def _position(shape: Any) -> dict[str, Any]:
    return {
        "left": round(float(shape.left), 2),
        "top": round(float(shape.top), 2),
        "width": round(float(shape.width), 2),
        "height": round(float(shape.height), 2),
    }


def _walk(shapes: Any) -> list[Any]:
    """Every shape, a group's members after the group."""
    out: list[Any] = []
    for shape in shapes:
        out.append(shape)
        out.extend(_walk(shape.children))
    return out


class _NotOnSheet(ToolError):
    """No shape by that name in the drawing pyOfficeEditor reads."""

    def __init__(self, message: str, sheet: str) -> None:
        super().__init__(message)
        self.sheet = sheet


def _find_shape(sheet: Any, name: str) -> Any:
    """A shape by its exact name, or by a name matched without case when that is unique."""
    everything = _walk(sheet.shapes)
    exact = [shape for shape in everything if shape.name == name]
    if exact:
        return exact[0]
    folded = [shape for shape in everything if shape.name.casefold() == name.casefold()]
    if len(folded) == 1:
        return folded[0]
    listed = ", ".join(shape.name for shape in everything) or "(none)"
    raise _NotOnSheet(
        f"No shape named {name!r} on {sheet.name}. Shapes there: {listed}.", sheet.name
    )


def _cell_origin(sheet: Any, reference: str) -> tuple[float, float]:
    """Where a cell's top-left corner is, in points, on pyOfficeEditor's own grid.

    The grid it turns a position back into an anchor with, a sheet's default
    column width included, so a shape placed at a cell lands on that cell. It is
    private until pyOfficeEditor takes a cell for a position
    (WilliamSmithEdward/pyOfficeEditor#4).
    """
    from pyofficeeditor.excel._shapes import SheetGrid

    from . import cells

    corner = cells.cell_reference(sheet, reference)
    grid = SheetGrid.of(sheet._root)
    return grid.x(corner.column - 1), grid.y(corner.row - 1)


# ------------------------------------------------ a control kept only in VML
#
# Excel 2007 writes a form control as a VML shape and nothing else, and later
# Excels keep reading it. pyOfficeEditor reaches a control through the DrawingML
# twin a later Excel writes beside the VML, so for these the macro is written
# where the reader above finds it, part by part, as this module always did.


def _set_legacy_macro(
    path: Path, sheet_name: str, shape_name: str, macro: str
) -> dict[str, Any] | None:
    """Point a VML-only form control at a macro; None when no such control has the name."""
    book = Workbook(path)
    canonical = book.canonical_sheet_name(sheet_name)
    sheet_part = next(s.part for s in book.sheets() if s.name == canonical)
    sheet_xml = book.part_text(sheet_part)
    relationships = book.part_relationships(sheet_part)
    written: list[str] = []
    previous = ""

    vml_part = _related_part(sheet_xml, relationships, "legacyDrawing", _RELATIONSHIP_VML)
    vml = book.optional_part_text(vml_part) if vml_part else None
    if vml_part and vml is not None:
        shape_ids = _control_shape_ids(sheet_xml, shape_name)
        updated, found, before = _set_vml_macro(vml, shape_name, shape_ids, macro)
        if found:
            previous = before
            if updated != vml:
                book.set_part_text(vml_part, updated)
            written.append("form control (VML)")

    updated_sheet, found, before = _set_control_macro(sheet_xml, shape_name, macro)
    if found:
        previous = previous or before
        if updated_sheet != sheet_xml:
            book.set_part_text(sheet_part, updated_sheet)
        written.append("form control (sheet entry)")

    if not written:
        return None
    _save_package(book, path)
    return {
        "sheet": canonical,
        "shape": shape_name,
        "macro": macro,
        "previous_macro": _display_macro(previous),
        "parts_written": written,
        "cleared": not macro,
    }


def _save_package(book: Workbook, path: Path) -> None:
    """Save through the package reader, as cells.save does through pyOfficeEditor."""
    from . import locks, xlide_vscode

    try:
        book.save()
    except XlsxError as exc:
        if isinstance(exc.__cause__, PermissionError):
            raise ToolError(locks.lock_message(path, "Excel")) from exc
        raise ToolError(str(exc)) from exc
    xlide_vscode.file_changed(path, "document")


def _stored_macro(wanted: str, previous: str) -> str:
    """The macro as the file should store it, keeping whatever prefix was there.

    Excel qualifies an OnAction to the workbook, as `[1]!Module1.Proc`, and the
    index is the workbook's own. Inventing one would be a guess; keeping the one
    already on the shape is not, and a shape that never had a macro gets a plain
    name, which Excel resolves.
    """
    if not wanted:
        return ""
    match = _WORKBOOK_PREFIX.match(previous or "")
    return (match.group(0) if match else "") + wanted


def _control_shape_ids(sheet_xml: str, shape_name: str) -> set[int]:
    """The VML shape ids the sheet's `<controls>` entries give this name."""
    found: set[int] = set()
    for start, end in _elements(sheet_xml, {"control"}):
        tag = _first_tag(sheet_xml[start:end], "control")
        if tag is not None and tag.get("name", "") == shape_name:
            found.add(_int(tag.get("shapeId")))
    return found


def _set_vml_macro(
    vml: str, shape_name: str, shape_ids: set[int], macro: str
) -> tuple[str, bool, str]:
    """Rewrite `<x:FmlaMacro>` on the VML shape this control owns.

    A VML shape carries no name of its own, so it is matched by the id the
    sheet's `<controls>` entry gave it. Excel 2007 writes no such entry, and then
    the name is the one this reader synthesized from the object type and the id.
    """
    for start, end in _elements(vml, {"v:shape"}):
        inner = vml[start:end]
        object_type_match = _OBJECT_TYPE.search(inner)
        if object_type_match is None or object_type_match.group(1) == "Note":
            continue
        head = inner[: inner.find(">") + 1]
        spid = _SPID.search(head)
        if spid is None:
            continue
        identifier = int(spid.group(1))
        synthesized = f"{object_type_match.group(1)} {identifier % 1024}"
        if identifier not in shape_ids and synthesized != shape_name:
            continue
        before = _client_data(inner, "FmlaMacro")
        stored = _stored_macro(macro, before)
        updated_inner = _set_client_data(inner, "FmlaMacro", stored)
        return vml[:start] + updated_inner + vml[end:], True, before
    return vml, False, ""


def _set_control_macro(sheet_xml: str, shape_name: str, macro: str) -> tuple[str, bool, str]:
    """Rewrite the `macro` attribute on this control's `<controlPr>`."""
    for start, end in _elements(sheet_xml, {"control"}):
        body = sheet_xml[start:end]
        tag = _first_tag(body, "control")
        if tag is None or tag.get("name", "") != shape_name:
            continue
        properties = next_tag(body, body.find("<controlPr"))
        if properties is None or properties.name != "controlPr":
            return sheet_xml, False, ""
        before = properties.attrs.get("macro", "")
        original = body[properties.start : properties.end]
        rebuilt = _with_attribute(original, "macro", _stored_macro(macro, before))
        updated_body = body[: properties.start] + rebuilt + body[properties.end :]
        return sheet_xml[:start] + updated_body + sheet_xml[end:], True, before
    return sheet_xml, False, ""


def _with_attribute(tag: str, name: str, value: str) -> str:
    """Set or remove one attribute on a start tag, leaving the rest as written."""
    pattern = re.compile(rf'\s+{re.escape(name)}="[^"]*"')
    stripped = pattern.sub("", tag)
    if not value:
        return stripped
    closing = "/>" if stripped.endswith("/>") else ">"
    head = stripped[: -len(closing)]
    return f'{head} {name}="{encode_attr(value)}"{closing}'


def _set_client_data(inner: str, name: str, value: str) -> str:
    """Set or remove one `<x:Name>` element inside a VML shape's ClientData."""
    pattern = re.compile(rf"\s*<x:{name}>.*?</x:{name}>", re.DOTALL)
    stripped = pattern.sub("", inner)
    if not value:
        return stripped
    # Put it back where Excel writes it: first inside the ClientData block.
    match = re.search(r"<x:ClientData\b[^>]*>", stripped)
    if match is None:
        return stripped
    insert_at = match.end()
    return (
        stripped[:insert_at]
        + f"\n   <x:{name}>{encode_xml(value)}</x:{name}>"
        + stripped[insert_at:]
    )


__all__ = [
    "ADDABLE_KINDS",
    "CHART_TYPES",
    "Shape",
    "ToolError",
    "XlsxError",
    "add_chart",
    "add_shape",
    "chart_summary",
    "control_states",
    "read_sheet_shapes",
    "remove_shape",
    "set_shape_macro",
]
