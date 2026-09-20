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

The one write here is narrow on purpose: an existing shape's macro can be changed
or cleared, and nothing is created or destroyed. That is what keeps it safe.
Adding or removing a form control means keeping four parts in agreement, and
getting that wrong produces a workbook that opens and then repairs itself, so it
is not offered. Changing a value on parts that already exist has no such failure,
and Excel was made to open the result and read the macro back.
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


# --------------------------------------------------------------------- writing


def set_shape_macro(
    path: Path, sheet_name: str, shape_name: str, macro: str
) -> dict[str, Any]:
    """Point an existing shape at a different macro, or at none.

    Deliberately narrow. This changes a value on parts that already exist and
    creates or destroys nothing, which is what keeps it safe: adding or removing
    a form control means keeping four parts in agreement, and getting that wrong
    produces a workbook that opens and then repairs itself.

    A form control stores its macro twice, in its VML shape and in the sheet's
    `<controls>` entry, and Excel reads both. Writing one and not the other
    leaves a button whose behaviour depends on which one Excel happens to trust.
    """
    book = Workbook(path)
    canonical = book.canonical_sheet_name(sheet_name)
    sheet_part = next(s.part for s in book.sheets() if s.name == canonical)
    sheet_xml = book.part_text(sheet_part)
    relationships = book.part_relationships(sheet_part)

    wanted = (macro or "").strip()
    written: list[str] = []
    previous = ""

    # The drawing part: the macro is an attribute on the shape's own element.
    drawing_part = _related_part(sheet_xml, relationships, "drawing", _RELATIONSHIP_DRAWING)
    if drawing_part:
        drawing_xml = book.optional_part_text(drawing_part)
        if drawing_xml is not None:
            updated, found, before = _set_drawing_macro(drawing_xml, shape_name, wanted)
            if found:
                previous = previous or before
                if updated != drawing_xml:
                    book.set_part_text(drawing_part, updated)
                written.append("drawing")

    # The form control: the same value in its VML shape and in the sheet entry.
    control_shape_ids = _control_shape_ids(sheet_xml, shape_name)
    vml_part = _related_part(sheet_xml, relationships, "legacyDrawing", _RELATIONSHIP_VML)
    if vml_part:
        vml = book.optional_part_text(vml_part)
        if vml is not None:
            updated, found, before = _set_vml_macro(
                vml, shape_name, control_shape_ids, wanted
            )
            if found:
                previous = previous or before
                if updated != vml:
                    book.set_part_text(vml_part, updated)
                written.append("form control (VML)")

    updated_sheet, found, before = _set_control_macro(sheet_xml, shape_name, wanted)
    if found:
        previous = previous or before
        if updated_sheet != sheet_xml:
            book.set_part_text(sheet_part, updated_sheet)
        written.append("form control (sheet entry)")

    if not written:
        existing = read_sheet_shapes(path, canonical).get(canonical, [])
        listed = ", ".join(shape.name for shape in existing) or "(none)"
        raise ToolError(
            f"No shape named {shape_name!r} on {canonical}. Shapes there: {listed}."
        )

    book.save()
    return {
        "sheet": canonical,
        "shape": shape_name,
        "macro": wanted,
        "previous_macro": _display_macro(previous),
        "parts_written": written,
        "cleared": not wanted,
    }


def _set_drawing_macro(xml: str, shape_name: str, macro: str) -> tuple[str, bool, str]:
    """Rewrite the `macro` attribute on the element that holds this cNvPr name."""
    for anchor_start, anchor_end in _elements(
        xml, {"xdr:twoCellAnchor", "xdr:oneCellAnchor", "xdr:absoluteAnchor",
              "mc:AlternateContent"}
    ):
        body = xml[anchor_start:anchor_end]
        for _, start, end in _drawing_children(body, 0):
            element = body[start:end]
            properties = _first_tag(element, "xdr:cNvPr")
            if properties is None or properties.get("name", "") != shape_name:
                continue
            opening = next_tag(element, 0)
            if opening is None:
                continue
            tag = element[opening.start : opening.end]
            before = opening.attrs.get("macro", "")
            rebuilt = _with_attribute(tag, "macro", _stored_macro(macro, before))
            updated_element = rebuilt + element[opening.end :]
            updated_body = body[:start] + updated_element + body[end:]
            return xml[:anchor_start] + updated_body + xml[anchor_end:], True, before
    return xml, False, ""


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


def _int(value: str | None) -> int:
    try:
        return int(value or 0)
    except ValueError:
        return 0


__all__ = ["Shape", "ToolError", "XlsxError", "read_sheet_shapes", "set_shape_macro"]
