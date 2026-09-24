"""The drawing layer: buttons, shapes, and the macros they run.

These run against `fixtures/shapes.xlsm`, the one binary fixture in the
repository. It is committed rather than built at test time because a workbook
with a Forms-toolbar button cannot be built without Excel: a form control lives
in a VML shape, a ctrlProp part, a `<controls>` entry and a hidden DrawingML
twin, and only Excel writes all four in agreement.

Excel 16 authored it by running `tests/test_shapes_live.py::test_rebuild_the_fixture`,
which also checks that the committed copy still reads the way a freshly built one
does. That test is what stops the fixture quietly becoming wrong.

It holds one sheet, Controls, with:
    RunButton  a Forms button at B2:C3, OnAction Module1.DoTheThing, caption "Run it"
    GoShape    a rectangle at E2:G4, OnAction Module1.DoTheThing, text "Go"
    Ready      a check box at B6:D7, caption "Ready?", linked to $D$6
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

# The `shapes_workbook` fixture lives in conftest.py, because the conformance
# runner needs it too: the corpus names the same fixture these tests use.


def shape(result: dict[str, Any], name: str) -> dict[str, Any]:
    for sheet in result["sheets"]:
        for entry in sheet["shapes"]:
            if entry["name"] == name:
                return entry
    raise AssertionError(f"no shape named {name}: {result}")


def test_a_forms_button_is_read_with_its_macro(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """The link an agent is usually after: a button on a sheet, and the Sub it runs."""
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    button = shape(result, "RunButton")

    assert button["kind"] == "button"
    assert button["macro"] == "Module1.DoTheThing"
    assert button["cells"] == "B2:C3"
    assert button["text"] == "Run it"


def test_a_drawing_shape_is_read_with_its_macro(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    go = shape(result, "GoShape")

    assert go["kind"] == "shape"
    assert go["macro"] == "Module1.DoTheThing"
    assert go["cells"] == "E2:G4"
    assert go["text"] == "Go"


def test_a_check_box_carries_its_linked_cell(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    ready = shape(result, "Ready")

    assert ready["kind"] == "checkBox"
    assert ready["linked_cell"] == "$D$6"
    assert ready["text"] == "Ready?"
    assert "macro" not in ready, "this check box runs nothing"


def test_a_visible_control_is_not_reported_hidden(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """Excel marks the DrawingML twin of every form control hidden, visible or
    not. Reading that flag off the twin reports every button on every sheet as
    hidden, which is worse than not reporting it at all."""
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    for name in ("RunButton", "Ready", "GoShape"):
        assert "hidden" not in shape(result, name), f"{name} is visible in the fixture"


def test_the_macros_shapes_run_are_lifted_out(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """The tree is the detail; this is the answer to the question usually asked."""
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    links = {(entry["shape"], entry["macro"]) for entry in result["macros_run_by_shapes"]}
    assert links == {
        ("RunButton", "Module1.DoTheThing"),
        ("GoShape", "Module1.DoTheThing"),
    }


def test_the_named_procedure_exists_in_the_project(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """The point of reading the link: an agent renaming DoTheThing has to know the
    button names it, because nothing rewrites an OnAction."""
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    macro = result["macros_run_by_shapes"][0]["macro"]
    module, procedure = macro.split(".", 1)

    listed = call("xlide_list_procedures", file_path=str(shapes_workbook), module_name=module)
    assert procedure in {entry["name"] for entry in listed["procedures"]}


def test_one_sheet_can_be_asked_for(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    result = call("xlide_list_shapes", file_path=str(shapes_workbook), sheet="controls")
    assert [entry["sheet"] for entry in result["sheets"]] == ["Controls"]
    assert result["shape_count"] == 3


def test_an_unknown_sheet_is_refused(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_list_shapes", file_path=str(shapes_workbook), sheet="Nope")
    assert "Controls" in refusal.value.message


def test_a_workbook_with_no_drawing_layer_reports_none(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    result = call("xlide_list_shapes", file_path=str(plain_workbook))
    assert result["shape_count"] == 0
    assert "macros_run_by_shapes" not in result


def test_a_shapes_macro_can_be_repointed(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """After renaming a Sub, nothing rewrites the OnAction that names it. This is
    how an agent closes that gap."""
    result = call(
        "xlide_set_shape_macro",
        file_path=str(shapes_workbook),
        sheet="Controls",
        shape_name="RunButton",
        macro="Module1.Renamed",
    )
    assert result["previous_macro"] == "Module1.DoTheThing"
    assert result["cleared"] is False

    after = call("xlide_list_shapes", file_path=str(shapes_workbook))
    assert shape(after, "RunButton")["macro"] == "Module1.Renamed"


def test_a_form_control_is_written_in_both_places(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """A Forms button stores its macro twice, in its VML shape and in the sheet's
    controls entry, and Excel reads both. Writing one and not the other leaves a
    button whose behaviour depends on which copy Excel happens to trust."""
    result = call(
        "xlide_set_shape_macro",
        file_path=str(shapes_workbook),
        sheet="Controls",
        shape_name="RunButton",
        macro="Module1.Renamed",
    )
    assert "form control (VML)" in result["parts_written"]
    assert "form control (sheet entry)" in result["parts_written"]
    # Read from the package, not from the result: the result is what the writer
    # says it did, and this is what Excel will find.
    vml, sheet_xml = _parts(shapes_workbook, "xl/drawings/vmlDrawing", "xl/worksheets/sheet")
    assert "Module1.Renamed" in vml
    assert "Module1.Renamed" in sheet_xml


def test_a_drawing_shape_is_written_in_its_own_part(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    result = call(
        "xlide_set_shape_macro",
        file_path=str(shapes_workbook),
        sheet="Controls",
        shape_name="GoShape",
        macro="Module1.Other",
    )
    assert result["parts_written"] == ["drawing"]
    after = call("xlide_list_shapes", file_path=str(shapes_workbook))
    assert shape(after, "GoShape")["macro"] == "Module1.Other"


def test_an_empty_macro_clears_the_link(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    call(
        "xlide_set_shape_macro",
        file_path=str(shapes_workbook),
        sheet="Controls",
        shape_name="RunButton",
        macro="",
    )
    after = call("xlide_list_shapes", file_path=str(shapes_workbook))
    assert "macro" not in shape(after, "RunButton")


def test_the_other_shapes_are_left_alone(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    before = call("xlide_list_shapes", file_path=str(shapes_workbook))
    call(
        "xlide_set_shape_macro",
        file_path=str(shapes_workbook),
        sheet="Controls",
        shape_name="RunButton",
        macro="Module1.Renamed",
    )
    after = call("xlide_list_shapes", file_path=str(shapes_workbook))

    assert after["shape_count"] == before["shape_count"]
    assert shape(after, "Ready") == shape(before, "Ready")
    assert shape(after, "GoShape") == shape(before, "GoShape")


def test_an_unknown_shape_names_the_ones_on_the_sheet(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_set_shape_macro",
            file_path=str(shapes_workbook),
            sheet="Controls",
            shape_name="Nope",
            macro="Module1.DoTheThing",
        )
    assert "RunButton" in refusal.value.message


# ------------------------------------------------------ adding and removing


def test_a_control_carries_its_state_and_every_shape_its_position(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    result = call("xlide_list_shapes", file_path=str(shapes_workbook))
    ready = shape(result, "Ready")
    assert ready["checked"] is False
    assert ready["position"]["width"] == 100.0
    assert shape(result, "RunButton")["position"]["height"] == 24.0


def test_a_button_is_added_at_a_cell_with_its_macro(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    added = call(
        "xlide_manage_shape",
        file_path=str(shapes_workbook),
        sheet="controls",
        action="add",
        shape_name="Refresh",
        kind="button",
        cell="E8",
        text="Refresh data",
        macro="Module1.DoTheThing",
    )
    assert added["sheet"] == "Controls"
    assert added["added"]["kind"] == "button"

    listed = shape(call("xlide_list_shapes", file_path=str(shapes_workbook)), "Refresh")
    assert listed["kind"] == "button"
    assert listed["macro"] == "Module1.DoTheThing"
    assert listed["text"] == "Refresh data"
    assert listed["cells"].startswith("E8"), "anchored on the cell's corner, not beside it"


def test_a_check_box_is_added_linked_to_its_cell(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    call(
        "xlide_manage_shape",
        file_path=str(shapes_workbook),
        sheet="Controls",
        action="add",
        shape_name="Include",
        kind="checkBox",
        cell="B10",
        text="Include totals",
        linked_cell="$D$10",
    )
    listed = shape(call("xlide_list_shapes", file_path=str(shapes_workbook)), "Include")
    assert listed["kind"] == "checkBox"
    assert listed["linked_cell"] == "$D$10"
    assert "macro" not in listed


def test_a_removed_control_takes_all_four_parts_with_it(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """A control is a drawing anchor, a sheet entry, a ctrlProp part and a VML
    shape. One left behind is a relationship to nothing, which Excel meets by
    repairing the whole workbook."""
    removed = call(
        "xlide_manage_shape",
        file_path=str(shapes_workbook),
        sheet="Controls",
        action="remove",
        shape_name="RunButton",
    )
    assert removed["had_macro"] == "Module1.DoTheThing"
    after = call("xlide_list_shapes", file_path=str(shapes_workbook))
    assert {entry["name"] for entry in after["sheets"][0]["shapes"]} == {"GoShape", "Ready"}

    import zipfile

    with zipfile.ZipFile(shapes_workbook) as package:
        names = package.namelist()
        properties = [name for name in names if name.startswith("xl/ctrlProps/")]
        vml, sheet_xml = _parts(shapes_workbook, "xl/drawings/vmlDrawing", "xl/worksheets/sheet")
    assert len(properties) == 1, "only the check box's control part is left"
    assert "RunButton" not in sheet_xml
    assert "Run it" not in vml


def test_a_name_already_on_the_sheet_is_refused(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_shape",
            file_path=str(shapes_workbook),
            sheet="Controls",
            action="add",
            shape_name="runbutton",
            cell="A20",
        )
    assert "already has a shape named 'RunButton'" in refusal.value.message


def test_a_picture_is_added_from_a_file_in_the_workspace(
    call: Callable[..., Any], shapes_workbook: Path, workspace: Path
) -> None:
    import base64

    image = workspace / "dot.png"
    image.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
    )
    call(
        "xlide_manage_shape",
        file_path=str(shapes_workbook),
        sheet="Controls",
        action="add",
        shape_name="Logo",
        kind="picture",
        cell="H2",
        image_path=str(image),
        text="Company logo",
    )
    listed = shape(call("xlide_list_shapes", file_path=str(shapes_workbook)), "Logo")
    assert listed["kind"] == "picture"


def test_removing_a_chart_is_refused_rather_than_half_done(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    import pyofficeeditor.excel as excel

    with excel.Workbook.open(shapes_workbook) as book:
        sheet = book.sheet("Controls")
        for reference, value in (("K1", "Amount"), ("J2", "a"), ("K2", 1), ("J3", "b"), ("K3", 2)):
            sheet.set_value(sheet[reference].reference, value)
        sheet.add_chart("column", "J1:K3", left=400, top=100, name="Totals")
        book.save()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_shape",
            file_path=str(shapes_workbook),
            sheet="Controls",
            action="remove",
            shape_name="Totals",
        )
    assert "chart part behind" in refusal.value.message


def test_a_shape_at_a_cell_lands_on_it_under_the_sheets_own_column_width(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """A sheet can set a standard column width of its own. Counting every column at
    Excel's 48 points, while pyOfficeEditor anchored against the sheet's width, put
    a button meant for C3 inside column A."""
    import re

    def wider(text: str) -> str:
        if "<sheetFormatPr" in text:
            return re.sub(r"<sheetFormatPr\b", '<sheetFormatPr defaultColWidth="20"', text, count=1)
        return text.replace(
            "<sheetData", '<sheetFormatPr defaultRowHeight="15" defaultColWidth="20"/><sheetData', 1
        )

    _rewrite(shapes_workbook, {"xl/worksheets/sheet1.xml": wider})
    call(
        "xlide_manage_shape",
        file_path=str(shapes_workbook),
        sheet="Controls",
        action="add",
        shape_name="Wide",
        kind="button",
        cell="C3",
    )
    listed = shape(call("xlide_list_shapes", file_path=str(shapes_workbook)), "Wide")
    assert listed["cells"].startswith("C3")


def test_a_control_excel_2007_saved_still_takes_a_macro(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """Excel 2007 keeps a form control in VML and nowhere else. pyOfficeEditor reaches
    a control through the DrawingML twin a later Excel writes beside it, so the
    button listed here was answered "no shape named" when it was repointed."""
    _vml_only(shapes_workbook)
    listed = call("xlide_list_shapes", file_path=str(shapes_workbook))
    button = next(
        entry for entry in listed["sheets"][0]["shapes"] if entry["kind"] == "button"
    )
    assert button["macro"] == "Module1.DoTheThing"

    result = call(
        "xlide_set_shape_macro",
        file_path=str(shapes_workbook),
        sheet="Controls",
        shape_name=button["name"],
        macro="Module1.Other",
    )
    assert result["parts_written"] == ["form control (VML)"]
    assert result["previous_macro"] == "Module1.DoTheThing"
    again = shape(call("xlide_list_shapes", file_path=str(shapes_workbook)), button["name"])
    assert again["macro"] == "Module1.Other"

    before = shapes_workbook.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_shape",
            file_path=str(shapes_workbook),
            sheet="Controls",
            action="remove",
            shape_name=button["name"],
        )
    assert "kept only in the sheet's VML drawing" in refusal.value.message
    assert shapes_workbook.read_bytes() == before


def test_a_group_is_removed_only_with_everything_in_it(
    call: Callable[..., Any], shapes_workbook: Path, workspace: Path
) -> None:
    """pyOfficeEditor takes a group off the drawing without its members' picture,
    chart and control parts, and does not reach a member at all."""
    import base64

    image = workspace / "dot.png"
    image.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
    )
    members = (("Box", "shape", {}), ("Logo", "picture", {"image_path": str(image)}))
    for name, kind, extra in members:
        call(
            "xlide_manage_shape",
            file_path=str(shapes_workbook),
            sheet="Controls",
            action="add",
            shape_name=name,
            kind=kind,
            cell="H2",
            **extra,
        )
    _group(shapes_workbook, "Pair", ("Box", "Logo"))
    before = shapes_workbook.read_bytes()

    for target, expected in (("Pair", "is a group holding 'Logo'"), ("Box", "inside the group")):
        with pytest.raises(ToolFailure) as refusal:
            call(
                "xlide_manage_shape",
                file_path=str(shapes_workbook),
                sheet="Controls",
                action="remove",
                shape_name=target,
            )
        assert expected in refusal.value.message
    assert shapes_workbook.read_bytes() == before


def _rewrite(
    workbook: Path, edits: dict[str, Callable[[str], str]], drop: tuple[str, ...] = ()
) -> None:
    """Rewrite some parts of a workbook's package, and leave out others."""
    import zipfile

    rewritten = workbook.with_suffix(".rewriting")
    with zipfile.ZipFile(workbook) as original, zipfile.ZipFile(
        rewritten, "w", zipfile.ZIP_DEFLATED
    ) as out:
        for item in original.infolist():
            if item.filename.startswith(drop):
                continue
            data = original.read(item.filename)
            edit = edits.get(item.filename)
            if edit is not None:
                data = edit(data.decode("utf-8")).encode("utf-8")
            out.writestr(item, data)
    rewritten.replace(workbook)


def _vml_only(workbook: Path) -> None:
    """The fixture's form controls as Excel 2007 saves them: the VML shapes alone."""
    import re

    def without_alternates(text: str) -> str:
        return re.sub(r"<mc:AlternateContent\b.*?</mc:AlternateContent>", "", text, flags=re.S)

    def without_controls(text: str) -> str:
        # The <controls> block nests an alternate per control inside its own.
        start = text.index("<mc:AlternateContent")
        end = text.rindex("</mc:AlternateContent>") + len("</mc:AlternateContent>")
        return text[:start] + text[end:]

    def without_control_links(text: str) -> str:
        return re.sub(r'<Relationship [^>]*Type="[^"]*/ctrlProp"[^>]*/>', "", text)

    def without_control_types(text: str) -> str:
        return re.sub(r'<Override PartName="/xl/ctrlProps/[^"]*"[^>]*/>', "", text)

    _rewrite(
        workbook,
        {
            "xl/drawings/drawing1.xml": without_alternates,
            "xl/worksheets/sheet1.xml": without_controls,
            "xl/worksheets/_rels/sheet1.xml.rels": without_control_links,
            "[Content_Types].xml": without_control_types,
        },
        drop=("xl/ctrlProps/",),
    )


def _group(workbook: Path, name: str, members: tuple[str, ...]) -> None:
    """Move shapes out of their own anchors into one group, as Excel's Group does."""
    import re

    anchor = re.compile(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", re.S)
    element = re.compile(r"<xdr:(sp|pic|cxnSp|graphicFrame)\b.*</xdr:\1>", re.S)

    def grouped(text: str) -> str:
        moved: list[str] = []
        for member in members:
            found = next(m for m in anchor.finditer(text) if f'name="{member}"' in m.group(0))
            moved.append(element.search(found.group(0)).group(0))  # type: ignore[union-attr]
            text = text[: found.start()] + text[found.end() :]
        group = (
            '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="914400" cy="914400"/>'
            f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="900" name="{name}"/><xdr:cNvGrpSpPr/>'
            '</xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/>'
            '<a:chExt cx="914400" cy="914400"/></a:xfrm></xdr:grpSpPr>'
            + "".join(moved)
            + "</xdr:grpSp><xdr:clientData/></xdr:absoluteAnchor>"
        )
        return text.replace("</xdr:wsDr>", group + "</xdr:wsDr>")

    _rewrite(workbook, {"xl/drawings/drawing1.xml": grouped})


def _parts(workbook: Path, *prefixes: str) -> tuple[str, ...]:
    """The text of the first part whose name starts with each prefix."""
    import zipfile

    with zipfile.ZipFile(workbook) as package:
        names = sorted(package.namelist())
        return tuple(
            package.read(next(n for n in names if n.startswith(prefix))).decode("utf-8")
            for prefix in prefixes
        )
