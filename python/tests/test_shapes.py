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
