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
