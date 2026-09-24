"""A file saved before its first macro: an ordinary file with no VBA project.

Excel, Word and PowerPoint write no vbaProject.bin into a macro-enabled file
until the first macro exists, and a binary .xls or .doc with no macros has no
VBA storage at all. Both are normal. Before pyOpenVBA 6.1 every such file read
as a failure, and through this server as an error with no message.

The fixtures are the files the applications themselves saved, copied from
pyOpenVBA's tests/fixtures/no_vba, where scripts/measure_no_vba.py makes them.
"""

from __future__ import annotations

import shutil
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE, ToolFailure

FIXTURES = Path(__file__).parent / "fixtures" / "no_vba"


@pytest.fixture
def no_vba(workspace: Path) -> Callable[[str], Path]:
    """A copy of one fixture in the workspace, so a write cannot touch the original."""

    def copy(name: str) -> Path:
        target = workspace / name
        shutil.copy2(FIXTURES / name, target)
        return target

    return copy


@pytest.mark.parametrize(
    "name", ["workbook.xlsm", "workbook.xls", "document.docm", "document.doc", "presentation.pptm"]
)
def test_the_listings_answer_empty_and_say_why(
    call: Callable[..., Any], no_vba: Callable[[str], Path], name: str
) -> None:
    path = str(no_vba(name))

    info = call("xlide_project_info", file_path=path)
    assert info["has_vba_project"] is False
    assert info["modules"] == []
    assert "no VBA project" in info["vba_note"]

    listed = call("xlide_list_modules", file_path=path)
    assert listed["has_vba_project"] is False
    assert listed["count"] == 0
    assert "no VBA project" in listed["note"]

    assert call("xlide_list_references", file_path=path)["count"] == 0
    assert call("xlide_list_forms", file_path=path)["count"] == 0
    assert call("xlide_search_modules", file_path=path, query="Sub")["match_count"] == 0

    analysis = call("xlide_analyze", file_path=path)
    assert analysis["modules_analyzed"] == 0
    assert "no VBA project" in analysis["note"]

    validated = call("xlide_validate_project", file_path=path)
    assert validated["verdict"] == "clean"
    assert validated["has_vba_project"] is False


@pytest.mark.parametrize("tool", ["xlide_read_module", "xlide_list_procedures"])
def test_a_named_read_says_the_file_has_no_project(
    call: Callable[..., Any], no_vba: Callable[[str], Path], tool: str
) -> None:
    path = no_vba("workbook.xlsm")
    with pytest.raises(ToolFailure) as refusal:
        call(tool, file_path=str(path), module_name="Module1")
    assert "workbook.xlsm has no VBA project yet" in refusal.value.message
    assert "xlide_write_module" in refusal.value.message


def test_the_document_surface_does_not_depend_on_the_vba_half(
    call: Callable[..., Any], no_vba: Callable[[str], Path]
) -> None:
    path = str(no_vba("workbook.xlsm"))
    sheets = call("xlide_list_sheets", file_path=path)["sheets"]
    assert [sheet["name"] for sheet in sheets] == ["Sheet1"]
    written = call("xlide_write_cells", file_path=path, sheet="Sheet1", start_cell="A1", data=[[1]])
    assert written["saved"] is True
    assert call("xlide_list_shapes", file_path=path)["shape_count"] == 0


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("workbook.xlsm", {"ThisWorkbook", "Sheet1", "Helpers"}),
        ("document.docm", {"ThisDocument", "Helpers"}),
        ("presentation.pptm", {"Helpers"}),
    ],
)
def test_the_first_module_gives_the_file_its_project(
    call: Callable[..., Any], no_vba: Callable[[str], Path], name: str, expected: set[str]
) -> None:
    """The project each application makes for a first macro: in Excel the
    document modules come with it, in Word ThisDocument, in PowerPoint nothing."""
    path = no_vba(name)
    written = call(
        "xlide_write_module", file_path=str(path), module_name="Helpers", source=SAMPLE_MODULE
    )
    assert written["created"] is True
    assert written["vba_project_created"] is True

    listed = call("xlide_list_modules", file_path=str(path))
    assert listed["has_vba_project"] is True
    assert {module["name"] for module in listed["modules"]} == expected
    assert call("xlide_analyze", file_path=str(path))["verdict"] == "clean"

    import pyopenvba

    opener = {".xlsm": pyopenvba.ExcelFile, ".docm": pyopenvba.WordFile}.get(
        path.suffix, pyopenvba.PowerPointFile
    )
    with opener(path) as reopened:
        assert reopened.has_vba_project()
        assert "AddNums" in reopened.get_module("Helpers")


def test_a_document_module_can_be_the_first_one_written(
    call: Callable[..., Any], no_vba: Callable[[str], Path]
) -> None:
    """ThisWorkbook exists the moment the project does, so writing it is a write
    to the module Excel made, not the creation of a standard module named after it."""
    path = no_vba("workbook.xlsm")
    written = call(
        "xlide_write_module",
        file_path=str(path),
        module_name="ThisWorkbook",
        source="Private Sub Workbook_Open()\nEnd Sub\n",
    )
    assert written["created"] is False
    assert written["kind"] == "document"


@pytest.mark.parametrize(
    ("name", "application"), [("workbook.xls", "Excel"), ("document.doc", "Word")]
)
def test_a_binary_file_with_no_project_is_refused_a_module(
    call: Callable[..., Any], no_vba: Callable[[str], Path], name: str, application: str
) -> None:
    path = no_vba(name)
    before = path.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_write_module", file_path=str(path), module_name="Helpers", source=SAMPLE_MODULE)
    assert "has no VBA project" in refusal.value.message
    assert f"written in {application}" in refusal.value.message
    assert path.read_bytes() == before


def test_create_project_adds_the_empty_project_word_keeps(
    call: Callable[..., Any], no_vba: Callable[[str], Path]
) -> None:
    path = no_vba("document.docm")
    added = call("xlide_create_project", file_path=str(path))
    assert added["created"] is False
    assert added["vba_project_created"] is True
    assert added["modules"] == ["ThisDocument"]
    assert call("xlide_project_info", file_path=str(path))["has_vba_project"] is True

    with pytest.raises(ToolFailure) as again:
        call("xlide_create_project", file_path=str(path))
    assert "already exists, with a VBA project" in again.value.message


@pytest.mark.parametrize("name", ["workbook.xlsm", "presentation.pptm"])
def test_create_project_says_why_excel_and_powerpoint_take_no_empty_project(
    call: Callable[..., Any], no_vba: Callable[[str], Path], name: str
) -> None:
    """Excel and PowerPoint write no project until it holds code, so an empty
    one would be reported and then not be there on the next read."""
    path = no_vba(name)
    before = path.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_create_project", file_path=str(path))
    assert "writes no project until it holds code" in refusal.value.message
    assert "xlide_write_module" in refusal.value.message
    assert path.read_bytes() == before


def test_a_form_or_an_import_gives_the_file_its_project_too(
    call: Callable[..., Any], no_vba: Callable[[str], Path], workspace: Path
) -> None:
    formed = no_vba("workbook.xlsm")
    created = call(
        "xlide_manage_form", file_path=str(formed), action="create", form_name="Wizard"
    )
    assert created["vba_project_created"] is True
    assert call("xlide_list_forms", file_path=str(formed))["count"] == 1

    imported = no_vba("document.docm")
    folder = workspace / "exported"
    folder.mkdir()
    (folder / "Helpers.bas").write_text(
        'Attribute VB_Name = "Helpers"\n' + SAMPLE_MODULE, encoding="utf-8"
    )
    preview = call("xlide_import_modules", file_path=str(imported), source_folder=str(folder))
    assert preview["applied"] is False
    applied = call(
        "xlide_import_modules", file_path=str(imported), source_folder=str(folder), apply=True
    )
    assert applied["vba_project_created"] is True
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(imported))["modules"]}
    assert names == {"ThisDocument", "Helpers"}
