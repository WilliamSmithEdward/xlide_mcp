"""Visual Basic 6 projects, read through the same surface as everything else.

A .vbp's modules are ordinary text files, so an agent with file tools can already
edit them. What it cannot do alone is read the project as a project: which files
are modules, what each is called inside VB, and above all analyze them together
so a call from a form into a standard module resolves.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure
from tests_vb6_sources import VB6_FORM, VB6_HELPERS, VB6_MANIFEST


@pytest.fixture
def vb6_project(workspace: Path) -> Path:
    (workspace / "Helpers.bas").write_text(VB6_HELPERS, encoding="cp1252", newline="")
    (workspace / "Form1.frm").write_text(VB6_FORM, encoding="cp1252", newline="")
    project = workspace / "Demo.vbp"
    project.write_text(VB6_MANIFEST, encoding="cp1252", newline="")
    return project


def test_a_vbp_is_discovered_and_readable(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    listed = call("xlide_list_projects")
    entry = next(item for item in listed["files"] if item["name"] == "Demo.vbp")
    assert entry["host"] == "vb6"
    assert entry["readable"] is True


def test_the_project_reports_its_modules_and_name(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    info = call("xlide_project_info", file_path=str(vb6_project))
    assert info["host"] == "vb6"
    assert info["project_name"] == "DemoProject"
    by_name = {module["name"]: module for module in info["modules"]}
    assert by_name["Helpers"]["kind"] == "standard"
    # The form's name is an attribute inside the file, not its file name.
    assert by_name["Form1"]["kind"] == "userform"
    # Nothing in a .vbp can hold a password or a signature.
    assert info["password_protected"] is False
    assert info["digitally_signed"] is False


def test_a_form_reads_as_code_not_as_markup(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    """The designer block is what VB draws the form from. Handing it to an agent
    as editable source invites an edit that stops the project loading."""
    read = call("xlide_read_module", file_path=str(vb6_project), module_name="Form1")

    assert read["source"].startswith("Option Explicit")
    assert "Begin VB.Form" not in read["source"]
    assert "Attribute VB_Name" not in read["source"]
    assert "Private Sub Form_Load" in read["source"]


def test_a_write_keeps_the_designer_block(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    read = call("xlide_read_module", file_path=str(vb6_project), module_name="Form1")
    call(
        "xlide_write_module",
        file_path=str(vb6_project),
        module_name="Form1",
        source=read["source"].replace("AddNums(1, 2)", "AddNums(3, 4)"),
        expected_content_token=read["content_token"],
    )

    on_disk = (vb6_project.parent / "Form1.frm").read_text(encoding="cp1252")
    assert on_disk.startswith("VERSION 5.00")
    assert "Begin VB.Form Form1" in on_disk
    assert 'Attribute VB_Name = "Form1"' in on_disk
    assert "AddNums(3, 4)" in on_disk


def test_analysis_resolves_a_call_across_the_project(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    """The thing a file tool cannot do. Form_Load calls AddNums, which lives in
    another file; analyzed one file at a time it is an undefined name."""
    report = call("xlide_analyze", file_path=str(vb6_project))
    assert report["counts"]["error"] == 0, report["problems"]


def test_analysis_still_finds_a_real_error(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    """A clean report has to mean something, so check the analyzer is looking."""
    call(
        "xlide_write_module",
        file_path=str(vb6_project),
        module_name="Helpers",
        source="Option Explicit\r\n\r\nPublic Sub Bad()\r\n    Dim n As Long\r\n"
        '    n = "not a number"\r\nEnd Sub\r\n',
    )
    report = call("xlide_analyze", file_path=str(vb6_project), min_severity="error")
    assert report["counts"]["error"] >= 1
    problem = next(p for p in report["problems"] if p["module"] == "Helpers")
    assert problem["line"] == 5


def test_procedures_and_search_work(call: Callable[..., Any], vb6_project: Path) -> None:
    listed = call("xlide_list_procedures", file_path=str(vb6_project), module_name="Form1")
    assert [p["name"] for p in listed["procedures"]] == ["Form_Load"]

    hits = call("xlide_search_modules", file_path=str(vb6_project), query="AddNums")
    assert {hit["module"] for hit in hits["matches"]} == {"Helpers", "Form1"}


def test_references_are_read_from_the_manifest(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    result = call("xlide_list_references", file_path=str(vb6_project))
    assert [entry["name"] for entry in result["references"]] == ["OLE Automation"]
    assert "stdole2.tlb" in result["references"][0]["path"]


def test_a_module_is_created_with_its_file_and_its_manifest_line(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    """VB6 keeps a module in two places. Writing one without the other leaves a
    file nothing loads, or a manifest naming a file that is not there."""
    call(
        "xlide_write_module",
        file_path=str(vb6_project),
        module_name="Extra",
        source="Option Explicit\r\n\r\nPublic Sub Hi()\r\nEnd Sub\r\n",
    )

    assert (vb6_project.parent / "Extra.bas").is_file()
    assert "Module=Extra; Extra.bas" in vb6_project.read_text(encoding="cp1252")
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(vb6_project))["modules"]}
    assert "Extra" in names


def test_a_rename_moves_the_name_the_file_and_the_manifest_together(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    call(
        "xlide_rename_module",
        file_path=str(vb6_project),
        module_name="Helpers",
        new_name="Tools",
    )

    assert (vb6_project.parent / "Tools.bas").is_file()
    assert not (vb6_project.parent / "Helpers.bas").exists()
    manifest = vb6_project.read_text(encoding="cp1252")
    assert "Module=Tools; Tools.bas" in manifest
    assert "Helpers" not in manifest
    assert 'Attribute VB_Name = "Tools"' in (
        vb6_project.parent / "Tools.bas"
    ).read_text(encoding="cp1252")


def test_a_delete_removes_the_file_and_the_manifest_line(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    call("xlide_delete_module", file_path=str(vb6_project), module_name="Helpers")

    assert not (vb6_project.parent / "Helpers.bas").exists()
    assert "Helpers" not in vb6_project.read_text(encoding="cp1252")


def test_the_manifest_keeps_the_settings_it_already_had(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    """A .vbp carries compiler switches and paths this server has no business
    reforming, so it is edited rather than regenerated."""
    call(
        "xlide_write_module",
        file_path=str(vb6_project),
        module_name="Extra",
        source="Option Explicit\r\n",
    )
    manifest = vb6_project.read_text(encoding="cp1252")
    assert 'Startup="Form1"' in manifest
    assert "MajorVer=1" in manifest
    assert "Reference=*\\G{00020430" in manifest


def test_validate_names_a_file_the_manifest_lost(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    (vb6_project.parent / "Helpers.bas").unlink()
    report = call("xlide_validate_project", file_path=str(vb6_project))
    assert report["problem_count"] == 1
    assert "Helpers.bas" in report["problems"][0]


def test_the_office_only_tools_refuse_a_vb6_project(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    for tool in (
        "xlide_list_sheets",
        "xlide_list_queries",
        "xlide_access_catalog",
        "xlide_git_changes",
    ):
        with pytest.raises(ToolFailure):
            call(tool, file_path=str(vb6_project))


def test_export_and_import_refuse_a_vb6_project(
    call: Callable[..., Any], vb6_project: Path, workspace: Path
) -> None:
    """They move modules between a container and files. A VB6 project's modules
    are already files, and an export would write a form out under .cls."""
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_export_modules",
            file_path=str(vb6_project),
            export_folder=str(workspace / "out"),
        )
    assert "already files" in refusal.value.message


def test_a_vb6_project_has_no_userform_designer(
    call: Callable[..., Any], vb6_project: Path
) -> None:
    """A VB6 form's design is text in its own .frm, not an MSForms control tree,
    so the designer tools report none rather than pretending."""
    listed = call("xlide_list_forms", file_path=str(vb6_project))
    assert listed["count"] == 0
