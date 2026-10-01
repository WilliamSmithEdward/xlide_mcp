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

from xlide_mcp import Settings


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


def test_a_module_path_no_file_can_have_is_missing(tmp_path: Path) -> None:
    # Found by fuzzing: a NUL in a module line raised ValueError out of
    # pathlib rather than being reported like any file the folder lacks.
    from xlide_mcp.vb6 import Vb6Project

    manifest = tmp_path / "Bad.vbp"
    manifest.write_bytes(b'Type=Exe\r\nModule=Helpers; Help\x00ers.bas\r\nName="Bad"\r\n')
    project = Vb6Project(manifest, settings=Settings(roots=(tmp_path.resolve(),)))
    assert project.module_names() == []
    assert len(project.validate()) == 1


# ------------------------------------------------- module files and the roots
#
# A .vbp is input this server does not control, so the module files it names are
# held to the workspace roots like any path a caller sends. A project naming one
# outside them is refused, for reading and so for writing: the module is never
# opened, and nothing is written back to it.


def _manifest(folder: Path, module_line: str) -> Path:
    manifest = folder / "Shared.vbp"
    manifest.write_text(
        f'Type=Exe\r\n{module_line}\r\nName="Shared"\r\n', encoding="cp1252", newline=""
    )
    return manifest


def _module(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(VB6_HELPERS, encoding="cp1252", newline="")
    return path


@pytest.fixture
def elsewhere(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """A folder outside the workspace root, as a sibling project's would be."""
    return tmp_path_factory.mktemp("elsewhere")


def test_a_module_in_a_subfolder_of_the_root_is_read(
    call: Callable[..., Any], workspace: Path
) -> None:
    _module(workspace / "shared" / "Helpers.bas")
    manifest = _manifest(workspace, "Module=Helpers; shared\\Helpers.bas")
    info = call("xlide_project_info", file_path=str(manifest))
    assert [module["name"] for module in info["modules"]] == ["Helpers"]


def test_an_absolute_module_path_outside_the_roots_is_refused(
    call: Callable[..., Any], workspace: Path, elsewhere: Path
) -> None:
    outside = _module(elsewhere / "Helpers.bas")
    manifest = _manifest(workspace, f"Module=Helpers; {outside.resolve()}")
    for tool, arguments in (
        ("xlide_project_info", {}),
        ("xlide_read_module", {"module_name": "Helpers"}),
    ):
        with pytest.raises(ToolFailure) as refused:
            call(tool, file_path=str(manifest), **arguments)
        assert "outside this server's workspace" in refused.value.message
        assert "--root" in refused.value.message


def test_a_dotdot_path_out_of_the_root_is_refused(
    call: Callable[..., Any], workspace: Path
) -> None:
    project_dir = workspace / "project"
    project_dir.mkdir()
    _module(workspace.parent / f"{workspace.name}-sibling" / "Helpers.bas")
    escape = f"..\\..\\{workspace.name}-sibling\\Helpers.bas"
    manifest = _manifest(project_dir, f"Module=Helpers; {escape}")
    with pytest.raises(ToolFailure) as refused:
        call("xlide_project_info", file_path=str(manifest))
    assert "Helpers.bas" in refused.value.message


def test_a_write_never_reaches_a_module_outside_the_roots(
    call: Callable[..., Any], workspace: Path, elsewhere: Path
) -> None:
    outside = _module(elsewhere / "Helpers.bas")
    before = outside.read_bytes()
    manifest = _manifest(workspace, f"Module=Helpers; {outside.resolve()}")
    with pytest.raises(ToolFailure):
        call(
            "xlide_write_module",
            file_path=str(manifest),
            module_name="Helpers",
            source="Option Explicit\r\n",
        )
    assert outside.read_bytes() == before


def test_a_symlink_inside_the_root_does_not_carry_a_module_out(
    call: Callable[..., Any], workspace: Path, elsewhere: Path
) -> None:
    outside = _module(elsewhere / "Helpers.bas")
    link = workspace / "Helpers.bas"
    try:
        link.symlink_to(outside)
    except (OSError, NotImplementedError):
        pytest.skip("this machine cannot create a symlink without privileges")
    manifest = _manifest(workspace, "Module=Helpers; Helpers.bas")
    with pytest.raises(ToolFailure) as refused:
        call("xlide_project_info", file_path=str(manifest))
    assert "outside this server's workspace" in refused.value.message


def test_allow_outside_roots_lets_a_shared_module_through(
    workspace: Path, elsewhere: Path
) -> None:
    from xlide_mcp.vb6 import Vb6Project

    outside = _module(elsewhere / "Helpers.bas")
    manifest = _manifest(workspace, f"Module=Helpers; {outside.resolve()}")
    roots = (workspace.resolve(),)
    with pytest.raises(Exception, match="outside this server's workspace"):
        Vb6Project(manifest, settings=Settings(roots=roots))
    project = Vb6Project(manifest, settings=Settings(roots=roots, allow_outside_roots=True))
    assert project.module_names() == ["Helpers"]
    # Adding the module's folder as a root is the ordinary way to grant it.
    project = Vb6Project(manifest, settings=Settings(roots=(*roots, elsewhere.resolve())))
    assert project.module_names() == ["Helpers"]
