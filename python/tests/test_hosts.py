"""The same surface against all four hosts, and one host per parameter.

Every host this server did not test turned out to have a bug in it. Access had
four: no close() on its container, a save that took no signature flag, a module
kind spelled as a string, a rename that addressed design modules only. All four
were shipped and green, because every fixture was Excel or Word.

So the core surface runs against Excel, Word, PowerPoint and Access here, from
one set of cases, and a host that behaves differently fails on the case it
differs on rather than in front of a user.

What each host does NOT have is as much the point as what it does: Power Query
and worksheet cells are Excel's, and a tool asked for them elsewhere refuses by
name rather than answering emptily.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

# A module every host compiles. Access wants Option Compare Database, and it is
# harmless on the others, but each host's analyzer measures against its own
# object model so the body stays to the language itself.
MODULE = (
    "Option Explicit\r\n"
    "\r\n"
    "Public Function Twice(ByVal n As Long) As Long\r\n"
    "    Twice = n * 2\r\n"
    "End Function\r\n"
)

ACCESS_MODULE = "Option Compare Database\r\n" + MODULE

HOSTS = [
    pytest.param(("excel", "Book.xlsm"), id="excel"),
    pytest.param(("word", "Doc.docm"), id="word"),
    pytest.param(("powerpoint", "Deck.pptm"), id="powerpoint"),
    pytest.param(("access", "App.accdb"), id="access"),
]


@pytest.fixture
def host_file(request: pytest.FixtureRequest, workspace: Path) -> tuple[str, Path]:
    """One file per host, each with the same module in it."""
    import pyopenvba

    host, file_name = request.param
    maker = {
        "excel": pyopenvba.ExcelFile,
        "word": pyopenvba.WordFile,
        "powerpoint": pyopenvba.PowerPointFile,
        "access": pyopenvba.AccessDatabase,
    }[host]
    path = workspace / file_name
    source = ACCESS_MODULE if host == "access" else MODULE
    with maker.create_new(path) as handle:
        handle.vba_project().add_module("Helpers", source)
        handle.save()
    return host, path


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_the_project_reads(call: Callable[..., Any], host_file: tuple[str, Path]) -> None:
    host, path = host_file
    info = call("xlide_project_info", file_path=str(path))

    assert info["host"] == host
    assert info["vba_readable"] is True
    assert "Helpers" in {module["name"] for module in info["modules"]}
    assert info["password_protected"] is False


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_a_module_round_trips(call: Callable[..., Any], host_file: tuple[str, Path]) -> None:
    _, path = host_file
    read = call("xlide_read_module", file_path=str(path), module_name="Helpers")
    assert read["kind"] == "standard"
    assert "Twice = n * 2" in read["source"]

    call(
        "xlide_write_module",
        file_path=str(path),
        module_name="Helpers",
        source=read["source"].replace("n * 2", "n * 3"),
        expected_content_token=read["content_token"],
    )
    again = call("xlide_read_module", file_path=str(path), module_name="Helpers")
    assert "n * 3" in again["source"]


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_a_module_is_created_renamed_and_deleted(
    call: Callable[..., Any], host_file: tuple[str, Path]
) -> None:
    host, path = host_file
    body = "Option Compare Database\r\n" if host == "access" else "Option Explicit\r\n"

    created = call(
        "xlide_write_module", file_path=str(path), module_name="Extra", source=body
    )
    assert created["created"] is True

    call(
        "xlide_rename_module", file_path=str(path), module_name="Extra", new_name="Renamed"
    )
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(path))["modules"]}
    assert "Renamed" in names and "Extra" not in names

    call("xlide_delete_module", file_path=str(path), module_name="Renamed")
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(path))["modules"]}
    assert "Renamed" not in names


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_analysis_measures_against_the_right_host(
    call: Callable[..., Any], host_file: tuple[str, Path]
) -> None:
    """Word code judged by Excel's surface produces confident nonsense, so the
    host has to travel with the file all the way into the analyzer."""
    host, path = host_file
    report = call("xlide_analyze", file_path=str(path))

    assert report["host"] == host
    assert report["counts"]["error"] == 0, report["problems"]


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_procedures_and_search(call: Callable[..., Any], host_file: tuple[str, Path]) -> None:
    _, path = host_file
    listed = call("xlide_list_procedures", file_path=str(path), module_name="Helpers")
    assert [p["name"] for p in listed["procedures"]] == ["Twice"]

    hits = call("xlide_search_modules", file_path=str(path), query="Twice")
    assert "Helpers" in {hit["module"] for hit in hits["matches"]}


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_references_are_read(call: Callable[..., Any], host_file: tuple[str, Path]) -> None:
    host, path = host_file
    result = call("xlide_list_references", file_path=str(path))
    names = {entry["name"] for entry in result["references"]}

    # Every VBA project references the OLE Automation library.
    assert "stdole" in names
    assert result["host"] == host


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_modules_export_to_files(
    call: Callable[..., Any], host_file: tuple[str, Path], workspace: Path
) -> None:
    host, path = host_file
    folder = workspace / f"{host}_vba"
    applied = call(
        "xlide_export_modules",
        file_path=str(path),
        export_folder=str(folder),
        apply=True,
    )
    assert applied["applied"] is True
    assert (folder / "Helpers.bas").is_file()


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_a_form_can_be_created_on_every_host_that_has_them(
    call: Callable[..., Any], host_file: tuple[str, Path]
) -> None:
    host, path = host_file
    result = call(
        "xlide_manage_form", file_path=str(path), action="create", form_name="Wizard"
    )
    assert result["created"] is True

    forms = call("xlide_list_forms", file_path=str(path))["forms"]
    assert "Wizard" in {entry["name"] for entry in forms}

    modules = call("xlide_list_modules", file_path=str(path))["modules"]
    if host != "access":
        # A UserForm's module carries its name; an Access design's is Form_X.
        assert {m["name"]: m["kind"] for m in modules}["Wizard"] == "userform"


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_a_userforms_module_is_protected_on_every_host(
    call: Callable[..., Any], host_file: tuple[str, Path]
) -> None:
    """The guard has to hold wherever forms exist, not only where it was found."""
    host, path = host_file
    call("xlide_manage_form", file_path=str(path), action="create", form_name="Wizard")
    module = "Form_Wizard" if host == "access" else "Wizard"

    with pytest.raises(ToolFailure):
        call(
            "xlide_rename_module",
            file_path=str(path),
            module_name=module,
            new_name="Moved",
        )


# --------------------------------------------- what a host does not have


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_the_excel_only_tools_refuse_the_others_by_name(
    call: Callable[..., Any], host_file: tuple[str, Path]
) -> None:
    """Power Query and the grid are Excel's. Answering emptily elsewhere would
    tell an agent the document has no queries, which is a different claim from
    the document cannot have any."""
    host, path = host_file
    for tool in ("xlide_list_queries", "xlide_list_sheets", "xlide_list_shapes"):
        if host == "excel":
            continue
        with pytest.raises(ToolFailure) as refusal:
            call(tool, file_path=str(path))
        assert "Excel" in refusal.value.message


@pytest.mark.parametrize("host_file", HOSTS, indirect=True)
def test_the_access_catalog_refuses_the_others(
    call: Callable[..., Any], host_file: tuple[str, Path]
) -> None:
    host, path = host_file
    if host == "access":
        assert call("xlide_access_catalog", file_path=str(path))["tables"] is not None
        return
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_access_catalog", file_path=str(path))
    assert "Access databases only" in refusal.value.message


def test_an_orphaned_design_is_reported_rather_than_counted(
    call: Callable[..., Any], workspace: Path
) -> None:
    """A designer storage with no module is what deleting a form's module on its
    own leaves behind. The editor does not show it, so counting it as an ordinary
    form reports a form the user cannot open.

    Built by deleting the module through pyOpenVBA directly, which still leaves
    the storage, rather than through this server, which refuses to. Until 6.1
    PowerPoint's template shipped in exactly this state and served as the
    fixture; the template was remade from a clean presentation."""
    import pyopenvba

    path = workspace / "Orphan.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        book.add_form("Wizard", caption="Setup")
        book.save()
    with pyopenvba.ExcelFile(path) as book:
        book.vba_project().delete_module("Wizard")
        book.save()

    listed = call("xlide_list_forms", file_path=str(path))
    orphans = [entry["name"] for entry in listed["forms"] if entry.get("orphaned")]
    assert orphans == ["Wizard"]
    assert "does not show it" in listed["note"]
