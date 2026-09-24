"""Access across the surface, because it is the host that is different.

Every other host keeps its VBA in a vbaProject.bin inside a package. Access keeps
it in the database itself, its container class releases differently, its edits
land as they are made rather than at save, and its forms are measured in twips.

Nothing here covered Access until a missing `close()` broke every Access tool in
the server and only the catalog tests noticed. These exist so the next difference
is caught by the suite rather than by a user.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

ACCESS_MODULE = (
    "Option Compare Database\r\n"
    "Option Explicit\r\n"
    "\r\n"
    "Public Function Twice(ByVal n As Long) As Long\r\n"
    "    Twice = n * 2\r\n"
    "End Function\r\n"
)


@pytest.fixture
def database(workspace: Path) -> Path:
    import pyopenvba

    path = workspace / "App.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        db.create_table("Orders", [pyopenvba.ColumnSpec("Id", "long")])
        db.vba_project().add_module("Helpers", ACCESS_MODULE)
        db.save()
    return path


def test_project_info_reads_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    info = call("xlide_project_info", file_path=str(database))

    assert info["host"] == "access"
    assert info["vba_readable"] is True
    assert "Helpers" in {module["name"] for module in info["modules"]}
    # Access keeps its VBA in the database, not in the signature streams this
    # server reads, so None is the honest answer rather than False.
    assert info["digitally_signed"] is None


def test_a_module_round_trips_in_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    read = call("xlide_read_module", file_path=str(database), module_name="Helpers")
    assert "Twice = n * 2" in read["source"]

    call(
        "xlide_write_module",
        file_path=str(database),
        module_name="Helpers",
        source=read["source"].replace("n * 2", "n * 3"),
        expected_content_token=read["content_token"],
    )
    again = call("xlide_read_module", file_path=str(database), module_name="Helpers")
    assert "n * 3" in again["source"]


def test_a_protected_access_project_is_a_refusal_that_names_the_flag(
    call: Callable[..., Any], database: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Access refuses the save with its own AccessError, not the VBAProjectError the
    other hosts raise, and that went past the handler as a bare "Error executing
    tool" rather than the refusal every other host gives."""
    import pyopenvba

    monkeypatch.setattr(pyopenvba.AccessDatabase, "vba_is_protected", lambda self: True)
    before = database.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_module",
            file_path=str(database),
            module_name="Helpers",
            source=ACCESS_MODULE.replace("n * 2", "n * 3"),
        )
    assert "password-protected" in refusal.value.message
    assert "allow_protected=true" in refusal.value.message
    assert database.read_bytes() == before


def test_a_module_can_be_created_and_deleted_in_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    created = call(
        "xlide_write_module",
        file_path=str(database),
        module_name="Extra",
        source="Option Compare Database\r\n\r\nPublic Sub Hi()\r\nEnd Sub\r\n",
    )
    assert created["created"] is True

    call("xlide_delete_module", file_path=str(database), module_name="Extra")
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(database))["modules"]}
    assert "Extra" not in names


def test_a_module_can_be_renamed_in_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    """Access renames through the database rather than the dir stream, which is a
    different call on a different object."""
    call(
        "xlide_rename_module",
        file_path=str(database),
        module_name="Helpers",
        new_name="Tools",
    )
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(database))["modules"]}
    assert "Tools" in names and "Helpers" not in names


def test_analysis_measures_access_code_against_access(
    call: Callable[..., Any], database: Path
) -> None:
    """`Option Compare Database` is a statement Excel's surface does not have.
    A host mix-up shows up here as an error on a line that is perfectly good."""
    report = call("xlide_analyze", file_path=str(database))
    assert report["host"] == "access"
    assert report["counts"]["error"] == 0, report["problems"]


def test_searching_works_in_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    hits = call("xlide_search_modules", file_path=str(database), query="Twice")
    assert hits["match_count"] >= 1
    assert hits["matches"][0]["module"] == "Helpers"


def test_procedures_are_listed_in_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    listed = call("xlide_list_procedures", file_path=str(database), module_name="Helpers")
    assert [p["name"] for p in listed["procedures"]] == ["Twice"]


def test_export_works_from_an_access_database(
    call: Callable[..., Any], database: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    applied = call(
        "xlide_export_modules",
        file_path=str(database),
        export_folder=str(folder),
        apply=True,
    )
    assert applied["applied"] is True
    assert (folder / "Helpers.bas").is_file()


def test_sheets_are_refused_on_an_access_database(
    call: Callable[..., Any], database: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_list_sheets", file_path=str(database))
    assert "Excel files only" in refusal.value.message
