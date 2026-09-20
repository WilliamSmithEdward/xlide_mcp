"""Discovery, forms, Power Query, and the round trip to files on disk."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


def test_list_projects_finds_what_is_in_the_workspace(
    call: Callable[..., Any], workspace: Path, workbook: Path, plain_workbook: Path
) -> None:
    listed = call("xlide_list_projects")
    by_name = {entry["name"]: entry for entry in listed["files"]}
    assert by_name["Budget.xlsm"]["readable"] is True
    assert by_name["Budget.xlsm"]["host"] == "excel"
    # A .xlsx is listed and marked unreadable, with the reason, rather than left
    # out: an agent that cannot see it will ask the user where the file went.
    assert by_name["Data.xlsx"]["readable"] is False
    assert "no VBA project by design" in by_name["Data.xlsx"]["reason"]


def test_list_projects_explains_a_recognized_but_unopenable_file(
    call: Callable[..., Any], workspace: Path
) -> None:
    (workspace / "Template.xltm").write_bytes(b"not really a template")
    entry = next(
        item for item in call("xlide_list_projects")["files"] if item["name"] == "Template.xltm"
    )
    assert entry["readable"] is False
    assert "Save a copy as .xlsm" in entry["reason"]


def test_list_projects_skips_office_lock_files(
    call: Callable[..., Any], workspace: Path, workbook: Path
) -> None:
    (workspace / "~$Budget.xlsm").write_bytes(b"lock")
    names = {entry["name"] for entry in call("xlide_list_projects")["files"]}
    assert "~$Budget.xlsm" not in names


def test_project_info_is_one_shot(
    call: Callable[..., Any], workbook: Path, plain_workbook: Path
) -> None:
    info = call("xlide_project_info", file_path=str(workbook))
    assert {"modules", "forms", "password_protected", "digitally_signed"} <= set(info)

    plain = call("xlide_project_info", file_path=str(plain_workbook))
    assert plain["vba_readable"] is False
    assert plain["power_query"]["count"] == 1
    assert plain["sheets"][0]["name"] == "Sheet1"


def test_validate_project_on_a_healthy_file(call: Callable[..., Any], workbook: Path) -> None:
    report = call("xlide_validate_project", file_path=str(workbook))
    assert report["supported"] is True
    assert report["verdict"] == "clean"


def test_create_project_never_overwrites(call: Callable[..., Any], workspace: Path) -> None:
    target = workspace / "Fresh.xlsm"
    created = call("xlide_create_project", file_path=str(target))
    assert created["created"] is True
    assert target.is_file()

    with pytest.raises(ToolFailure) as refusal:
        call("xlide_create_project", file_path=str(target))
    assert "already exists" in refusal.value.message


def test_create_project_refuses_an_extension_it_cannot_write(
    call: Callable[..., Any], workspace: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_create_project", file_path=str(workspace / "Nope.xltm"))
    assert "Save a copy as .xlsm" in refusal.value.message


def test_export_previews_before_it_writes(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    preview = call(
        "xlide_export_modules", file_path=str(workbook), export_folder=str(folder)
    )
    assert preview["applied"] is False
    assert not folder.exists(), "a preview writes nothing"
    assert all(entry["action"] == "create" for entry in preview["plan"])

    applied = call(
        "xlide_export_modules",
        file_path=str(workbook),
        export_folder=str(folder),
        apply=True,
    )
    assert applied["applied"] is True
    assert (folder / "Helpers.bas").is_file()
    # A document module exports as .cls, which is what the VBA editor writes.
    assert (folder / "ThisWorkbook.cls").is_file()


def test_export_reports_an_unchanged_file_as_unchanged(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    call(
        "xlide_export_modules",
        file_path=str(workbook),
        export_folder=str(folder),
        apply=True,
    )
    again = call("xlide_export_modules", file_path=str(workbook), export_folder=str(folder))
    assert all(entry["action"] == "unchanged" for entry in again["plan"])


def test_import_round_trips_an_edit(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    call(
        "xlide_export_modules",
        file_path=str(workbook),
        export_folder=str(folder),
        apply=True,
    )
    exported = folder / "Helpers.bas"
    exported.write_text(
        exported.read_text(encoding="utf-8").replace("a + b", "a + b + 1"),
        encoding="utf-8",
        newline="",
    )

    preview = call(
        "xlide_import_modules", file_path=str(workbook), source_folder=str(folder)
    )
    assert preview["applied"] is False
    changing = [e for e in preview["plan"] if e["action"] == "update"]
    assert [e["module"] for e in changing] == ["Helpers"]

    # Nothing changed inside the file until the import is applied. That is the
    # rule an agent most often gets wrong, so it is asserted rather than assumed.
    unchanged = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    assert "a + b + 1" not in unchanged["source"]

    applied = call(
        "xlide_import_modules",
        file_path=str(workbook),
        source_folder=str(folder),
        apply=True,
    )
    assert applied["modules_changed"] == 1
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    assert "a + b + 1" in read["source"]


def test_import_creates_a_module_from_a_new_file(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    folder.mkdir()
    (folder / "Extra.bas").write_text(
        "Option Explicit\r\n\r\nPublic Sub Extra()\r\nEnd Sub\r\n",
        encoding="utf-8",
        newline="",
    )
    applied = call(
        "xlide_import_modules",
        file_path=str(workbook),
        source_folder=str(folder),
        apply=True,
    )
    assert applied["modules_changed"] == 1
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(workbook))["modules"]}
    assert "Extra" in names


def test_import_refuses_an_empty_folder(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    empty = workspace / "empty"
    empty.mkdir()
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_import_modules", file_path=str(workbook), source_folder=str(empty))
    assert "xlide_export_modules" in refusal.value.message


def test_power_query_read_and_write(call: Callable[..., Any], plain_workbook: Path) -> None:
    listed = call("xlide_list_queries", file_path=str(plain_workbook))
    assert [q["name"] for q in listed["queries"]] == ["Numbers"]

    read = call("xlide_read_query", file_path=str(plain_workbook), query_name="numbers")
    assert read["query"] == "Numbers"
    assert "let" in read["formula"]

    call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="set",
        query_name="Numbers",
        formula="let Source = {1..20} in Source",
    )
    again = call("xlide_read_query", file_path=str(plain_workbook), query_name="Numbers")
    assert "{1..20}" in again["formula"]

    created = call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="set",
        query_name="Totals",
        formula="let Source = 1 in Source",
        group="Staging",
    )
    assert created["created"] is True
    assert "Staging" in call("xlide_list_queries", file_path=str(plain_workbook))["groups"]

    call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="rename",
        query_name="Totals",
        new_name="GrandTotals",
    )
    call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="remove",
        query_name="GrandTotals",
    )
    remaining = call("xlide_list_queries", file_path=str(plain_workbook))["queries"]
    assert [q["name"] for q in remaining] == ["Numbers"]


def test_power_query_is_refused_on_a_word_file(
    call: Callable[..., Any], workspace: Path
) -> None:
    import pyopenvba

    document = workspace / "Report.docm"
    with pyopenvba.WordFile.create_new(document) as doc:
        doc.save()
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_list_queries", file_path=str(document))
    assert "Excel packages" in refusal.value.message


def test_forms_round_trip(call: Callable[..., Any], workbook: Path) -> None:
    import pyopenvba

    with pyopenvba.ExcelFile(workbook) as book:
        form = book.add_form("Wizard", caption="Setup", width=300, height=200)
        form.add_control("CommandButton", "OkButton", left=12, top=12)
        book.save()

    listed = call("xlide_list_forms", file_path=str(workbook))
    assert listed["geometry_unit"] == "points"
    assert listed["forms"][0]["name"] == "Wizard"

    read = call("xlide_read_form", file_path=str(workbook), form_name="wizard")
    assert [c["name"] for c in read["controls"]] == ["OkButton"]

    call(
        "xlide_edit_form",
        file_path=str(workbook),
        form_name="Wizard",
        action="set_property",
        control_name="OkButton",
        property_name="Caption",
        property_value="Save",
    )
    after = call("xlide_read_form", file_path=str(workbook), form_name="Wizard")
    assert after["controls"][0]["properties"]["Caption"] == "Save"

    call(
        "xlide_edit_form",
        file_path=str(workbook),
        form_name="Wizard",
        action="add_control",
        control_name="Hint",
        control_type="Label",
        left=12,
        top=60,
        width=200,
    )
    after_add = call("xlide_read_form", file_path=str(workbook), form_name="Wizard")
    assert {c["name"] for c in after_add["controls"]} == {"OkButton", "Hint"}

    call(
        "xlide_edit_form",
        file_path=str(workbook),
        form_name="Wizard",
        action="remove_control",
        control_name="Hint",
    )
    after_remove = call("xlide_read_form", file_path=str(workbook), form_name="Wizard")
    assert {c["name"] for c in after_remove["controls"]} == {"OkButton"}


def test_unknown_form_names_the_ones_that_exist(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_read_form", file_path=str(workbook), form_name="Nope")
    assert "No form or report named" in refusal.value.message
