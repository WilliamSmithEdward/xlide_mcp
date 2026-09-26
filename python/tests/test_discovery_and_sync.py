"""Discovery, forms, Power Query, and the round trip to files on disk."""

from __future__ import annotations

import os
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


def test_office_file_pages_reach_later_files(
    call: Callable[..., Any], workbook: Path, plain_workbook: Path
) -> None:
    first = call("xlide_list_projects", max_results=1)
    assert first["count"] >= 2
    assert first["next_offset"] == 1
    second = call("xlide_list_projects", max_results=1, offset=first["next_offset"])
    assert second["files"][0]["path"] != first["files"][0]["path"]


def test_project_info_is_one_shot(
    call: Callable[..., Any], workbook: Path, plain_workbook: Path
) -> None:
    info = call("xlide_project_info", file_path=str(workbook))
    assert {"modules", "forms", "password_protected", "digitally_signed"} <= set(info)

    plain = call("xlide_project_info", file_path=str(plain_workbook))
    assert plain["vba_readable"] is False
    assert plain["power_query"]["count"] == 1
    assert plain["sheets"][0]["name"] == "Sheet1"


def test_failed_form_and_query_reads_are_unknown_not_empty(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], workbook: Path
) -> None:
    import pyopenvba

    from xlide_mcp.tools import discovery

    class BrokenForms:
        def forms(self) -> list[Any]:
            raise ValueError("form storage is damaged")

    form_result = discovery._form_names(BrokenForms())
    assert form_result["forms"] is None
    assert "form storage is damaged" in form_result["forms_error"]

    def broken_queries(_path: Path) -> Any:
        raise pyopenvba.PowerQueryError("query metadata is damaged")

    monkeypatch.setattr(pyopenvba, "PowerQueryWorkbook", broken_queries)
    query_result = discovery._query_summary(workbook)
    assert query_result["count"] is None
    assert query_result["queries"] is None
    assert "query metadata is damaged" in query_result["error"]

    monkeypatch.setattr(discovery, "_form_names", lambda _handle: form_result)
    info = call("xlide_project_info", file_path=str(workbook))
    assert info["forms"] is None
    assert info["power_query"]["count"] is None
    assert "forms_error" in info
    assert "error" in info["power_query"]


def test_unreadable_applied_steps_are_not_reported_as_empty() -> None:
    from xlide_mcp.tools.powerquery import _steps

    class BrokenQuery:
        @property
        def steps(self) -> list[str]:
            raise ValueError("M expression is incomplete")

    result = _steps(BrokenQuery())
    assert result["steps"] is None
    assert "M expression is incomplete" in result["steps_error"]
    assert "Read the M formula" in result["steps_error"]


def test_doctor_answers(call: Callable[..., Any]) -> None:
    """Nothing called xlide_doctor, so when pyVBAharness 1.1.3 renamed the field it
    reads, the tool raised for everyone with the live extra and the suite stayed
    green. On a machine with Office this runs the registry probe it broke in."""
    report = call("xlide_doctor")
    assert report["server_version"]
    assert "available" in report["execution"]
    assert set(report["layers"]) >= {"files", "analysis", "execution"}


def test_the_office_settings_key_is_read_under_either_name() -> None:
    from types import SimpleNamespace

    from xlide_mcp.tools.execution import _settings_key

    assert _settings_key(SimpleNamespace(registry_key="Excel")) == "Excel"
    assert _settings_key(SimpleNamespace(security_key="Word")) == "Word"


def test_validate_project_on_a_healthy_file(call: Callable[..., Any], workbook: Path) -> None:
    report = call("xlide_validate_project", file_path=str(workbook))
    assert report["supported"] is True
    assert report["verdict"] == "clean"


def test_structural_problems_can_be_read_in_pages(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], workbook: Path
) -> None:
    from contextlib import nullcontext

    from xlide_mcp import project as project_layer

    class DamagedProject:
        def validate(self) -> list[str]:
            return [f"problem {index}" for index in range(305)]

    monkeypatch.setattr(project_layer, "open_project", lambda *_args: nullcontext(DamagedProject()))
    monkeypatch.setattr(project_layer, "has_project", lambda *_args: True)

    first = call("xlide_validate_project", file_path=str(workbook), max_results=2)
    assert first["problem_count"] == 305
    assert first["verdict"] == "problems found"
    assert first["problems"] == ["problem 0", "problem 1"]
    assert first["next_offset"] == 2
    assert "offset=2" in first["note"]

    last = call("xlide_validate_project", file_path=str(workbook), offset=304)
    assert last["problems"] == ["problem 304"]
    assert last["next_offset"] is None


def test_structural_validation_failure_is_not_called_unsupported(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], workbook: Path
) -> None:
    from contextlib import nullcontext

    from xlide_mcp import project as project_layer

    class DamagedProject:
        def validate(self) -> list[str]:
            raise AttributeError("directory record is missing")

    monkeypatch.setattr(project_layer, "open_project", lambda *_args: nullcontext(DamagedProject()))
    monkeypatch.setattr(project_layer, "has_project", lambda *_args: True)

    with pytest.raises(ToolFailure) as refusal:
        call("xlide_validate_project", file_path=str(workbook))
    assert "could not complete" in refusal.value.message
    assert "directory record is missing" in refusal.value.message


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


def test_applying_unchanged_export_preserves_file_mtime(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    call(
        "xlide_export_modules", file_path=str(workbook),
        export_folder=str(folder), apply=True,
    )
    exported = folder / "Helpers.bas"
    old_time = 946684800
    os.utime(exported, (old_time, old_time))

    result = call(
        "xlide_export_modules", file_path=str(workbook),
        export_folder=str(folder), apply=True,
    )
    assert result["files_written"] == 0
    assert result["files_unchanged"] == len(result["plan"])
    assert exported.stat().st_mtime == old_time


def test_identical_import_skips_saving(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    call(
        "xlide_export_modules", file_path=str(workbook),
        export_folder=str(folder), apply=True,
    )
    original_file = workbook.read_bytes()
    result = call(
        "xlide_import_modules", file_path=str(workbook),
        source_folder=str(folder), apply=True,
    )
    assert result["applied"] is True
    assert result["modules_changed"] == 0
    assert result["saved"] is False
    assert workbook.read_bytes() == original_file


def test_export_refuses_a_linked_module_outside_the_workspace(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    folder.mkdir()
    outside = workspace.parent / "outside-export.bas"
    outside.write_text("keep this", encoding="utf-8")
    try:
        (folder / "Helpers.bas").symlink_to(outside)
    except OSError:
        pytest.skip("creating a file symlink needs privilege on this machine")
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_export_modules", file_path=str(workbook), export_folder=str(folder),
             apply=True)
    assert "outside this server's workspace" in refusal.value.message
    assert outside.read_text(encoding="utf-8") == "keep this"


def test_import_refuses_a_linked_module_outside_the_workspace(
    call: Callable[..., Any], workbook: Path, workspace: Path
) -> None:
    folder = workspace / "vba"
    folder.mkdir()
    outside = workspace.parent / "outside-import.bas"
    outside.write_text("Option Explicit\n", encoding="utf-8")
    try:
        (folder / "Extra.bas").symlink_to(outside)
    except OSError:
        pytest.skip("creating a file symlink needs privilege on this machine")
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_import_modules", file_path=str(workbook), source_folder=str(folder))
    assert "outside this server's workspace" in refusal.value.message


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
    assert read["content_token"].startswith("xlide1:")

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
    first_page = call(
        "xlide_list_queries", file_path=str(plain_workbook), max_results=1,
    )
    assert first_page["count"] == 2
    assert first_page["next_offset"] == 1
    second_page = call(
        "xlide_list_queries", file_path=str(plain_workbook),
        max_results=1, offset=first_page["next_offset"],
    )
    assert second_page["queries"][0]["name"] != first_page["queries"][0]["name"]
    assert second_page["next_offset"] is None

    totals_read = call(
        "xlide_read_query", file_path=str(plain_workbook), query_name="Totals",
    )
    renamed = call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="rename",
        query_name="Totals",
        new_name="GrandTotals",
        expected_content_token=totals_read["content_token"],
    )
    assert renamed["content_token"] == call(
        "xlide_read_query", file_path=str(plain_workbook), query_name="GrandTotals",
    )["content_token"]
    call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="remove",
        query_name="GrandTotals",
    )
    remaining = call("xlide_list_queries", file_path=str(plain_workbook))["queries"]
    assert [q["name"] for q in remaining] == ["Numbers"]


def test_power_query_read_slices_and_guards_a_write(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    formula = "let\n    Source = {1..10}\nin\n    Source\n"
    call(
        "xlide_write_query", file_path=str(plain_workbook), action="set",
        query_name="Numbers", formula=formula,
    )
    full = call("xlide_read_query", file_path=str(plain_workbook), query_name="Numbers")
    assert full["total_lines"] == 4
    sliced = call(
        "xlide_read_query", file_path=str(plain_workbook), query_name="Numbers",
        start_line=2, end_line=2,
    )
    assert sliced["formula"] == "    Source = {1..10}"
    assert sliced["content_token"] == full["content_token"]
    changed = call(
        "xlide_write_query", file_path=str(plain_workbook), action="set",
        query_name="Numbers", formula=formula.replace("1..10", "1..20"),
        expected_content_token=sliced["content_token"],
    )
    assert changed["content_token"] != sliced["content_token"]
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_query", file_path=str(plain_workbook), action="set",
            query_name="Numbers", formula=formula,
            expected_content_token=sliced["content_token"],
        )
    assert "changed since it was read" in refusal.value.message


def test_setting_a_query_reports_the_m_that_changed(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    """M is code the same way VBA is, so a change to it gets the same diff. A
    workbook's logic can move entirely in here with no module touched."""
    written = call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="set",
        query_name="Numbers",
        formula="let Source = {1..99} in Source",
    )
    assert "-let Source = {1..10} in Source" in written["diff"]
    assert "+let Source = {1..99} in Source" in written["diff"]

    quiet = call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="set",
        query_name="Numbers",
        formula="let Source = {1..5} in Source",
        include_diff=False,
    )
    assert "diff" not in quiet


def test_an_action_with_nothing_to_diff_reports_no_diff(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    renamed = call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="rename",
        query_name="Numbers",
        new_name="Digits",
    )
    assert "diff" not in renamed


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
