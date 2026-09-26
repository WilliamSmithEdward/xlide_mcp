"""Creating, renaming and deleting a form, and the guard that stops half of it.

A form is two things that have to agree: a designer storage, and a code module of
the same name. The module tools could move one of them, and would: renaming a
UserForm's module left a storage with no module, which the host does not show, and
a module with no storage, which is a class. Deleting it left the storage behind
entirely. Either way the form silently went, and the file still opened.

Access moves both halves together, so there it is allowed. For a UserForm nothing
here can move the storage, so it is refused rather than done badly.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


@pytest.fixture
def with_form(workspace: Path) -> Path:
    import pyopenvba

    path = workspace / "Forms.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        form = book.add_form("Wizard", caption="Setup", width=300, height=200)
        form.add_control("CommandButton", "Ok", left=12, top=12)
        book.save()
    return path


def test_an_unreadable_access_report_list_is_not_reported_as_empty() -> None:
    from xlide_mcp.errors import ToolError
    from xlide_mcp.tools.forms import _forms

    class BrokenReports:
        def forms(self) -> list[Any]:
            return []

        def reports(self) -> list[Any]:
            raise ValueError("report directory is damaged")

    with pytest.raises(ToolError) as refusal:
        _forms(BrokenReports(), "Access")
    assert "report directory is damaged" in str(refusal.value)
    assert "list is incomplete" in str(refusal.value)


def test_a_form_storage_error_is_not_called_an_unsupported_host() -> None:
    from xlide_mcp.errors import ToolError
    from xlide_mcp.tools.forms import _forms

    class BrokenForms:
        def forms(self) -> list[Any]:
            raise AttributeError("designer record is missing")

    with pytest.raises(ToolError) as refusal:
        _forms(BrokenForms(), "Excel")
    assert "could not be read" in str(refusal.value)
    assert "designer record is missing" in str(refusal.value)


def test_unreadable_form_properties_report_why_they_are_missing() -> None:
    from xlide_mcp.tools.forms import _safe_properties

    class BrokenProperties:
        def properties(self) -> dict[str, Any]:
            raise ValueError("property stream is damaged")

    result = _safe_properties(BrokenProperties())
    assert "property stream is damaged" in result["_read_error"]


def test_unreadable_access_sections_are_not_reported_as_empty() -> None:
    from xlide_mcp.tools.forms import _sections

    class BrokenSections:
        @property
        def sections(self) -> list[Any]:
            raise ValueError("section directory is damaged")

    sections, error = _sections(BrokenSections())
    assert sections is None
    assert "section directory is damaged" in (error or "")


def test_form_list_and_read_expose_section_errors(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], with_form: Path
) -> None:
    from xlide_mcp.tools import forms

    monkeypatch.setattr(
        forms, "_sections", lambda _design: (None, "The design sections could not be read")
    )
    listed = call("xlide_list_forms", file_path=str(with_form))["forms"][0]
    assert listed["sections"] is None
    assert "could not be read" in listed["sections_error"]

    read = call("xlide_read_form", file_path=str(with_form), form_name="Wizard")
    assert read["sections"] is None
    assert "could not be read" in read["sections_error"]


def test_a_userforms_module_is_reported_as_a_userform(
    call: Callable[..., Any], with_form: Path
) -> None:
    """Its text is a class. What makes it a form is the storage beside it, and a
    kind read from the text alone calls every form a class - which is a thing
    this server will happily rename."""
    modules = call("xlide_list_modules", file_path=str(with_form))["modules"]
    by_name = {module["name"]: module for module in modules}
    assert by_name["Wizard"]["kind"] == "userform"


def test_form_pages_reach_later_forms(call: Callable[..., Any], with_form: Path) -> None:
    call(
        "xlide_manage_form", file_path=str(with_form),
        action="create", form_name="Second",
    )
    first = call("xlide_list_forms", file_path=str(with_form), max_results=1)
    assert first["count"] == 2
    assert first["next_offset"] == 1
    second = call(
        "xlide_list_forms", file_path=str(with_form),
        max_results=1, offset=first["next_offset"],
    )
    assert second["forms"][0]["name"] != first["forms"][0]["name"]
    assert second["next_offset"] is None


def test_renaming_a_userforms_module_is_refused(
    call: Callable[..., Any], with_form: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_rename_module",
            file_path=str(with_form),
            module_name="Wizard",
            new_name="Renamed",
        )
    assert "design" in refusal.value.message
    # The form is still whole.
    forms = call("xlide_list_forms", file_path=str(with_form))["forms"]
    assert [entry["name"] for entry in forms] == ["Wizard"]
    assert "Wizard" in module_names(call, with_form)


def test_deleting_a_userforms_module_is_refused(
    call: Callable[..., Any], with_form: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_delete_module", file_path=str(with_form), module_name="Wizard")
    assert "orphaned" in refusal.value.message

    read = call("xlide_read_form", file_path=str(with_form), form_name="Wizard")
    assert [c["name"] for c in read["controls"]] == ["Ok"]


def test_a_userforms_code_can_still_be_written(
    call: Callable[..., Any], with_form: Path
) -> None:
    """The guard is about the name, not about the code. An agent still writes the
    event procedures."""
    call(
        "xlide_write_module",
        file_path=str(with_form),
        module_name="Wizard",
        source="Option Explicit\r\n\r\nPrivate Sub Ok_Click()\r\n    Me.Hide\r\nEnd Sub\r\n",
    )
    read = call("xlide_read_module", file_path=str(with_form), module_name="Wizard")
    assert "Ok_Click" in read["source"]


def test_a_userform_is_created_with_its_module(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Both halves, which is why this is not xlide_write_module: a module with no
    storage is a class, not a form."""
    result = call(
        "xlide_manage_form",
        file_path=str(workbook),
        action="create",
        form_name="Setup",
        caption="Set it up",
        width=300,
        height=200,
    )
    assert result["created"] is True

    forms = call("xlide_list_forms", file_path=str(workbook))["forms"]
    assert [entry["name"] for entry in forms] == ["Setup"]
    modules = call("xlide_list_modules", file_path=str(workbook))["modules"]
    assert {m["name"]: m["kind"] for m in modules}["Setup"] == "userform"


def test_creating_over_an_existing_form_is_refused(
    call: Callable[..., Any], with_form: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_form",
            file_path=str(with_form),
            action="create",
            form_name="Wizard",
        )
    assert "already exists" in refusal.value.message


def test_renaming_a_userform_through_the_form_tool_is_also_refused(
    call: Callable[..., Any], with_form: Path
) -> None:
    """The refusal is about what this server can do, not about which tool asked."""
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_form",
            file_path=str(with_form),
            action="rename",
            form_name="Wizard",
            new_name="Renamed",
        )
    assert "Excel editor" in refusal.value.message


def test_a_report_is_refused_outside_access(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_form",
            file_path=str(workbook),
            action="create",
            form_name="Monthly",
            design="report",
        )
    assert "Access databases" in refusal.value.message


# ---------------------------------------------------------------- Access


def test_an_access_form_and_report_are_created(
    call: Callable[..., Any], access_database: Path
) -> None:
    call(
        "xlide_manage_form",
        file_path=str(access_database),
        action="create",
        form_name="Summary",
        caption="Totals",
        width=8000,
        height=3000,
    )
    call(
        "xlide_manage_form",
        file_path=str(access_database),
        action="create",
        form_name="Monthly",
        design="report",
    )
    listed = call("xlide_list_forms", file_path=str(access_database))["forms"]
    by_name = {entry["name"]: entry["design"] for entry in listed}
    assert by_name == {"Summary": "form", "Monthly": "report"}


def test_an_access_form_renames_with_its_code(
    call: Callable[..., Any], access_designs: Path
) -> None:
    """Access moves the design and the module together, which is the whole reason
    rename is allowed there and refused for a UserForm."""
    assert "Form_Summary" in module_names(call, access_designs)

    call(
        "xlide_manage_form",
        file_path=str(access_designs),
        action="rename",
        form_name="Summary",
        new_name="Totals",
    )

    forms = form_names(call, access_designs)
    assert "Totals" in forms and "Summary" not in forms
    after = module_names(call, access_designs)
    assert "Form_Totals" in after and "Form_Summary" not in after


def test_an_access_report_renames_and_deletes(
    call: Callable[..., Any], access_designs: Path
) -> None:
    call(
        "xlide_manage_form",
        file_path=str(access_designs),
        action="rename",
        form_name="Monthly",
        new_name="Quarterly",
    )
    result = call(
        "xlide_manage_form",
        file_path=str(access_designs),
        action="delete",
        form_name="Quarterly",
    )
    assert result["design"] == "report"

    forms = form_names(call, access_designs)
    assert "Quarterly" not in forms and "Monthly" not in forms


def test_an_access_design_module_cannot_be_renamed_on_its_own(
    call: Callable[..., Any], access_designs: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_rename_module",
            file_path=str(access_designs),
            module_name="Form_Summary",
            new_name="Form_Other",
        )
    assert "xlide_manage_form" in refusal.value.message


def test_an_unknown_design_is_refused_by_name(
    call: Callable[..., Any], access_designs: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_form",
            file_path=str(access_designs),
            action="delete",
            form_name="Nope",
        )
    assert "Summary" in refusal.value.message


def test_a_bad_action_is_refused(call: Callable[..., Any], workbook: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_form",
            file_path=str(workbook),
            action="duplicate",
            form_name="Wizard",
        )
    assert "action must be" in refusal.value.message


def module_names(call: Callable[..., Any], path: Path) -> set[str]:
    return {m["name"] for m in call("xlide_list_modules", file_path=str(path))["modules"]}


def form_names(call: Callable[..., Any], path: Path) -> set[str]:
    return {e["name"] for e in call("xlide_list_forms", file_path=str(path))["forms"]}


def test_a_vb6_project_is_refused(call: Callable[..., Any], vb6_project: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_form",
            file_path=str(vb6_project),
            action="create",
            form_name="Form2",
        )
    assert ".frm" in refusal.value.message
