"""Access forms and reports, which the form tools claimed to handle and did not.

The tool descriptions said "the forms and reports in an Access database" while
the reader called forms() alone, so every report in every database came back as
not existing. A description that claims something the code does not do is worse
than a missing feature: an agent reports back that the database has no reports.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


@pytest.fixture
def designed(workspace: Path) -> Path:
    """A database with one form, one report, and a control on each."""
    import pyopenvba

    path = workspace / "Designs.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        form = db.add_form("Summary", caption="Totals", width=8000, height=3000)
        form.add_control(
            "Label", "Title", left=240, top=240, width=2000, height=300, caption="Hello"
        )
        report = db.add_report("Monthly")
        report.add_control("Label", "Banner", section="PageHeaderSection", caption="Header")
        db.save()
    return path


def test_reports_are_listed_beside_forms(
    call: Callable[..., Any], designed: Path
) -> None:
    listed = call("xlide_list_forms", file_path=str(designed))
    by_name = {entry["name"]: entry for entry in listed["forms"]}

    assert set(by_name) == {"Summary", "Monthly"}
    assert by_name["Summary"]["design"] == "form"
    assert by_name["Monthly"]["design"] == "report"
    assert listed["geometry_unit"] == "twips"


def test_an_access_design_names_the_sections_a_control_can_go_in(
    call: Callable[..., Any], designed: Path
) -> None:
    """Without these an agent adding a control to a report has to guess the band,
    and a wrong one is refused by name only after the call."""
    listed = call("xlide_list_forms", file_path=str(designed))
    report = next(e for e in listed["forms"] if e["name"] == "Monthly")
    assert "PageHeaderSection" in report["sections"]
    assert "Detail" in report["sections"]


def test_a_report_reads_its_controls(call: Callable[..., Any], designed: Path) -> None:
    read = call("xlide_read_form", file_path=str(designed), form_name="monthly")
    assert read["design"] == "report"
    assert [c["name"] for c in read["controls"]] == ["Banner"]
    assert read["controls"][0]["properties"]["Caption"] == "Header"


def test_unnamed_properties_are_counted_rather_than_listed(
    call: Callable[..., Any], designed: Path
) -> None:
    """An Access design stores more property ids the library cannot name than it
    can. Listing them buries the ones a reader came for; hiding them silently
    would be a different lie."""
    read = call("xlide_read_form", file_path=str(designed), form_name="Summary")
    properties = read["properties"]

    assert not [key for key in properties if key.startswith("Unidentified")]
    assert properties["_unnamed_property_count"] > 0


def test_a_property_change_round_trips_on_an_access_form(
    call: Callable[..., Any], designed: Path
) -> None:
    call(
        "xlide_edit_form",
        file_path=str(designed),
        form_name="Summary",
        action="set_property",
        control_name="Title",
        property_name="Caption",
        property_value="Changed",
    )
    read = call("xlide_read_form", file_path=str(designed), form_name="Summary")
    assert read["controls"][0]["properties"]["Caption"] == "Changed"


def test_a_control_can_be_added_to_a_report_band(
    call: Callable[..., Any], designed: Path
) -> None:
    call(
        "xlide_edit_form",
        file_path=str(designed),
        form_name="Monthly",
        action="add_control",
        control_name="Footer",
        control_type="Label",
        section="PageFooterSection",
        caption="Page footer",
        left=240,
        top=120,
    )
    read = call("xlide_read_form", file_path=str(designed), form_name="Monthly")
    by_name = {c["name"]: c for c in read["controls"]}
    assert set(by_name) == {"Banner", "Footer"}
    assert by_name["Footer"]["properties"]["Caption"] == "Page footer"


def test_a_control_can_be_removed_from_an_access_design(
    call: Callable[..., Any], designed: Path
) -> None:
    call(
        "xlide_edit_form",
        file_path=str(designed),
        form_name="Summary",
        action="remove_control",
        control_name="Title",
    )
    read = call("xlide_read_form", file_path=str(designed), form_name="Summary")
    assert read["controls"] == []


def test_an_unknown_design_names_both_forms_and_reports(
    call: Callable[..., Any], designed: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_read_form", file_path=str(designed), form_name="Nope")
    assert "Summary" in refusal.value.message
    assert "Monthly" in refusal.value.message
