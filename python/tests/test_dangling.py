"""The rest of the two-part things, and what happens when only one part moves.

Four bugs of one shape turned up while building this: a write that moved one half
of something made of two. A form and its module. An Access design and its module.
A Forms button's macro, stored twice. And a Power Query loaded onto a sheet, which
is four parts rather than two.

These cover the two that remained, and they are here as a set rather than spread
across the tool files, because the shape is the point.
"""

from __future__ import annotations

import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest


@pytest.fixture
def loaded_query(workspace: Path) -> Path:
    """A query loaded onto a sheet: a definition, a connection, a query table
    and the table itself, all naming each other."""
    import pyopenvba

    path = workspace / "Loaded.xlsx"
    with pyopenvba.PowerQueryWorkbook.create_new(path) as book:
        book.add_query("Numbers", "let Source = {1..10} in Source")
        book.load_to_sheet("Numbers", ["Value"], cell="A1")
        book.save()
    return path


def plumbing(path: Path) -> list[str]:
    """The parts a loaded query leaves in the package besides its definition."""
    with zipfile.ZipFile(path) as archive:
        return sorted(
            name
            for name in archive.namelist()
            if "connection" in name.lower()
            or "queryTable" in name
            or name.startswith("xl/tables/")
        )


def test_the_fixture_really_is_loaded(loaded_query: Path) -> None:
    """A test that passes because the fixture never had the plumbing proves
    nothing, so check it is there before checking it goes."""
    assert plumbing(loaded_query), "the fixture should carry a connection and a table"


def test_removing_a_loaded_query_takes_its_plumbing_with_it(
    call: Callable[..., Any], loaded_query: Path
) -> None:
    """Removing only the definition leaves a connection pointing at a query that
    no longer exists. Excel meets that on the next refresh rather than on open,
    so nothing says so at the time."""
    result = call(
        "xlide_write_query",
        file_path=str(loaded_query),
        action="remove",
        query_name="Numbers",
    )
    assert result["unloaded_from_sheet"] is True
    assert plumbing(loaded_query) == []

    listed = call("xlide_list_queries", file_path=str(loaded_query))
    assert listed["queries"] == []


def test_removing_an_unloaded_query_says_it_was_not_loaded(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    result = call(
        "xlide_write_query",
        file_path=str(plain_workbook),
        action="remove",
        query_name="Numbers",
    )
    assert result["unloaded_from_sheet"] is False


def test_deleting_a_module_warns_about_the_buttons_that_called_it(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    """Deleting a module a button calls is a legitimate thing to do, so this
    warns rather than refusing. What it prevents is the silent version: nothing
    rewrites an OnAction, and the user finds out by clicking."""
    result = call(
        "xlide_delete_module", file_path=str(shapes_workbook), module_name="Module1"
    )

    assert result["deleted"] == "Module1"
    orphaned = {entry["shape"] for entry in result["shapes_now_calling_nothing"]}
    assert orphaned == {"RunButton", "GoShape"}
    assert "xlide_set_shape_macro" in result["warning"]


def test_renaming_a_module_warns_about_the_same_thing(
    call: Callable[..., Any], shapes_workbook: Path
) -> None:
    result = call(
        "xlide_rename_module",
        file_path=str(shapes_workbook),
        module_name="Module1",
        new_name="Renamed",
    )
    assert result["renamed_to"] == "Renamed"
    assert len(result["shapes_still_naming_the_old_module"]) == 2


def test_no_warning_when_nothing_pointed_at_the_module(
    call: Callable[..., Any], workbook: Path
) -> None:
    """A warning on every delete is a warning nobody reads."""
    result = call("xlide_delete_module", file_path=str(workbook), module_name="Helpers")
    assert "shapes_now_calling_nothing" not in result
    assert "warning" not in result


def test_the_warning_survives_a_workbook_with_no_drawing_layer(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Looking for shapes must not fail a delete on a workbook that has none."""
    result = call("xlide_rename_module", file_path=str(workbook), module_name="Helpers",
                  new_name="Tools")
    assert result["renamed_to"] == "Tools"
