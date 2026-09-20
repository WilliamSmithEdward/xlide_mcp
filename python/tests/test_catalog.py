"""References, and the rest of an Access database."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure


@pytest.fixture
def database(workspace: Path) -> Path:
    """An Access database with a table, a saved query and some code."""
    import pyopenvba

    path = workspace / "App.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        db.create_table(
            "Orders",
            [
                pyopenvba.ColumnSpec("Id", "long"),
                pyopenvba.ColumnSpec("Customer", "text", size=50),
            ],
        )
        db.create_query("AllOrders", "SELECT * FROM Orders;")
        db.save()
    return path


def test_an_excel_project_lists_its_references(
    call: Callable[..., Any], workbook: Path
) -> None:
    result = call("xlide_list_references", file_path=str(workbook))
    by_name = {entry["name"]: entry for entry in result["references"]}

    # Every Excel project carries these two from the template.
    assert "stdole" in by_name
    assert "Office" in by_name
    office = by_name["Office"]
    assert office["guid"].startswith("{")
    assert office["path"].lower().endswith(".dll")
    assert "Object Library" in office["description"]


def test_a_reference_carries_the_path_that_breaks_on_another_machine(
    call: Callable[..., Any], workbook: Path
) -> None:
    """The reason the tool exists: a reference is a path plus a version, and the
    path is the only thing that explains code compiling here and not there."""
    result = call("xlide_list_references", file_path=str(workbook))
    stdole = next(e for e in result["references"] if e["name"] == "stdole")

    assert stdole["version"] == "2.0"
    assert "stdole2.tlb" in stdole["path"]
    assert stdole["libid"].startswith("*\\G{")


def test_an_access_database_lists_its_references(
    call: Callable[..., Any], database: Path
) -> None:
    """Access keeps references in the database rather than the dir stream, and
    the reader has to answer the same way for both."""
    result = call("xlide_list_references", file_path=str(database))
    by_name = {entry["name"]: entry for entry in result["references"]}

    assert "DAO" in by_name
    assert "ACEDAO.DLL" in by_name["DAO"]["path"].upper()


def test_the_access_catalog_reads_tables_and_queries(
    call: Callable[..., Any], database: Path
) -> None:
    result = call("xlide_access_catalog", file_path=str(database))

    orders = next(table for table in result["tables"] if table["name"] == "Orders")
    columns = {column["name"]: column for column in orders["columns"]}
    assert columns["Id"]["type"] == "Long"
    assert columns["Customer"]["type"] == "Text"
    assert columns["Customer"]["size"] == 50

    query = next(entry for entry in result["queries"] if entry["name"] == "AllOrders")
    assert "SELECT * FROM Orders" in query["sql"]


def test_system_objects_are_left_out_unless_asked_for(
    call: Callable[..., Any], database: Path
) -> None:
    """Every database carries the navigation-pane tables and the relationships
    between them. Listing those by default buries whatever the user made."""
    ordinary = call("xlide_access_catalog", file_path=str(database))
    assert not [t for t in ordinary["tables"] if t["name"].startswith("MSys")]
    assert ordinary["relationships"] == []

    with_system = call(
        "xlide_access_catalog", file_path=str(database), include_system=True
    )
    assert [t for t in with_system["tables"] if t["name"].startswith("MSys")]
    assert with_system["relationships"], "the MSys relationships are real and should show"


def test_the_catalog_can_be_narrowed(call: Callable[..., Any], database: Path) -> None:
    only_queries = call("xlide_access_catalog", file_path=str(database), include="queries")
    assert "queries" in only_queries
    assert "tables" not in only_queries


def test_a_bad_include_is_refused(call: Callable[..., Any], database: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_access_catalog", file_path=str(database), include="everything")
    assert "include must be" in refusal.value.message


def test_the_catalog_is_refused_on_another_host(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_access_catalog", file_path=str(workbook))
    assert "Access databases only" in refusal.value.message
