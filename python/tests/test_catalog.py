"""References, and the rest of an Access database."""

from __future__ import annotations

import sys
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


# ----------------------------------------------------------- changing references

WORD_CODE = """Option Explicit

Public Sub Report()
    Dim doc As Word.Document
    Set doc = Nothing
End Sub
"""


def _analyzed_codes(call: Callable[..., Any], path: Path) -> set[str]:
    return {problem["code"] for problem in call("xlide_analyze", file_path=str(path))["problems"]}


def test_a_missing_library_is_reported_and_adding_it_clears_the_report(
    call: Callable[..., Any], workbook: Path
) -> None:
    """The analyzer knows the project's references now, so Word.Document in a
    workbook with no reference to Word is the compile error it is in Excel, and
    adding the reference is the fix."""
    call("xlide_write_module", file_path=str(workbook), module_name="Reports", source=WORD_CODE)
    assert "missing-library-reference" in _analyzed_codes(call, workbook)

    added = call("xlide_manage_reference", file_path=str(workbook), action="add", library="Word")
    assert added["changed"] is True
    assert added["added"]["name"] == "Word"
    assert added["added"]["guid"] == "{00020905-0000-0000-C000-000000000046}"
    listed = call("xlide_list_references", file_path=str(workbook))
    assert "Word" in {entry["name"] for entry in listed["references"]}
    assert "missing-library-reference" not in _analyzed_codes(call, workbook)

    removed = call(
        "xlide_manage_reference", file_path=str(workbook), action="remove", library="word"
    )
    assert removed["removed"] == ["Word"]
    assert "missing-library-reference" in _analyzed_codes(call, workbook)


def test_adding_a_library_twice_writes_nothing(call: Callable[..., Any], workbook: Path) -> None:
    call("xlide_manage_reference", file_path=str(workbook), action="add", library="Word")
    before = workbook.read_bytes()
    again = call("xlide_manage_reference", file_path=str(workbook), action="add", library="Word")
    assert again["changed"] is False
    assert workbook.read_bytes() == before


def test_the_hosts_own_library_is_implicit(call: Callable[..., Any], workbook: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_reference", file_path=str(workbook), action="add", library="Excel")
    assert "without a reference" in refusal.value.message


def test_an_access_databases_own_library_is_implicit_too(
    call: Callable[..., Any], access_database: Path
) -> None:
    """Access refuses with its own AccessError, not the VBAProjectError the other hosts
    raise, and that went past the handler as a bare "Error executing tool"."""
    before = access_database.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_manage_reference",
            file_path=str(access_database),
            action="add",
            library="Access",
        )
    assert "without a reference" in refusal.value.message
    assert access_database.read_bytes() == before


def test_microsoft_forms_stays_while_a_form_needs_it(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Creating a form references Microsoft Forms, as the VBA editor does when it
    inserts one; without it a KeyPress handler's MSForms.ReturnInteger does not
    compile. And once a form needs it, removing it is refused."""
    created = call(
        "xlide_manage_form", file_path=str(workbook), action="create", form_name="Wizard"
    )
    assert created["references_added"] == ["MSForms"]
    listed = call("xlide_list_references", file_path=str(workbook))["references"]
    assert "MSForms" in {entry["name"] for entry in listed}
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_reference", file_path=str(workbook), action="remove", library="MSForms")
    assert "Wizard" in refusal.value.message
    assert "stays while a form does" in refusal.value.message


def test_removing_what_is_not_there_names_what_is(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_reference", file_path=str(workbook), action="remove", library="Nope")
    assert "no reference named 'Nope'" in refusal.value.message
    assert "stdole" in refusal.value.message


def test_a_library_can_be_given_by_guid_on_any_machine(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Off Windows there is no registry to look a name up in, so the GUID and
    version are what an agent passes, with the name code uses."""
    added = call(
        "xlide_manage_reference",
        file_path=str(workbook),
        action="add",
        library="Scripting",
        guid="420b2830-e718-11cf-893d-00a0c9054228",
        version="1.0",
    )
    assert added["added"]["name"] == "Scripting"
    assert added["added"]["guid"] == "{420B2830-E718-11CF-893D-00A0C9054228}"
    assert added["added"]["version"] == "1.0"


def test_an_ambiguous_name_is_refused_with_the_candidates(
    call: Callable[..., Any], workbook: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlide_mcp import typelibs

    candidates = [
        typelibs.RegisteredLibrary("{00000000-0000-0000-0000-00000000000A}", 2, 8, "Lib A", ""),
        typelibs.RegisteredLibrary("{00000000-0000-0000-0000-00000000000B}", 6, 1, "Lib B", ""),
    ]
    monkeypatch.setattr(typelibs, "available", lambda: True)
    monkeypatch.setattr(typelibs, "resolve", lambda text: candidates)
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_manage_reference", file_path=str(workbook), action="add", library="Lib")
    assert "matches 2 registered libraries" in refusal.value.message
    assert "Lib B ({00000000-0000-0000-0000-00000000000B} 6.1)" in refusal.value.message


ADO_28 = "{2A75196C-D9EB-4129-B803-931327F72D5C}"
ADO_61 = "{B691E011-1797-432E-907A-4D8C69339129}"


def _registry(monkeypatch: pytest.MonkeyPatch, entries: list[tuple[str, str, str, str]]) -> None:
    """A registry of (guid, version, description), with each library's own name last."""
    from xlide_mcp import typelibs

    libraries = []
    names = {}
    for guid, version, description, name in entries:
        major, minor = typelibs.parse_version(version) or (1, 0)
        libraries.append(typelibs.RegisteredLibrary(guid, major, minor, description, ""))
        names[guid] = name
    monkeypatch.setattr(typelibs, "available", lambda: True)
    monkeypatch.setattr(typelibs, "registered", lambda: libraries)
    monkeypatch.setattr(typelibs, "library_name", lambda lib: names[lib.guid])


def test_a_library_resolves_by_the_name_code_writes(
    call: Callable[..., Any], workbook: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """ADODB is in no description: the References dialog says Microsoft ActiveX Data
    Objects. Each ADO version has its own GUID under the same name, and the newest
    is the one to add."""
    from xlide_mcp import typelibs

    _registry(
        monkeypatch,
        [
            (ADO_28, "2.8", "Microsoft ActiveX Data Objects 2.8 Library", "ADODB"),
            (ADO_61, "6.1", "Microsoft ActiveX Data Objects 6.1 Library", "ADODB"),
            ("{420B2830-E718-11CF-893D-00A0C9054228}", "1.0", "Microsoft Scripting Runtime",
             "Scripting"),
            ("{565783C6-CB41-11D1-8B02-00600806D9B6}", "1.2",
             "Microsoft WMI Scripting V1.2 Library", "WbemScripting"),
            ("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "2.8",
             "Microsoft Office 16.0 Object Library", "Office"),
            ("{00000000-0000-0000-0000-0000000000AA}", "2.0", "Microsoft Office Smart Tags 2.0",
             "SmartTagLib"),
            ("{0D452EE1-E08F-101A-852E-02608C4D0BB4}", "2.0",
             "Microsoft Forms 2.0 Object Library", "MSForms"),
            ("{8F825DF8-D817-472B-A783-0F7948CF838C}", "2.0",
             "Microsoft Forms 2.0 Object Library", "MSForms"),
        ],
    )
    assert typelibs.resolve("adodb").guid == ADO_61
    assert typelibs.resolve("Scripting").description == "Microsoft Scripting Runtime"
    assert typelibs.resolve("Office").description == "Microsoft Office 16.0 Object Library"
    tied = typelibs.resolve("MSForms")
    assert isinstance(tied, list) and len(tied) == 2

    added = call("xlide_manage_reference", file_path=str(workbook), action="add", library="ADODB")
    assert added["added"]["name"] == "ADODB"
    assert added["added"]["guid"] == ADO_61
    assert added["added"]["version"] == "6.1"


@pytest.mark.skipif(sys.platform != "win32", reason="the registry lookup is Windows only")
def test_adodb_resolves_against_this_machines_registry() -> None:
    pytest.importorskip("pythoncom")
    from xlide_mcp import typelibs

    found = typelibs.resolve("ADODB")
    if isinstance(found, list):
        pytest.skip("ADO is not registered on this machine")
    assert "ActiveX Data Objects" in found.description


@pytest.mark.skipif(sys.platform != "win32", reason="the registry lookup is Windows only")
def test_a_registered_library_resolves_by_its_description(
    call: Callable[..., Any], workbook: Path
) -> None:
    # The name code writes, Scripting, is read from the type library itself,
    # which needs pywin32: the live extra, not the dev one CI installs.
    pytest.importorskip("pythoncom")
    from xlide_mcp import typelibs

    if isinstance(typelibs.resolve("Microsoft Scripting Runtime"), list):
        pytest.skip("the Scripting Runtime is not registered on this machine")
    added = call(
        "xlide_manage_reference",
        file_path=str(workbook),
        action="add",
        library="Microsoft Scripting Runtime",
    )
    assert added["added"]["name"] == "Scripting"
    assert added["added"]["guid"] == "{420B2830-E718-11CF-893D-00A0C9054228}"
    assert "scrrun.dll" in added["added"]["path"].lower()


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
    assert query["sql_truncated"] is False
    assert query["sql_chars"] == len(query["sql"])


def test_a_saved_access_query_can_be_read_in_character_pages(
    call: Callable[..., Any], database: Path
) -> None:
    first = call(
        "xlide_read_access_query", file_path=str(database), query_name="allorders",
        max_chars=6,
    )
    assert first["query"] == "AllOrders"
    assert first["sql"] == "SELECT"
    assert first["next_offset"] == 6
    assert first["content_token"].startswith("xlide1:")

    second = call(
        "xlide_read_access_query", file_path=str(database), query_name="AllOrders",
        offset=first["next_offset"],
    )
    assert "SELECT * FROM Orders" in first["sql"] + second["sql"]
    assert second["content_token"] == first["content_token"]
    assert second["next_offset"] is None
    assert second["total_chars"] == len(first["sql"] + second["sql"])

    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_access_query", file_path=str(database), query_name="AllOrders",
            offset=second["total_chars"] + 1,
        )
    assert "past the end" in refusal.value.message

    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_access_query", file_path=str(database), query_name="Missing",
        )
    assert "AllOrders" in refusal.value.message


def test_the_access_catalog_previews_long_query_sql() -> None:
    from xlide_mcp.tools.catalog import _queries

    class Query:
        name = "Long"
        sql = "X" * 9_000

    class Handle:
        def queries(self) -> list[Query]:
            return [Query()]

    shown, count, next_offset = _queries(Handle(), 0, 1)
    assert count == 1
    assert next_offset is None
    assert len(shown[0]["sql"]) == 8_000
    assert shown[0]["sql_chars"] == 9_000
    assert shown[0]["sql_truncated"] is True


def test_one_unreadable_saved_query_does_not_hide_the_catalog() -> None:
    from xlide_mcp.tools.catalog import _queries

    class BrokenQuery:
        name = "Broken"

        @property
        def sql(self) -> str:
            raise ValueError("invalid SQL record")

    class GoodQuery:
        name = "Good"
        sql = "SELECT 1;"

    class Handle:
        def queries(self) -> list[Any]:
            return [BrokenQuery(), GoodQuery()]

    shown, count, next_offset = _queries(Handle(), 0, 2)
    assert count == 2
    assert next_offset is None
    assert shown[0]["sql"] is None
    assert shown[0]["sql_chars"] is None
    assert shown[0]["sql_truncated"] is None
    assert "invalid SQL record" in shown[0]["sql_error"]
    assert shown[1]["sql"] == "SELECT 1;"


def test_access_table_page_reads_only_its_requested_specs() -> None:
    from xlide_mcp.tools.catalog import _tables

    class Handle:
        def __init__(self) -> None:
            self.read_specs: list[str] = []

        def table_names(self, *, include_system: bool) -> list[str]:
            assert include_system is False
            return [f"Table{index}" for index in range(5)]

        def table_specs(self, name: str) -> tuple[list[Any], list[Any]]:
            self.read_specs.append(name)
            return [], []

    handle = Handle()
    shown, count, next_offset = _tables(handle, False, 2, 1)
    assert count == 5
    assert [entry["name"] for entry in shown] == ["Table2"]
    assert next_offset == 3
    assert handle.read_specs == ["Table2"]


def test_access_query_page_reads_only_its_requested_sql() -> None:
    from xlide_mcp.tools.catalog import _queries

    read_sql: list[str] = []

    class Query:
        def __init__(self, name: str) -> None:
            self.name = name

        @property
        def sql(self) -> str:
            read_sql.append(self.name)
            return f"SELECT * FROM {self.name}"

    class Handle:
        def queries(self) -> list[Query]:
            return [Query(f"Query{index}") for index in range(5)]

    shown, count, next_offset = _queries(Handle(), 2, 1)
    assert count == 5
    assert [entry["name"] for entry in shown] == ["Query2"]
    assert next_offset == 3
    assert read_sql == ["Query2"]


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

    first = call(
        "xlide_access_catalog", file_path=str(database), include="tables",
        include_system=True, max_results=1,
    )
    assert first["counts"]["tables"] > 1
    assert first["next_offsets"]["tables"] == 1
    second = call(
        "xlide_access_catalog", file_path=str(database), include="tables",
        include_system=True, offset=first["next_offsets"]["tables"], max_results=1,
    )
    assert len(second["tables"]) == 1
    assert second["tables"] != first["tables"]


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
