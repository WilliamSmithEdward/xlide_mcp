"""The VBA workflow end to end, against a workbook built for each test."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE, ToolFailure


def test_project_info_reports_modules_and_guards(call: Callable[..., Any], workbook: Path) -> None:
    info = call("xlide_project_info", file_path=str(workbook))

    assert info["host"] == "excel"
    assert info["vba_readable"] is True
    names = {module["name"] for module in info["modules"]}
    assert "Helpers" in names
    assert info["password_protected"] is False
    # A fresh workbook from the template carries no signature, and the answer has
    # to be False rather than missing: an agent deciding whether to ask the user
    # needs the fact, not its absence.
    assert info["digitally_signed"] is False


def test_read_module_strips_the_attribute_header(
    call: Callable[..., Any], workbook: Path
) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")

    assert "Attribute VB_Name" not in read["source"]
    assert "Public Function AddNums" in read["source"]
    assert read["content_token"].startswith("xlide1:")

    with_header = call(
        "xlide_read_module", file_path=str(workbook), module_name="Helpers", include_header=True
    )
    assert "Attribute VB_Name" in with_header["source"]


def test_module_name_matching_ignores_case(call: Callable[..., Any], workbook: Path) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="hELPers")
    assert read["module"] == "Helpers"


def test_read_module_slice(call: Callable[..., Any], workbook: Path) -> None:
    sliced = call(
        "xlide_read_module",
        file_path=str(workbook),
        module_name="Helpers",
        start_line=3,
        end_line=5,
    )
    assert sliced["first_line"] == 3
    assert sliced["last_line"] == 5
    assert len(sliced["source"].splitlines()) == 3
    assert "note" in sliced


def test_write_module_round_trip(call: Callable[..., Any], workbook: Path) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    edited = read["source"].replace("a + b", "a + b + 0")

    written = call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Helpers",
        source=edited,
        expected_content_token=read["content_token"],
    )
    assert written["created"] is False
    assert written["saved"] is True
    assert written["content_token"] != read["content_token"]

    again = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    assert "a + b + 0" in again["source"]
    assert again["content_token"] == written["content_token"]


def test_a_write_reports_the_diff_of_what_landed(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Against the read-back, not against what was sent. That is the difference
    between reporting the edit and reporting what the file now holds."""
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    written = call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Helpers",
        source=read["source"].replace("a + b", "a + b + 0"),
        expected_content_token=read["content_token"],
    )

    assert "-    AddNums = a + b" in written["diff"]
    assert "+    AddNums = a + b + 0" in written["diff"]
    assert written["lines_added"] == 1
    assert written["lines_removed"] == 1


def test_a_created_module_diffs_against_nothing(
    call: Callable[..., Any], workbook: Path
) -> None:
    written = call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Fresh",
        source="Option Explicit\r\n\r\nPublic Sub Go()\r\nEnd Sub\r\n",
    )
    assert written["created"] is True
    assert "+Public Sub Go()" in written["diff"]
    assert written["lines_removed"] == 0


def test_the_write_diff_can_be_turned_off(
    call: Callable[..., Any], workbook: Path
) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    written = call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Helpers",
        source=read["source"].replace("a + b", "a - b"),
        expected_content_token=read["content_token"],
        include_diff=False,
    )
    assert "diff" not in written
    assert written["lines_added"] == 1


def test_stale_token_refuses_the_write(call: Callable[..., Any], workbook: Path) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Helpers",
        source=read["source"] + "\n' someone else got here first\n",
        expected_content_token=read["content_token"],
    )

    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_module",
            file_path=str(workbook),
            module_name="Helpers",
            source=read["source"] + "\n' my edit\n",
            expected_content_token=read["content_token"],
        )
    assert "changed since it was read" in refusal.value.message
    # The refusal hands back the current token so the caller can recover in one
    # step instead of guessing what to do next.
    assert "xlide1:" in refusal.value.message


def test_write_creates_a_missing_module(call: Callable[..., Any], workbook: Path) -> None:
    created = call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="NewThing",
        source="Option Explicit\r\n\r\nPublic Sub Hi()\r\nEnd Sub\r\n",
    )
    assert created["created"] is True
    assert created["kind"] == "standard"

    as_class = call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Widget",
        source="Option Explicit\r\n",
        kind="class",
    )
    assert as_class["kind"] == "class"


def test_creating_a_module_rejects_a_stale_token(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_module",
            file_path=str(workbook),
            module_name="NotThere",
            source="Option Explicit\r\n",
            expected_content_token="xlide1:0000",
        )
    assert "No module named" in refusal.value.message


@pytest.mark.parametrize("name", ["1Bad", "Has Space", "Sub", "", "x" * 32])
def test_invalid_module_names_are_refused(
    call: Callable[..., Any], workbook: Path, name: str
) -> None:
    with pytest.raises(ToolFailure):
        call(
            "xlide_write_module",
            file_path=str(workbook),
            module_name=name,
            source="Option Explicit\r\n",
        )


def test_duplicate_name_is_refused_without_case(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_rename_module",
            file_path=str(workbook),
            module_name="Helpers",
            new_name="HELPERS",
        )
    assert "already exists" in refusal.value.message


def test_document_modules_cannot_be_renamed_or_deleted(
    call: Callable[..., Any], workbook: Path
) -> None:
    modules = call("xlide_list_modules", file_path=str(workbook))["modules"]
    document = next((m for m in modules if m["kind"] == "document"), None)
    assert document is not None, "the template workbook should carry ThisWorkbook"

    with pytest.raises(ToolFailure) as rename_refusal:
        call(
            "xlide_rename_module",
            file_path=str(workbook),
            module_name=document["name"],
            new_name="Renamed",
        )
    assert "document module" in rename_refusal.value.message

    with pytest.raises(ToolFailure) as delete_refusal:
        call("xlide_delete_module", file_path=str(workbook), module_name=document["name"])
    assert "document module" in delete_refusal.value.message


def test_rename_and_delete(call: Callable[..., Any], workbook: Path) -> None:
    renamed = call(
        "xlide_rename_module", file_path=str(workbook), module_name="Helpers", new_name="Tools"
    )
    assert renamed["renamed_to"] == "Tools"

    names = {m["name"] for m in call("xlide_list_modules", file_path=str(workbook))["modules"]}
    assert "Tools" in names and "Helpers" not in names

    deleted = call("xlide_delete_module", file_path=str(workbook), module_name="Tools")
    assert deleted["deleted"] == "Tools"
    names = {m["name"] for m in call("xlide_list_modules", file_path=str(workbook))["modules"]}
    assert "Tools" not in names


def test_list_procedures_finds_both_kinds(call: Callable[..., Any], workbook: Path) -> None:
    listed = call("xlide_list_procedures", file_path=str(workbook), module_name="Helpers")
    by_name = {p["name"]: p for p in listed["procedures"]}
    assert by_name["AddNums"]["kind"] == "Function"
    assert by_name["Greet"]["kind"] == "Sub"
    assert by_name["AddNums"]["line"] == SAMPLE_MODULE.splitlines().index(
        "Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long"
    ) + 1


def test_list_procedures_joins_a_continued_signature(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Wrapped",
        source=(
            "Option Explicit\r\n\r\n"
            "Public Function Long_(ByVal a As Long, _\r\n"
            "                      ByVal b As Long) As Long\r\n"
            "    Long_ = a + b\r\n"
            "End Function\r\n"
        ),
    )
    listed = call("xlide_list_procedures", file_path=str(workbook), module_name="Wrapped")
    names = [p["name"] for p in listed["procedures"]]
    assert names == ["Long_"], "a continued declaration is one procedure, not two"


def test_search_modules(call: Callable[..., Any], workbook: Path) -> None:
    hits = call("xlide_search_modules", file_path=str(workbook), query="addnums")
    assert hits["match_count"] >= 1
    assert all(hit["module"] == "Helpers" for hit in hits["matches"])

    none = call(
        "xlide_search_modules", file_path=str(workbook), query="addnums", match_case=True
    )
    assert none["match_count"] == 0

    regex = call(
        "xlide_search_modules",
        file_path=str(workbook),
        query=r"Public (Sub|Function)",
        is_regex=True,
    )
    assert regex["match_count"] == 2


def test_search_rejects_a_bad_pattern(call: Callable[..., Any], workbook: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_search_modules", file_path=str(workbook), query="(unclosed", is_regex=True)
    assert "regular expression" in refusal.value.message


def test_missing_module_names_the_ones_that_exist(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_read_module", file_path=str(workbook), module_name="Nope")
    assert "Helpers" in refusal.value.message
