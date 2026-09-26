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


def test_read_empty_module_returns_an_empty_body(
    call: Callable[..., Any], workbook: Path
) -> None:
    call("xlide_write_module", file_path=str(workbook), module_name="Blank", source="")
    read = call("xlide_read_module", file_path=str(workbook), module_name="Blank")
    assert read["source"] == ""
    assert read["total_lines"] == 0
    assert read["first_line"] == 0
    assert read["last_line"] == 0
    assert read["content_token"]


def test_read_module_refuses_a_reversed_slice(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_module", file_path=str(workbook), module_name="Helpers",
            start_line=5, end_line=3,
        )
    assert "before start_line" in refusal.value.message


def test_read_selected_ranges_and_edit_them_together(
    call: Callable[..., Any], workbook: Path
) -> None:
    selected = call(
        "xlide_read_module", file_path=str(workbook), module_name="Helpers",
        line_ranges=[[3, 4], [7, 10]],
    )
    assert len(selected["sections"]) == 2
    assert "AddNums = a + b" in selected["sections"][0]["source"]
    assert "Debug.Print" in selected["sections"][1]["source"]
    written = call(
        "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
        expected_content_token=selected["content_token"],
        edits=[
            {"start_line": 4, "end_line": 4, "replacement": "    AddNums = a * b"},
            {"start_line": 10, "end_line": 10,
             "replacement": '    Debug.Print "hi " & who'},
        ],
    )
    assert written["edits_applied"] == 2
    again = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    assert "AddNums = a * b" in again["source"]
    assert 'Debug.Print "hi " & who' in again["source"]


def test_edit_preview_keeps_file_and_token_usable(
    call: Callable[..., Any], workbook: Path
) -> None:
    selected = call(
        "xlide_read_module", file_path=str(workbook), module_name="Helpers",
    )
    original_file = workbook.read_bytes()
    edit = {"start_line": 4, "end_line": 4, "replacement": "    AddNums = a * b"}
    preview = call(
        "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
        expected_content_token=selected["content_token"], edits=[edit],
        preview_only=True,
    )
    assert preview["applied"] is False
    assert preview["saved"] is False
    assert preview["content_token"] == selected["content_token"]
    assert preview["lines_added"] == 1
    assert "+    AddNums = a * b" in preview["diff"]
    assert workbook.read_bytes() == original_file

    applied = call(
        "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
        expected_content_token=preview["content_token"], edits=[edit],
    )
    assert applied["applied"] is True
    assert applied["saved"] is True
    assert "+    AddNums = a * b" in applied["diff"]


def test_selected_ranges_reject_header_line_numbers(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_read_module", file_path=str(workbook), module_name="Helpers",
            line_ranges=[[1, 2]], include_header=True,
        )
    assert "body line numbers" in refusal.value.message


def test_edit_refuses_overlap_and_stale_token(
    call: Callable[..., Any], workbook: Path
) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    with pytest.raises(ToolFailure) as overlap:
        call(
            "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
            expected_content_token=read["content_token"],
            edits=[{"start_line": 3, "end_line": 4, "replacement": "x"},
                   {"start_line": 4, "end_line": 5, "replacement": "y"}],
        )
    assert "overlap" in overlap.value.message
    call(
        "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
        expected_content_token=read["content_token"],
        edits=[{"start_line": 4, "end_line": 4, "replacement": "    AddNums = a * b"}],
    )
    with pytest.raises(ToolFailure) as stale:
        call(
            "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
            expected_content_token=read["content_token"],
            edits=[{"start_line": 4, "end_line": 4, "replacement": "    AddNums = a - b"}],
        )
    assert "changed since it was read" in stale.value.message


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


def test_identical_module_write_skips_saving(
    call: Callable[..., Any], workbook: Path
) -> None:
    read = call("xlide_read_module", file_path=str(workbook), module_name="Helpers")
    original_file = workbook.read_bytes()
    result = call(
        "xlide_write_module", file_path=str(workbook), module_name="Helpers",
        source=read["source"].replace("\r\n", "\n"),
        expected_content_token=read["content_token"],
    )
    assert result["changed"] is False
    assert result["saved"] is False
    assert result["diff"] == "(no change)"
    assert result["content_token"] == read["content_token"]
    assert workbook.read_bytes() == original_file


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


def test_delete_refuses_when_shape_callers_cannot_be_checked(
    monkeypatch: pytest.MonkeyPatch, call: Callable[..., Any], workbook: Path
) -> None:
    from xlide_mcp import shapes

    def broken_shapes(_path: Path) -> Any:
        raise ValueError("drawing part is damaged")

    monkeypatch.setattr(shapes, "read_sheet_shapes", broken_shapes)
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_delete_module", file_path=str(workbook), module_name="Helpers")
    assert "Could not check which shapes call" in refusal.value.message
    assert "Nothing was deleted" in refusal.value.message
    assert "Helpers" in {
        module["name"]
        for module in call("xlide_list_modules", file_path=str(workbook))["modules"]
    }


def test_list_procedures_finds_both_kinds(call: Callable[..., Any], workbook: Path) -> None:
    listed = call("xlide_list_procedures", file_path=str(workbook), module_name="Helpers")
    assert listed["content_token"] == call(
        "xlide_read_module", file_path=str(workbook), module_name="Helpers"
    )["content_token"]
    by_name = {p["name"]: p for p in listed["procedures"]}
    assert by_name["AddNums"]["kind"] == "Function"
    assert by_name["Greet"]["kind"] == "Sub"
    assert by_name["AddNums"]["line"] == SAMPLE_MODULE.splitlines().index(
        "Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long"
    ) + 1
    first = call(
        "xlide_list_procedures", file_path=str(workbook), module_name="Helpers",
        max_results=1,
    )
    assert first["count"] == 2
    assert first["next_offset"] == 1
    second = call(
        "xlide_list_procedures", file_path=str(workbook), module_name="Helpers",
        offset=first["next_offset"], max_results=1,
    )
    assert second["procedures"][0]["name"] == "Greet"
    assert second["next_offset"] is None


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
    assert hits["content_tokens"]["Helpers"]

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


def test_search_token_can_guard_a_line_edit(
    call: Callable[..., Any], workbook: Path
) -> None:
    hits = call(
        "xlide_search_modules", file_path=str(workbook), query="AddNums = a + b",
    )
    assert hits["matches"] == [
        {"module": "Helpers", "line": 4, "text": "AddNums = a + b"}
    ]
    edited = call(
        "xlide_edit_module", file_path=str(workbook), module_name="Helpers",
        expected_content_token=hits["content_tokens"]["Helpers"],
        edits=[{"start_line": hits["matches"][0]["line"],
                "end_line": hits["matches"][0]["line"],
                "replacement": "    AddNums = a * b"}],
    )
    assert edited["applied"] is True
    assert "+    AddNums = a * b" in edited["diff"]


def test_search_pages_through_all_matches(
    call: Callable[..., Any], workbook: Path
) -> None:
    first = call(
        "xlide_search_modules", file_path=str(workbook), query="AddNums",
        max_results=1,
    )
    assert first["match_count"] == 1
    assert first["total_match_count"] == 2
    assert first["truncated"] is True
    assert first["next_offset"] == 1

    second = call(
        "xlide_search_modules", file_path=str(workbook), query="AddNums",
        max_results=1, offset=first["next_offset"],
    )
    assert second["match_count"] == 1
    assert second["total_match_count"] == 2
    assert second["matches"][0]["line"] != first["matches"][0]["line"]
    assert second["truncated"] is False
    assert second["next_offset"] is None


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
