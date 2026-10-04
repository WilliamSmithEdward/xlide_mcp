"""Analysis: the build gate, and the promise that a clean file reports clean."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

BROKEN = """Option Explicit

Public Sub Broken()
    Dim n As Long
    n = "not a number"
End Sub
"""

CALLS_HELPERS = """Option Explicit

Public Sub Use()
    Debug.Print AddNums(1, 2)
End Sub
"""


def test_a_clean_module_reports_no_errors(call: Callable[..., Any], workbook: Path) -> None:
    report = call("xlide_analyze", file_path=str(workbook))
    assert report["counts"]["error"] == 0
    assert report["verdict"] == "clean"


def test_a_type_mismatch_is_reported_with_a_line(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Broken",
        source=BROKEN.replace("\n", "\r\n"),
    )
    report = call("xlide_analyze", file_path=str(workbook), min_severity="error")

    assert report["verdict"] == "errors found"
    problem = next(p for p in report["problems"] if p["module"] == "Broken")
    assert problem["severity"] == "error"
    assert problem["line"] == 5, "the assignment is on line 5 of the body"
    assert "not a number" in problem["text"]
    assert "next_step" in report


def test_analysis_pages_keep_the_full_count(
    call: Callable[..., Any], workbook: Path
) -> None:
    source = BROKEN.replace(
        '    n = "not a number"',
        '    n = "first"\n    n = "second"',
    )
    call(
        "xlide_write_module", file_path=str(workbook), module_name="Broken",
        source=source,
    )
    first = call(
        "xlide_analyze", file_path=str(workbook), min_severity="error",
        module_name="Broken", max_results=1,
    )
    assert first["matching_count"] >= 2
    assert first["reported"] == 1
    assert first["next_offset"] == 1
    second = call(
        "xlide_analyze", file_path=str(workbook), min_severity="error",
        module_name="Broken", max_results=1, offset=first["next_offset"],
    )
    assert second["counts"] == first["counts"]
    assert second["problems"][0]["line"] != first["problems"][0]["line"]

    source_first = call("xlide_analyze_source", source=source, max_results=1)
    assert source_first["matching_count"] >= 2
    source_second = call(
        "xlide_analyze_source", source=source, max_results=1,
        offset=source_first["next_offset"],
    )
    assert source_second["problems"][0]["line"] != source_first["problems"][0]["line"]


def test_min_severity_filters_without_hiding_the_count(
    call: Callable[..., Any], workbook: Path
) -> None:
    # A module with no Option Explicit produces a warning and no error.
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Loose",
        source="Public Sub Whatever()\r\n    x = 1\r\nEnd Sub\r\n",
    )
    everything = call("xlide_analyze", file_path=str(workbook))
    errors_only = call("xlide_analyze", file_path=str(workbook), min_severity="error")

    assert everything["counts"]["warning"] >= 1
    # The counts are of the whole project either way; only the list is filtered.
    assert errors_only["counts"] == everything["counts"]
    assert all(p["severity"] == "error" for p in errors_only["problems"])


def test_analyze_one_module_still_resolves_the_rest(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Caller",
        source=CALLS_HELPERS.replace("\n", "\r\n"),
    )
    # Deliberately give another module a problem, so "no problems reported for
    # Caller" cannot be satisfied by the project simply being clean.
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Elsewhere",
        source=BROKEN.replace("\n", "\r\n"),
    )
    report = call("xlide_analyze", file_path=str(workbook), module_name="Caller")

    assert all(p["module"] == "Caller" for p in report["problems"])
    assert not [p for p in report["problems"] if p["severity"] == "error"], (
        "AddNums lives in Helpers and must resolve when the whole project is analyzed"
    )
    # The counts stay project-wide, so the agent still learns the file is not clean.
    assert report["counts"]["error"] >= 1
    assert report["verdict"] == "errors found"


def test_reported_lines_match_the_lines_a_read_returns(
    call: Callable[..., Any], workbook: Path
) -> None:
    """The regression that matters most: a diagnostic has to point at the line the
    agent would edit. The project stores an `Attribute VB_Name` header that a read
    strips, so an unshifted analyzer position lands one line late on every module."""
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Broken",
        source=BROKEN.replace("\n", "\r\n"),
    )
    report = call("xlide_analyze", file_path=str(workbook), min_severity="error")
    problem = next(p for p in report["problems"] if p["module"] == "Broken")

    read = call("xlide_read_module", file_path=str(workbook), module_name="Broken")
    lines = read["source"].splitlines()
    assert lines[problem["line"] - 1].strip() == 'n = "not a number"'


def test_analyze_source_needs_no_file(call: Callable[..., Any]) -> None:
    report = call("xlide_analyze_source", source=BROKEN, host="excel")
    assert report["verdict"] == "errors found"
    assert report["counts"]["error"] >= 1


def test_analyze_source_against_a_project_resolves_its_calls(
    call: Callable[..., Any], workbook: Path
) -> None:
    source = CALLS_HELPERS.replace("\n", "\r\n")
    alone = call("xlide_analyze_source", source=source, module_name="Draft", host="excel")
    with_project = call(
        "xlide_analyze_source",
        source=source,
        module_name="Draft",
        file_path=str(workbook),
    )
    assert with_project["counts"]["error"] == 0
    assert with_project["host"] == "excel"
    # Alone, AddNums is a name nothing defines; the point of file_path is that it
    # stops being one.
    assert alone["counts"]["error"] >= with_project["counts"]["error"]


def test_analyze_source_against_a_project_reports_what_analyze_would(
    call: Callable[..., Any], workbook: Path
) -> None:
    # A draft checked against a project has to get the findings it would get once
    # written into it. The project also holds a module with problems of its own,
    # which are not the draft's: the comparison below is against the draft's
    # findings alone, so one of Elsewhere's reaching the report would fail it.
    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Elsewhere",
        source=BROKEN.replace("\n", "\r\n"),
    )
    draft = (CALLS_HELPERS + BROKEN.replace("Option Explicit\n", "")).replace("\n", "\r\n")
    as_draft = call(
        "xlide_analyze_source", source=draft, module_name="Draft", file_path=str(workbook)
    )
    assert as_draft["counts"]["error"] >= 1

    call("xlide_write_module", file_path=str(workbook), module_name="Draft", source=draft)
    written = call("xlide_analyze", file_path=str(workbook), module_name="Draft")

    def key(problem: dict[str, Any]) -> tuple[Any, ...]:
        return tuple(problem.get(k) for k in ("line", "column", "severity", "code", "message"))

    assert as_draft["matching_count"] == len(as_draft["problems"])
    assert sorted(map(key, as_draft["problems"])) == sorted(map(key, written["problems"]))
    assert not [p for p in as_draft["problems"] if "AddNums" in p["message"]], (
        "AddNums lives in Helpers and must resolve against the file"
    )


def test_analyze_source_against_a_project_analyzes_only_the_draft(
    call: Callable[..., Any], workbook: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # The other modules are indexed so the draft's calls resolve, and not analyzed:
    # their findings are never reported here, and on a large project working them
    # out was most of what the call cost.
    import pyvbaanalysis
    import pyvbaanalysis.project as analyzer

    analyzed: list[str] = []
    indexed: list[str] = []
    analyze_module = analyzer.analyze_module
    build_project_index = analyzer.build_project_index

    def spy_analyze(source: str, options: Any) -> Any:
        analyzed.append(options.module_name)
        return analyze_module(source, options)

    def spy_index(modules: Any, **kwargs: Any) -> Any:
        listed = list(modules)
        indexed.extend(m.module_name for m in listed)
        return build_project_index(listed, **kwargs)

    monkeypatch.setattr(analyzer, "analyze_module", spy_analyze)
    monkeypatch.setattr(analyzer, "build_project_index", spy_index)
    assert pyvbaanalysis.analyze_project is analyzer.analyze_project

    report = call(
        "xlide_analyze_source",
        source=CALLS_HELPERS.replace("\n", "\r\n"),
        module_name="Draft",
        file_path=str(workbook),
    )

    assert report["counts"]["error"] == 0
    assert analyzed == ["Draft"]
    assert "Helpers" in indexed and "Draft" in indexed


def test_analyze_source_replaces_the_module_of_its_name_whatever_the_case(
    call: Callable[..., Any], workbook: Path
) -> None:
    # VBA compares names without case, so a draft called "helpers" stands in for
    # Helpers rather than colliding with it, and is the module reported on.
    report = call(
        "xlide_analyze_source",
        source=BROKEN.replace("\n", "\r\n"),
        module_name="helpers",
        file_path=str(workbook),
    )
    assert report["module"] == "helpers"
    assert report["counts"]["error"] >= 1
    assert report["verdict"] == "errors found"


def test_bad_severity_and_host_are_refused(call: Callable[..., Any], workbook: Path) -> None:
    with pytest.raises(ToolFailure):
        call("xlide_analyze", file_path=str(workbook), min_severity="loud")
    with pytest.raises(ToolFailure):
        call("xlide_analyze_source", source="Sub A()\r\nEnd Sub\r\n", host="outlook")


def test_rules_catalogue(call: Callable[..., Any]) -> None:
    everything = call("xlide_rules")
    assert everything["count"] > 100

    one = call("xlide_rules", code="unterminated-string")
    assert one["rule"]["default_severity"] == "error"
    assert one["rule"]["compile_error_equivalent"] is True

    narrowed = call("xlide_rules", search="option-explicit")
    assert narrowed["count"] >= 1
    assert all("option-explicit" in r["code"] for r in narrowed["rules"])

    with pytest.raises(ToolFailure):
        call("xlide_rules", code="no-such-rule")
