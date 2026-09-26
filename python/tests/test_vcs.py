"""Diffing an Office file against a git revision.

Every test builds a real repository and commits a real workbook, because the
thing under test is exactly the interaction with git: a mocked git would pass
while the tool failed on the first repository it met.
"""

from __future__ import annotations

import shutil
import subprocess
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE, ToolFailure

pytestmark = pytest.mark.skipif(
    shutil.which("git") is None, reason="these tests are about git, and git is not installed"
)


def git(repository: Path, *arguments: str) -> None:
    subprocess.run(
        ["git", "-C", str(repository), *arguments],
        check=True,
        capture_output=True,
        text=True,
    )


@pytest.fixture
def committed(workspace: Path, workbook: Path) -> Path:
    """The fixture workbook, committed to a repository at the workspace root."""
    git(workspace, "init", "-b", "main")
    git(workspace, "config", "user.email", "tests@example.invalid")
    git(workspace, "config", "user.name", "Tests")
    # Office files are binary; a normalizing filter would corrupt the container.
    (workspace / ".gitattributes").write_text("*.xlsm binary\n", encoding="utf-8")
    git(workspace, "add", "-A")
    git(workspace, "commit", "-m", "the workbook as it was")
    return workbook


def find(report: dict[str, Any], name: str) -> dict[str, Any]:
    return next(entry for entry in report["changes"] if entry["name"] == name)


def test_an_unchanged_file_reports_no_changes(
    call: Callable[..., Any], committed: Path
) -> None:
    report = call("xlide_git_changes", file_path=str(committed))
    assert report["changed"] == 0
    assert report["verdict"] == "no changes"
    assert all(entry["status"] == "unchanged" for entry in report["changes"])


def test_the_result_says_what_it_does_not_compare(
    call: Callable[..., Any], committed: Path
) -> None:
    """A reviewer who does not know the scope reads silence as "nothing changed",
    so every report carries it rather than leaving it in the documentation."""
    report = call("xlide_git_changes", file_path=str(committed))
    assert "cell values are not compared" in report["covers"].casefold()


def test_a_modified_module_is_reported_with_its_diff(
    call: Callable[..., Any], committed: Path
) -> None:
    call(
        "xlide_write_module",
        file_path=str(committed),
        module_name="Helpers",
        source=SAMPLE_MODULE.replace("\n", "\r\n").replace("a + b", "a + b + 1"),
    )
    report = call("xlide_git_changes", file_path=str(committed))

    assert report["changed"] == 1
    entry = find(report, "Helpers")
    assert entry["kind"] == "module"
    assert entry["status"] == "modified"
    assert entry["lines_added"] == 1
    assert entry["lines_removed"] == 1
    assert "-    AddNums = a + b" in entry["diff"]
    assert "+    AddNums = a + b + 1" in entry["diff"]


def test_an_added_module_is_reported_as_added(
    call: Callable[..., Any], committed: Path
) -> None:
    call(
        "xlide_write_module",
        file_path=str(committed),
        module_name="Brand",
        source="Option Explicit\r\n\r\nPublic Sub New_()\r\nEnd Sub\r\n",
    )
    report = call("xlide_git_changes", file_path=str(committed))
    entry = find(report, "Brand")
    assert entry["status"] == "added"
    assert entry["lines_removed"] == 0
    assert entry["lines_added"] > 0


def test_a_deleted_module_is_reported_as_removed(
    call: Callable[..., Any], committed: Path
) -> None:
    call("xlide_delete_module", file_path=str(committed), module_name="Helpers")
    report = call("xlide_git_changes", file_path=str(committed))
    entry = find(report, "Helpers")
    assert entry["status"] == "removed"
    assert entry["lines_added"] == 0
    assert "AddNums" in entry["diff"]


def test_a_power_query_change_is_reported_beside_the_vba(
    call: Callable[..., Any], committed: Path
) -> None:
    """The reason this tool reports sections rather than modules. A workbook's
    logic can move entirely in its M with the VBA untouched, and a review that
    only looked at modules would call that commit empty."""
    call(
        "xlide_write_query",
        file_path=str(committed),
        action="set",
        query_name="Orders",
        formula="let Source = {1..10} in Source",
    )
    report = call("xlide_git_changes", file_path=str(committed))

    entry = find(report, "Orders")
    assert entry["kind"] == "query"
    assert entry["status"] == "added"
    assert "let Source = {1..10} in Source" in entry["diff"]
    assert all(e["status"] == "unchanged" for e in report["changes"] if e["kind"] == "module")


def test_one_section_can_be_asked_for(call: Callable[..., Any], committed: Path) -> None:
    call(
        "xlide_write_module",
        file_path=str(committed),
        module_name="Helpers",
        source=SAMPLE_MODULE.replace("\n", "\r\n").replace("a + b", "a - b"),
    )
    report = call("xlide_git_changes", file_path=str(committed), section="helpers")
    assert [entry["name"] for entry in report["changes"]] == ["Helpers"]


def test_the_diff_can_be_left_out(call: Callable[..., Any], committed: Path) -> None:
    call(
        "xlide_write_module",
        file_path=str(committed),
        module_name="Helpers",
        source=SAMPLE_MODULE.replace("\n", "\r\n").replace("a + b", "a - b"),
    )
    report = call("xlide_git_changes", file_path=str(committed), include_diff=False)
    entry = find(report, "Helpers")
    assert entry["lines_added"] == 1
    assert "diff" not in entry


def test_a_long_diff_reports_how_much_it_withheld() -> None:
    """The branch an ordinary fixture never reaches. It once reported the number
    of lines it kept rather than the number it dropped, which reads as though a
    400-line diff had 400 more lines behind it."""
    from xlide_mcp.tools._common import MAX_DIFF_LINES, unified_diff

    before = "\r\n".join(f"    x = {i}" for i in range(500))
    after = "\r\n".join(f"    y = {i}" for i in range(500))
    text, was_cut = unified_diff(before, after, label="Big", narrower="Ask for it alone.")

    lines = text.splitlines()
    assert was_cut
    assert len(lines) == MAX_DIFF_LINES + 1
    assert lines[-1].startswith("... 603 more diff lines")


def test_long_change_counts_lines_beyond_the_display_limit() -> None:
    from types import SimpleNamespace

    from xlide_mcp.tools.vcs import _change

    before = SimpleNamespace(name="Big", kind="module",
                             source="\n".join(f"x = {i}" for i in range(500)))
    after = SimpleNamespace(name="Big", kind="module",
                            source="\n".join(f"y = {i}" for i in range(500)))
    change = _change(after, before, after)
    assert change.truncated is True
    assert change.added == 500
    assert change.removed == 500


def test_a_file_outside_a_repository_says_so(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_git_changes", file_path=str(workbook))
    assert "not inside a git repository" in refusal.value.message


def test_an_untracked_file_says_everything_in_it_is_new(
    call: Callable[..., Any], workspace: Path, committed: Path
) -> None:
    import pyopenvba

    fresh = workspace / "Later.xlsm"
    with pyopenvba.ExcelFile.create_new(fresh) as book:
        book.save()
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_git_changes", file_path=str(fresh))
    assert "not tracked" in refusal.value.message


def test_an_unknown_revision_says_which_one(
    call: Callable[..., Any], committed: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_git_changes", file_path=str(committed), revision="no-such-branch")
    assert "no-such-branch" in refusal.value.message


def test_a_name_that_exists_in_neither_version_is_refused(
    call: Callable[..., Any], committed: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_git_changes", file_path=str(committed), section="Ghost")
    assert "Helpers" in refusal.value.message
