"""An Office file rendered as text, and the git textconv driver built on it.

The driver's contract with git is unusual and worth pinning: git aborts the whole
diff on a non-zero exit, hands over a temporary file with no useful name, and
compares whatever comes back on stdout. Two of the tests here are about files
this cannot read at all, because that is where a renderer quietly reports "no
change" for two files that share nothing.
"""

from __future__ import annotations

import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from xlide_mcp.hosts import sniff_extension
from xlide_mcp.textual import render, sections


def run_textconv(path: Path) -> subprocess.CompletedProcess:
    """The driver as git runs it: a subprocess, one argument, stdout is the text."""
    return subprocess.run(
        [sys.executable, "-m", "xlide_mcp", "--textconv", str(path)],
        capture_output=True,
        text=True,
        check=False,
    )


# ------------------------------------------------------------------ rendering


def test_a_workbook_renders_its_modules(workbook: Path) -> None:
    text = render(workbook)
    assert "' ===== module Helpers =====" in text
    assert "AddNums" in text


def test_power_query_renders_beside_the_vba(
    call: Callable[..., Any], workbook: Path
) -> None:
    call(
        "xlide_write_query",
        file_path=str(workbook),
        action="set",
        query_name="Orders",
        formula="let Source = {1..10} in Source",
    )
    text = render(workbook)
    assert "' ===== module Helpers =====" in text
    assert "' ===== query Orders =====" in text
    assert "let Source = {1..10} in Source" in text


def test_the_rendered_text_says_what_it_leaves_out(workbook: Path) -> None:
    """Without this line a reader takes an empty diff for an unchanged workbook,
    when every number on every sheet may have moved."""
    assert "cell values are not compared" in render(workbook).splitlines()[0].casefold()


def test_sheets_are_listed_so_a_structural_change_shows(workbook: Path) -> None:
    text = render(workbook)
    assert "' ===== sheets =====" in text
    assert "Sheet1" in text


def test_sections_come_back_in_an_order_that_does_not_move(
    call: Callable[..., Any], workbook: Path
) -> None:
    """Container order is not a stable order. If it reached the diff, the first
    host that reordered its streams would read as a rewrite of the whole file."""
    for name in ("Zeta", "Alpha", "Mid"):
        call(
            "xlide_write_module",
            file_path=str(workbook),
            module_name=name,
            source=f"Option Explicit\r\n\r\nPublic Sub {name}_()\r\nEnd Sub\r\n",
        )
    names = [s.name for s in sections(workbook) if s.kind == "module"]
    assert names == sorted(names, key=str.casefold)


def test_a_line_ending_change_is_not_a_change(
    call: Callable[..., Any], workbook: Path
) -> None:
    """VBA is stored CRLF and edited LF by half the tools that touch it. A
    renderer that passed that through would make every diff a whole-file rewrite."""
    before = render(workbook)
    unchanged_but_for_endings = next(
        s.source for s in sections(workbook) if s.name == "Helpers"
    ).replace("\r\n", "\n")

    call(
        "xlide_write_module",
        file_path=str(workbook),
        module_name="Helpers",
        source=unchanged_but_for_endings,
    )
    assert render(workbook) == before


# ------------------------------------------------------------------ the driver


def test_the_driver_prints_the_text_and_succeeds(workbook: Path) -> None:
    result = run_textconv(workbook)
    assert result.returncode == 0
    assert "' ===== module Helpers =====" in result.stdout


def test_the_driver_reads_a_file_with_no_extension(workbook: Path, tmp_path: Path) -> None:
    """The case that matters. git writes the blob from history to a temporary
    file, and every reader in this server is chosen by extension."""
    blob = tmp_path / "0123abcd"
    blob.write_bytes(workbook.read_bytes())

    assert sniff_extension(blob) == ".xlsm"
    result = run_textconv(blob)
    assert result.returncode == 0
    assert "AddNums" in result.stdout


def test_a_file_this_cannot_read_still_succeeds(tmp_path: Path) -> None:
    """A non-zero exit here aborts git's whole diff, including the files it could
    have shown."""
    junk = tmp_path / "notes.bin"
    junk.write_bytes(b"\x00\x01\x02 not an office file at all")
    result = run_textconv(junk)
    assert result.returncode == 0
    assert "xlide-mcp" in result.stdout


def test_two_unreadable_files_do_not_render_the_same(tmp_path: Path) -> None:
    """The trap in the line above. A renderer that returns one fixed sentence for
    everything it cannot read makes git report no change between two files with
    nothing in common, which is a wrong answer rather than a missing one."""
    first = tmp_path / "a.bin"
    second = tmp_path / "b.bin"
    first.write_bytes(b"\x00\x01\x02 one")
    second.write_bytes(b"\x00\x01\x02 two")

    assert run_textconv(first).stdout != run_textconv(second).stdout


def test_a_missing_file_says_so_without_failing(tmp_path: Path) -> None:
    result = run_textconv(tmp_path / "gone.xlsm")
    assert result.returncode == 0
    assert "no such file" in result.stdout


# ------------------------------------------------------------- end to end, git


@pytest.mark.skipif(
    subprocess.run(["git", "--version"], capture_output=True).returncode != 0,
    reason="this test is about git",
)
def test_git_diff_shows_the_vba_through_the_driver(workspace: Path, workbook: Path) -> None:
    """The whole point, proved the way a user meets it: `git diff` on a workbook.

    Without the driver git says "Binary files differ" and stops. This asserts
    both halves, because a test that only checked the good case would pass on a
    repository where the driver was never wired up at all.
    """

    def git(*arguments: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            ["git", "-C", str(workspace), *arguments],
            capture_output=True,
            text=True,
            check=False,
        )

    git("init", "-b", "main")
    git("config", "user.email", "tests@example.invalid")
    git("config", "user.name", "Tests")
    # `binary diff=vba`, not `diff=vba`. The binary macro is -diff -merge -text,
    # and the later diff=vba overrides only its -diff: the file keeps -text, so
    # git never applies end-of-line conversion to a container it would corrupt.
    (workspace / ".gitattributes").write_text("*.xlsm binary diff=vba\n", encoding="utf-8")
    git("add", "-A")
    git("commit", "-m", "before")

    attributes = git("check-attr", "text", "merge", "diff", "--", "Budget.xlsm").stdout
    assert "text: unset" in attributes
    assert "merge: unset" in attributes
    assert "diff: vba" in attributes

    import pyopenvba

    with pyopenvba.ExcelFile(workbook) as book:
        book.set_module("Helpers", "Option Explicit\r\n\r\nPublic Sub Later()\r\nEnd Sub\r\n")
        book.save()

    assert "Binary files" in git("diff").stdout

    # git runs a textconv driver through a shell, which eats Windows backslashes.
    interpreter = Path(sys.executable).as_posix()
    git("config", "diff.vba.textconv", f'"{interpreter}" -m xlide_mcp --textconv')

    diff = git("diff").stdout
    assert "Binary files" not in diff
    assert "+Public Sub Later()" in diff
    assert "-Public Function AddNums" in diff
