"""Fixtures: a real workbook, and a way to call a tool the way a client does.

Tests go through `call`, which runs the tool through the server exactly as the
protocol does, rather than reaching for the Python function behind it. That is the
only way the schema, the validation and the error path are covered at all: a
direct call skips every one of them, which is where the bugs an agent would hit
actually live.
"""

from __future__ import annotations

import asyncio
import json
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from xlide_mcp import Settings, build_server

SAMPLE_MODULE = """Option Explicit

Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function

Public Sub Greet()
    Dim who As String
    who = "world"
    Debug.Print "hello " & who
End Sub
"""
"""Clean VBA on purpose: a fixture that analyzes dirty makes every analysis test
argue with the fixture instead of with the code. PyVbaLog and the assertion
helpers are harness names that exist only at run time, so they live in the live
execution tests and nowhere else."""


class ToolFailure(Exception):
    """A tool answered with isError. Carries the message the agent would see."""

    def __init__(self, tool: str, message: str) -> None:
        super().__init__(f"{tool}: {message}")
        self.tool = tool
        self.message = message


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--rebuild-fixture",
        action="store_true",
        default=False,
        help="Let a live test overwrite a committed binary fixture. Off by default.",
    )


@pytest.fixture
def workspace(tmp_path: Path) -> Path:
    return tmp_path


@pytest.fixture
def settings(workspace: Path) -> Settings:
    return Settings(roots=(workspace.resolve(),))


@pytest.fixture
def call(settings: Settings) -> Callable[..., Any]:
    """Call a tool by name and get its structured result, or raise ToolFailure."""
    server = build_server(settings)

    def invoke(name: str, **arguments: Any) -> Any:
        try:
            result = asyncio.run(server.call_tool(name, arguments))
        except ToolError as exc:
            # In-process, the SDK re-raises a deliberate tool failure; over the
            # wire the kernel turns the same thing into isError with this text.
            # Tests assert on the message either way.
            raise ToolFailure(name, str(exc)) from exc
        text = "\n".join(
            block.text for block in getattr(result, "content", []) if hasattr(block, "text")
        )
        if getattr(result, "isError", False) or getattr(result, "is_error", False):
            raise ToolFailure(name, text)
        structured = getattr(result, "structuredContent", None) or getattr(
            result, "structured_content", None
        )
        if structured is not None:
            return structured
        try:
            return json.loads(text)
        except ValueError:
            return text

    return invoke


@pytest.fixture
def workbook(workspace: Path) -> Path:
    """A macro-enabled workbook with one standard module in it."""
    import pyopenvba

    path = workspace / "Budget.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        book.vba_project().add_module("Helpers", SAMPLE_MODULE)
        book.save()
    return path


@pytest.fixture
def plain_workbook(workspace: Path) -> Path:
    """A workbook with no macros, for the sheets and Power Query paths."""
    import pyopenvba

    path = workspace / "Data.xlsx"
    with pyopenvba.PowerQueryWorkbook.create_new(path) as book:
        book.add_query("Numbers", "let Source = {1..10} in Source")
        book.save()
    return path


@pytest.fixture
def word_document(workspace: Path) -> Path:
    """A host that is not Excel, so the host-specific refusals have a real subject."""
    import pyopenvba

    path = workspace / "Report.docm"
    with pyopenvba.WordFile.create_new(path) as document:
        document.save()
    return path


@pytest.fixture
def committed_workbook(workspace: Path, workbook: Path) -> Path:
    """The fixture workbook, committed to a git repository at the workspace root.

    The file is marked binary in .gitattributes: a normalizing filter would
    rewrite bytes inside the container, so a fixture that skips that step tests a
    corrupted workbook rather than the tool.
    """
    import subprocess

    def git(*arguments: str) -> None:
        subprocess.run(
            ["git", "-C", str(workspace), *arguments],
            check=True,
            capture_output=True,
            text=True,
        )

    git("init", "-b", "main")
    git("config", "user.email", "tests@example.invalid")
    git("config", "user.name", "Tests")
    (workspace / ".gitattributes").write_text("*.xlsm binary\n", encoding="utf-8")
    git("add", "-A")
    git("commit", "-m", "the workbook as it was")
    return workbook


@pytest.fixture
def shapes_workbook(workspace: Path) -> Path:
    """The shipped shapes fixture, copied into the workspace so its path resolves.

    Committed as a binary because a Forms-toolbar button cannot be built without
    Excel; tests/test_shapes_live.py rebuilds it and checks it still reads the same.
    """
    import shutil

    source = Path(__file__).parent / "fixtures" / "shapes.xlsm"
    if not source.is_file():
        pytest.skip(f"{source} is missing")
    target = workspace / "Shapes.xlsm"
    shutil.copy2(source, target)
    return target


ACCESS_MODULE = """Option Compare Database
Option Explicit

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
"""


@pytest.fixture
def access_database(workspace: Path) -> Path:
    """The host that is different: VBA in the database, not in a package."""
    import pyopenvba

    path = workspace / "App.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        db.create_table("Orders", [pyopenvba.ColumnSpec("Id", "long")])
        db.vba_project().add_module("Helpers", ACCESS_MODULE.replace("\n", "\r\n"))
        db.save()
    return path


ACCESS_FORM_CODE = """Option Compare Database

Private Sub Form_Load()
End Sub
"""


@pytest.fixture
def access_designs(workspace: Path) -> Path:
    """A database holding both kinds of design, which live in separate collections."""
    import pyopenvba

    path = workspace / "Designs.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        form = db.add_form("Summary", caption="Totals", width=8000, height=3000)
        form.add_control(
            "Label", "Title", left=240, top=240, width=2000, height=300, caption="Hello"
        )
        # A design with code behind it, because that is where the interesting
        # case is: the module and the design have to move together.
        form.set_code(ACCESS_FORM_CODE.replace("\n", "\r\n"))
        report = db.add_report("Monthly")
        report.add_control("Label", "Banner", section="PageHeaderSection", caption="Header")
        db.save()
    return path


@pytest.fixture
def vb6_project(workspace: Path) -> Path:
    """A .vbp with a standard module and a form that calls into it.

    The files are written in the ANSI code page, which is what VB6 writes.
    """
    from tests_vb6_sources import VB6_FORM, VB6_HELPERS, VB6_MANIFEST

    (workspace / "Helpers.bas").write_text(VB6_HELPERS, encoding="cp1252", newline="")
    (workspace / "Form1.frm").write_text(VB6_FORM, encoding="cp1252", newline="")
    project = workspace / "Demo.vbp"
    project.write_text(VB6_MANIFEST, encoding="cp1252", newline="")
    return project


@pytest.fixture
def workbook_with_a_form(workspace: Path) -> Path:
    """A workbook holding a UserForm: a designer storage and a module of one name."""
    import pyopenvba

    path = workspace / "Forms.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        form = book.add_form("Wizard", caption="Setup", width=300, height=200)
        form.add_control("CommandButton", "Ok", left=12, top=12)
        book.save()
    return path


@pytest.fixture
def loaded_query(workspace: Path) -> Path:
    """A query loaded onto a sheet: four parts that all name each other."""
    import pyopenvba

    path = workspace / "Loaded.xlsx"
    with pyopenvba.PowerQueryWorkbook.create_new(path) as book:
        book.add_query("Numbers", "let Source = {1..10} in Source")
        book.load_to_sheet("Numbers", ["Value"], cell="A1")
        book.save()
    return path


@pytest.fixture
def crowded_workbook(workspace: Path) -> Path:
    """More modules than any listing returns, so the bound has something to bind."""
    import pyopenvba

    from xlide_mcp.tools._common import MAX_ITEMS

    body = "Option Explicit\r\n\r\nPublic Sub Only()\r\nEnd Sub\r\n"
    path = workspace / "Crowded.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        for i in range(MAX_ITEMS["modules"] + 20):
            book.vba_project().add_module(f"Mod{i:03d}", body)
        book.save()
    return path
