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
