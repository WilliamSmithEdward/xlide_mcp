"""The workspace boundary: what a path argument is allowed to reach.

A path argument is untrusted input. The model that supplies it can be steered by
anything it has read, so these are the tests that matter most for a server that
writes to files.
"""

from __future__ import annotations

import os
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

from xlide_mcp import Settings, build_server
from xlide_mcp.errors import ToolError
from xlide_mcp.paths import resolve_path


def test_a_path_inside_a_root_resolves(tmp_path: Path) -> None:
    settings = Settings(roots=(tmp_path.resolve(),))
    target = tmp_path / "book.xlsm"
    target.write_bytes(b"")
    assert resolve_path(str(target), settings) == target.resolve()


def test_a_path_outside_every_root_is_refused(tmp_path: Path) -> None:
    settings = Settings(roots=((tmp_path / "inside").resolve(),))
    (tmp_path / "inside").mkdir()
    outside = tmp_path / "outside.xlsm"
    outside.write_bytes(b"")

    with pytest.raises(ToolError) as refusal:
        resolve_path(str(outside), settings)
    assert "outside this server's workspace" in str(refusal.value)
    assert "--root" in str(refusal.value), "the message has to say how to fix it"


def test_dot_dot_traversal_is_refused(tmp_path: Path) -> None:
    inside = tmp_path / "inside"
    inside.mkdir()
    settings = Settings(roots=(inside.resolve(),))
    (tmp_path / "secret.xlsm").write_bytes(b"")

    with pytest.raises(ToolError):
        resolve_path(str(inside / ".." / "secret.xlsm"), settings)


@pytest.mark.skipif(os.name == "nt", reason="creating a symlink on Windows needs privilege")
def test_a_symlink_out_of_a_root_is_refused(tmp_path: Path) -> None:
    """Resolution happens before the containment check, so a link inside a root
    that points outside one does not smuggle the target in."""
    inside = tmp_path / "inside"
    inside.mkdir()
    outside = tmp_path / "outside.xlsm"
    outside.write_bytes(b"")
    (inside / "link.xlsm").symlink_to(outside)

    settings = Settings(roots=(inside.resolve(),))
    with pytest.raises(ToolError):
        resolve_path(str(inside / "link.xlsm"), settings)


def test_a_null_byte_is_refused(tmp_path: Path) -> None:
    settings = Settings(roots=(tmp_path.resolve(),))
    with pytest.raises(ToolError) as refusal:
        resolve_path("book\x00.xlsm", settings)
    assert "null byte" in str(refusal.value)


def test_a_relative_path_resolves_against_the_first_root(tmp_path: Path) -> None:
    """Not the process working directory: a server started with --root and run
    from somewhere else would otherwise answer differently for the same call."""
    settings = Settings(roots=(tmp_path.resolve(),))
    (tmp_path / "book.xlsm").write_bytes(b"")
    assert resolve_path("book.xlsm", settings) == (tmp_path / "book.xlsm").resolve()


def test_allow_outside_roots_opens_the_door_deliberately(tmp_path: Path) -> None:
    outside = tmp_path / "outside.xlsm"
    outside.write_bytes(b"")
    settings = Settings(roots=((tmp_path / "inside").resolve(),), allow_outside_roots=True)
    (tmp_path / "inside").mkdir()
    assert resolve_path(str(outside), settings) == outside.resolve()


def test_read_only_refuses_writes_but_not_reads(tmp_path: Path, workbook: Path) -> None:
    import asyncio

    settings = Settings(roots=(workbook.parent.resolve(),), read_only=True)
    server = build_server(settings)

    read = asyncio.run(
        server.call_tool("xlide_list_modules", {"file_path": str(workbook)})
    )
    assert read.structured_content["count"] > 0

    with pytest.raises(ToolError) as refusal:
        asyncio.run(
            server.call_tool(
                "xlide_write_module",
                {
                    "file_path": str(workbook),
                    "module_name": "Helpers",
                    "source": "Option Explicit\r\n",
                },
            )
        )
    assert "read-only" in str(refusal.value)


def test_a_missing_file_says_how_to_find_one(call: Callable[..., Any], workspace: Path) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_list_modules", file_path=str(workspace / "nope.xlsm"))
    assert "xlide_list_projects" in refusal.value.message


def test_settings_read_from_the_environment(tmp_path: Path) -> None:
    env = {
        "XLIDE_MCP_ROOTS": os.pathsep.join([str(tmp_path), str(tmp_path / "other")]),
        "XLIDE_MCP_READ_ONLY": "yes",
        "XLIDE_MCP_TIMEOUT": "45",
    }
    settings = Settings.from_environment(env)
    assert settings.read_only is True
    assert settings.default_timeout == 45
    assert len(settings.roots) == 2


def test_a_nonsense_timeout_falls_back_rather_than_crashing() -> None:
    settings = Settings.from_environment({"XLIDE_MCP_TIMEOUT": "soon"})
    assert settings.default_timeout > 0


def test_timeout_is_held_at_the_ceiling() -> None:
    from xlide_mcp.config import MAX_TIMEOUT_SECONDS, clamp_timeout

    settings = Settings()
    assert clamp_timeout(10_000, settings) == MAX_TIMEOUT_SECONDS
    assert clamp_timeout(0.1, settings) == 1.0
    assert clamp_timeout(None, settings) == settings.default_timeout


def test_roots_can_be_positional_as_well_as_flagged(tmp_path: Path) -> None:
    """A launcher that mounts the caller's folders somewhere of its own choosing
    appends them as plain arguments; it cannot repeat a flag in front of each.

    Before this, `--root /a /b` exited 2 on the unrecognised second path, which
    meant a server that would not start at all rather than one with the wrong
    roots.
    """
    from xlide_mcp.__main__ import build_parser
    from xlide_mcp.config import roots_from_argv

    first = tmp_path / "one"
    second = tmp_path / "two"
    first.mkdir()
    second.mkdir()

    args = build_parser().parse_args(["--root", str(first), str(second)])
    roots = roots_from_argv(list(args.root or []) + list(args.roots or []))

    assert set(roots) == {first.resolve(), second.resolve()}
