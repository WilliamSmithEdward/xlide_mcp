"""Telling a running XLIDE for VS Code what a tool changed.

The extension's side is a loopback endpoint behind a token, advertised in a
discovery file. These tests stand one up, record what this server posts to it,
and check that a tool with no XLIDE running behaves exactly as before.
"""

from __future__ import annotations

import json
import socket
import sys
import threading
from collections.abc import Callable, Iterator
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE

from xlide_mcp import xlide_vscode

TOKEN = "t0ken"


class _StandIn:
    """What the extension would run: hello, and a record of every notice."""

    def __init__(self) -> None:
        self.notices: list[tuple[str, dict[str, Any]]] = []
        stand_in = self

        class Handler(BaseHTTPRequestHandler):
            def _route(self) -> str | None:
                prefix = f"/{TOKEN}/"
                return self.path[len(prefix):] if self.path.startswith(prefix) else None

            def _answer(self, status: int, body: dict[str, Any]) -> None:
                data = json.dumps(body).encode()
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def do_GET(self) -> None:
                if self._route() == "hello":
                    self._answer(200, {"product": "xlide_vscode", "protocol": 1})
                else:
                    self._answer(404, {})

            def do_POST(self) -> None:
                route = self._route()
                length = int(self.headers.get("Content-Length", "0"))
                body = json.loads(self.rfile.read(length) or b"{}")
                if route is None:
                    self._answer(403, {})
                    return
                stand_in.notices.append((route, body))
                reply = {"shown": True}
                if route == "agent-edit":
                    reply["review"] = "pending"
                self._answer(200, reply)

            def log_message(self, *_args: Any) -> None:
                return None

        self.server = HTTPServer(("127.0.0.1", 0), Handler)
        self.port = self.server.server_address[1]
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def close(self) -> None:
        self.server.shutdown()
        self.server.server_close()


@pytest.fixture
def xlide(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Iterator[_StandIn]:
    """A stand-in XLIDE window, found through a discovery file as the real one is."""
    stand_in = _StandIn()
    folder = tmp_path / "profile" / xlide_vscode.DISCOVERY_DIRECTORY
    folder.mkdir(parents=True)
    (folder / "xlide-api-4242.json").write_text(
        json.dumps({"pid": 4242, "port": stand_in.port, "token": TOKEN, "product": "xlide_vscode"}),
        encoding="utf-8",
    )
    monkeypatch.setattr(xlide_vscode, "_discovery_directory", lambda: folder)
    xlide_vscode.forget()
    yield stand_in
    stand_in.close()
    xlide_vscode.forget()


def test_the_discovery_folder_is_the_one_the_extension_writes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Undo conftest's redirection and check the real location, per platform."""
    monkeypatch.undo()
    if sys.platform == "win32":
        monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    else:
        monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path))
    assert xlide_vscode._discovery_directory() == tmp_path / "xlide_vscode"


def test_a_module_write_reaches_xlides_review_with_its_before_and_after(
    call: Callable[..., Any], workbook: Path, xlide: _StandIn
) -> None:
    changed = SAMPLE_MODULE.replace('"world"', '"there"')
    written = call(
        "xlide_write_module", file_path=str(workbook), module_name="Helpers", source=changed
    )
    assert written["xlide_vscode"]["notified"] is True
    assert written["xlide_vscode"]["review"] == "pending"

    route, notice = xlide.notices[-1]
    assert route == "agent-edit"
    assert notice["file"] == str(workbook)
    assert notice["module"] == "Helpers"
    assert notice["beforeExisted"] is True
    assert '"world"' in notice["before"]
    assert '"there"' in notice["after"]
    assert "Attribute VB_Name" not in notice["after"]
    assert notice["tool"] == "xlide_write_module"
    assert notice["server"] == "xlide_mcp"


def test_a_created_module_says_it_did_not_exist_before(
    call: Callable[..., Any], workbook: Path, xlide: _StandIn
) -> None:
    call("xlide_write_module", file_path=str(workbook), module_name="Fresh", source=SAMPLE_MODULE)
    _route, notice = xlide.notices[-1]
    assert notice["beforeExisted"] is False
    assert notice["before"] == ""


def test_a_deleted_module_carries_its_text_so_it_can_be_put_back(
    call: Callable[..., Any], workbook: Path, xlide: _StandIn
) -> None:
    call("xlide_delete_module", file_path=str(workbook), module_name="Helpers")
    route, notice = xlide.notices[-1]
    assert route == "agent-edit"
    assert notice["afterExists"] is False
    assert "AddNums" in notice["before"]
    assert notice["kind"] == "standard"


def test_a_rename_moves_the_review(
    call: Callable[..., Any], workbook: Path, xlide: _StandIn
) -> None:
    call("xlide_rename_module", file_path=str(workbook), module_name="Helpers", new_name="Tools")
    route, notice = xlide.notices[-1]
    assert (route, notice["from"], notice["to"]) == ("module-renamed", "Helpers", "Tools")


def test_a_cell_write_refreshes_the_tree(
    call: Callable[..., Any], plain_workbook: Path, xlide: _StandIn
) -> None:
    call(
        "xlide_write_cells",
        file_path=str(plain_workbook),
        sheet="Sheet1",
        start_cell="A1",
        data=[[1]],
    )
    route, notice = xlide.notices[-1]
    assert (route, notice["what"]) == ("file-changed", "document")


def test_with_no_xlide_running_nothing_changes(call: Callable[..., Any], workbook: Path) -> None:
    written = call(
        "xlide_write_module", file_path=str(workbook), module_name="Helpers", source=SAMPLE_MODULE
    )
    assert "xlide_vscode" not in written


def test_a_window_that_does_not_answer_is_skipped(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A discovery file outlives the window that wrote it."""
    folder = tmp_path / xlide_vscode.DISCOVERY_DIRECTORY
    folder.mkdir()
    (folder / "xlide-api-1.json").write_text(
        json.dumps({"pid": 1, "port": 9, "token": "gone"}), encoding="utf-8"
    )
    monkeypatch.setattr(xlide_vscode, "_discovery_directory", lambda: folder)
    xlide_vscode.forget()
    try:
        assert xlide_vscode.windows() == []
    finally:
        xlide_vscode.forget()


def test_whatever_holds_a_stale_port_cannot_fail_the_write(
    call: Callable[..., Any], plain_workbook: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """VS Code crashed and left its discovery file, and something that does not speak
    HTTP took the port. The notice goes after the save, so a failure there turned a
    written cell into a bare "Error executing tool", and did so on every write."""
    listener = socket.socket()
    listener.bind(("127.0.0.1", 0))
    listener.listen()

    def answer_badly() -> None:
        while True:
            try:
                connection, _ = listener.accept()
            except OSError:
                return
            with connection:
                connection.recv(4096)
                connection.sendall(b"Content-Length: 5\r\n\r\nhello")

    threading.Thread(target=answer_badly, daemon=True).start()
    folder = tmp_path / "stale" / xlide_vscode.DISCOVERY_DIRECTORY
    folder.mkdir(parents=True)
    (folder / "xlide-api-99.json").write_text(
        json.dumps({"pid": 99, "port": listener.getsockname()[1], "token": TOKEN}),
        encoding="utf-8",
    )
    monkeypatch.setattr(xlide_vscode, "_discovery_directory", lambda: folder)
    xlide_vscode.forget()
    try:
        written = call(
            "xlide_write_cells",
            file_path=str(plain_workbook),
            sheet="Sheet1",
            start_cell="A1",
            data=[[7]],
        )
        assert written["cells_written"] == 1
        assert "xlide_vscode" not in written
        read = call(
            "xlide_read_cells", file_path=str(plain_workbook), sheet="Sheet1", cell_range="A1"
        )
        assert read["values"] == [[7]]
    finally:
        listener.close()
        xlide_vscode.forget()


def test_a_proxy_in_the_environment_is_never_used(
    call: Callable[..., Any], workbook: Path, xlide: _StandIn, proxy_in_environment: str
) -> None:
    """urllib's default opener sends 127.0.0.1 through HTTP_PROXY like anywhere else,
    which hands the proxy the token and the module text. The proxy refuses, so a
    notice that went through it would not arrive."""
    written = call(
        "xlide_write_module", file_path=str(workbook), module_name="Helpers",
        source=SAMPLE_MODULE.replace("a + b", "a * b"),
    )
    assert written["xlide_vscode"]["notified"] is True
    assert xlide.notices[-1][0] == "agent-edit"


def test_only_a_port_on_this_machine_is_ever_told(monkeypatch: pytest.MonkeyPatch) -> None:
    """Module source goes nowhere but loopback, whatever the environment says."""
    monkeypatch.setenv(xlide_vscode.ENVIRONMENT, "http://example.com:8080/token")
    xlide_vscode.forget()
    try:
        assert all(
            not window.base.startswith("http://example.com") for window in xlide_vscode.windows()
        )
    finally:
        xlide_vscode.forget()
