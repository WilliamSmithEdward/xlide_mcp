"""The live tools' side of the loopback: what they reach, and how they fail.

A live session is xlide_vbide's local API, a port on this machine behind a token
and advertised in a discovery file. These tests stand one in and reach it the way
the tools do, which needs no Office: nothing here calls into the editor.
"""

from __future__ import annotations

import json
import socket
import threading
from collections.abc import Callable, Iterator
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

TOKEN = "t0ken"


def _advertise(root: Path, port: int, pid: int) -> None:
    folder = root / "xlide_vbide"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"xlide-api-{pid}.json").write_text(
        json.dumps({"pid": pid, "port": port, "token": TOKEN, "host": "EXCEL"}),
        encoding="utf-8",
    )


@pytest.fixture
def session(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Iterator[int]:
    """A stand-in session that answers the agent route, found as a real one is."""

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:
            ok = self.path == f"/{TOKEN}/agent"
            data = json.dumps({"routes": ["agent"]} if ok else {}).encode()
            self.send_response(200 if ok else 404)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def log_message(self, *_args: Any) -> None:
            return None

    server = HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    _advertise(tmp_path, server.server_address[1], pid=4242)
    yield 4242
    server.shutdown()
    server.server_close()


def test_a_proxy_in_the_environment_is_never_used(
    call: Callable[..., Any], session: int, proxy_in_environment: str
) -> None:
    """urllib's default opener sends 127.0.0.1 through HTTP_PROXY like anywhere else,
    which hands the proxy the session's token. The proxy refuses, so a request that
    went through it would not arrive."""
    answered = call("xlide_live_request", route="agent", pid=session)
    assert answered["response"] == {"routes": ["agent"]}


def test_live_module_can_be_read_in_slices_with_one_token(
    call: Callable[..., Any], monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlide_mcp.tools import live

    class Session:
        def describe(self) -> dict[str, Any]:
            return {"pid": 4242, "host": "EXCEL"}

    source = "Option Explicit\r\nPublic Sub Go()\r\n    Debug.Print 1\r\nEnd Sub\r\n"
    monkeypatch.setattr(live, "_pick", lambda _pid: Session())
    monkeypatch.setattr(live, "_request", lambda *_args: {"text": source})
    full = call("xlide_live_read_module", module_name="Tools")
    sliced = call(
        "xlide_live_read_module", module_name="Tools", start_line=2, end_line=3,
    )
    assert full["source"].endswith("\n")
    assert full["total_lines"] == 4
    assert sliced["source"] == "Public Sub Go()\n    Debug.Print 1"
    assert sliced["content_token"] == full["content_token"]
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_live_read_module", module_name="Tools", start_line=3, end_line=2)
    assert "before start_line" in refusal.value.message


def test_a_stranger_on_a_stale_port_is_a_session_that_closed(
    call: Callable[..., Any], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """The editor closed and left its discovery file, and something that does not
    speak HTTP took the port. Its answer raised past the handler, so the tool
    failed with a bare "Error executing tool" instead of saying the session is gone."""
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
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    _advertise(tmp_path, listener.getsockname()[1], pid=99)
    try:
        with pytest.raises(ToolFailure) as refusal:
            call("xlide_live_request", route="agent", pid=99)
        assert "No answer from the session at pid 99" in refusal.value.message
    finally:
        listener.close()
