"""The bridge to a running xlide_vbide session inside the Visual Basic Editor.

Every other tool in this server works on the file on disk. This group works on
what is open in front of the user right now: which module the editor is showing,
whether the project is in break mode, what the analyzer is reporting live. It is
the only way to answer "what am I looking at" rather than "what is in the file".

How it connects. A live session writes a discovery file to
%LOCALAPPDATA%\\xlide_vbide\\xlide-api-{pid}.json carrying a loopback port and a
token; every request goes to http://127.0.0.1:{port}/{token}/{route}. The token is
the whole of the security model: it is loopback only, and anything running as the
user can read the file. Nothing here opens that door - the user turns the API on
from the add-in's own agent card, and until they do, these tools report that it is
shut and say where the switch is.

A discovery file outlives a killed host process, so answering is the only proof of
life: every instance is probed rather than trusted.
"""

from __future__ import annotations

import http.client
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from ..config import Settings
from ..errors import ToolError
from ._common import read_only, truncate

DISCOVERY_DIRECTORY = "xlide_vbide"
DISCOVERY_GLOB = "xlide-api-*.json"

# The add-in answers on the host thread with a three second deadline of its own.
# Anything past this is a wedged VBE, not slow work.
REQUEST_TIMEOUT = 10.0

# urllib's default opener sends 127.0.0.1 through HTTP_PROXY or the system proxy
# like any other host, which would hand the proxy the session's token.
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

# Routes a caller may reach through xlide_live_request. Everything that reads the
# session is here; the page-driving routes are not, because "drive the editor the
# user is typing in" is not something an agent should reach for on its own.
_READ_ROUTES = frozenset(
    {
        "agent", "agent/routes", "agent/examples", "analyzer", "doctor", "engine",
        "model", "native", "project", "projects", "state", "stats", "windows",
    }
)


@dataclass(frozen=True)
class Instance:
    pid: int
    port: int
    token: str
    host: str
    product: str
    file: Path

    @property
    def base(self) -> str:
        return f"http://127.0.0.1:{self.port}/{self.token}"

    def describe(self) -> dict[str, Any]:
        return {"pid": self.pid, "host": self.host, "product": self.product, "port": self.port}


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_live_sessions",
        title="Live VBE sessions",
        annotations=read_only("Live VBE sessions"),
        description=(
            "Lists the running xlide_vbide sessions this machine can reach: one per Office "
            "process that has opened the Visual Basic Editor with the add-in loaded. Use it to "
            "find out whether the user has a live editor before asking about what is on their "
            "screen. Windows only, and the add-in's local API has to be switched on by the "
            "user. Each session is probed before it is listed, because a discovery file outlives "
            "the process that wrote it."
        ),
    )
    def live_sessions() -> dict[str, Any]:
        return live_report()

    @server.tool(
        name="xlide_live_state",
        title="Live editor state",
        annotations=read_only("Live editor state"),
        description=(
            "What the running Visual Basic Editor is doing right now: the module and project "
            "on screen, whether the project is in break mode, whether there are unsaved edits, "
            "the procedure the caret is in, and whether the analyzer engine is answering. Use "
            "it when the user asks about what they are looking at, or before suggesting an "
            "edit to a module they may have unsaved changes in."
        ),
    )
    def live_state(
        pid: Annotated[
            int,
            Field(
                default=0,
                description="Which session, by process id. 0 uses the only one, or refuses.",
            ),
        ] = 0,
    ) -> dict[str, Any]:
        instance = _pick(pid)
        state = _request(instance, "state")
        native = _request(instance, "native")
        return {
            "session": instance.describe(),
            "state": state,
            "native_editor": native,
            "note": (
                "Unsaved edits in the editor are not in the file on disk. A read through the "
                "file tools returns what was last saved, not what is on screen."
            ),
        }

    @server.tool(
        name="xlide_live_request",
        title="Query a live session",
        annotations=read_only("Query a live session"),
        description=(
            "Calls one read route on a running xlide_vbide session and returns its JSON. "
            "Start with route='agent', which answers with the session's own route table and "
            "what each one is for. Allowed routes: "
            + ", ".join(sorted(_READ_ROUTES))
            + ". Routes that drive the editor are deliberately not reachable here."
        ),
    )
    def live_request(
        route: Annotated[str, Field(description="One of the read routes listed above.")],
        pid: Annotated[
            int, Field(default=0, description="Which session, by process id.")
        ] = 0,
        query: Annotated[
            str,
            Field(
                default="",
                description="Query string for the route, such as 'name=Module1' or 'type=Range'.",
            ),
        ] = "",
    ) -> dict[str, Any]:
        wanted = (route or "").strip().strip("/")
        if wanted not in _READ_ROUTES:
            raise ToolError(
                f"{route!r} is not a route this server will call. Allowed: "
                f"{', '.join(sorted(_READ_ROUTES))}. Call route='agent' for what each one does."
            )
        instance = _pick(pid)
        payload = _request(instance, wanted, query)
        return {"session": instance.describe(), "route": wanted, "response": payload}

    @server.tool(
        name="xlide_live_read_module",
        title="Read a module from the editor",
        annotations=read_only("Read a module from the editor"),
        description=(
            "Reads a module's text as the running editor holds it, including edits the user "
            "has not saved. This is the one read that can differ from xlide_read_module, and "
            "the difference is exactly the user's unsaved work. Use it to see what they are "
            "actually working on; use xlide_read_module for what is in the file. "
            "surface_only=true reads the modern editor's own copy instead, which exists only "
            "for a module the user has open in a tab."
        ),
    )
    def live_read_module(
        module_name: Annotated[str, Field(description="Module name as the editor shows it.")],
        pid: Annotated[int, Field(default=0, description="Which session, by process id.")] = 0,
        project: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Which project, when two documents are open and both hold a module of "
                    "this name."
                ),
            ),
        ] = "",
        surface_only: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Read the modern editor's own copy, which only exists for an open tab. "
                    "Off reads the module through the editor, which always answers."
                ),
            ),
        ] = False,
    ) -> dict[str, Any]:
        instance = _pick(pid)
        # `live=1` asks the Monaco surface for its copy, and the surface holds text
        # only for a module the user has opened a tab on: measured against a live
        # session, it refuses outright for every other module. The default asks the
        # editor, which answers for any module in the project.
        parts = [("name", module_name), ("project", project)]
        if surface_only:
            parts.append(("live", "1"))
        query = urllib.parse.urlencode({k: v for k, v in parts if v})
        try:
            payload = _request(instance, "module", query)
        except ToolError as exc:
            if surface_only and "no text" in str(exc):
                raise ToolError(
                    f"The editor surface holds no text for {module_name}, which means the "
                    "user does not have it open in a tab. Call again without surface_only."
                ) from exc
            raise
        source = ""
        if isinstance(payload, dict):
            for key in ("text", "source", "content"):
                if isinstance(payload.get(key), str):
                    source = payload[key]
                    break
        text, was_cut = truncate(source)
        return {
            "session": instance.describe(),
            "module": module_name,
            "truncated": was_cut,
            "source": text,
            "raw": payload if not source else None,
            "note": "This is the editor's live text, which may differ from the saved file.",
        }


def live_report() -> dict[str, Any]:
    """Every live session, probed. Shared with xlide_doctor."""
    if sys.platform != "win32":
        return {
            "available": False,
            "reason": "xlide_vbide runs inside the Windows Visual Basic Editor.",
        }
    files = list(_discovery_files())
    if not files:
        return {
            "available": False,
            "count": 0,
            "reason": (
                "No xlide_vbide session is advertising a local API. Either the add-in is not "
                "installed, no Office application has opened the Visual Basic Editor, or the "
                "API is switched off. The user turns it on from the robot button on the xlide "
                "toolbar inside the VBE; it ships shut."
            ),
        }
    live: list[dict[str, Any]] = []
    dead = 0
    for instance in files:
        try:
            _request(instance, "agent")
        except ToolError:
            dead += 1
            continue
        live.append(instance.describe())
    result: dict[str, Any] = {
        "available": bool(live),
        "count": len(live),
        "sessions": live,
    }
    if dead:
        result["stale_discovery_files"] = dead
    if not live:
        result["reason"] = (
            f"{dead} discovery files were found and none answered. A file outlives the process "
            "that wrote it, so these are sessions that have closed."
        )
    return result


def _discovery_root() -> Path:
    local = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(local) / DISCOVERY_DIRECTORY


def _discovery_files() -> list[Instance]:
    root = _discovery_root()
    if not root.is_dir():
        return []
    found: list[Instance] = []
    for path in sorted(root.glob(DISCOVERY_GLOB)):
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        port = raw.get("port")
        token = raw.get("token")
        if not isinstance(port, int) or not isinstance(token, str) or not token:
            continue
        found.append(
            Instance(
                pid=int(raw.get("pid") or 0),
                port=port,
                token=token,
                host=str(raw.get("host") or "unknown"),
                product=str(raw.get("product") or "xlide_vbide"),
                file=path,
            )
        )
    return found


def _pick(pid: int) -> Instance:
    """The session the caller means, or a refusal that says how to choose."""
    candidates = _discovery_files()
    if not candidates:
        raise ToolError(
            "No xlide_vbide session is advertising a local API. Call xlide_live_sessions for "
            "what that means and how the user turns it on."
        )
    if pid:
        for instance in candidates:
            if instance.pid == pid:
                return instance
        listed = ", ".join(str(c.pid) for c in candidates)
        raise ToolError(f"No session with pid {pid}. Discovered pids: {listed}.")
    if len(candidates) == 1:
        return candidates[0]

    # A discovery file outlives the process that wrote it, so several files do not
    # mean several sessions. Probing first is what stops a corpse from forcing the
    # caller to disambiguate against the one session that is actually running.
    alive = [instance for instance in candidates if _answers(instance)]
    if len(alive) == 1:
        return alive[0]
    if not alive:
        raise ToolError(
            f"{len(candidates)} discovery files were found and none answered. Those sessions "
            "have closed. Ask the user to open the Visual Basic Editor again."
        )
    listed = ", ".join(f"{c.pid} ({c.host})" for c in alive)
    raise ToolError(
        f"{len(alive)} sessions are live and none was named. Pass pid. Sessions: {listed}."
    )


def _answers(instance: Instance) -> bool:
    try:
        _request(instance, "agent")
    except ToolError:
        return False
    return True


def _request(instance: Instance, route: str, query: str = "") -> Any:
    """One GET against a live session. Loopback only, by construction."""
    url = f"{instance.base}/{route}"
    if query:
        url += "?" + query.lstrip("?")
    request = urllib.request.Request(url, method="GET")
    try:
        with _OPENER.open(request, timeout=REQUEST_TIMEOUT) as response:
            body = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise ToolError(
                f"The session at pid {instance.pid} refused the request. A wrong token reads as "
                "404, and so does a route this build does not carry. Call "
                "xlide_live_request(route='agent/routes') for the routes it has."
            ) from exc
        raise ToolError(f"The live session answered {exc.code} for {route}.") from exc
    except (urllib.error.URLError, http.client.HTTPException, OSError) as exc:
        # HTTPException is a stale discovery file's port now held by something
        # that does not speak HTTP, which is a closed session like any other.
        raise ToolError(
            f"No answer from the session at pid {instance.pid}: {exc}. The host process may "
            "have closed, or the Visual Basic Editor may be busy running the user's code."
        ) from exc
    try:
        payload = json.loads(body)
    except ValueError:
        return {"raw": body[:4000]}
    if isinstance(payload, dict) and payload.get("error"):
        raise ToolError(f"The live session reported: {payload['error']}")
    return payload


__all__ = ["live_report", "register"]
